using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DatatransferProto.Databse;
using System.IO;
using System.Web.Services;
using System.Text.Json;
using iTextSharp.text;
using iTextSharp.text.pdf;
using OfficeOpenXml;
using ClosedXML.Excel;
using System.Data.Entity.Migrations;
namespace DatatransferProto
{
    public partial class Data : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        [WebMethod]
        public static string GetWholeData()
        {
            try
            {
                using (HimanshuEntities db = new HimanshuEntities())
                {
                    var res = db.SaveDatas.ToList();
                    if (res.Count > 0)
                    {
                        return JsonSerializer.Serialize(res);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "Error";
            }
            return "Error";
        }

        [WebMethod]
        private static void Import(List<DatatransferProto.Databse.SaveData> dataList)
        {
            using (HimanshuEntities db = new HimanshuEntities())
            {
                foreach (var data in dataList)
                {
                   
                 
                    db.SaveDatas.Add(data); 
                }
                db.SaveChanges(); 
            }
        }
        [WebMethod]
        public static string SaveDataToDatabase(List<SaveData> data)
        {
            try
            {
                Import(data);
                return "Success";
            }
            catch (Exception ex)
            {
            
                return "Error: " + ex.Message;
            }
        }

        private static void Import(List<SaveData> data)
        {
            throw new NotImplementedException();
        }

        [WebMethod]
        public static string GetWholeDataFromExcel()
        {
            try
            {
                string filePath = @"E:\DataSave.xlsx";

                var workbook = new XLWorkbook(filePath);
                var worksheet = workbook.Worksheet(1); // Assuming the first worksheet

                List<SaveData> dataList = new List<SaveData>();

                // Assuming the first row is header, starting from row 2
                var rows = worksheet.RowsUsed().Skip(1);

                foreach (var row in rows)
                {
                    SaveData data = new SaveData
                    {
                        ID = row.Cell(1).GetValue<int>(),
                        ModelName = row.Cell(2).GetValue<string>(),
                        VarientName = row.Cell(3).GetValue<string>(),
                        QR_Data = row.Cell(4).GetValue<string>(),
                        Status = row.Cell(5).GetValue<string>(),
                        //QRPrintTime = row.Cell(6).GetValue<DateTime>(),
                        LineName = row.Cell(7).GetValue<string>()
                    };

                    dataList.Add(data);
                }

                return JsonSerializer.Serialize(dataList);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "Error";
            }
        }

        public class SaveData
        {
            public int ID { get; set; }
            public string ModelName { get; set; }
            public string VarientName { get; set; }
            public string QR_Data { get; set; }
            public string Status { get; set; }
            public DateTime QRPrintTime { get; set; }
            public string LineName { get; set; }

        }
    }
}

