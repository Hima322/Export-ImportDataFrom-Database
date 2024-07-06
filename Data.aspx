<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Data.aspx.cs" Inherits="DatatransferProto.Data" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="//code.jquery.com/ui/1.10.4/jquery-ui.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.4/xlsx.full.min.js"></script>

<head runat="server">
    <title>Data Import & Export</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            font-family: 'Lucida Sans', 'Lucida Sans Regular', 'Lucida Grande', 'Lucida Sans Unicode', Geneva, Verdana, sans-serif
        }

        .NavBar {
            width: 100%;
            background-color: cadetblue;
            color: white;
            font-size: large;
            text-align: center;
            padding: 10px 0;
        }

        .container {
            display: flex;
            flex-direction: column;
            align-items: center;
            margin-top: 20px;
        }

        .for-import, .for-export {
            margin: 20px;
            text-align: center;
        }

        .button-container {
            display: flex;
            justify-content: center;
            gap: 35rem;
            margin-top: 20px;
        }

        button {
            background-color: blue;
            color: white;
            border: none;
            padding: 1rem 2rem;
            font-size: large;
            cursor: pointer;
            transition: background-color 0.3s ease;
            border: 2px solid white;
            border-radius: 25px;
        }

        .Show-Data {
            width: 95%;
            margin: 20px auto;
            align-items: center;
            border-collapse: collapse;
            display: none;
        }

            .Show-Data table {
                width: 100%;
                border-collapse: collapse;
            }

            .Show-Data th, .Show-Data td {
                border: 1px solid #ddd;
                padding: 8px;
                text-align: center;
            }

            .Show-Data th {
                background-color: cadetblue;
                color: white;
                padding-top: 12px;
                padding-bottom: 12px;
            }

            .Show-Data tr:nth-child(even) {
                background-color: #f2f2f2;
            }

            .Show-Data tr:hover {
                background-color: #ddd;
            }


        #download-button {
            background-color: blue;
            color: white;
            border: none;
            padding: 1rem 3rem;
            cursor: pointer;
            transition: background-color 0.3s ease;
            border: 2px solid white;
            border-radius: 25px;
            margin: auto;
            display: block;
        }

        #Import-button1 {
            background-color: blue;
            color: white;
            border: none;
            padding: 1rem 3rem;
            cursor: pointer;
            transition: background-color 0.3s ease;
            border: 2px solid white;
            border-radius: 25px;
            margin: auto;
            display: block;
        }

        #download-button:hover {
            background-color: darkblue;
        }

        @media (min-width: 768px) {
            .container {
                flex-direction: row;
                justify-content: space-around;
            }

            .for-import {
                position: relative;
                left: 0;
            }

            .for-export {
                position: relative;
                right: 0;
            }

            #ImportSave {
                position: static;
            }

            #ExportSave {
                position: static;
            }
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div class="NavBar">
            <h1>Data Import & Export (Excel -> Database)</h1>
        </div>
        <div class="container">
            <div class="for-import">
                <input type="radio" id="Import-button" name="fav_language" />
                <label for="Import-button" style="font-size: large;">Import Data into Database</label>
            </div>

            <div class="for-export">
                <input type="radio" id="export-button" name="fav_language" />
                <label for="export-button" style="font-size: large;">Export Data from Database</label>
            </div>
        </div>
        <div class="button-container">
            <button id="ImportSave" type="button" onclick="importdata()">ImportSave</button>
            <button id="ExportSave" type="button" onclick="ExportData()">ExportSave</button>
        </div>
        <div class="Show-Data" id="ShowData">
            <table id="table">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Model</th>
                        <th>Variant</th>
                        <th>QRCode</th>
                        <th>Status</th>
                        <th>QRPrintTime</th>
                        <th>LineName</th>
                    </tr>
                </thead>
                <tbody>
                </tbody>
            </table>

        </div>
        <button id="download-button" style="font-size: large; font-weight: 600; color: white" onclick="ExportToExcel()">Download&#128070</button>
        <button id="Import-button1" style="font-size: large; font-weight: 600; color: white" onclick="Import()">Import&#128070</button>


    </form>
  <script>
      $(document).ready(function () {
          $("#download-button").hide();
          $("#Import-button1").hide(); // Corrected selector for hiding the button on load
      });

      const ExportData = () => {
          let isExportSelected = document.getElementById('export-button').checked;

          if (!isExportSelected) {
              alert("Please select whether you want to export!");
              return;
          }

          $.ajax({
              type: "POST",
              url: "Data.aspx/GetWholeData",
              data: '{}',
              contentType: "application/json; charset=utf-8",
              dataType: "json",
              async: true,
              cache: false,
              success: function (res) {
                  if (res.d != "Error") {
                      let data = JSON.parse(res.d);
                      $("#table tbody").html(
                          data.map(e => `
                            <tr>
                                <td>${e.ID}</td>
                                <td>${e.ModelName}</td>
                                <td>${e.VarientName}</td>
                                <td>${e.QR_Data}</td>
                                <td>${e.Status}</td>
                                <td>${e.QRPrintTime}</td>
                                <td>${e.LineName}</td>
                            </tr>
                        `).join('')
                      );
                      $("#ShowData").show();
                      $("#download-button").show();
                  } else {
                      alert("No data to export!");
                  }
              },
              error: function (xhr, status, error) {
                  console.log("Error: " + error);
              }
          });
      };

      function ExportToExcel() {
          alert("Are You Sure Download This Report");
          var wb = XLSX.utils.table_to_book(document.getElementById('table'), { sheet: "Sheet JS" });
          XLSX.writeFile(wb, 'ExportedData.xlsx');
      }

      const importdata = () => {
          let isImportSelected = document.getElementById('Import-button').checked;

          if (!isImportSelected) {
              alert("Please select whether you want to import!");
              return;
          }

          $.ajax({
              type: "POST",
              url: "Data.aspx/GetWholeDataFromExcel",
              data: '{}',
              contentType: "application/json; charset=utf-8",
              dataType: "json",
              async: true,
              cache: false,
              success: function (res) {
                  if (res.d != "Error") {
                      let data = JSON.parse(res.d);
                      $("#table tbody").html(
                          data.map(e => `
                            <tr>
                                <td>${e.ID}</td>
                                <td>${e.ModelName}</td>
                                <td>${e.VarientName}</td>
                                <td>${e.QR_Data}</td>
                                <td>${e.Status}</td>
                                <td>${e.QRPrintTime}</td>
                                <td>${e.LineName}</td>
                            </tr>
                        `).join('')
                      );
                      $("#ShowData").show();
                      $("#Import-button1").show(); // Show import button when data is loaded
                  } else {
                      alert("No data to import!");
                  }
              },
              error: function (xhr, status, error) {
                  console.log("Error: " + error);
              }
          });
      };
  </script>

</body>
</html>
