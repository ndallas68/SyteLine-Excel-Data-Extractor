using Microsoft.Office.Interop.Excel;
using Mongoose.IDO.Protocol;
using System;
using System.Collections;
using System.Data;
using System.Globalization;
using System.IO;
using DataTable = System.Data.DataTable;

namespace WorkCenterDemandExtract
{
    internal class WorkCenterDemandExtract
    {
        public string fileName;
        public Session session;
        public DataTable finalTable;
        public WorkCenterDemandExtract(string url, string username, string password, string configuration, string fileName)
        {
            this.fileName = fileName + " " + DateTime.Now.ToString("yyyy-MM-dd");
            session = new Session();
            session.Login(url, username, password, configuration);
        }
        public void GenerateDataTable()
        {
            finalTable = CreateDataTable("WorkcenterDemand", "Item,ItemDescription,OperationDescription," +
                "Operation,WC,WCDescription,RunDuration,RunMachineHours,RunLaborHours,CustomerOrder," +
                "CustomerOrderLine,Customer,DueDate,Status,QtyOnHand,QtyOrdered,QtyPicked,QtyPacked," +
                "QtyShipped,OrderLinePrice,OrderMinimum,OrderMaximum,OrderMultiple");

            LoadCollectionResponseData res;

            LoadCollectionRequestData parentRequest = new LoadCollectionRequestData
            {
                IDOName = "SLCoitems",
                Filter = "Stat<>'C'",
                RecordCap = 0,
                OrderBy = "CoNum,CoLine",
                PropertyList = new PropertyList("CoLine,CoNum,CoCustNum,DueDate,Item,QtyPacked,QtyPicked,QtyShipped,QtyOrdered,Stat,PriceConv")
            };

            LoadCollectionRequestData childRequest = new LoadCollectionRequestData
            {
                IDOName = "SLJobRoutes",
                Filter = "(JobSuffix = '0' OR JobSuffix > 1) AND Type = 'S' AND ((DerRevision<> '') OR (DerRevision = '' AND Job = ItwhseJob))",
                RecordCap = 0,
                OrderBy = "",
                PropertyList = new PropertyList("DerRunLbrHrs,DerRunMchHrs,ItmItem,Job,Efficiency,Suffix,OperNum,RunDur,ue_STY_OperationDescription,Wc,WcDescription")
            };

            childRequest.SetLinkBy("Item", "ItmItem");
            parentRequest.AddNestedRequest(childRequest);

            childRequest = new LoadCollectionRequestData
            {
                IDOName = "SLItems",
                Filter = "",
                RecordCap = 0,
                OrderBy = "",
                PropertyList = new PropertyList("Item,Description,DerQtyOnHand,OrderMax,OrderMin,OrderMult")
            };

            childRequest.SetLinkBy("Item", "Item");
            parentRequest.AddNestedRequest(childRequest);

            res = session.client.LoadCollection(parentRequest);

            if (res.Items.Count <= 0)
            {
                return;
            }

            for (int i = 0; i < res.Items.Count; i++)
            {
                IEnumerator resCollections = res.Items[i].NestedResponses.GetEnumerator();
                resCollections.MoveNext();
                LoadCollectionResponseData currentOperationsResponse = (LoadCollectionResponseData)resCollections.Current;
                resCollections.MoveNext();
                LoadCollectionResponseData itemsResponse = (LoadCollectionResponseData)resCollections.Current;

                if (currentOperationsResponse.Items.Count <= 0)
                {
                    continue;
                }

                for (int j = 0; j < currentOperationsResponse.Items.Count; j++)
                {
                    DataRow row = finalTable.NewRow();
                    row.SetField("CustomerOrderLine", res[i, "CoLine"].GetValue(""));
                    row.SetField("CustomerOrder", res[i, "CoNum"].GetValue(""));
                    row.SetField("Customer", res[i, "CoCustNum"].GetValue(""));
                    row.SetField("DueDate", DateTime.ParseExact(res[i, "DueDate"].GetValue(""), "yyyyMMdd HH:mm:ss.fff", CultureInfo.InvariantCulture).ToString("yyyy-MM-dd"));
                    row.SetField("QtyPacked", res[i, "QtyPacked"].GetValue(""));
                    row.SetField("QtyPicked", res[i, "QtyPicked"].GetValue(""));
                    row.SetField("QtyShipped", res[i, "QtyShipped"].GetValue(""));
                    row.SetField("QtyOrdered", res[i, "QtyOrdered"].GetValue(""));
                    row.SetField("Status", res[i, "Stat"].GetValue(""));
                    row.SetField("OrderLinePrice", res[i, "PriceConv"].GetValue(""));
                    row.SetField("RunLaborHours", currentOperationsResponse[j, "DerRunLbrHrs"].GetValue(""));
                    row.SetField("RunMachineHours", currentOperationsResponse[j, "DerRunMchHrs"].GetValue(""));
                    row.SetField("Item", currentOperationsResponse[j, "ItmItem"].GetValue(""));
                    row.SetField("Operation", currentOperationsResponse[j, "OperNum"].GetValue(""));
                    row.SetField("OperationDescription", currentOperationsResponse[j, "ue_STY_OperationDescription"].GetValue(""));
                    row.SetField("WC", currentOperationsResponse[j, "Wc"].GetValue(""));
                    row.SetField("WCDescription", currentOperationsResponse[j, "WcDescription"].GetValue(""));
                    row.SetField("RunDuration", currentOperationsResponse[j, "RunDur"].GetValue(""));

                    if (itemsResponse.Items.Count <= 0)
                    {
                        continue;
                    }

                    row.SetField("ItemDescription", itemsResponse[0, "Description"].GetValue(""));
                    row.SetField("QtyOnHand", itemsResponse[0, "DerQtyOnHand"].GetValue(""));
                    row.SetField("OrderMaximum", itemsResponse[0, "OrderMax"].GetValue(""));
                    row.SetField("OrderMinimum", itemsResponse[0, "OrderMin"].GetValue(""));
                    row.SetField("OrderMultiple", itemsResponse[0, "OrderMult"].GetValue(""));

                    finalTable.Rows.Add(row);
                }
            }
        }

        public void ExportToCsv()
        {
            fileName += ".csv";
            string[] textForFile = new string[finalTable.Rows.Count + 1];

            bool first = true;
            foreach (DataColumn columns in finalTable.Columns)
            {
                if (first)
                {
                    textForFile[0] += columns.Caption;
                    first = false;
                }
                else
                {
                    textForFile[0] += "," + columns.Caption;
                }
            }

            int i = 1;
            foreach (DataRow rows in finalTable.Rows)
            {
                first = true;
                foreach (var item in rows.ItemArray)
                {
                    if (first)
                    {
                        textForFile[i] += item.ToString().Replace(",", " ");
                        first = false;
                    }
                    else
                    {
                        textForFile[i] += "," + item.ToString().Replace(",", " ");
                    }
                }
                i++;
            }
            File.WriteAllLines(fileName, textForFile);
        }

        public void ExportToExcel()
        {
            fileName += ".xlsx";
            Application excelApp = new Application
            {
                Visible = false
            };

            Workbook workbook = excelApp.Workbooks.Add();
            Worksheet worksheet = workbook.Sheets[1];
            worksheet.Tab.Color = XlRgbColor.rgbBlue;
            worksheet.Name = finalTable.TableName;

            int indx = 1;
            foreach (DataColumn columns in finalTable.Columns)
            {
                worksheet.Cells[1, indx] = columns.Caption;
                indx++;
            }

            int row = 2;
            int column;
            foreach (DataRow rows in finalTable.Rows)
            {
                column = 1;
                foreach (var item in rows.ItemArray)
                {
                    worksheet.Cells[row, column] = item;
                    column++;
                }
                row++;
            }

            ListObject table = worksheet.ListObjects.Add(
                XlListObjectSourceType.xlSrcRange,
                worksheet.Range["A1:W" + (row - 1).ToString()],
                Type.Missing,
                XlYesNoGuess.xlYes,
                Type.Missing);

            table.TableStyle = "TableStyleMedium1";

            worksheet.Application.ActiveWindow.SplitRow = 1;
            worksheet.Application.ActiveWindow.FreezePanes = true;

            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            workbook.SaveAs(fileName);
            workbook.Close();
            excelApp.Quit();
        }

        private DataTable CreateDataTable(string tableName, string properties)
        {
            DataTable dt = new DataTable(tableName);
            string[] propertyList = properties.Split(',');

            foreach (string property in propertyList)
            {
                dt.Columns.Add(property);
            }
            return dt;
        }
    }
}
