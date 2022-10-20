using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.EnterpriseServices;
using System.IO;
using System.Runtime.InteropServices;
using HtmlAgilityPack;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace NpoiExcelCom
{
    [ComVisible(true)]
    [Guid("6B8C4DF8-7729-4398-9716-7EFB7007FA8A")]
    [ProgId("NpoiExcelCom.ExcelHelper")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [SecurityRole("User")]
    public class ExcelHelper : ServicedComponent
    {
        static log4net.ILog log;

        public byte[] ExcelBinary { get; set; }
        public string FilePath { get; set; }

        public static void SetupLog4Net()
        {
            log4net.Repository.Hierarchy.Hierarchy hierarchy = (log4net.Repository.Hierarchy.Hierarchy)log4net.LogManager.GetRepository();

            log4net.Layout.PatternLayout patternLayout = new log4net.Layout.PatternLayout();
            patternLayout.ConversionPattern = "%date [%thread] %-5level %logger - %message%newline";
            patternLayout.ActivateOptions();

            log4net.Appender.RollingFileAppender roller = new log4net.Appender.RollingFileAppender();
            roller.AppendToFile = true;
            roller.File = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), @"LccCom\AppLog.txt");
            roller.Layout = patternLayout;
            roller.MaxSizeRollBackups = 5;
            roller.MaximumFileSize = "5120KB";
            roller.RollingStyle = log4net.Appender.RollingFileAppender.RollingMode.Size;
            roller.StaticLogFileName = true;
            roller.ActivateOptions();
            hierarchy.Root.AddAppender(roller);

            log4net.Appender.MemoryAppender memory = new log4net.Appender.MemoryAppender();
            memory.ActivateOptions();
            hierarchy.Root.AddAppender(memory);

            hierarchy.Root.Level = log4net.Core.Level.Debug;
            hierarchy.Configured = true;
        }

        public ExcelHelper()
        {
            SetupLog4Net();

            log = log4net.LogManager.GetLogger(typeof(ExcelHelper));
            log.Debug("log4net initialize completed");
        }

        public string Health()
        {
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }

        public string TestFile(string filename)
        {
            string directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location).ToString();
            FilePath = Path.Combine(directory, filename);

            ExcelBinary = File.ReadAllBytes(FilePath);
            return FilePath;
        }

        public void ConvertHtmlToExcel(string sheetName, string htmlText)
        {
            try
            {
                IWorkbook wb = new XSSFWorkbook();

                HtmlDocument document = new HtmlDocument();
                document.LoadHtml(htmlText);

                var tableList = document.DocumentNode.SelectNodes("//table");
                for (int t = 0; t < tableList.Count; t++)
                {
                    HtmlAttribute tableAttr = tableList[t].Attributes["id"];
                    if (tableAttr != null)
                    {
                        sheetName = tableAttr.Value;
                    }
                    tableAttr = tableList[t].Attributes["class"];
                    bool addFilter = tableAttr != null && tableAttr.Value.ToLower() == "filter";

                    ISheet ws = wb.CreateSheet(sheetName);

                    IFont font = wb.CreateFont();
                    font.Boldweight = (short)FontBoldWeight.Normal;

                    ICellStyle style = wb.CreateCellStyle();
                    style.BorderBottom = BorderStyle.Thin;
                    style.BorderTop = BorderStyle.Thin;
                    style.BorderLeft = BorderStyle.Thin;
                    style.BorderRight = BorderStyle.Thin;
                    style.WrapText = true;
                    style.SetFont(font);

                    IFont headerFont = wb.CreateFont();
                    headerFont.Boldweight = (short)FontBoldWeight.Bold;

                    ICellStyle headerStyle = wb.CreateCellStyle();
                    headerStyle.BorderBottom = BorderStyle.Thin;
                    headerStyle.BorderTop = BorderStyle.Thin;
                    headerStyle.BorderLeft = BorderStyle.Thin;
                    headerStyle.BorderRight = BorderStyle.Thin;
                    headerStyle.WrapText = true;

                    IDataFormat dateFormat = wb.CreateDataFormat();

                    ICellStyle dateStyle = wb.CreateCellStyle();
                    dateStyle.BorderBottom = BorderStyle.Thin;
                    dateStyle.BorderTop = BorderStyle.Thin;
                    dateStyle.BorderLeft = BorderStyle.Thin;
                    dateStyle.BorderRight = BorderStyle.Thin;
                    dateStyle.WrapText = true;
                    dateStyle.SetFont(font);
                    dateStyle.DataFormat = dateFormat.GetFormat("mm/dd/yyyy");

                    IDataFormat stringFormat = wb.CreateDataFormat();

                    ICellStyle stringStyle = wb.CreateCellStyle();
                    stringStyle.BorderBottom = BorderStyle.Thin;
                    stringStyle.BorderTop = BorderStyle.Thin;
                    stringStyle.BorderLeft = BorderStyle.Thin;
                    stringStyle.BorderRight = BorderStyle.Thin;
                    stringStyle.WrapText = true;
                    stringStyle.SetFont(font);
                    stringStyle.DataFormat = stringFormat.GetFormat("@");

                    List<int> columnWidthList = new List<int>();
                    List<string> columnDataTypeList = new List<string>();
                    bool hasColumn = false;

                    var trList = tableList[t].SelectNodes("tr");
                    for (int i = 0; i < trList.Count; i++)
                    {
                        IRow row = ws.CreateRow(i);
                        if (i == 0)
                        {
                            //get header row
                            var thList = trList[i].SelectNodes("th");
                            for (int j = 0; j < thList.Count; j++)
                            {
                                hasColumn = true;
                                ICell cell = row.CreateCell(j);
                                cell.SetCellValue(thList[j].InnerHtml);
                                cell.CellStyle = headerStyle;
                                HtmlAttribute attr = thList[j].Attributes["width"];
                                if (attr != null)
                                    columnWidthList.Add(int.Parse(attr.Value));
                                attr = thList[j].Attributes["class"];
                                if (attr != null)
                                    columnDataTypeList.Add(attr.Value);
                                else
                                    columnDataTypeList.Add("string");
                            }
                        }
                        else
                        {
                            var tdList = trList[i].SelectNodes("td");
                            for (int j = 0; j < tdList.Count; j++)
                            {
                                ICell cell = row.CreateCell(j);
                                string cellType = columnDataTypeList[j];
                                HtmlAttribute attr = tdList[j].Attributes["class"];
                                if (attr != null)
                                    cellType = attr.Value;
                                switch (cellType)
                                {
                                    case "datetime":
                                        DateTime date = DateTime.MinValue;
                                        if (DateTime.TryParseExact(tdList[j].InnerHtml, "yyyy/M/d", null, System.Globalization.DateTimeStyles.None, out date))
                                            cell.SetCellValue(date);
                                        else
                                            cell.SetCellValue(tdList[j].InnerHtml);
                                        cell.CellStyle = dateStyle;
                                        break;
                                    case "money":
                                    case "int":
                                        int data = int.MinValue;
                                        if (int.TryParse(tdList[j].InnerHtml, out data))
                                            cell.SetCellValue(data);
                                        else
                                            cell.SetCellValue(tdList[j].InnerHtml);
                                        cell.SetCellType(CellType.Numeric);
                                        cell.CellStyle = style;
                                        break;
                                    default:
                                        cell.SetCellValue(tdList[j].InnerHtml);
                                        cell.SetCellType(CellType.String);
                                        cell.CellStyle = stringStyle;
                                        break;
                                }
                            }
                        }
                    }

                    if (hasColumn && addFilter)
                    {
                        //freeze column header row
                        ws.CreateFreezePane(0, 1);
                        ws.SetAutoFilter(new CellRangeAddress(0, ws.LastRowNum, 0, columnWidthList.Count - 1));
                    }

                    for (int c = 0; c < columnWidthList.Count; c++)
                    {
                        //set column width
                        ws.SetColumnWidth(c, columnWidthList[c] * 256);
                    }
                }

                MemoryStream stream = new MemoryStream();
                wb.Write(stream);
                ExcelBinary = stream.ToArray();
                stream.Close();
            }
            catch (Exception ex)
            {
                using (EventLog eventLog = new EventLog("Application"))
                {
                    eventLog.Source = "LccCom.ExcelHelper";
                    eventLog.WriteEntry(ex.Message, EventLogEntryType.Error, 101, 1);
                }
            }
        }

        public void HtmlToExcel(string html)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Data");

            

            using (MemoryStream stream = new MemoryStream())
            {
                workbook.Write(stream);
                ExcelBinary = stream.ToArray();
            }
        }
    }
}
