using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;
using RevitTest.Schedules.Misc;
using RevitTest.Schedules.Utilities;

namespace RevitTest.Schedules.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ExportScheduleDataCommand : IExternalCommand
    {
        private Misc.Settings settings = Misc.Settings.Instance;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            ExportScheduleData(uidoc);

            return Result.Succeeded;
        }

        private void ExportScheduleData(UIDocument uidoc)
        {
            Document doc = uidoc.Document;

            List<ElementId> selectedItemIds = uidoc.Selection.GetElementIds().ToList();

            if (selectedItemIds.Count() == 1)
            {
                Element element = doc.GetElement(selectedItemIds.First());
                if (element is ScheduleSheetInstance)
                {
                    var scheduleInstance = element as ScheduleSheetInstance;
                    ElementId scheduleId = scheduleInstance.ScheduleId;
                    ViewSchedule schedule = doc.GetElement(scheduleId) as ViewSchedule;

                    TableData tableData = schedule.GetTableData();
                    TableSectionData tableBodySection = tableData.GetSectionData(SectionType.Body);
                    int numOfRows = tableBodySection.NumberOfRows;
                    int numOfCols = tableBodySection.NumberOfColumns;
                    var cellData = new string[numOfRows, numOfCols];

                    for (int i = 0; i < numOfRows; i++)
                    {
                        string rowLine = "";
                        for (int j = 0; j < numOfCols; j++)
                        {
                            cellData[i, j] = tableBodySection.GetCellText(i, j);
                            rowLine += $"{cellData[i, j]}   ";
                        }
                        Console.WriteLine(rowLine);
                    }

                    string exportFileName = Path.Combine(settings.EXCEL_EXPORT_FOLDER, $"{schedule.Name}.csv");

                    ExportSchedule(schedule, exportFileName);

                    var rng = new Random();
                    var rngs = new string[] { "apple", "orange", "kiwi", "L", "W", "air", "metal" };

                    var testData = new List<TestData>();
                    for (int i = 0; i < rng.Next(90, 120); i++)
                    {
                        testData.Add(new TestData { Val1 = $"{rngs[rng.Next(0, rngs.Length)]} {i}" });
                    }
                    CSVUtilities.ExportList(settings.EXCEL_TEST_DATA_PATH, testData);

                    var testData2 = new List<TestData>();
                    for (int i = 0; i < rng.Next(10, 40); i++)
                    {
                        testData2.Add(new TestData { Val1 = $"{rngs[rng.Next(0, rngs.Length)]} {i}", Val2 = $"{rngs[rng.Next(0, rngs.Length)]} {i}", Val3 = $"{rngs[rng.Next(0, rngs.Length)]} {i}" });
                    }
                    CSVUtilities.ExportList(settings.EXCEL_TEST_DATA_2_PATH, testData2);

                    CreateExcelSpreadsheet(new string[] { exportFileName, settings.EXCEL_TEST_DATA_PATH, settings.EXCEL_TEST_DATA_2_PATH });
                    //TestCreateExcelSpreadsheet();
                }
                else
                {
                    TaskDialog.Show("Notice", "Selected element should be schedule instance");
                }
            }
            else
            {
                TaskDialog.Show("Notice", "Please select a single element");
            }
        }

        private static void ExportSchedule(ViewSchedule schedule, string exportFileName)
        {
            var exportOptions = new ViewScheduleExportOptions
            {
                Title = false,
                ColumnHeaders = ExportColumnHeaders.None,
                HeadersFootersBlanks = true,
                FieldDelimiter = ",",
                TextQualifier = ExportTextQualifier.DoubleQuote
            };

            schedule.Export(Path.GetDirectoryName(exportFileName), Path.GetFileName(exportFileName), exportOptions);
        }

        #region example
        private void TestCreateExcelSpreadsheet()
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "First Name";
                oSheet.Cells[1, 2] = "Last Name";
                oSheet.Cells[1, 3] = "Full Name";
                oSheet.Cells[1, 4] = "Salary";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "D1").Font.Bold = true;
                oSheet.get_Range("A1", "D1").VerticalAlignment =
                Excel.XlVAlign.xlVAlignCenter;

                // Create an array to multiple values at once.
                string[,] saNames = new string[5, 2];

                saNames[0, 0] = "John";
                saNames[0, 1] = "Smith";
                saNames[1, 0] = "Tom";
                saNames[1, 1] = "Brown";
                saNames[2, 0] = "Sue";
                saNames[2, 1] = "Thomas";
                saNames[3, 0] = "Jane";
                saNames[3, 1] = "Jones";
                saNames[4, 0] = "Adam";
                saNames[4, 1] = "Johnson";

                //Fill A2:B6 with an array of values (First and Last Names).
                oSheet.get_Range("A2", "B6").Value2 = saNames;

                //Fill C2:C6 with a relative formula (=A2 & " " & B2).
                oRng = oSheet.get_Range("C2", "C6");
                oRng.Formula = "=A2 & \" \" & B2";

                //Fill D2:D6 with a formula(=RAND()*100000) and apply format.
                oRng = oSheet.get_Range("D2", "D6");
                oRng.Formula = "=RAND()*100000";
                oRng.NumberFormat = "$0.00";

                //AutoFit columns A:D.
                oRng = oSheet.get_Range("A1", "D1");
                oRng.EntireColumn.AutoFit();

                //Manipulate a variable number of columns for Quarterly Sales Data.
                DisplayQuarterlySales(oSheet);

                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                oXL.Visible = true;
                oXL.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        private void DisplayQuarterlySales(Excel._Worksheet oWS)
        {
            Excel._Workbook oWB;
            Excel.Series oSeries;
            Excel.Range oResizeRange;
            Excel._Chart oChart;
            String sMsg;
            int iNumQtrs;

            //Determine how many quarters to display data for.
            for (iNumQtrs = 4; iNumQtrs >= 2; iNumQtrs--)
            {
                sMsg = "Enter sales data for ";
                sMsg = String.Concat(sMsg, iNumQtrs);
                sMsg = String.Concat(sMsg, " quarter(s)?");

                DialogResult iRet = MessageBox.Show(sMsg, "Quarterly Sales?",
                MessageBoxButtons.YesNo);
                if (iRet == DialogResult.Yes)
                    break;
            }

            sMsg = "Displaying data for ";
            sMsg = String.Concat(sMsg, iNumQtrs);
            sMsg = String.Concat(sMsg, " quarter(s).");

            MessageBox.Show(sMsg, "Quarterly Sales");

            //Starting at E1, fill headers for the number of columns selected.
            oResizeRange = oWS.get_Range("E1", "E1").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=\"Q\" & COLUMN()-4 & CHAR(10) & \"Sales\"";

            //Change the Orientation and WrapText properties for the headers.
            oResizeRange.Orientation = 38;
            oResizeRange.WrapText = true;

            //Fill the interior color of the headers.
            oResizeRange.Interior.ColorIndex = 36;

            //Fill the columns with a formula and apply a number format.
            oResizeRange = oWS.get_Range("E2", "E6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=RAND()*100";
            oResizeRange.NumberFormat = "$0.00";

            //Apply borders to the Sales data and headers.
            oResizeRange = oWS.get_Range("E1", "E6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            //Add a Totals formula for the sales data and apply a border.
            oResizeRange = oWS.get_Range("E8", "E8").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=SUM(E2:E6)";
            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle
            = Excel.XlLineStyle.xlDouble;
            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight
            = Excel.XlBorderWeight.xlThick;

            //Add a Chart for the selected data.
            oWB = (Excel._Workbook)oWS.Parent;
            oChart = (Excel._Chart)oWB.Charts.Add(Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);

            //Use the ChartWizard to create a new chart from the selected data.
            oResizeRange = oWS.get_Range("E2:E6", Missing.Value).get_Resize(
            Missing.Value, iNumQtrs);
            oChart.ChartWizard(oResizeRange, Excel.XlChartType.xl3DColumn, Missing.Value,
            Excel.XlRowCol.xlColumns, Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            oSeries = (Excel.Series)oChart.SeriesCollection(1);
            oSeries.XValues = oWS.get_Range("A2", "A6");
            for (int iRet = 1; iRet <= iNumQtrs; iRet++)
            {
                oSeries = (Excel.Series)oChart.SeriesCollection(iRet);
                String seriesName;
                seriesName = "=\"Q";
                seriesName = String.Concat(seriesName, iRet);
                seriesName = String.Concat(seriesName, "\"");
                oSeries.Name = seriesName;
            }

            oChart.Location(Excel.XlChartLocation.xlLocationAsObject, oWS.Name);

            //Move the chart so as not to cover your data.
            oResizeRange = (Excel.Range)oWS.Rows.get_Item(10, Missing.Value);
            oWS.Shapes.Item("Chart 1").Top = (float)(double)oResizeRange.Top;
            oResizeRange = (Excel.Range)oWS.Columns.get_Item(2, Missing.Value);
            oWS.Shapes.Item("Chart 1").Left = (float)(double)oResizeRange.Left;
        }
        #endregion

        private void CreateExcelSpreadsheet(string[] paths/*, ViewSchedule schedule*/)
        {


            //var lines = File.ReadAllLines(fileName);

            ////Number of rows including header
            //var numOfRows = lines.Count();

            ////Number of rows columns
            //var numOfCols = lines[0].Split(',').Count();

            Excel.Application oXL;
            Excel.Workbook oWB;
            Excel.Worksheet oSheet;

            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = false;

                //Get a new workbook.
                oWB = oXL.Workbooks.Open(settings.EXCEL_TEMPLATE_PATH);
                oSheet = (Excel.Worksheet)oWB.ActiveSheet;

                //Excel.PageSetup pageSetup = oSheet.PageSetup;
                //var leftMargin = pageSetup.LeftMargin;
                //var rightMargin = pageSetup.RightMargin;
                //var topMargin = pageSetup.TopMargin;
                //var botMargin = pageSetup.BottomMargin;

                //Excel.Range topRange = oSheet.Cells[leftMargin, rightMargin];
                //Excel.Range sideRange = oSheet.Cells[topMargin, botMargin];
                //Excel.Range testRange = oSheet.Range[topRange, sideRange];
                //testRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Salmon);

                Excel.Range pageRange = oSheet.Range["A1", "K17"];
                //pageRange.Copy();

                //Excel.Range startCellRange = pageRange.EntireColumn.Find("{first-cell}", Missing.Value, Excel.XlFindLookIn.xlValues, Missing.Value, Missing.Value, Excel.XlSearchDirection.xlNext, false, false, Missing.Value);
                //startCellRange.Value2 = "";
                //Excel.Range endCellRange = pageRange.EntireColumn.Find("{last-cell}", Missing.Value, Excel.XlFindLookIn.xlValues, Missing.Value, Missing.Value, Excel.XlSearchDirection.xlNext, false, false, Missing.Value);
                //endCellRange.Value2 = "";

                //Excel.Range bodyRange = oSheet.Range[startCellRange, endCellRange];

                var test = new SchedulePage(oSheet, "A1", "K17", paths);
                test.CreatePages();

                //int itemsPerPage = bodyRange.Rows.Count;
                ////int pageCount = numOfRows / itemsPerPage;
                //int pageCount = (numOfRows + itemsPerPage - 1) / itemsPerPage;



                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                oXL.Visible = true;
                oXL.UserControl = true;
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        private List<SchedulePage> FindTableCells(Excel.Range searchRange, string tableName)
        {
            var cells = new List<Excel.Range>();
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;

            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            currentFind = searchRange.Find($"{{first-cell:'{tableName}'}}", Missing.Value,
                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                Missing.Value, Missing.Value);

            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                    cells.Add(currentFind);
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                //handle found element
                currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Salmon);

                currentFind = searchRange.FindNext(currentFind);
                if (currentFind != null)
                {
                    cells.Add(currentFind);
                }
            }

            return null;
        }

        public void ImportCSV(string importFileName, Excel.Worksheet destinationSheet, Excel.Range destinationRange, int[] columnDataTypes)
        {
            destinationSheet.QueryTables.Add(
                "TEXT;" + Path.GetFullPath(importFileName),
            destinationRange, Type.Missing);
            destinationSheet.QueryTables[1].Name = Path.GetFileNameWithoutExtension(importFileName);
            destinationSheet.QueryTables[1].FieldNames = true;
            destinationSheet.QueryTables[1].RowNumbers = false;
            destinationSheet.QueryTables[1].FillAdjacentFormulas = false;
            destinationSheet.QueryTables[1].PreserveFormatting = true;
            destinationSheet.QueryTables[1].RefreshOnFileOpen = false;
            destinationSheet.QueryTables[1].RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells;
            destinationSheet.QueryTables[1].SavePassword = false;
            destinationSheet.QueryTables[1].SaveData = true;
            destinationSheet.QueryTables[1].AdjustColumnWidth = true;
            destinationSheet.QueryTables[1].RefreshPeriod = 0;
            destinationSheet.QueryTables[1].TextFilePromptOnRefresh = false;
            destinationSheet.QueryTables[1].TextFilePlatform = 437;
            destinationSheet.QueryTables[1].TextFileStartRow = 1;
            destinationSheet.QueryTables[1].TextFileParseType = Excel.XlTextParsingType.xlDelimited;
            destinationSheet.QueryTables[1].TextFileTextQualifier = Excel.XlTextQualifier.xlTextQualifierDoubleQuote;
            destinationSheet.QueryTables[1].TextFileConsecutiveDelimiter = false;
            destinationSheet.QueryTables[1].TextFileTabDelimiter = false;
            destinationSheet.QueryTables[1].TextFileSemicolonDelimiter = false;
            destinationSheet.QueryTables[1].TextFileCommaDelimiter = true;
            destinationSheet.QueryTables[1].TextFileSpaceDelimiter = false;
            destinationSheet.QueryTables[1].TextFileColumnDataTypes = columnDataTypes;

            //Logger.GetInstance().WriteLog("Importing data...");
            destinationSheet.QueryTables[1].Refresh(false);
            destinationSheet.QueryTables[1].Destination.EntireColumn.AutoFit();
            //destinationSheet.QueryTables[1].Destination.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            //destinationSheet.QueryTables[1].Destination.EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

            // cleanup
            //this.ActiveSheet.QueryTables[1].Delete();
        }
    }

    public class RevitTableData
    {
        public string DataFilePath { get; set; }

        public int RowCount { get; set; }
        public int ColumnCount { get; set; }

        public Excel.Range StartCell { get; set; }

        public int ItemsPerPage { get; set; }
        public int ColumnOffset { get; set; }

        public int PageCount => (RowCount + ItemsPerPage - 1) / ItemsPerPage;
    }

    public class SchedulePage
    {
        private Misc.Settings settings = Misc.Settings.Instance;

        public Dictionary<string, RevitTableData> RevitTables { get; set; } = new Dictionary<string, RevitTableData>();

        public Excel.Worksheet Sheet { get; set; }
        public Excel.Range PageRange { get; set; }

        public int PageRows { get; set; }
        public int PageColumns { get; set; }

        public int PageCount { get; set; }

        public SchedulePage(Excel.Worksheet sheet, object pageCell1, object pageCell2, string[] paths)
        {
            Sheet = sheet;
            PageRange = Sheet.Range[pageCell1, pageCell2];

            PageRows = PageRange.Rows.Count;
            PageColumns = PageRange.Columns.Count;

            Excel.Range pageFirstCell = PageRange.Cells[1, 1];
            Excel.Range pageCopyCell = pageFirstCell.Offset[0, PageColumns];
            PageRange.Copy(pageCopyCell);
            Excel.Range pageCopyLastCell = pageCopyCell.Offset[PageRows - 1, PageColumns - 1];
            Excel.Range copyRange = Sheet.Range[pageCopyCell, pageCopyLastCell];

            foreach (var filePath in paths)
            {
                if (Path.GetExtension(filePath) == ".csv")
                {
                    string fileName = Path.GetFileNameWithoutExtension(filePath);

                    var lines = File.ReadAllLines(filePath);

                    //Number of rows including header
                    int rowCount = lines.Count();

                    //Number of rows columns
                    int columnCount = lines[0].Split(',').Count();

                    //items per page 
                    Excel.Range startCell = PageRange.EntireColumn.Find($"{{first-cell:'{fileName}'}}", Missing.Value, Excel.XlFindLookIn.xlValues, Missing.Value, Missing.Value, Excel.XlSearchDirection.xlNext, false, false, Missing.Value);
                    Excel.Range endCell = PageRange.EntireColumn.Find($"{{last-cell:'{fileName}'}}", Missing.Value, Excel.XlFindLookIn.xlValues, Missing.Value, Missing.Value, Excel.XlSearchDirection.xlNext, false, false, Missing.Value);
                    Excel.Range tableRange = PageRange.Range[startCell, endCell];

                    int itemsPerPage = tableRange.Rows.Count;
                    int pageCount = (rowCount + itemsPerPage - 1) / itemsPerPage;

                    //copy offset
                    Excel.Range copyStartCell = copyRange.EntireColumn.Find($"{{first-cell:'{fileName}'}}", Missing.Value, Excel.XlFindLookIn.xlValues, Missing.Value, Missing.Value, Excel.XlSearchDirection.xlNext, false, false, Missing.Value);
                    Excel.Range copyEndCell = copyRange.EntireColumn.Find($"{{last-cell:'{fileName}'}}", Missing.Value, Excel.XlFindLookIn.xlValues, Missing.Value, Missing.Value, Excel.XlSearchDirection.xlNext, false, false, Missing.Value);
                    Excel.Range copyOffsetRange = Sheet.Range[startCell, copyStartCell];

                    int copyOffset = copyOffsetRange.Columns.Count;

                    RevitTables.Add(fileName, new RevitTableData { DataFilePath = filePath, RowCount = rowCount, ColumnCount = columnCount, StartCell = startCell, ItemsPerPage = itemsPerPage, ColumnOffset = copyOffset - 1 });

                    //cleanup
                    startCell.Value2 = "";
                    endCell.Value2 = "";
                }
                else
                {
                    TaskDialog.Show("Warning", $" File {Path.GetFileName(filePath)} not supported. Only csv data files are supported.");
                }
            }

            PageCount = RevitTables.Values.Max(t => t.PageCount);

            //cleanup
            copyRange.Clear();
        }

        //public int GetPageCount()
        //{
        //    return (numOfRows + itemsPerPage - 1) / itemsPerPage;
        //}

        public void CreatePages()
        {
            //populate clipboard
            PageRange.Copy();

            //copy template to pages
            CopyTemplateToPages();

            //clear clipboard
            Clipboard.Clear();

            foreach (var table in RevitTables)
            {
                //populate table data
                PopulateTableData(table.Value);
            }

        }

        private void PopulateTableData(RevitTableData table)
        {
            using (TextFieldParser parser = new TextFieldParser(table.DataFilePath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                Excel.Range prevRowCell = null;
                int itemNum = 1;

                var startCellRange = table.StartCell;

                //fill pages
                while (!parser.EndOfData)
                {
                    //Process row
                    string[] fields = parser.ReadFields();

                    Excel.Range rowCell = null;
                    if (prevRowCell == null)
                    {
                        rowCell = startCellRange;
                    }
                    else
                    {
                        rowCell = prevRowCell.Offset[1, 0];
                    }

                    string[] subFields = fields.Skip(1).Take(fields.Length - 1).ToArray();
                    //header
                    if (!String.IsNullOrEmpty(fields[0]) && subFields.All(f => String.IsNullOrEmpty(f)))
                    {
                        Excel.Range rowEndCell = rowCell.Offset[0, table.ColumnCount - 1];
                        Excel.Range row = Sheet.Range[rowCell, rowEndCell];
                        row.Merge();

                        row.Value2 = fields[0];
                    }
                    //record
                    else
                    {
                        Excel.Range prevColCell = null;
                        foreach (string field in fields)
                        {
                            //TODO: Process field
                            Excel.Range colCell = null;
                            if (prevColCell == null)
                            {
                                colCell = rowCell;
                            }
                            else
                            {
                                colCell = prevColCell.Next;
                            }

                            colCell.Value2 = field;
                            if (colCell.Value2 is string && Path.GetExtension(colCell.Value2) == ".jpg")
                            {
                                var imageFileName = colCell.Value2 as string;

                                //remove cell value
                                colCell.Value2 = "";

                                string picPath = Path.Combine(settings.IMAGES_FOLDER, imageFileName);

                                float Left = (float)((double)colCell.Left) + settings.IMAGE_PADDING;
                                float Top = (float)((double)colCell.Top) + settings.IMAGE_PADDING;
                                float ImageWidth = (float)((double)colCell.Width) - 2 * settings.IMAGE_PADDING;
                                float ImageHeight = (float)((double)colCell.Height) - 2 * settings.IMAGE_PADDING;

                                Excel.Shape oShape = Sheet.Shapes.AddPicture(picPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, ImageWidth, ImageHeight);

                                //handle instance
                                oShape.Placement = Excel.XlPlacement.xlMoveAndSize;
                            }

                            prevColCell = colCell;
                        }

                    }

                    if (itemNum % table.ItemsPerPage == 0)
                    {
                        bool isMerged = startCellRange.MergeCells;
                        if (isMerged)
                        {
                            startCellRange.UnMerge();
                            Excel.Range prevStartCell = startCellRange;
                            startCellRange = startCellRange.Offset[0, table.ColumnOffset];

                            Excel.Range rowEndCell = prevStartCell.Offset[0, table.ColumnCount - 1];
                            Excel.Range row = Sheet.Range[prevStartCell, rowEndCell];
                            row.Merge();

                            rowCell = null;
                        }
                        else
                        {
                            startCellRange = startCellRange.Offset[0, table.ColumnOffset];

                            rowCell = null;
                        }


                    }

                    prevRowCell = rowCell;
                    itemNum++;
                }
            }
        }

        private void CopyTemplateToPages()
        {
            if (Clipboard.ContainsText())
            {
                Excel.Range prevCopyCell = null;
                for (int i = 0; i < PageCount; i++)
                {
                    Excel.Range copyCell = null;
                    if (prevCopyCell == null)
                    {
                        copyCell = PageRange.Cells[1, 1];
                    }
                    else
                    {
                        copyCell = prevCopyCell.Offset[0, PageColumns];
                    }

                    string clipText = Clipboard.GetText();
                    if (!String.IsNullOrEmpty(clipText))
                    {
                        copyCell.PasteSpecial(Excel.XlPasteType.xlPasteColumnWidths, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, true, false);
                        copyCell.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, true, false);
                    }

                    prevCopyCell = copyCell;
                }
            }
        }
    }

}
