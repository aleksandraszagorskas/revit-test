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

                    string exportFileName = Path.Combine(settings.EXCEL_EXPORT_FOLDER, $"{schedule.Name}.csv");

                    ExportSchedule(schedule, exportFileName);

                    #region test-data
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
                    #endregion

                    CreateExcelSpreadsheet(settings.EXCEL_TEMPLATE_PATH, new string[] { exportFileName, settings.EXCEL_TEST_DATA_PATH, settings.EXCEL_TEST_DATA_2_PATH });
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

        private void CreateExcelSpreadsheet(string templatePath, string[] dataFilePaths)
        {
            Excel.Application excelApp;
            Excel.Workbook workbook;
            Excel.Worksheet sheet;

            try
            {
                //Start Excel and get Application object.
                excelApp = new Excel.Application();
                excelApp.Visible = false;

                //Get a new workbook.
                workbook = excelApp.Workbooks.Open(templatePath);
                sheet = (Excel.Worksheet)workbook.ActiveSheet;

                var schedulePage = new SchedulePage(sheet, "A1", "K17", dataFilePaths);
                schedulePage.CreatePages();

                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                excelApp.Visible = true;
                excelApp.UserControl = true;
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
