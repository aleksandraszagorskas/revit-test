using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitTest.Schedules.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class TestScheduleCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                var scheduleName = "Wall schedule";

                ViewSchedule schedule = CreateSchedule(doc, scheduleName, ref message);
                HandleSchedule(doc, schedule);
                ViewSheet sheet = CreateScheduleSheet(doc, schedule, scheduleName);
                uidoc.ActiveView = sheet;

                ExportSheetToPDF(doc, sheet, scheduleName);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void HandleSchedule(Document doc, ViewSchedule schedule)
        {
            ScheduleField field = null;
            try
            {
                field = schedule.Definition.GetField(4);
            }
            catch (Exception ex)
            {
                field = null;
                throw new Exception("Page parameter does not exist");
            }

            if (field != null)
            {
                using (Transaction transaction = new Transaction(doc, "Handle schedule"))
                {
                    transaction.Start();

                    field.IsHidden = true;

                    try
                    {
                        var filter = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, 2);
                        schedule.Definition.AddFilter(filter);
                    }
                    catch (Exception ex)
                    {
                        //Filter not added
                        TaskDialog.Show("Warning", "Filter not applied");
                    }



                    transaction.Commit();
                }
            }


        }

        /// <summary>
        /// Export sheet to PDF
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="sheet"></param>
        /// <param name="fileName"></param>
        private void ExportSheetToPDF(Document doc, ViewSheet sheet, string fileName)
        {
            //print
            var printManager = doc.PrintManager;
            printManager.SelectNewPrintDriver("Microsoft Print To PDF");
            printManager.PrintRange = PrintRange.Select;
            printManager.CombinedFile = true;

            ViewSheetSetting viewSheetSetting = printManager.ViewSheetSetting;

            var viewSet = new ViewSet();
            viewSet.Insert(sheet);

            using (Transaction transaction = new Transaction(doc, "Export sheet"))
            {
                transaction.Start();
                viewSheetSetting.InSession.Views = viewSet;
                transaction.Commit();
            }

            printManager.SubmitPrint();
        }


        /// <summary>
        /// Create sheet for schedule
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="schedule"></param>
        /// <param name="scheduleName"></param>
        /// <returns></returns>
        private ViewSheet CreateScheduleSheet(Document doc, ViewSchedule schedule, string scheduleName)
        {
            ViewSheet sheet = null;
            using (Transaction transaction = new Transaction(doc, "Create sheet"))
            {
                transaction.Start();

                //ElementId titleBlockId = GetTitleBlockId(doc);

                sheet = ViewSheet.Create(doc, ElementId.InvalidElementId);
                sheet.Name = scheduleName;
                var scheduleInstance = ScheduleSheetInstance.Create(doc, sheet.Id, schedule.Id, new XYZ());

                transaction.Commit();
            }

            return sheet;
        }

        //private ElementId GetTitleBlockId(Document doc)
        //{
        //    var titleBlockName = "Test Titile Block";
        //    FilteredElementCollector collector = new FilteredElementCollector(doc);
        //    collector.OfClass(typeof(FamilySymbol));
        //    collector.OfCategory(BuiltInCategory.OST_TitleBlocks);

        //    List<Element> titleBlocks = collector.ToList();
        //    return titleBlocks.Where(tb => tb.Name == titleBlockName).First().Id;
        //}

        /// <summary>
        /// Create schedule
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="scheduleName"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        private ViewSchedule CreateSchedule(Document doc, string scheduleName, ref string message)
        {
            ViewSchedule schedule = GetSchedule(doc, scheduleName);
            if (schedule == null)
            {
                using (Transaction transaction = new Transaction(doc, "Create schedule"))
                {
                    transaction.Start();

                    schedule = ViewSchedule.CreateSchedule(doc, new ElementId(BuiltInCategory.OST_Walls));
                    schedule.Name = scheduleName;

                    var schedulableFields = schedule.Definition.GetSchedulableFields();
                    var schedulableInstances = schedulableFields.Where(f => f.FieldType == ScheduleFieldType.Instance).ToList();
                    var allowedFields = new string[] { "Type", "Family", "Volume", "Area", "Page" };

                    foreach (SchedulableField schField in schedulableInstances)
                    {
                        var name = schField.GetName(doc);
                        if (allowedFields.Contains(name))
                        {
                            var field = schedule.Definition.AddField(schField);
                            //if (name == "Page")
                            //{
                            //    //filter
                            //    field.IsHidden = true;

                            //    var filter = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, 2);
                            //    schedule.Definition.AddFilter(filter);
                            //}
                        }
                    }

                    doc.Regenerate();

                    //group headers
                    if (schedule.CanGroupHeaders(0, 2, 0, 3))
                    {
                        schedule.GroupHeaders(0, 2, 0, 3, "Measurements");
                    }

                    transaction.Commit();
                }
            }
            return schedule;
        }

        /// <summary>
        /// Get existing schedule
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="scheduleName"></param>
        /// <returns></returns>
        private ViewSchedule GetSchedule(Document doc, string scheduleName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewSchedule));
            collector.OfCategory(BuiltInCategory.OST_Schedules);

            List<Element> schedules = collector.ToList();

            var scheduleElements = schedules.Where(s => s.Name.Contains(scheduleName)).ToList();

            if (scheduleElements.Count == 1)
            {
                return scheduleElements.First() as ViewSchedule;
            }

            return null;
        }
    }
}
