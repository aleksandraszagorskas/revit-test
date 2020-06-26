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
                if (schedule != null)
                {
                    ViewSheet sheet = CreateScheduleSheet(doc, schedule, scheduleName);
                    if (sheet != null)
                    {
                        uidoc.ActiveView = sheet;

                        ExportSheetToPDF(doc, sheet, scheduleName);
                    }
                    else
                    {
                        return Result.Failed;
                    }
                }
                else
                {
                    return Result.Failed;
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void ExportSheetToPDF(Document doc, ViewSheet sheet, string fileName)
        {
            using (Transaction transaction = new Transaction(doc, "Export sheet"))
            {
                transaction.Start();

                //print
                var printManager = doc.PrintManager;
                printManager.SelectNewPrintDriver("Microsoft Print To PDF");
                printManager.PrintRange = PrintRange.Select;
                printManager.CombinedFile = true;

                ViewSheetSetting viewSheetSetting = printManager.ViewSheetSetting;

                var viewSet = new ViewSet();
                viewSet.Insert(sheet);

                viewSheetSetting.CurrentViewSheetSet.Views = viewSet;
                viewSheetSetting.Save();

                printManager.SubmitPrint();

                transaction.Commit();
            }
        }

        private ViewSheet CreateScheduleSheet(Document doc, ViewSchedule schedule, string scheduleName)
        {
            ViewSheet sheet = null;
            using (Transaction transaction = new Transaction(doc, "Create sheet"))
            {
                transaction.Start();

                ElementId titleBlockId = GetTitleBlockId(doc);

                sheet = ViewSheet.Create(doc, ElementId.InvalidElementId);
                sheet.Name = scheduleName;
                var scheduleInstance = ScheduleSheetInstance.Create(doc, sheet.Id, schedule.Id, new XYZ());

                transaction.Commit();
            }

            return sheet;
        }

        private ElementId GetTitleBlockId(Document doc)
        {
            var titleBlockName = "Test Titile Block";
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);

            List<Element> titleBlocks = collector.ToList();
            return titleBlocks.Where(tb => tb.Name == titleBlockName).First().Id;
        }

        private ViewSchedule CreateSchedule(Document doc, string scheduleName, ref string message)
        {
            ViewSchedule schedule = null;
            using (Transaction transaction = new Transaction(doc, "Create schedule"))
            {
                transaction.Start();

                schedule = ViewSchedule.CreateSchedule(doc, new ElementId(BuiltInCategory.OST_Walls));
                try
                {
                    schedule.Name = scheduleName;
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                    doc.Delete(schedule.Id);
                    return null;
                }

                var schedulableFields = schedule.Definition.GetSchedulableFields();
                var schedulableInstances = schedulableFields.Where(f => f.FieldType == ScheduleFieldType.Instance).ToList();
                var allowedFields = new string[] { "Type", "Family", "Volume", "Area", "Page" };
                //ScheduleField pageField = null;
                foreach (SchedulableField field in schedulableInstances)
                {
                    var name = field.GetName(doc);
                    if (allowedFields.Contains(name))
                    {
                        //if (name == "Page")
                        //{
                        //    pageField = schedule.Definition.AddField(field);
                        //}
                        //else
                        //{
                        //    schedule.Definition.AddField(field);
                        //}
                        schedule.Definition.AddField(field);
                    }
                }

                doc.Regenerate();

                //group headers
                if (schedule.CanGroupHeaders(0, 2, 0, 3))
                {
                    schedule.GroupHeaders(0, 2, 0, 3, "Measurements");
                }

                //filter
                ScheduleField pageField = schedule.Definition.GetField(4);
                pageField.IsHidden = true;

                int val = 2;
                var filter = new ScheduleFilter(pageField.FieldId, ScheduleFilterType.Equal, val);
                //pageField.Definition.AddFilter(filter);
                schedule.Definition.AddFilter(filter);


                transaction.Commit();
            }

            return schedule;
        }
    }
}
