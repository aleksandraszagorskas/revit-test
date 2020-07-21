using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitTest.Schedules.Utilities;
using RevitTest.Schedules.Utilities.Revit.DB;
using RevitTest.Schedules.Utilities.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using SystemPoint = System.Drawing.Point;

namespace RevitTest.Schedules.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SplitScheduleCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            SplitSelectedSchedule(uidoc);

            return Result.Succeeded;
        }

        private void SplitSelectedSchedule(UIDocument uidoc)
        {
            Document doc = uidoc.Document;

            List<ElementId> selectedItemIds = uidoc.Selection.GetElementIds().ToList();
            //ViewSchedule schedule = doc.ActiveView as ViewSchedule;

            if (selectedItemIds.Count() == 1)
            {
                Element element = doc.GetElement(selectedItemIds.First());
                if (element is ScheduleSheetInstance)
                {
                    var scheduleInstance = element as ScheduleSheetInstance;

                    ViewScheduleUtilities.SplitSchedule(uidoc, scheduleInstance);
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

    }
}
