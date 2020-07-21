using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitTest.Schedules.Utilities.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using Application = Autodesk.Revit.ApplicationServices.Application;

namespace RevitTest.Schedules.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ChangeSchedulePageCommand : IExternalCommand
    {
        private readonly int WORKING_PAGE_NUMBER = 3;
        private readonly bool IS_NEW_PAGE = false;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            SetSelectionPage(uidoc, IS_NEW_PAGE);


            //TaskDialog.Show("Test", "ChangeSchedulePageCommand");

            return Result.Succeeded;
        }

        private void SetSelectionPage(UIDocument uidoc, bool isNewPage)
        {
            Document doc = uidoc.Document;

            if (doc.ActiveView is ViewSchedule)
            {
                List<ElementId> selectedItemIds = uidoc.Selection.GetElementIds().ToList();
                ViewSchedule schedule = doc.ActiveView as ViewSchedule;

                if (selectedItemIds.Any())
                {
                    if (isNewPage)
                    {
                        ResetPage(doc, schedule, WORKING_PAGE_NUMBER);
                    }

                    SetPage(doc, selectedItemIds, WORKING_PAGE_NUMBER);
                }
            }
            else
            {
                TaskDialog.Show("Test", "Please change paging from schedule view.");
            }
        }

        private void ResetPage(Document doc, ViewSchedule schedule, int pageToReset)
        {
            using (var tran = new Transaction(doc, "Handle element"))
            {
                tran.Start();
                List<Element> elements = ViewScheduleUtilities.GetScheduleElements(doc, schedule);

                foreach (Element element in elements)
                {
                    var pageParamList = element.GetParameters("My Page");

                    if (pageParamList.Count == 1)
                    {
                        Parameter pageParam = pageParamList.First();
                        if (pageParam.StorageType == StorageType.Integer && pageParam.AsInteger() == pageToReset)
                        {
                            pageParam.Set(-1);
                        }
                    }
                }

                tran.Commit();
            }
        }

        private void SetPage(Document doc, List<ElementId> selectedIds, int pageNumber)
        {
            foreach (ElementId id in selectedIds)
            {
                Element element = doc.GetElement(id);

                var pageParamList = element.GetParameters("My Page");

                if (pageParamList.Count == 1)
                {
                    using (var tran = new Transaction(doc, "Handle element"))
                    {
                        tran.Start();

                        Parameter pageParam = pageParamList.First();
                        if (pageParam.StorageType == StorageType.Integer)
                        {
                            pageParam.Set(pageNumber);
                        }

                        tran.Commit();
                    }
                }

            }
        }

    }
}
