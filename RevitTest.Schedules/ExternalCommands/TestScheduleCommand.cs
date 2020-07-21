using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using RevitTest.Schedules.Utilities;
using RevitTest.Schedules.Utilities.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Application = Autodesk.Revit.ApplicationServices.Application;

namespace RevitTest.Schedules.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class TestScheduleCommand : IExternalCommand
    {
        private readonly int PAGE_ITEM_COUNT = 12;
        private readonly string SHARED_PARAMETER_FILE_NAME = @"C:\Users\AleksandrasZagorskas\Desktop\revit_csd\schedules\HelloSharedParameterWorld.txt";
        private readonly string SHARED_PARAMETER_GROUP_NAME = "My Shared Parameter Group";
        private readonly string MAIN_TITLEBLOCK_NAME = "Test Titile Block";
        private readonly string EXPANDABLE_TITLEBLOCK_NAME = "Test Expandable Titile Block";
        private readonly string MAIN_LIST_NAME = "Title schedule";
        private readonly string EXPANDABLE_LIST_NAME = "Expandable schedule";
        private readonly string ANNOTATION_TEMPLATE_NAME = "rebar-test-image";
        private readonly string ROOT_FOLDER = @"C:\Users\AleksandrasZagorskas\Desktop\revit_csd\schedules";
        private readonly string IMAGES_FOLDER = @"C:\Users\AleksandrasZagorskas\Desktop\revit_csd\schedules\images";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                Application app = doc.Application;
                app.DocumentChanged += App_DocumentChanged;

                #region data
                FilteredElementCollector rebarCollector = new FilteredElementCollector(doc);
                rebarCollector.OfClass(typeof(Rebar));
                rebarCollector.OfCategory(BuiltInCategory.OST_Rebar);

                List<Element> rebars = rebarCollector.ToList();

                var rebarTypeIds = new List<ElementId>();
                var rebarTypes = new List<Element>();
                foreach (var rebar in rebars)
                {
                    ElementId typeId = rebar.GetTypeId();
                    if (!rebarTypeIds.Contains(typeId))
                    {
                        var rebarBarType = doc.GetElement(typeId);
                        rebarTypes.Add(rebarBarType);
                        rebarTypeIds.Add(typeId);
                    }
                }

                int pageCount = (rebars.Count() + PAGE_ITEM_COUNT - 1) / PAGE_ITEM_COUNT;
                #endregion


                var fileNames = new string[] { Path.Combine(ROOT_FOLDER, $"{MAIN_TITLEBLOCK_NAME}.rfa"), Path.Combine(ROOT_FOLDER, $"{EXPANDABLE_TITLEBLOCK_NAME}.rfa"), Path.Combine(ROOT_FOLDER, $"{ANNOTATION_TEMPLATE_NAME}.rfa") };
                var imageFilePath = Path.Combine(IMAGES_FOLDER, $"{ANNOTATION_TEMPLATE_NAME}.png");


                //debug
                //ImageUtilities.DeleteUnusedImages(doc);

                ImageUtilities.AddTypeImages(doc, rebarTypes, imageFilePath);
                ImageUtilities.AddInstanceImages(uidoc, rebars);

                FamilyUtilities.LoadRequiredFamilies(doc, fileNames);
                ParameterUtilities.AddSharedParameters(doc, SHARED_PARAMETER_FILE_NAME, SHARED_PARAMETER_GROUP_NAME);

                var sheets = new List<ViewSheet>();

                #region TitleBlock
                ViewSchedule titleSchedule = ViewScheduleUtilities.GetSchedule(doc, MAIN_LIST_NAME);
                if (titleSchedule == null)
                {
                    titleSchedule = ViewScheduleUtilities.CreateSchedule(doc, BuiltInCategory.OST_Walls, MAIN_LIST_NAME);
                }

                ViewSheet titleSheet = ViewSheetUtilities.CreateTitleSheet(doc, MAIN_TITLEBLOCK_NAME, titleSchedule, MAIN_LIST_NAME);
                sheets.Add(titleSheet);
                #endregion

                //Set active view
                uidoc.ActiveView = titleSheet;

                #region ExpandingBlock
                List<ViewSchedule> expandingSchedules = ViewScheduleUtilities.GetSchedules(doc, EXPANDABLE_LIST_NAME);
                if (expandingSchedules == null)
                {
                    expandingSchedules = ViewScheduleUtilities.CreateSchedules(doc, BuiltInCategory.OST_Rebar, EXPANDABLE_LIST_NAME, pageCount);
                }

                var expandingSheets = new List<ViewSheet>();
                for (int i = 0; i < expandingSchedules.Count; i++)
                {
                    ViewSheet expandingSheet = ViewSheetUtilities.CreateExpandableSheet(uidoc, EXPANDABLE_TITLEBLOCK_NAME, expandingSchedules[i], EXPANDABLE_LIST_NAME, i + 1);
                    expandingSheets.Add(expandingSheet);
                }
                sheets.AddRange(expandingSheets);
                #endregion

                PDFUtilities.ExportSheets(doc, sheets);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void App_DocumentChanged(object sender, Autodesk.Revit.DB.Events.DocumentChangedEventArgs e)
        {
            List<ElementId> addedElementIds = e.GetAddedElementIds().ToList();
            List<ElementId> deletedElementIds = e.GetDeletedElementIds().ToList();
        }
    }
}
