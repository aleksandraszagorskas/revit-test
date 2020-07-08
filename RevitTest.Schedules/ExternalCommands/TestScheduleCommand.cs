using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
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
    public class WallImage
    {
        public ElementId WallId { get; set; }
        public string FileBaseName { get; set; }
        public int A { get; set; }
        public int B { get; set; }
        public int C { get; set; }
        public int D { get; set; }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class TestScheduleCommand : IExternalCommand
    {
        private readonly int PAGE_ITEM_COUNT = 5;
        private readonly string MAIN_TITLEBLOCK_NAME = "Test Titile Block";
        private readonly string EXPANDABLE_TITLEBLOCK_NAME = "Test Expandable Titile Block";
        private readonly string ANNOTATION_TEMPLATE_NAME = "test image fam2";
        private readonly string ROOT_FOLDER = @"C:\Users\AleksandrasZagorskas\Desktop\revit_csd\schedules";
        private readonly string IMAGES_FOLDER = @"C:\Users\AleksandrasZagorskas\Desktop\revit_csd\schedules\images";

        //List<ViewSchedule> expandingSchedules = null;
        //List<ViewSheet> sheets = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                Application app = doc.Application;
                app.DocumentChanged += App_DocumentChanged;

                //var folderName = @"C:\Users\AleksandrasZagorskas\Desktop\revit_csd\schedules";
                var fileNames = new string[] { Path.Combine(ROOT_FOLDER, $"{MAIN_TITLEBLOCK_NAME}.rfa"), Path.Combine(ROOT_FOLDER, $"{EXPANDABLE_TITLEBLOCK_NAME}.rfa"), Path.Combine(ROOT_FOLDER, $"{ANNOTATION_TEMPLATE_NAME}.rfa") };


                AddWallTypeImages(doc);
                AddWallImages(uidoc);

                LoadRequiredFamilies(doc, fileNames);
                AddSharedParameters(doc);
                //AddInstances(doc);

                var titleScheduleName = "Title schedule";
                var expandableScheduleName = "Expandable schedule";

                var sheets = new List<ViewSheet>();

                #region TitleBlock
                Element titleBlock = GetTitleBlock(doc, MAIN_TITLEBLOCK_NAME);

                ViewSchedule titleSchedule = ViewScheduleUtilities.GetSchedule(doc, titleScheduleName);
                if (titleSchedule == null)
                {
                    titleSchedule = ViewScheduleUtilities.CreateSchedule(doc, titleScheduleName);
                }

                ViewSheet titleSheet = ViewSheetUtilities.CreateTitleSheet(doc, titleSchedule, titleScheduleName, titleBlock.Id);
                sheets.Add(titleSheet);
                #endregion

                //Set active view
                uidoc.ActiveView = titleSheet;

                #region ExpandingBlock
                Element expandableTitleBlock = GetTitleBlock(doc, EXPANDABLE_TITLEBLOCK_NAME);

                List<ViewSchedule> expandingSchedules = ViewScheduleUtilities.GetSchedules(doc, expandableScheduleName);
                if (expandingSchedules == null)
                {
                    expandingSchedules = ViewScheduleUtilities.CreateSchedules(doc, expandableScheduleName, PAGE_ITEM_COUNT);
                }

                //AddImageToSchedules(uidoc, expandingSchedules);
                
                var expandingSheets = new List<ViewSheet>();
                for (int i = 0; i < expandingSchedules.Count; i++)
                {
                    ViewSheet expandingSheet = ViewSheetUtilities.CreateExpandableSheet(doc, expandingSchedules[i], expandableScheduleName, i + 1, expandableTitleBlock.Id);
                    //sheets.Add(expandingSheet);
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

        private void AddWallTypeImages(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Wall));
            collector.OfCategory(BuiltInCategory.OST_Walls);

            var walls = collector.Cast<Wall>().ToList();
            var wallTypeIds = new List<ElementId>();
            foreach (var wall in walls)
            {
                if (!wallTypeIds.Contains(wall.WallType.Id))
                {
                    wallTypeIds.Add(wall.WallType.Id);
                }
            }

            using (var tran = new Transaction(doc, "Set wall type images"))
            {
                tran.Start();

                foreach (var wallTypeId in wallTypeIds)
                {
                    var wallType = doc.GetElement(wallTypeId) as WallType;
                    var imgParam = wallType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_IMAGE);
                    if (imgParam.AsElementId() == ElementId.InvalidElementId)
                    {
                        var imageFilePath = Path.Combine(IMAGES_FOLDER, "rebar-test-image.png");
                        var image = ImageType.Create(doc, imageFilePath);

                        imgParam.Set(image.Id);
                    }
                }

                tran.Commit();
            }
        }

        private void AddImageToSchedules(UIDocument uidoc, List<ViewSchedule> schedules)
        {
            var doc = uidoc.Document;
            //FilteredElementCollector annotationCollector = new FilteredElementCollector(doc);
            //annotationCollector.OfCategory(BuiltInCategory.OST_GenericAnnotation);
            //annotationCollector.OfClass(typeof(FamilySymbol));

            foreach (ViewSchedule schedule in schedules)
            {
                //elements
                var elementCollector = new FilteredElementCollector(doc, schedule.Id);
                elementCollector.OfCategory(BuiltInCategory.OST_Walls);
                var elements = elementCollector.ToList();

                CreateImages(uidoc, elements);

                using (var tran = new Transaction(doc, "Schedule cleanup"))
                {
                    tran.Start();

                    schedule.ImageRowHeight = 0.01;

                    tran.Commit();
                }
            }
        }

        private void AddWallImages(UIDocument uidoc)
        {
            var doc = uidoc.Document;

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Wall));
            collector.OfCategory(BuiltInCategory.OST_Walls);

            var wallsWithoutImage = new Dictionary<string, List<Element>>();

            var walls = collector.Cast<Wall>().ToList();
            foreach (var wall in walls)
            {
                Parameter imgParam = wall.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_IMAGE);
                if (imgParam.AsElementId() != ElementId.InvalidElementId)
                {
                    ElementId imgId = imgParam.AsElementId();
                    var img = doc.GetElement(imgId) as ImageType;

                    var baseFileName = Path.GetFileNameWithoutExtension(img.Path);

                    if (!wallsWithoutImage.ContainsKey(baseFileName))
                    {
                        var list = new List<Element>();
                        list.Add(wall);

                        wallsWithoutImage.Add(baseFileName, list);
                    }
                    else
                    {
                        wallsWithoutImage[baseFileName].Add(wall);
                    }
                }
            }

            foreach (var item in wallsWithoutImage)
            {
                //var path = Path.Combine(ROOT_FOLDER, $"{ANNOTATION_TEMPLATE_NAME}.rfa");
                var path = Path.Combine(ROOT_FOLDER, $"{item.Key}.rfa");

                //debug
                if (item.Key == "rebar-test-image")
                {

                    path = Path.Combine(ROOT_FOLDER, $"{ANNOTATION_TEMPLATE_NAME}.rfa");
                }

                CreateImages(uidoc, item.Key, item.Value, path);
            }
        }

        private void CreateImages(UIDocument uidoc, List<Element> elements)
        {
            UIApplication app = uidoc.Application;
            Document doc = uidoc.Document;
            var rng = new Random();

            var path = Path.Combine(ROOT_FOLDER, $"{ANNOTATION_TEMPLATE_NAME}.rfa");

            UIDocument familyUiDoc = app.OpenAndActivateDocument(path);
            Document familyDoc = familyUiDoc.Document;

            if (familyDoc.IsFamilyDocument)
            {
                FilteredElementCollector textCollector = new FilteredElementCollector(familyDoc);
                textCollector.OfClass(typeof(TextNote));
                textCollector.OfCategory(BuiltInCategory.OST_TextNotes);

                List<TextNote> textElements = textCollector.Cast<TextNote>().ToList();
                var A = textElements.Find(tn => tn.Text.Contains("A"));
                var B = textElements.Find(tn => tn.Text.Contains("B"));
                var C = textElements.Find(tn => tn.Text.Contains("C"));
                var D = textElements.Find(tn => tn.Text.Contains("D"));

                foreach (var element in elements)
                {
                    double randA = rng.Next(10, 100);
                    double randB = rng.Next(10, 100);
                    double randC = rng.Next(10, 100);
                    double randD = rng.Next(10, 100);

                    using (var tran = new Transaction(familyDoc, "Modify parameters"))
                    {
                        tran.Start();

                        A.Text = randA.ToString();
                        B.Text = randB.ToString();
                        C.Text = randC.ToString();
                        D.Text = randD.ToString();

                        tran.Commit();
                    }

                    var baseFileName = "testWallImg";
                    var imageFileName = $"{baseFileName}_{randA}_{randB}_{randC}_{randD}.jpg";
                    var imageFilePath = Path.Combine(IMAGES_FOLDER, imageFileName);

                    if (!File.Exists(imageFilePath))
                    {
                        ExportImage(familyDoc, imageFilePath);
                    }

                    using (var tran = new Transaction(doc, "Assign instance image"))
                    {
                        tran.Start();

                        Parameter imgParam = element.get_Parameter(BuiltInParameter.ALL_MODEL_IMAGE);

                        var image = ImageType.Create(doc, imageFilePath);
                        imgParam.Set(image.Id);

                        tran.Commit();
                    }
                }

                app.OpenAndActivateDocument(doc.PathName);
                familyDoc.Close(false);
            }
        }


        private void CreateImages(UIDocument uidoc, string baseFileName, List<Element> elements, string annotationDocPath)
        {
            UIApplication app = uidoc.Application;
            Document doc = uidoc.Document;
            var rng = new Random();

            UIDocument familyUiDoc = app.OpenAndActivateDocument(annotationDocPath);
            Document familyDoc = familyUiDoc.Document;

            if (familyDoc.IsFamilyDocument)
            {
                FilteredElementCollector textCollector = new FilteredElementCollector(familyDoc);
                textCollector.OfClass(typeof(TextNote));
                textCollector.OfCategory(BuiltInCategory.OST_TextNotes);

                List<TextNote> textElements = textCollector.Cast<TextNote>().ToList();
                var A = textElements.Find(tn => tn.Text.Contains("A"));
                var B = textElements.Find(tn => tn.Text.Contains("B"));
                var C = textElements.Find(tn => tn.Text.Contains("C"));
                var D = textElements.Find(tn => tn.Text.Contains("D"));

                foreach (var element in elements)
                {
                    double randA = rng.Next(10, 100);
                    double randB = rng.Next(10, 100);
                    double randC = rng.Next(10, 100);
                    double randD = rng.Next(10, 100);

                    using (var tran = new Transaction(familyDoc, "Modify parameters"))
                    {
                        tran.Start();

                        A.Text = randA.ToString();
                        B.Text = randB.ToString();
                        C.Text = randC.ToString();
                        D.Text = randD.ToString();

                        tran.Commit();
                    }

                    //var baseFileName = "testWallImg";
                    var imageFileName = $"{baseFileName}_{randA}_{randB}_{randC}_{randD}.jpg";
                    var imageFilePath = Path.Combine(IMAGES_FOLDER, imageFileName);

                    if (!File.Exists(imageFilePath))
                    {
                        ExportImage(familyDoc, imageFilePath);
                    }

                    using (var tran = new Transaction(doc, "Assign instance image"))
                    {
                        tran.Start();

                        Parameter imgParam = element.get_Parameter(BuiltInParameter.ALL_MODEL_IMAGE);

                        var image = ImageType.Create(doc, imageFilePath);
                        imgParam.Set(image.Id);

                        tran.Commit();
                    }
                }

                app.OpenAndActivateDocument(doc.PathName);
                familyDoc.Close(false);
            }
        }


        private static void ExportImage(Document familyDoc, string filePath)
        {
            var exportOptions = new ImageExportOptions
            {
                ZoomType = ZoomFitType.Zoom,
                Zoom = 100,
                FilePath = filePath,
                FitDirection = FitDirectionType.Horizontal,
                HLRandWFViewsFileType = ImageFileType.JPEGLossless,
                ImageResolution = ImageResolution.DPI_600
            };

            familyDoc.ExportImage(exportOptions);
        }

        private void HandleImgInstance(FamilyInstance instance)
        {
            var rng = new Random();
            var myNumberParamList = instance.GetParameters("My Number");
            if (myNumberParamList.Count == 1)
            {
                Parameter myNumberParam = myNumberParamList.First();
                if (myNumberParam.StorageType == StorageType.Double)
                {
                    myNumberParam.Set(rng.NextDouble());
                }
            }
        }

        private void App_DocumentChanged(object sender, Autodesk.Revit.DB.Events.DocumentChangedEventArgs e)
        {
            List<ElementId> addedElementIds = e.GetAddedElementIds().ToList();
            List<ElementId> deletedElementIds = e.GetDeletedElementIds().ToList();
        }

        private void AddSharedParameters(Document doc)
        {
            var app = doc.Application;

            //open shared parameter file
            app.SharedParametersFilename = @"C:\Users\AleksandrasZagorskas\Desktop\revit_csd\schedules\HelloSharedParameterWorld.txt";
            DefinitionFile currentDefinitionFile = app.OpenSharedParameterFile();

            //get groups
            DefinitionGroups defGroups = currentDefinitionFile.Groups;
            DefinitionGroup defGroup = defGroups.get_Item("My Shared Parameter Group");
            Definitions allDefinitions = defGroup.Definitions;

            if (defGroup != null)
            {
                //get categories
                Category sheetCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Sheets);
                CategorySet sheetCategories = app.Create.NewCategorySet();
                sheetCategories.Insert(sheetCategory);

                Category wallCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Walls);
                CategorySet wallCategories = app.Create.NewCategorySet();
                wallCategories.Insert(wallCategory);

                InstanceBinding sheetBinding = app.Create.NewInstanceBinding(sheetCategories);
                InstanceBinding wallBinding = app.Create.NewInstanceBinding(wallCategories);

                BindingMap bindingMap = doc.ParameterBindings;

                using (Transaction tran = new Transaction(doc, "Adding shared parameters"))
                {
                    tran.Start();

                    foreach (Definition item in allDefinitions)
                    {
                        if (item.Name == "My Text" || item.Name == "My Number")
                        {
                            bindingMap.Insert(item, sheetBinding, BuiltInParameterGroup.PG_DATA);
                        }
                        else if (item.Name == "My Page")
                        {
                            bindingMap.Insert(item, wallBinding, BuiltInParameterGroup.PG_DATA);
                        }
                    }

                    tran.Commit();
                }
            }

            //Initial values new parameters
            HandleWalls(doc, PAGE_ITEM_COUNT);
        }

        private void LoadRequiredFamilies(Document doc, string[] fileNameArray)
        {

            using (var tran = new Transaction(doc, "Loading title blocks"))
            {
                tran.Start();

                foreach (var fileName in fileNameArray)
                {
                    doc.LoadFamily(fileName);
                }

                tran.Commit();
            }
        }

        private void HandleWalls(Document doc, int pageItemCount)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Wall));
            collector.OfCategory(BuiltInCategory.OST_Walls);

            List<Element> walls = collector.ToList();

            var pageNumber = 1;

            for (int i = 0; i < walls.Count; i++)
            {
                var myPageParamList = walls[i].GetParameters("My Page");

                if (myPageParamList.Count == 1)
                {
                    using (var tran = new Transaction(doc, "Handle wall"))
                    {
                        tran.Start();

                        Parameter myPageParam = myPageParamList.First();
                        if (myPageParam.StorageType == StorageType.Integer)
                        {
                            myPageParam.Set(pageNumber);
                        }

                        tran.Commit();
                    }
                }

                if ((i + 1) % pageItemCount == 0)
                {
                    pageNumber++;
                }
            }
        }


        /// <summary>
        /// Get title block for ViewSheet
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="titleBlockName"></param>
        /// <returns></returns>
        private Element GetTitleBlock(Document doc, string titleBlockName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);

            List<Element> titleBlocks = collector.ToList();
            return titleBlocks.Where(tb => tb.Name.Equals(titleBlockName)).First();
        }

    }
}
