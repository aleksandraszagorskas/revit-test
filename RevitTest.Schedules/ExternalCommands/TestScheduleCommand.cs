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
        private readonly string ANNOTATION_TEMPLATE_NAME = "rebar-test-image";
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

                FilteredElementCollector rebarCollector = new FilteredElementCollector(doc);
                rebarCollector.OfClass(typeof(Rebar));
                rebarCollector.OfCategory(BuiltInCategory.OST_Rebar);
                //rebarCollector.

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

                //Type rebarType = typeof(Rebar);
                //BuiltInCategory rebarCategory = BuiltInCategory.OST_Rebar;

                var fileNames = new string[] { Path.Combine(ROOT_FOLDER, $"{MAIN_TITLEBLOCK_NAME}.rfa"), Path.Combine(ROOT_FOLDER, $"{EXPANDABLE_TITLEBLOCK_NAME}.rfa"), Path.Combine(ROOT_FOLDER, $"{ANNOTATION_TEMPLATE_NAME}.rfa") };
                var imageFilePath = Path.Combine(IMAGES_FOLDER, $"{ANNOTATION_TEMPLATE_NAME}.png");


                //debug
                //DeleteUnusedImages(doc);

                AddTypeImages(doc, rebarTypes, imageFilePath);
                AddInstanceImages(uidoc, rebars);

                LoadRequiredFamilies(doc, fileNames);
                ParameterUtilities.AddSharedParameters(doc, SHARED_PARAMETER_FILE_NAME, SHARED_PARAMETER_GROUP_NAME);
                HandleInstances(doc, rebars, PAGE_ITEM_COUNT);

                var titleScheduleName = "Title schedule";
                var expandableScheduleName = "Expandable schedule";

                var sheets = new List<ViewSheet>();

                #region TitleBlock
                Element titleBlock = GetTitleBlock(doc, MAIN_TITLEBLOCK_NAME);

                ViewSchedule titleSchedule = ViewScheduleUtilities.GetSchedule(doc, titleScheduleName);
                if (titleSchedule == null)
                {
                    //titleSchedule = ViewScheduleUtilities.CreateSchedule(doc, titleScheduleName);
                    titleSchedule = ViewScheduleUtilities.CreateSchedule(doc, BuiltInCategory.OST_Walls, titleScheduleName);
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
                    //expandingSchedules = ViewScheduleUtilities.CreateSchedules(doc, expandableScheduleName, PAGE_ITEM_COUNT);
                    expandingSchedules = ViewScheduleUtilities.CreateSchedules(doc, BuiltInCategory.OST_Rebar, expandableScheduleName, pageCount);
                }

                var expandingSheets = new List<ViewSheet>();
                for (int i = 0; i < expandingSchedules.Count; i++)
                {
                    ViewSheet expandingSheet = ViewSheetUtilities.CreateExpandableSheet(doc, expandingSchedules[i], expandableScheduleName, i + 1, expandableTitleBlock.Id);
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

        private void DeleteUnusedImages(Document doc)
        {
            var imageCollector = new FilteredElementCollector(doc);
            imageCollector.OfCategory(BuiltInCategory.OST_RasterImages);
            imageCollector.OfClass(typeof(ImageType));
            List<ImageType> images = imageCollector.Cast<ImageType>().ToList();

            using (var tran = new Transaction(doc, "Project cleanup"))
            {
                tran.Start();

                foreach (var image in images)
                {
                    doc.Delete(image.Id);
                }

                tran.Commit();
            }


        }

        private void AddTypeImages(Document doc, List<Element> typeElements, string defaultImagePath)
        {
            //FilteredElementCollector collector = new FilteredElementCollector(doc);
            //collector.OfClass(type);
            //collector.OfCategory(category);
            //List<Element> instanceElements = collector.ToList();

            //var typeIds = new List<ElementId>();
            //var typeElements = new List<Element>();
            //foreach (var instance in instanceElements)
            //{
            //    ElementId typeId = instance.GetTypeId();
            //    if (!typeIds.Contains(typeId))
            //    {
            //        var typeElement = doc.GetElement(typeId);
            //        typeElements.Add(typeElement);
            //        typeIds.Add(typeId);
            //    }
            //}

            using (var tran = new Transaction(doc, "Set type images"))
            {
                tran.Start();

                foreach (Element typeElement in typeElements)
                {
                    var imageCollector = new FilteredElementCollector(doc);
                    imageCollector.OfCategory(BuiltInCategory.OST_RasterImages);
                    imageCollector.OfClass(typeof(ImageType));
                    List<ImageType> images = imageCollector.Cast<ImageType>().ToList();

                    var imgParam = typeElement.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_IMAGE);
                    if (imgParam.AsElementId() == ElementId.InvalidElementId)
                    {

                        //var image = ImageType.Create(doc, defaultImagePath);
                        ImageType image = null;
                        string fileName = Path.GetFileName(defaultImagePath);
                        List<ImageType> imageFilter = images.Where(i => Path.GetFileName(i.Name).Contains(fileName)).ToList();//?.First();
                        if (imageFilter.Any())
                        {
                            image = imageFilter.First();
                        }
                        else
                        {
                            image = ImageType.Create(doc, defaultImagePath);
                        }

                        imgParam.Set(image.Id);
                    }
                }

                tran.Commit();
            }
        }

        private void AddInstanceImages(UIDocument uidoc, List<Element> instanceElements)
        {
            var doc = uidoc.Document;
            //FilteredElementCollector collector = new FilteredElementCollector(doc);
            //collector.OfClass(type);
            //collector.OfCategory(category);
            //List<Element> instanceElements = collector.ToList();

            var elementsWithoutImage = new Dictionary<string, List<Element>>();
            foreach (var instanceElement in instanceElements)
            {
                ElementId instanceTypeId = instanceElement.GetTypeId();
                Element instanceType = doc.GetElement(instanceTypeId);

                Parameter imgParam = instanceType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_IMAGE);
                if (imgParam.AsElementId() != ElementId.InvalidElementId)
                {
                    ElementId imgId = imgParam.AsElementId();
                    var img = doc.GetElement(imgId) as ImageType;

                    var baseFileName = Path.GetFileNameWithoutExtension(img.Path);

                    if (!elementsWithoutImage.ContainsKey(baseFileName))
                    {
                        var list = new List<Element>();
                        list.Add(instanceElement);

                        elementsWithoutImage.Add(baseFileName, list);
                    }
                    else
                    {
                        elementsWithoutImage[baseFileName].Add(instanceElement);
                    }
                }
            }

            foreach (var item in elementsWithoutImage)
            {
                var path = Path.Combine(ROOT_FOLDER, $"{item.Key}.rfa");

                CreateImages(uidoc, item.Key, item.Value, path);
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

                string paramAName = "A";
                string paramBName = "B";
                string paramCName = "C";
                string paramDName = "D";
                string paramEName = "E";
                string paramFName = "F";
                string paramGName = "G";
                string paramHName = "H";
                string paramJName = "J";
                string paramKName = "K";
                string paramOName = "O";
                string paramRName = "R";

                var A = textElements.Find(tn => tn.Text.Contains(paramAName));
                var B = textElements.Find(tn => tn.Text.Contains(paramBName));
                var C = textElements.Find(tn => tn.Text.Contains(paramCName));
                var D = textElements.Find(tn => tn.Text.Contains(paramDName));
                var E = textElements.Find(tn => tn.Text.Contains(paramEName));
                var F = textElements.Find(tn => tn.Text.Contains(paramFName));
                var G = textElements.Find(tn => tn.Text.Contains(paramGName));
                var H = textElements.Find(tn => tn.Text.Contains(paramHName));
                var J = textElements.Find(tn => tn.Text.Contains(paramJName));
                var K = textElements.Find(tn => tn.Text.Contains(paramKName));
                var O = textElements.Find(tn => tn.Text.Contains(paramOName));
                var R = textElements.Find(tn => tn.Text.Contains(paramRName));

                foreach (var element in elements)
                {
                    Parameter paramA = element.GetParameters(paramAName).Count() == 1 ? element.GetParameters(paramAName).First() : null;
                    Parameter paramB = element.GetParameters(paramBName).Count() == 1 ? element.GetParameters(paramBName).First() : null;
                    Parameter paramC = element.GetParameters(paramCName).Count() == 1 ? element.GetParameters(paramCName).First() : null;
                    Parameter paramD = element.GetParameters(paramDName).Count() == 1 ? element.GetParameters(paramDName).First() : null;
                    Parameter paramE = element.GetParameters(paramEName).Count() == 1 ? element.GetParameters(paramEName).First() : null;
                    Parameter paramF = element.GetParameters(paramFName).Count() == 1 ? element.GetParameters(paramFName).First() : null;
                    Parameter paramG = element.GetParameters(paramGName).Count() == 1 ? element.GetParameters(paramGName).First() : null;
                    Parameter paramH = element.GetParameters(paramHName).Count() == 1 ? element.GetParameters(paramHName).First() : null;
                    Parameter paramJ = element.GetParameters(paramJName).Count() == 1 ? element.GetParameters(paramJName).First() : null;
                    Parameter paramK = element.GetParameters(paramKName).Count() == 1 ? element.GetParameters(paramKName).First() : null;
                    Parameter paramO = element.GetParameters(paramOName).Count() == 1 ? element.GetParameters(paramOName).First() : null;
                    Parameter paramR = element.GetParameters(paramRName).Count() == 1 ? element.GetParameters(paramRName).First() : null;

                    string paramAs = FormatParam(paramA);
                    string paramBs = FormatParam(paramB);
                    string paramCs = FormatParam(paramC);
                    string paramDs = FormatParam(paramD);
                    string paramEs = FormatParam(paramE);
                    string paramFs = FormatParam(paramF);
                    string paramGs = FormatParam(paramG);
                    string paramHs = FormatParam(paramH);
                    string paramJs = FormatParam(paramJ);
                    string paramKs = FormatParam(paramK);
                    string paramOs = FormatParam(paramO);
                    string paramRs = FormatParam(paramR);

                    using (var tran = new Transaction(familyDoc, "Modify parameters"))
                    {
                        tran.Start();

                        if (A != null && paramAs != null)
                        {
                            A.Text = paramAs;
                        }
                        if (B != null && paramBs != null)
                        {
                            B.Text = paramBs;
                        }
                        if (C != null && paramCs != null)
                        {
                            C.Text = paramCs;
                        }
                        if (D != null && paramDs != null)
                        {
                            D.Text = paramDs;
                        }
                        if (E != null && paramEs != null)
                        {
                            E.Text = paramEs;
                        }
                        if (F != null && paramFs != null)
                        {
                            F.Text = paramFs;
                        }
                        if (G != null && paramGs != null)
                        {
                            G.Text = paramGs;
                        }
                        if (H != null && paramHs != null)
                        {
                            H.Text = paramHs;
                        }
                        if (J != null && paramJs != null)
                        {
                            J.Text = paramJs;
                        }
                        if (K != null && paramKs != null)
                        {
                            K.Text = paramKs;
                        }
                        if (O != null && paramOs != null)
                        {
                            O.Text = paramOs;
                        }
                        if (R != null && paramRs != null)
                        {
                            R.Text = paramRs;
                        }

                        tran.Commit();
                    }

                    var imageFileName = $"{baseFileName}_{paramAs}_{paramBs}_{paramCs}_{paramDs}.jpg";
                    var imageFilePath = Path.Combine(IMAGES_FOLDER, imageFileName);

                    if (!File.Exists(imageFilePath))
                    {
                        ExportImage(familyDoc, imageFilePath);
                    }

                    var imageCollector = new FilteredElementCollector(doc);
                    imageCollector.OfCategory(BuiltInCategory.OST_RasterImages);
                    imageCollector.OfClass(typeof(ImageType));
                    List<ImageType> images = imageCollector.Cast<ImageType>().ToList();

                    using (var tran = new Transaction(doc, "Assign instance image"))
                    {
                        tran.Start();

                        Parameter imgParam = element.get_Parameter(BuiltInParameter.ALL_MODEL_IMAGE);

                        ImageType image = null;
                        List<ImageType> imageFilter = images.Where(i => Path.GetFileName(i.Name).Contains(imageFileName)).ToList();//?.First();
                        if (imageFilter.Any())
                        {
                            image = imageFilter.First();
                        }
                        else
                        {
                            image = ImageType.Create(doc, imageFilePath);
                        }

                        imgParam.Set(image.Id);

                        tran.Commit();
                    }
                }

                app.OpenAndActivateDocument(doc.PathName);
                familyDoc.Close(false);
            }
        }

        private string FormatParam(Parameter param)
        {
            if (param != null && param.StorageType == StorageType.Double)
            {
                double val = Math.Round(param.AsDouble(), 2);
                int convVal = (int)(val * 100);
                return convVal.ToString();
            }
            return null;
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

        private void App_DocumentChanged(object sender, Autodesk.Revit.DB.Events.DocumentChangedEventArgs e)
        {
            List<ElementId> addedElementIds = e.GetAddedElementIds().ToList();
            List<ElementId> deletedElementIds = e.GetDeletedElementIds().ToList();
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

        private void HandleInstances(Document doc, List<Element> instanceElements, int pageItemCount)
        {
            //FilteredElementCollector collector = new FilteredElementCollector(doc);
            //collector.OfClass(type);
            //collector.OfCategory(category);
            //List<Element> instanceElements = collector.ToList();

            var pageNumber = 1;
            for (int i = 0; i < instanceElements.Count; i++)
            {
                var myPageParamList = instanceElements[i].GetParameters("My Page");

                if (myPageParamList.Count == 1)
                {
                    using (var tran = new Transaction(doc, "Handle rebar"))
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
