using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitTest.Schedules.Utilities
{
    public class ImageUtilities
    {
        private static readonly string ANNOTATIONS_FOLDER = @"C:\Users\AleksandrasZagorskas\Desktop\revit_csd\schedules";
        private static readonly string IMAGES_FOLDER = @"C:\Users\AleksandrasZagorskas\Desktop\revit_csd\schedules\images";

        /// <summary>
        /// Add default type images
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="typeElements"></param>
        /// <param name="defaultImagePath"></param>
        public static void AddTypeImages(Document doc, List<Element> typeElements, string defaultImagePath)
        {
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

        /// <summary>
        /// Generate instance images based on type images
        /// </summary>
        /// <param name="uidoc"></param>
        /// <param name="instanceElements"></param>
        public static void AddInstanceImages(UIDocument uidoc, List<Element> instanceElements)
        {
            var doc = uidoc.Document;

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
                var path = Path.Combine(ANNOTATIONS_FOLDER, $"{item.Key}.rfa");

                CreateImages(uidoc, item.Key, item.Value, path);
            }
        }

        /// <summary>
        /// Create parametrized
        /// </summary>
        /// <param name="uidoc"></param>
        /// <param name="baseFileName"></param>
        /// <param name="elements"></param>
        /// <param name="annotationDocPath"></param>
        private static void CreateImages(UIDocument uidoc, string baseFileName, List<Element> elements, string annotationDocPath)
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

                    string paramAs = ParameterUtilities.FormatParam(doc, paramA);
                    string paramBs = ParameterUtilities.FormatParam(doc, paramB);
                    string paramCs = ParameterUtilities.FormatParam(doc, paramC);
                    string paramDs = ParameterUtilities.FormatParam(doc, paramD);
                    string paramEs = ParameterUtilities.FormatParam(doc, paramE);
                    string paramFs = ParameterUtilities.FormatParam(doc, paramF);
                    string paramGs = ParameterUtilities.FormatParam(doc, paramG);
                    string paramHs = ParameterUtilities.FormatParam(doc, paramH);
                    string paramJs = ParameterUtilities.FormatParam(doc, paramJ);
                    string paramKs = ParameterUtilities.FormatParam(doc, paramK);
                    string paramOs = ParameterUtilities.FormatParam(doc, paramO);
                    string paramRs = ParameterUtilities.FormatParam(doc, paramR);

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

        /// <summary>
        /// Cleanup images if generation clutters workspace
        /// </summary>
        /// <param name="doc"></param>
        public static void DeleteUnusedImages(Document doc)
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


    }
}
