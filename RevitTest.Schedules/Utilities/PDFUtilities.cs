using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitTest.Schedules.Utilities
{
    public class PDFUtilities
    {
        /// <summary>
        /// Export sheets to PDF
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="sheet"></param>
        /// <param name="fileName"></param>
        public static void ExportSheets(Document doc, List<ViewSheet> sheets)
        {
            //print
            var printManager = doc.PrintManager;
            printManager.SelectNewPrintDriver("Microsoft Print To PDF");
            printManager.PrintRange = PrintRange.Select;
            printManager.CombinedFile = true;

            PrintSetup printSetup = printManager.PrintSetup;
            using (Transaction tran = new Transaction(doc, "Handle setup options"))
            {
                tran.Start();

                PrintParameters parameters = printSetup.InSession.PrintParameters;
                //PrintParameters parameters = printSetup.CurrentPrintSetting.PrintParameters;
                foreach (PaperSize paperSize in printManager.PaperSizes)
                {
                    if (paperSize.Name.Equals("A4"))
                    {
                        parameters.PaperSize = paperSize;
                        parameters.PaperPlacement = PaperPlacementType.Center;
                        parameters.PageOrientation = PageOrientationType.Portrait;
                        parameters.ZoomType = ZoomType.Zoom;
                        parameters.Zoom = 100;
                        parameters.HideCropBoundaries = true;
                        break;
                    }
                }

                printSetup.CurrentPrintSetting = printSetup.InSession;
                //printSetup.SaveAs("!temp");

                tran.Commit();
            }

            var viewSet = new ViewSet();
            foreach (var sheet in sheets)
            {
                viewSet.Insert(sheet);
            }

            ViewSheetSetting viewSheetSetting = printManager.ViewSheetSetting;
            using (Transaction tran = new Transaction(doc, "Set export views"))
            {
                tran.Start();

                viewSheetSetting.InSession.Views = viewSet;
                //viewSheetSetting.CurrentViewSheetSet.Views = viewSet;

                viewSheetSetting.CurrentViewSheetSet = viewSheetSetting.InSession;
                //viewSheetSetting.SaveAs("!temp");

                tran.Commit();
            }

            printManager.SubmitPrint();

            ////cleanup
            //using (Transaction tran = new Transaction(doc, "Print settings Cleanup"))
            //{
            //    tran.Start();

            //    viewSheetSetting.Delete();
            //    printSetup.Delete();

            //    tran.Commit();
            //}


        }

    }
}
