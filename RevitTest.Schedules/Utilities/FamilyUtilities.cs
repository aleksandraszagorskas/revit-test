using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;


namespace RevitTest.Schedules.Utilities
{
    class FamilyUtilities
    {
        public static void LoadRequiredFamilies(Document doc, string[] fileNameArray)
        {
            using (var tran = new Transaction(doc, "Loading families"))
            {
                tran.Start();

                foreach (var fileName in fileNameArray)
                {
                    doc.LoadFamily(fileName);
                }

                tran.Commit();
            }
        }

    }
}
