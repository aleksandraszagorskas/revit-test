using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RevitTest.Schedules.Utilities.Revit.DB
{
    public class ViewSheetUtilities
    {
        /// <summary>
        /// Create title sheet
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="schedule"></param>
        /// <param name="scheduleName"></param>
        /// <returns></returns>
        public static ViewSheet CreateTitleSheet(Document doc, ViewSchedule schedule, string scheduleName, ElementId titleBlockTypeId)
        {
            ViewSheet sheet = null;
            using (Transaction transaction = new Transaction(doc, "Create sheet"))
            {
                transaction.Start();

                sheet = ViewSheet.Create(doc, titleBlockTypeId);
                sheet.Name = scheduleName;

                FamilyInstance titleblockInstance = new FilteredElementCollector(doc, sheet.Id).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.Symbol.Id == titleBlockTypeId).First();

                //test
                var bbx = titleblockInstance.get_BoundingBox(sheet);
                XYZ min = bbx.Min;
                XYZ max = bbx.Max;

                Line botLine = Line.CreateBound(min, new XYZ(max.X, min.Y, 0));
                Line leftLine = Line.CreateBound(min, new XYZ(min.X, max.Y, 0));

                XYZ schCoord1 = new XYZ(botLine.Evaluate(0.05, true).X, leftLine.Evaluate(0.75, true).Y, 0);
                XYZ schCoord2 = new XYZ(botLine.Evaluate(0.52, true).X, leftLine.Evaluate(0.75, true).Y, 0);
                XYZ schCoord3 = new XYZ(botLine.Evaluate(0.05, true).X, leftLine.Evaluate(0.38, true).Y, 0);
                XYZ schCoord4 = new XYZ(botLine.Evaluate(0.52, true).X, leftLine.Evaluate(0.38, true).Y, 0);

                var scheduleInstance1 = ScheduleSheetInstance.Create(doc, sheet.Id, schedule.Id, schCoord1);
                var scheduleInstance2 = ScheduleSheetInstance.Create(doc, sheet.Id, schedule.Id, schCoord2);
                var scheduleInstance3 = ScheduleSheetInstance.Create(doc, sheet.Id, schedule.Id, schCoord3);
                var scheduleInstance4 = ScheduleSheetInstance.Create(doc, sheet.Id, schedule.Id, schCoord4);

                transaction.Commit();
            }

            using (var tran = new Transaction(doc, "Handle sheet"))
            {
                tran.Start();

                var myTextParamList = sheet.GetParameters("My Text");
                if (myTextParamList.Count == 1)
                {
                    Parameter myTextParam = myTextParamList.First();
                    if (myTextParam.StorageType == StorageType.String)
                    {
                        myTextParam.Set("Test from code");
                    }
                }

                var myNumberParamList = sheet.GetParameters("My Number");
                if (myNumberParamList.Count == 1)
                {
                    Parameter myNumberParam = myNumberParamList.First();
                    if (myNumberParam.StorageType == StorageType.Double)
                    {
                        myNumberParam.Set(258.0);
                    }
                }

                tran.Commit();
            }

            return sheet;
        }

        /// <summary>
        /// Create expandable sheet
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="schedule"></param>
        /// <param name="scheduleName"></param>
        /// <returns></returns>
        public static ViewSheet CreateExpandableSheet(Document doc, ViewSchedule schedule, string scheduleName, int pageNumber, ElementId titleBlockTypeId)
        {
            ViewSheet sheet = null;
            using (Transaction transaction = new Transaction(doc, "Create sheet"))
            {
                transaction.Start();

                sheet = ViewSheet.Create(doc, titleBlockTypeId);
                sheet.Name = $"{scheduleName} Page {pageNumber}";

                FamilyInstance titleblockInstance = new FilteredElementCollector(doc, sheet.Id).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.Symbol.Id == titleBlockTypeId).First();

                var bbx = titleblockInstance.get_BoundingBox(sheet);
                XYZ min = bbx.Min;
                XYZ max = bbx.Max;

                Line botLine = Line.CreateBound(min, new XYZ(max.X, min.Y, 0));
                Line leftLine = Line.CreateBound(min, new XYZ(min.X, max.Y, 0));

                XYZ schCoord = new XYZ(botLine.Evaluate(0.25, true).X, leftLine.Evaluate(0.75, true).Y, 0);

                var scheduleInstance = ScheduleSheetInstance.Create(doc, sheet.Id, schedule.Id, schCoord);

                transaction.Commit();
            }

            using (var tran = new Transaction(doc, "Handle sheet"))
            {
                tran.Start();

                var myTextParamList = sheet.GetParameters("My Text");
                if (myTextParamList.Count == 1)
                {
                    Parameter myTextParam = myTextParamList.First();
                    if (myTextParam.StorageType == StorageType.String)
                    {
                        myTextParam.Set($"Page {pageNumber}");
                    }
                }

                var rng = new Random();
                var myNumberParamList = sheet.GetParameters("My Number");
                if (myNumberParamList.Count == 1)
                {
                    Parameter myNumberParam = myNumberParamList.First();
                    if (myNumberParam.StorageType == StorageType.Double)
                    {
                        myNumberParam.Set(rng.NextDouble());
                    }
                }

                tran.Commit();
            }

            return sheet;
        }

    }
}
