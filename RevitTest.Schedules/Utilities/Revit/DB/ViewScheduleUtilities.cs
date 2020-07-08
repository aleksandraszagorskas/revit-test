using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RevitTest.Schedules.Utilities.Revit.DB
{
    public class ViewScheduleUtilities
    {
        /// <summary>
        /// Get existing schedules
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="scheduleName"></param>
        /// <returns></returns>
        public static List<ViewSchedule> GetSchedules(Document doc, string scheduleName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewSchedule));
            collector.OfCategory(BuiltInCategory.OST_Schedules);

            List<Element> scheduleElements = collector.ToList();

            List<ViewSchedule> schedules = scheduleElements.Where(s => s.Name.Contains($"{scheduleName} Page")).Cast<ViewSchedule>().ToList();

            if (schedules.Count > 0)
            {
                return schedules;
            }

            return null;
        }

        /// <summary>
        /// Get existing schedule
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="scheduleName"></param>
        /// <returns></returns>
        public static ViewSchedule GetSchedule(Document doc, string scheduleName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewSchedule));
            collector.OfCategory(BuiltInCategory.OST_Schedules);

            List<Element> scheduleElements = collector.ToList();

            List<ViewSchedule> schedules = scheduleElements.Where(s => s.Name.Equals(scheduleName)).Cast<ViewSchedule>().ToList();

            if (schedules.Count > 0)
            {
                return schedules.First();
            }

            return null;
        }

        /// <summary>
        /// Create schedule
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="scheduleName"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static ViewSchedule CreateSchedule(Document doc, string scheduleName)
        {
            //CleanupSchedule(doc, scheduleName);
            ViewSchedule schedule = null;

            //create new schedule
            using (Transaction transaction = new Transaction(doc, "Create schedule"))
            {
                transaction.Start();

                schedule = ViewSchedule.CreateSchedule(doc, new ElementId(BuiltInCategory.OST_Walls));
                schedule.Name = scheduleName;

                transaction.Commit();
            }

            if (schedule != null)
            {
                HandleSchedule(doc, schedule, 1);
            }

            return schedule;
        }

        /// <summary>
        /// Create schedules
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="scheduleName"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static List<ViewSchedule> CreateSchedules(Document doc, string scheduleName, int pageItemCount)
        {
            //CleanupSchedule(doc, scheduleName);
            var schedules = new List<ViewSchedule>();

            //create new schedule
            using (Transaction transaction = new Transaction(doc, "Create schedule"))
            {
                transaction.Start();

                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfClass(typeof(Wall));
                collector.OfCategory(BuiltInCategory.OST_Walls);

                //FilteredElementCollector testCollector = new FilteredElementCollector(doc);
                ////testCollector.OfClass(typeof(Element));
                //testCollector.OfCategory(BuiltInCategory.OST_Walls);
                //List<Element> testElements = testCollector.ToList();

                int pageCount = (collector.GetElementCount() + pageItemCount - 1) / pageItemCount;
                for (int i = 0; i < pageCount; i++)
                {
                    var schedule = ViewSchedule.CreateSchedule(doc, new ElementId(BuiltInCategory.OST_Walls));
                    schedule.Name = $"{scheduleName} Page {i + 1}";
                    schedules.Add(schedule);
                }

                //var schedule = ViewSchedule.CreateSchedule(doc, new ElementId(BuiltInCategory.OST_Walls));
                ////schedule.Name = scheduleName;
                ////HandleSchedule(doc, schedule, 1);
                //schedules.Add(schedule);



                //if (schedule.Definition.GetFieldCount() > pageItemCount)
                //{
                //    //schedule.Name = $"{scheduleName} Page 1";

                //    int pageCount = (schedule.Definition.GetFieldCount() + pageItemCount - 1) / pageItemCount;
                //    for (int i = 1; i < pageCount; i++)
                //    {
                //        var nextSchedule = ViewSchedule.CreateSchedule(doc, new ElementId(BuiltInCategory.OST_Walls));
                //        //nextSchedule.Name = $"{scheduleName} Page {i + 1}";
                //        //HandleSchedule(doc, nextSchedule, i + 1);
                //        schedules.Add(nextSchedule);
                //    }
                //}

                transaction.Commit();
            }

            //if (schedules.Count == 1)
            //{
            //    schedules[0].Name = scheduleName;
            //    HandleSchedule(doc, schedules[0], 1);
            //}

            if (schedules != null)
            {
                for (int i = 0; i < schedules.Count; i++)
                {
                    HandleExpandableSchedule(doc, schedules[i], i + 1);
                }
            }

            return schedules;
        }



        private static void HandleSchedule(Document doc, ViewSchedule schedule, int pageNumber)
        {
            using (var tran = new Transaction(doc, "Handle schedule"))
            {
                tran.Start();

                var schedulableFields = schedule.Definition.GetSchedulableFields();
                var schedulableInstances = schedulableFields.Where(f => f.FieldType == ScheduleFieldType.Instance).ToList();
                var allowedFields = new string[] { "Type", "Family", "Volume", "Area", "My Page" };

                foreach (SchedulableField schField in schedulableInstances)
                {
                    var name = schField.GetName(doc);
                    if (allowedFields.Contains(name))
                    {
                        var field = schedule.Definition.AddField(schField);
                        if (field.GetName() == "My Page")
                        {
                            field.IsHidden = true;

                            try
                            {
                                var filter = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, pageNumber);
                                schedule.Definition.AddFilter(filter);
                            }
                            catch (Exception ex)
                            {
                                //Filter not added
                                TaskDialog.Show("Warning", "Filter not applied");
                            }
                        }
                        else if (field.GetName() == "Volume" || field.GetName() == "Area")
                        {
                            field.GridColumnWidth = 0.75 * field.GridColumnWidth;
                        }
                    }
                }

                doc.Regenerate();

                //group headers
                if (schedule.CanGroupHeaders(0, 2, 0, 3))
                {
                    schedule.GroupHeaders(0, 2, 0, 3, "Measurements");
                }

                tran.Commit();
            }
        }

        private static void HandleExpandableSchedule(Document doc, ViewSchedule schedule, int pageNumber)
        {
            using (var tran = new Transaction(doc, "Handle schedule"))
            {
                tran.Start();

                var schedulableFields = schedule.Definition.GetSchedulableFields();
                var schedulableInstances = schedulableFields.Where(f => f.FieldType == ScheduleFieldType.Instance).ToList();

                //debug
                var schedulableInstanceNames = new List<string>();

                var allowedFields = new string[] { "Type", "Family", "Volume", "Area", "Image", "My Page" };

                foreach (SchedulableField schField in schedulableInstances)
                {
                    var name = schField.GetName(doc);
                    //debug
                    schedulableInstanceNames.Add(name);

                    if (allowedFields.Contains(name))
                    {
                        var field = schedule.Definition.AddField(schField);
                        if (field.GetName() == "My Page")
                        {
                            field.IsHidden = true;

                            try
                            {
                                var filter = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, pageNumber);
                                schedule.Definition.AddFilter(filter);
                            }
                            catch (Exception ex)
                            {
                                //Filter not added
                                TaskDialog.Show("Warning", "Filter not applied");
                            }
                        }
                        else if (field.GetName() == "Volume" || field.GetName() == "Area")
                        {
                            field.GridColumnWidth = 1.3 * field.GridColumnWidth;
                        }
                    }
                }

                doc.Regenerate();

                //group headers
                if (schedule.CanGroupHeaders(0, 3, 0, 4))
                {
                    schedule.GroupHeaders(0, 3, 0, 4, "Measurements");
                }

                //schedule.ImageRowHeight = 2.0;//2 * schedule.ImageRowHeight;

                tran.Commit();
            }
        }

        /// <summary>
        /// Delete schedule
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="scheduleName"></param>
        /// <returns></returns>
        private void CleanupSchedule(Document doc, string scheduleName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewSchedule));
            collector.OfCategory(BuiltInCategory.OST_Schedules);

            List<Element> schedules = collector.ToList();

            var scheduleElements = schedules.Where(s => s.Name.Contains(scheduleName)).ToList();

            if (scheduleElements.Count == 1)
            {
                using (var tran = new Transaction(doc, "Cleanup schedule"))
                {
                    tran.Start();

                    doc.Delete(scheduleElements.First().Id);

                    tran.Commit();
                }
            }
        }
    }
}
