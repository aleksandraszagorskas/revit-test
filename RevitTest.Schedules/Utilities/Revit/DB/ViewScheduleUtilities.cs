using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitTest.Schedules.Utilities.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using Application = Autodesk.Revit.ApplicationServices.Application;
using SystemPoint = System.Drawing.Point;

namespace RevitTest.Schedules.Utilities.Revit.DB
{
    public class ViewScheduleUtilities
    {
        private static readonly int PAGE_ITEM_COUNT = 12;

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
        public static ViewSchedule CreateSchedule(Document doc, BuiltInCategory scheduleCategory, string scheduleName)
        {
            //CleanupSchedule(doc, scheduleName);
            ViewSchedule schedule = null;

            //create new schedule
            using (Transaction transaction = new Transaction(doc, "Create schedule"))
            {
                transaction.Start();

                schedule = ViewSchedule.CreateSchedule(doc, new ElementId(scheduleCategory));
                schedule.Name = scheduleName;

                transaction.Commit();
            }

            return schedule;
        }

        public static void HandleTestSchedule(Document doc, ViewSchedule schedule)
        {
            using (var tran = new Transaction(doc, "Handle schedule"))
            {
                tran.Start();

                var schedulableFields = schedule.Definition.GetSchedulableFields();
                var schedulableInstances = schedulableFields.Where(f => f.FieldType == ScheduleFieldType.Instance).ToList();
                var allowedFields = new string[] { "Type", "Family", "Volume", "Area" };

                foreach (SchedulableField schField in schedulableInstances)
                {
                    var name = schField.GetName(doc);
                    if (allowedFields.Contains(name))
                    {
                        var field = schedule.Definition.AddField(schField);

                        if (field.GetName() == "Volume" || field.GetName() == "Area")
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

        /// <summary>
        /// Create schedules
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="scheduleBaseName"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static List<ViewSchedule> CreateSchedules(Document doc, BuiltInCategory category, string scheduleBaseName, int pageCount)
        {
            var schedules = new List<ViewSchedule>();

            //create new schedule
            using (Transaction transaction = new Transaction(doc, "Create schedule"))
            {
                transaction.Start();

                for (int i = 0; i < pageCount; i++)
                {
                    var schedule = ViewSchedule.CreateSchedule(doc, new ElementId(category));
                    schedule.Name = $"{scheduleBaseName} Page {i + 1}";
                    schedules.Add(schedule);
                }


                transaction.Commit();
            }

            return schedules;
        }

        /// <summary>
        /// Create schedules
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="scheduleBaseName"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static List<ViewSchedule> CreateExpandableSchedulePages(Document doc, ViewSchedule schedule, int pageCount)
        {
            var schedules = new List<ViewSchedule>();

            //create new schedule
            using (Transaction tran = new Transaction(doc, "Create schedule pages"))
            {
                tran.Start();

                for (int i = 0; i < pageCount; i++)
                {
                    if (schedule.CanViewBeDuplicated(ViewDuplicateOption.Duplicate))
                    {
                        ElementId scheduleId = schedule.Duplicate(ViewDuplicateOption.Duplicate);
                        ViewSchedule dupSchedule = doc.GetElement(scheduleId) as ViewSchedule;
                        schedules.Add(dupSchedule);
                    }


                }

                tran.Commit();
            }

            return schedules;
        }

        public static void HandleExpandableSchedule(Document doc, ViewSchedule schedule, int pageNumber = 0)
        {
            using (var tran = new Transaction(doc, "Handle schedule"))
            {
                tran.Start();

                var allSchedulableFields = schedule.Definition.GetSchedulableFields();
                var schedulableInstances = allSchedulableFields.Where(f => f.FieldType == ScheduleFieldType.Instance).ToList();

                var allowedFields = new string[] { "Rebar Number", "Quantity", "Bar Diameter", "Bar Length", "Total Bar Length", "Image", "Comments", "My Page", "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "O", "R" };
                var schedulableFields = new SchedulableField[allowedFields.Count()];

                //filter schedulable fields 
                for (int i = 0; i < schedulableInstances.Count; i++)
                {
                    SchedulableField field = schedulableInstances[i];
                    var name = field.GetName(doc);
                    if (allowedFields.Contains(name))
                    {
                        int index = Array.IndexOf<string>(allowedFields, name);

                        if (schedulableFields[index] == null)
                        {
                            schedulableFields[index] = field;
                        }

                        //schedulableFields.Add(field);
                    }
                }
                //foreach (SchedulableField field in schedulableInstances)
                //{
                //    var name = field.GetName(doc);
                //    if (allowedFields.Contains(name) && )
                //    {
                //        schedulableFields.Add(field);
                //    }
                //}

                //sort


                //var scheduleFields = new List<SchedulableField>();

                //config sche fields
                foreach (SchedulableField schField in schedulableFields)
                {
                    //"Rebar Number", "Quantity", "Bar Diameter", "Bar Length", "Total Bar Length", "Image", "Comments", "My Page", "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "O", "R"
                    ScheduleField field = schedule.Definition.AddField(schField);
                    var name = field.GetName();

                    if (name == "Rebar Number")
                    {
                        field.ColumnHeading = "POS";
                        // instance grouping
                        var sortGroupField = new ScheduleSortGroupField(field.FieldId, ScheduleSortOrder.Ascending);
                        schedule.Definition.AddSortGroupField(sortGroupField);
                    }
                    else if (name == "Quantity")
                    {
                        field.ColumnHeading = "QTY";
                    }
                    else if (name == "Bar Diameter")
                    {
                        field.ColumnHeading = "D \r\n mm";
                    }
                    else if (name == "Bar Length")
                    {
                        field.ColumnHeading = "REBAR";
                    }
                    else if (name == "Total Bar Length")
                    {
                        field.ColumnHeading = "TOTAL";
                    }
                    else if (name == "Image")
                    {
                        field.ColumnHeading = "SHAPE \r\n (OUTSIDE DIMENSIONS IN cm)";
                    }
                    else if (name == "Comments")
                    {
                        field.ColumnHeading = "COMMENTS";
                    }
                    else if (name == "My Page")
                    {
                        field.ColumnHeading = "PAGE";
                        //field.IsHidden = true;

                        //if (pageNumber != 0)
                        //{
                        //    try
                        //    {
                        //        var filter = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, pageNumber);
                        //        schedule.Definition.AddFilter(filter);
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        //Filter not added
                        //        TaskDialog.Show("Notice", "Page filter not applied.");
                        //    }
                        //}
                    }
                    //else if (
                    //    name == "A"
                    //    || name == "B"
                    //    || name == "C"
                    //    || name == "D"
                    //    || name == "E"
                    //    || name == "F"
                    //    || name == "G"
                    //    || name == "H"
                    //    || name == "J"
                    //    || name == "K"
                    //    || name == "O"
                    //    || name == "R")
                    //{
                    //    field.IsHidden = true;
                    //}

                }

                doc.Regenerate();

                schedule.Definition.IsItemized = false;

                tran.Commit();
            }
        }



        private static void HandleExpandableScheduleInstances(Document doc, ViewSchedule schedule, int pageNumber)
        {
            int startIndex = (pageNumber - 1) * PAGE_ITEM_COUNT;
            List<Element> scheduleElements = GetScheduleElements(doc, schedule);
            List<int> rebarNumbers = new List<int>();
            List<Element> instanceElements2 = new List<Element>();
            foreach (Element element in scheduleElements)
            {
                var paramList = element.GetParameters("Rebar Number");

                if (paramList.Count == 1)
                {
                    Parameter param = paramList.First();
                    int number = param.AsInteger();
                    if (!rebarNumbers.Contains(number))
                    {
                        rebarNumbers.Add(number);
                    }

                }
            }

            List<Element> instanceElements = GetScheduleElements(doc, schedule).GetRange(startIndex, PAGE_ITEM_COUNT);
            foreach (Element element in instanceElements)
            {
                //SetElementPage(doc,element, pageNumber);
                var myPageParamList = element.GetParameters("My Page");

                if (myPageParamList.Count == 1)
                {
                    using (var tran = new Transaction(doc, "Handle element"))
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
            }
        }

        private static void HandleScheduledInstances(Document doc, List<ViewSchedule> schedules)
        {
            foreach (ViewSchedule schedule in schedules)
            {
                List<Element> scheduleElements = GetScheduleElements(doc, schedule);

            }
        }


        public static List<Element> GetScheduleElements(Document doc, ViewSchedule schedule)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc, schedule.Id);

            return collector.ToList();
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

        /// <summary>
        /// Load template schedule from donor file
        /// </summary>
        /// <param name="uidoc"></param>
        /// <param name="donorDocPath"></param>
        /// <param name="scheduleName"></param>
        /// <returns></returns>
        public static ViewSchedule LoadTemplateSchedule(UIDocument uidoc, string donorDocPath, string scheduleName)
        {
            Document toDoc = uidoc.Document;
            UIApplication uiapp = uidoc.Application;
            Application app = uiapp.Application;
            Document fromDoc = app.OpenDocumentFile(donorDocPath);

            var templateScheduleId = new ElementId(2413);
            var elementsToCopy = new List<ElementId>();
            elementsToCopy.Add(templateScheduleId);

            ViewSchedule schedule = null;
            //ElementId copiedScheduleId = null;
            using (Transaction tran = new Transaction(toDoc, "Duplicate template schedule"))
            {
                tran.Start();

                // Set options for copy-paste to hide the duplicate types dialog
                CopyPasteOptions options = new CopyPasteOptions();
                options.SetDuplicateTypeNamesHandler(new HideAndAcceptDuplicateTypeNamesHandler());

                // Copy the input elements.
                var copiedIds = ElementTransformUtils.CopyElements(fromDoc, elementsToCopy, toDoc, Transform.Identity, options).ToList();
                if (copiedIds.Count == 1)
                {
                    ElementId copiedScheduleId = copiedIds.First();

                    schedule = toDoc.GetElement(copiedScheduleId) as ViewSchedule;
                    schedule.Name = scheduleName;
                }

                fromDoc.Close();

                // Set failure handler to hide duplicate types warnings which may be posted.
                FailureHandlingOptions failureOptions = tran.GetFailureHandlingOptions();
                failureOptions.SetFailuresPreprocessor(new HidePasteDuplicateTypesPreprocessor());
                tran.Commit(failureOptions);
            }

            return schedule;
        }


        /// <summary>
        /// A handler to accept duplicate types names created by the copy/paste operation.
        /// </summary>
        class HideAndAcceptDuplicateTypeNamesHandler : IDuplicateTypeNamesHandler
        {
            #region IDuplicateTypeNamesHandler Members

            /// <summary>
            /// Implementation of the IDuplicateTypeNameHandler
            /// </summary>
            /// <param name="args"></param>
            /// <returns></returns>
            public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
            {
                // Always use duplicate destination types when asked
                return DuplicateTypeAction.UseDestinationTypes;
            }

            #endregion
        }

        /// <summary>
        /// A failure preprocessor to hide the warning about duplicate types being pasted.
        /// </summary>
        class HidePasteDuplicateTypesPreprocessor : IFailuresPreprocessor
        {
            #region IFailuresPreprocessor Members

            /// <summary>
            /// Implementation of the IFailuresPreprocessor.
            /// </summary>
            /// <param name="failuresAccessor"></param>
            /// <returns></returns>
            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                foreach (FailureMessageAccessor failure in failuresAccessor.GetFailureMessages())
                {
                    // Delete any "Can't paste duplicate types.  Only non duplicate types will be pasted." warnings
                    if (failure.GetFailureDefinitionId() == BuiltInFailures.CopyPasteFailures.CannotCopyDuplicates)
                    {
                        failuresAccessor.DeleteWarning(failure);
                    }
                }

                // Handle any other errors interactively
                return FailureProcessingResult.Continue;
            }

            #endregion
        }

        public static void SplitSchedule(UIDocument uidoc, ScheduleSheetInstance scheduleInstance)
        {
            Document doc = uidoc.Document;
            var ownerView = doc.GetElement(scheduleInstance.OwnerViewId) as ViewSheet;
            uidoc.ActiveView = ownerView;

            BoundingBoxXYZ schBB = scheduleInstance.get_BoundingBox(doc.ActiveView);
            if (schBB != null)
            {
                //using (var tran = new Transaction(doc, "Split schedule"))
                //{

                //}
                Line schRightLine = Line.CreateBound(new XYZ(schBB.Max.X, schBB.Min.Y, 0), schBB.Max);
                XYZ lineMidpoint = schRightLine.Evaluate(0.5, true);

                UIView uiView = UIViewUtilities.GetActiveUiView(uidoc);
                uiView.ZoomToFit();

                //Rectangle rect = uiView.GetWindowRectangle();

                SystemPoint startCursorPos = System.Windows.Forms.Cursor.Position;


                SystemPoint scheduleSplitPos = Win32Utilities.Revit2Screen(uidoc, uiView, lineMidpoint);

                System.Windows.Forms.Cursor.Position = scheduleSplitPos;

                Win32Utilities.PressMouseLeftButton();

                //reset mouse pos
                System.Windows.Forms.Cursor.Position = startCursorPos;
            }
            else
            {
                TaskDialog.Show("Error", "Failed to read schedule bounding box.");
            }
        }

    }
}
