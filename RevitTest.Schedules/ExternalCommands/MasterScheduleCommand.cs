﻿using Autodesk.Revit.Attributes;
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
    class MasterScheduleCommand : IExternalCommand
    {
        private Misc.Settings settings = Misc.Settings.Instance;

        //List<ViewSchedule> expandingSchedules = null;
        //List<ViewSheet> sheets = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                Units units = doc.GetUnits();
                Application app = doc.Application;
                app.DocumentChanged += App_DocumentChanged;

                var familiesToLoad = new string[] {
                    settings.MAIN_TITLEBLOCK_FILE_NAME,
                    settings.EXPANDABLE_TITLEBLOCK_FILE_NAME,
                    settings.DEFAULT_ANNOTATION_FILE_NAME,

                };

                //debug
                //ImageUtilities.DeleteUnusedImages(doc);

                FamilyUtilities.LoadRequiredFamilies(doc, familiesToLoad);
                ParameterUtilities.AddSharedParameters(doc, settings.SHARED_PARAMETER_FILE_NAME, settings.SHARED_PARAMETER_GROUP_NAME);

                var sheets = new List<ViewSheet>();

                ViewSheet titleSheet = CreateMainSheet(doc);
                sheets.Add(titleSheet);

                //Set active view
                //uidoc.ActiveView = titleSheet;

                //ViewSchedule expandableSchedule = FamilyUtilities.LoadTemplateSchedule(uidoc, SCHEDULE_TEMPLATE_FILE_NAME);
                ViewSchedule expandableSchedule = CreateExpandableSchedule(uidoc);
                HandleExpandableScheduleItems(uidoc, expandableSchedule, settings.GENERATE_IMAGES, out int pageCount);

                List<ViewSheet> expandingSheets = CreateExpandingSheetPages(uidoc, expandableSchedule, pageCount);
                sheets.AddRange(expandingSheets);

                PDFUtilities.ExportSheets(doc, sheets);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private ViewSchedule CreateExpandableSchedule(Document doc)
        {
            ViewSchedule expandableSchedule = ViewScheduleUtilities.GetSchedule(doc, settings.EXPANDABLE_LIST_NAME);
            if (expandableSchedule == null)
            {
                expandableSchedule = ViewScheduleUtilities.CreateSchedule(doc, BuiltInCategory.OST_Rebar, settings.EXPANDABLE_LIST_NAME);
                ViewScheduleUtilities.HandleExpandableSchedule(doc, expandableSchedule);

            }

            return expandableSchedule;
        }
        private ViewSchedule CreateExpandableSchedule(UIDocument uidoc)
        {
            Document doc = uidoc.Document;


            ViewSchedule expandableSchedule = ViewScheduleUtilities.GetSchedule(doc, settings.EXPANDABLE_LIST_NAME);
            if (expandableSchedule == null)
            {
                expandableSchedule = ViewScheduleUtilities.LoadTemplateSchedule(uidoc, settings.SCHEDULE_TEMPLATE_FILE_NAME, settings.EXPANDABLE_LIST_NAME);
                //expandableSchedule = ViewScheduleUtilities.CreateSchedule(doc, BuiltInCategory.OST_Rebar, EXPANDABLE_LIST_NAME);
                //ViewScheduleUtilities.HandleExpandableSchedule(doc, expandableSchedule);

            }

            return expandableSchedule;
        }
        private void HandleExpandableScheduleItems(UIDocument uidoc, ViewSchedule expandableSchedule, bool generateImages, out int pageNumber)
        {
            var doc = uidoc.Document;

            //get data
            var scheduleElements = ViewScheduleUtilities.GetScheduleElements(doc, expandableSchedule);

            var rebarTypeIds = new List<ElementId>();
            var rebarTypes = new List<Element>();
            var rebarElementIds = new List<ElementId>();
            var rebarElements = new List<Element>();
            foreach (Element scheduleElement in scheduleElements)
            {
                if (scheduleElement is Rebar || scheduleElement is RebarInSystem)
                {
                    rebarElementIds.Add(scheduleElement.Id);
                    rebarElements.Add(scheduleElement);

                    ElementId typeId = scheduleElement.GetTypeId();
                    if (!rebarTypeIds.Contains(typeId))
                    {
                        var rebarBarType = doc.GetElement(typeId);
                        rebarTypes.Add(rebarBarType);
                        rebarTypeIds.Add(typeId);
                    }
                }

            }

            if (generateImages)
            {
                //add images
                ImageUtilities.AddTypeImages(doc, rebarTypes, settings.DEFAULT_ANNOTATION_IMAGE_FILE_NAME);
                ImageUtilities.AddInstanceImages(uidoc, rebarElements);
            }

            #region celltext
            //var tableData = expandableSchedule.GetTableData();
            //TableSectionData tableBodySection = tableData.GetSectionData(SectionType.Body);
            //int numOfRows = tableBodySection.NumberOfRows;
            //int numOfCols = tableBodySection.NumberOfColumns;
            //var cellData = new string[numOfRows, numOfCols];
            ////expandableSchedule.
            //for (int i = 0; i < numOfRows; i++)
            //{
            //    for (int j = 0; j < numOfCols; j++)
            //    {
            //        cellData[i, j] = tableBodySection.GetCellText(i, j);
            //    }
            //}

            //var testParamId = tableBodySection.GetCellParamId(4, 5);
            //var testParam = doc.GetElement(testParamId);

            //var exportOptions = new ViewScheduleExportOptions();
            //expandableSchedule.Export(EXCEL_EXPORT_FOLDER, "schedule.csv", exportOptions);
            #endregion

            #region handle grouping
            //List<ScheduleSortGroupField> sortGroupFields = expandableSchedule.Definition.GetSortGroupFields().ToList();
            //foreach (var field in sortGroupFields)
            //{
            //    ScheduleField groupField = expandableSchedule.Definition.GetField(field.FieldId);
            //    string groupFieldName = groupField.GetName();
            //    ScheduleSortOrder groupFieldSortOrder = field.SortOrder;

            //    StorageType groupValueStorageType = StorageType.None;
            //    var groupValues = new List<object>();
            //    foreach (Element element in rebarElements)
            //    {
            //        var paramList = element.GetParameters(groupFieldName);

            //        if (paramList.Count == 1)
            //        {
            //            Parameter param = paramList.First();
            //            groupValueStorageType = param.StorageType;
            //            switch (groupValueStorageType)
            //            {
            //                case StorageType.Integer:
            //                    groupValues.Add(param.AsInteger());
            //                    break;
            //                case StorageType.Double:
            //                    groupValues.Add(param.AsDouble());
            //                    break;
            //                case StorageType.String:
            //                    groupValues.Add(param.AsString());
            //                    break;
            //                case StorageType.None:
            //                case StorageType.ElementId:
            //                default:
            //                    TaskDialog.Show("Grouping not supported", "Please group manualy");
            //                    break;
            //            }
            //        }
            //    }

            //    switch (groupValueStorageType)
            //    {
            //        case StorageType.Integer:
            //            HandleGroupValues<int>(doc, rebarElementIds, groupFieldName, groupFieldSortOrder, groupValues.Cast<int>().ToList(), groupValueStorageType);
            //            break;
            //        case StorageType.Double:
            //            HandleGroupValues<double>(doc, rebarElementIds, groupFieldName, groupFieldSortOrder, groupValues.Cast<double>().ToList(), groupValueStorageType);
            //            break;
            //        case StorageType.String:
            //            HandleGroupValues<string>(doc, rebarElementIds, groupFieldName, groupFieldSortOrder, groupValues.Cast<string>().ToList(), groupValueStorageType);
            //            break;
            //        case StorageType.None:
            //        case StorageType.ElementId:
            //        default:
            //            break;
            //    }
            //}
            #endregion

            //add pages
            pageNumber = 1;
            //int pageCount = (rebarElements.Count() + PAGE_ITEM_COUNT - 1) / PAGE_ITEM_COUNT;
            using (var tran = new Transaction(doc, "Add pages"))
            {
                tran.Start();

                for (int i = 0; i < rebarElements.Count; i++)
                {
                    var rebarElement = rebarElements[i];

                    var paramList = rebarElement.GetParameters("My Page");

                    if (paramList.Count == 1)
                    {
                        Parameter param = paramList.First();
                        if (param.StorageType == StorageType.Integer)
                        {
                            param.Set(-1);
                        }
                    }

                    if ((i + 1) % settings.PAGE_ITEM_COUNT == 0)
                    {
                        //pageNumber++;
                    }
                }

                tran.Commit();
            }
        }

        #region grouping
        //private static void HandleGroupValues<T>(Document doc, List<ElementId> rebarElementIds, string groupFieldName, ScheduleSortOrder groupFieldSortOrder, List<T> groupValues, StorageType storageType)
        //{
        //    if (groupValues.Any())
        //    {
        //        List<T> uniqueValues = groupValues.Cast<T>().Distinct().ToList();

        //        int pageCount = uniqueValues.Count / PAGE_ITEM_COUNT;
        //        for (int i = 0; i < pageCount; i++)
        //        {
        //            int startIndex = i * PAGE_ITEM_COUNT;
        //            List<T> workValues = uniqueValues.GetRange(startIndex, PAGE_ITEM_COUNT);

        //            List<ElementId> sortedRebarElementIds = null;
        //            switch (groupFieldSortOrder)
        //            {
        //                case ScheduleSortOrder.Ascending:
        //                    sortedRebarElementIds = rebarElementIds.OrderBy(r => r).ToList();
        //                    break;
        //                case ScheduleSortOrder.Descending:
        //                    sortedRebarElementIds = rebarElementIds.OrderByDescending(r => r).ToList();
        //                    break;
        //                default:
        //                    break;
        //            }

        //            var workElements = new List<Element>();
        //            foreach (ElementId elementId in sortedRebarElementIds)
        //            {
        //                Element element = doc.GetElement(elementId);

        //                var paramList = element.GetParameters(groupFieldName);

        //                if (paramList.Count == 1)
        //                {
        //                    Parameter param = paramList.First();

        //                    T paramVal = default;
        //                    switch (storageType)
        //                    {
        //                        case StorageType.Integer:
        //                            paramVal = param.AsInteger();
        //                            break;
        //                        case StorageType.Double:

        //                            break;
        //                        case StorageType.String:

        //                            break;
        //                        case StorageType.None:
        //                        case StorageType.ElementId:
        //                        default:
        //                            break;
        //                    }

        //                    if (workValues.Contains(paramVal))
        //                    {
        //                        workElements.Add(element);
        //                    }
        //                }
        //            }

        //            foreach (Element element in workElements)
        //            {
        //                var paramList = element.GetParameters("My Page");

        //                if (paramList.Count == 1)
        //                {
        //                    Parameter param = paramList.First();
        //                    if (param.StorageType == StorageType.Integer)
        //                    {
        //                        param.Set(i + 1);
        //                    }
        //                }
        //            }
        //        }
        //    }
        //}
        #endregion


        private List<ViewSheet> CreateExpandingSheetPages(UIDocument uidoc, ViewSchedule schedule, int pageCount)
        {
            Document doc = uidoc.Document;

            List<ViewSchedule> expandingSchedules = ViewScheduleUtilities.GetSchedules(doc, settings.EXPANDABLE_LIST_NAME);
            if (expandingSchedules == null)
            {
                expandingSchedules = ViewScheduleUtilities.CreateExpandableSchedulePages(doc, schedule, pageCount);
                if (expandingSchedules != null)
                {
                    using (var tran = new Transaction(doc, "Filter schedule pages"))
                    {
                        tran.Start();

                        for (int i = 0; i < expandingSchedules.Count; i++)
                        {
                            var expandingSchedule = expandingSchedules[i];
                            expandingSchedule.Name = $"{schedule.Name} Page {i + 1}";

                            HandleExpandingSchedulePage(doc, expandingSchedule, i + 1);
                        }

                        tran.Commit();
                    }
                }
            }

            var expandingSheets = new List<ViewSheet>();
            for (int i = 0; i < expandingSchedules.Count; i++)
            {
                ViewSheet expandingSheet = ViewSheetUtilities.CreateExpandableSheet(uidoc, settings.EXPANDABLE_TITLEBLOCK_NAME, expandingSchedules[i], settings.EXPANDABLE_LIST_NAME, i + 1);
                expandingSheets.Add(expandingSheet);
            }

            return expandingSheets;
        }

        private static void HandleExpandingSchedulePage(Document doc, ViewSchedule expandingSchedule, int pageNumber)
        {
            for (int i = 0; i < expandingSchedule.Definition.GetFieldCount(); i++)
            {
                ScheduleField field = expandingSchedule.Definition.GetField(i);
                string fieldName = field.GetName();
                if (fieldName == "My Page")
                {
                    field.IsHidden = true;

                    //filter
                    //var filter = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, pageNumber);
                    //expandingSchedule.Definition.AddFilter(filter);
                }
                else if (
                    fieldName == "A"
                    || fieldName == "B"
                    || fieldName == "C"
                    || fieldName == "D"
                    || fieldName == "E"
                    || fieldName == "F"
                    || fieldName == "G"
                    || fieldName == "H"
                    || fieldName == "J"
                    || fieldName == "K"
                    || fieldName == "O"
                    || fieldName == "R")
                {
                    field.IsHidden = true;
                }

            }
        }

        private ViewSheet CreateMainSheet(Document doc)
        {
            ViewSchedule titleSchedule = ViewScheduleUtilities.GetSchedule(doc, settings.MAIN_LIST_NAME);
            if (titleSchedule == null)
            {
                titleSchedule = ViewScheduleUtilities.CreateSchedule(doc, BuiltInCategory.OST_Walls, settings.MAIN_LIST_NAME);
                if (titleSchedule != null)
                {
                    ViewScheduleUtilities.HandleTestSchedule(doc, titleSchedule);
                }
            }

            ViewSheet titleSheet = ViewSheetUtilities.CreateTitleSheet(doc, settings.MAIN_TITLEBLOCK_NAME, titleSchedule, settings.MAIN_LIST_NAME);

            return titleSheet;
        }

        private void App_DocumentChanged(object sender, Autodesk.Revit.DB.Events.DocumentChangedEventArgs e)
        {
            List<ElementId> addedElementIds = e.GetAddedElementIds().ToList();
            List<ElementId> deletedElementIds = e.GetDeletedElementIds().ToList();
        }
    }
}
