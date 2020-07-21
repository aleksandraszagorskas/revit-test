using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitTest.Schedules.Utilities
{
    public class ParameterUtilities
    {
        public static void AddSharedParameters(Document doc, string sharedParametersFilename, string sharedParametersGroupName)
        {
            var app = doc.Application;

            //open shared parameter file
            app.SharedParametersFilename = sharedParametersFilename;//@"C:\Users\AleksandrasZagorskas\Desktop\revit_csd\schedules\HelloSharedParameterWorld.txt";
            DefinitionFile currentDefinitionFile = app.OpenSharedParameterFile();

            //get groups
            DefinitionGroups defGroups = currentDefinitionFile.Groups;
            DefinitionGroup defGroup = defGroups.get_Item(sharedParametersGroupName);
            Definitions allDefinitions = defGroup.Definitions;

            if (defGroup != null)
            {
                //get categories
                Category sheetCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Sheets);
                CategorySet sheetCategories = app.Create.NewCategorySet();
                sheetCategories.Insert(sheetCategory);

                Category rebarCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rebar);
                CategorySet rebarCategories = app.Create.NewCategorySet();
                rebarCategories.Insert(rebarCategory);

                InstanceBinding sheetBinding = app.Create.NewInstanceBinding(sheetCategories);
                InstanceBinding rebarBinding = app.Create.NewInstanceBinding(rebarCategories);

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
                            bindingMap.Insert(item, rebarBinding, BuiltInParameterGroup.PG_DATA);
                        }
                    }

                    tran.Commit();
                }
            }

            //Initial values new parameters
            //HandleWalls(doc, PAGE_ITEM_COUNT);
            //HandleRebars(doc, PAGE_ITEM_COUNT);
        }

        public static string FormatParam(Parameter param)
        {
            if (param != null && param.StorageType == StorageType.Double)
            {
                double val = Math.Round(param.AsDouble(), 2);
                int convVal = (int)(val * 100);
                return convVal.ToString();
            }
            return null;
        }

        public static string FormatParam(Document doc, Parameter param)
        {
            Units units = doc.GetUnits();
            if (param != null && param.StorageType == StorageType.Double)
            {
                double val = param.AsDouble();
                string formatedVal = UnitFormatUtils.Format(units, UnitType.UT_Length, val, false, false);

                //double val = Math.Round(param.AsDouble(), 2);
                //int convVal = (int)(val * 100);
                return formatedVal;
            }
            return null;
        }
    }
}
