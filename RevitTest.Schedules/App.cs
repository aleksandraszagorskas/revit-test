using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using RevitTest.Schedules.ExternalCommands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace RevitTest.Schedules
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class App : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            CreateAppUI(application);
            return Result.Succeeded;
        }

        private void CreateAppUI(UIControlledApplication application)
        {
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("RevitTest");
            ribbonPanel?.AddItem(new PushButtonData("TestSchedule","Test Schedule", Assembly.GetExecutingAssembly().Location, typeof(TestScheduleCommand).FullName));
        }
    }
}
