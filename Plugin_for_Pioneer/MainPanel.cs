using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Plugin_for_Pioneer.Properties;
using System.IO;
using System.Windows.Media.Imaging;
using System.Security.AccessControl;

namespace Plugin_for_Pioneer
{
    internal class MainPanel : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public string GetExeDirectory()
        {
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            path = Path.GetDirectoryName(path);
            return path;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            //Путь к картинке
            string absPath = GetExeDirectory();
            //string pathImg = Path.Combine(absPath, @"Resources\klassificator.png");

            Autodesk.Revit.UI.RibbonPanel ribbonPanel = null;
            foreach (Autodesk.Revit.UI.RibbonPanel ribbonPanel2 in application.GetRibbonPanels(Tab.AddIns))
            {
                if (ribbonPanel2.Name.Equals("Pioneer"))
                {
                    ribbonPanel = ribbonPanel2;
                    break;
                }
            }

            application.GetRibbonPanels(Tab.AddIns);
            ribbonPanel = application.CreateRibbonPanel("Pioneer");
            
            var PulldownButtonData = new PulldownButtonData("Классификация", "Классификация");
            var group = ribbonPanel.AddItem(PulldownButtonData) as PulldownButton;

            var PushButtonData_1 = new PushButtonData("Импорт", "Импорт", Assembly.GetExecutingAssembly().Location, "Plugin_for_Pioneer.Import");
            var NewButton_1 = group.AddPushButton(PushButtonData_1) as PushButton;

            var PushButtonData_2 = new PushButtonData("Экспорт", "Экспорт", Assembly.GetExecutingAssembly().Location, "Plugin_for_Pioneer.Export");
            var NewButton_2 = group.AddPushButton(PushButtonData_2) as PushButton;

            //PulldownButtonData.Image = new BitmapImage(new Uri("/Resources/button_icon.png", UriKind.RelativeOrAbsolute));

            /*Uri uri = new Uri(pathImg, UriKind.Absolute);
            BitmapImage bitmap = new BitmapImage(uri);
            PulldownButtonData.LargeImage = bitmap;*/

            return Result.Succeeded;
        }
    }
}
