using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Plugin_for_Pioneer
{
    [Transaction(TransactionMode.Manual)]
    internal class Export : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            //Чтение параметра
            try
            {
                var selectedRef = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Выберите элементы");
                var elementList = new List<Element>();

                foreach (var seleсtedElement in selectedRef)
                {
                    Element element = doc.GetElement(seleсtedElement);
                    if ((BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_IOSModelGroups ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Sections ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Views ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Levels)
                        continue;
                    elementList.Add(element);
                }

                //Запись параметров в файл
                /*string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string excelPath = Path.Combine(desktopPath, "pioneer_plugin.xlsx");*/

                var saveDialog = new SaveFileDialog
                {
                    OverwritePrompt = true,
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    Filter = "All files (*.*)|*.*",
                    FileName = "klassificator_info.csv",
                    DefaultExt = ".csv"
                };

                string selectedFilePath = string.Empty;
                
                if(saveDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = saveDialog.FileName;
                }

                if(selectedFilePath == string.Empty)
                {
                    return Result.Succeeded;
                }

                using (FileStream stream = new FileStream(selectedFilePath, FileMode.Create, FileAccess.Write))
                {
                    IWorkbook workbook = new XSSFWorkbook();
                    ISheet sheet = workbook.CreateSheet("Лист 1");

                    int rowIndex = 0;
                    foreach(var element in elementList)
                    {
                        if (element.LookupParameter("PNR_Код по классификатору") == null || element.LookupParameter("PNR_Описание по классификатору") == null)
                            continue;
                        sheet.SetCellValue(rowIndex, columnIndex: 0, element.LookupParameter("PNR_Код по классификатору").AsString());
                        sheet.SetCellValue(rowIndex, columnIndex: 1, element.LookupParameter("PNR_Описание по классификатору").AsString());
                        sheet.SetCellValue(rowIndex, columnIndex: 2, element.UniqueId);
                        rowIndex++;
                    }

                    workbook.Write(stream);
                    workbook.Close();
                }

                System.Diagnostics.Process.Start(selectedFilePath);

            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }

            return Result.Succeeded;
        }
    }
}
