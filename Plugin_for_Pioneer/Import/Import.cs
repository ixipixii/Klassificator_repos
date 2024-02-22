using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks; // Библиотека для работы параллельными задачами
using System.Collections.Concurrent; //Библиотека, содержащая потокобезопасные коллекции
using System.Security.Policy;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Menu;
using System.Collections;
using System.Diagnostics;

namespace Plugin_for_Pioneer
{
    [Transaction(TransactionMode.Manual)]
    internal class Import : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            ImportParameter(uiapp, uidoc, doc);
            return Result.Succeeded;
        }
        public Result ImportParameter(UIApplication uiapp, UIDocument uidoc, Document doc)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            try //Добавление и изменение параметра
            {
                //Чтение файла
                //Лист-строк из Excel
                List<Data> listDataExcel = new List<Data>();

                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                openFileDialog.Filter = "All files(*.*)|*.*";

                string filePath = string.Empty;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                    filePath = openFileDialog.FileName;

                if (string.IsNullOrEmpty(filePath))
                    return Result.Cancelled;

                //Кэширование файла
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(filePath);
                    ISheet sheet = workbook.GetSheetAt(index: 0);

                    int rowIndex = 0;

                    while (sheet.GetRow(rowIndex) != null)
                    {
                        if (sheet.GetRow(rowIndex).GetCell(0) == null ||
                            sheet.GetRow(rowIndex).GetCell(1) == null ||
                            sheet.GetRow(rowIndex).GetCell(2) == null)
                        {
                            rowIndex++;
                            continue;
                        }

                        //Создаём объект-строку из Excel и добавляем в лист объектов-строк 
                        Data excelData = new Data();
                        excelData.pnr_1 = sheet.GetRow(rowIndex).GetCell(0).StringCellValue;
                        excelData.pnr_2 = sheet.GetRow(rowIndex).GetCell(1).StringCellValue;
                        excelData.guid = sheet.GetRow(rowIndex).GetCell(2).StringCellValue;
                        listDataExcel.Add(excelData);
                        rowIndex++;
                    }
                }

                //Чтение элементов модели
                var selectedRef = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Выберите элементы");
                var elementList = new List<Element>();

                List<Data> listDataElement = new List<Data>();

                //Кэширование модели
                foreach (var seleсtedElement in selectedRef)
                {
                    Element element = doc.GetElement(seleсtedElement);

                    if ((BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_IOSModelGroups ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Sections ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Views ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Levels ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_VolumeOfInterest ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_BoundaryConditions ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Grids ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_GridChains ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Dimensions ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_WeakDims ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_ShaftOpeningHiddenLines ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_gbXML_OpeningAir ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_GbXML_Opening ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_MassOpening ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_ArcWallRectOpening ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_DormerOpeningIncomplete ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_SWallRectOpening ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_ShaftOpening ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_StructuralFramingOpening ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_ColumnOpening ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_FloorOpening ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_RoofOpening ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_IOSOpening ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_DoorsOpeningProjection ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_DoorsOpeningCut ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_WindowsOpeningProjection ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_WindowsOpeningCut ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_RoomSeparationLines ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Lines ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_CurtainGridsWall ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_CeilingOpening
                        )
                        continue;

                    elementList.Add(element);

                    //Закидываем данные элемента в объект Data
                    Data elementData = new Data();
                    if (element.LookupParameter("PNR_Код по классификатору") != null
                        || element.LookupParameter("PNR_Описание по классификатору") != null)
                    {
                        elementData.pnr_1 = element.LookupParameter("PNR_Код по классификатору").AsString();
                        elementData.pnr_2 = element.LookupParameter("PNR_Описание по классификатору").AsString();
                    }
                    elementData.guid = element.UniqueId;
                    elementData.element = element;
                    listDataElement.Add(elementData);
                }

/*                //Без кэширования и многопогточности
                Transaction tr = new Transaction(doc, "Заносим параметр");
                tr.Start();
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(filePath);
                    ISheet sheet = workbook.GetSheetAt(index: 0);

                    int rowIndex = 0;

                    while (sheet.GetRow(rowIndex) != null)
                    {
                        if (sheet.GetRow(rowIndex).GetCell(0) == null ||
                            sheet.GetRow(rowIndex).GetCell(1) == null ||
                            sheet.GetRow(rowIndex).GetCell(2) == null)
                        {
                            rowIndex++;
                            continue;
                        }

                        //Создаём объект-строку из Excel и добавляем в лист объектов-строк 
                        Data excelData = new Data();
                        excelData.pnr_1 = sheet.GetRow(rowIndex).GetCell(0).StringCellValue;
                        excelData.pnr_2 = sheet.GetRow(rowIndex).GetCell(1).StringCellValue;
                        excelData.guid = sheet.GetRow(rowIndex).GetCell(2).StringCellValue;
                        listDataExcel.Add(excelData);

                        var element = elementList.FirstOrDefault(r => r.UniqueId == sheet.GetRow(rowIndex).GetCell(2).StringCellValue);

                        if (element != null)
                        {
                            element.LookupParameter("test").Set(sheet.GetRow(rowIndex).GetCell(2).StringCellValue);
                        }
                        rowIndex++;
                    }
                }
                tr.Commit();*/


                ConcurrentBag<Data> listDataElementTrue = new ConcurrentBag<Data>(); ///Многопоточная коллекция

                var listDataElementSorted = listDataElement.OrderBy(pr => pr.guid).ToList();
                var listDataExcelSorted = listDataExcel.OrderBy(pr => pr.guid).ToList();

                //Многопоточность
                Parallel.For(0, listDataExcelSorted.Count, X =>
                {
                    var element = listDataElementSorted.FirstOrDefault(x => x.guid == listDataExcelSorted[X].guid);

                    if (element != null)
                    {
                        //Если парамтеры не равны, добавляем в массив изменяемых элементов
                        if (element.pnr_1 != listDataExcelSorted[X].pnr_1
                        || element.pnr_2 != listDataExcelSorted[X].pnr_2)
                        {
                            if (listDataExcelSorted[X].pnr_1 != "" & listDataExcelSorted[X].pnr_2 != "")
                                listDataElementTrue.Add(element);
                        }
                    }
                });

                var listDataElementTrueSorted = listDataElementTrue.OrderBy(pr => pr.guid).ToList();

                //Заносим значения в параметр из листа нужных элементов 
                if (listDataElementTrue.Count > 0)
                {
                    Transaction transaction = new Transaction(doc, "Заносим значения в параметр");
                    transaction.Start();
                    foreach (var desieredElementTrue in listDataElementTrue)
                    {
                        if (desieredElementTrue == null)
                            continue;
                        var excelElement = listDataExcel.FirstOrDefault(r => r.guid == desieredElementTrue.guid);
                        if (desieredElementTrue != null)
                        {
                            if (desieredElementTrue.element.GroupId.IntegerValue != -1)
                                continue;

                            //Заносим значение в параметр
                            try
                            {
                                desieredElementTrue.element.LookupParameter("test").Set(excelElement.pnr_1);
                                //desieredElementTrue.element.LookupParameter("PNR_Описание по классификатору").Set(excelElement.pnr_2);
                            }
                            catch (System.NullReferenceException)
                            {
                                TaskDialog.Show("1", $"{desieredElementTrue.element.Category.Name}");
                                TaskDialog.Show("", $"{desieredElementTrue.element.Id}");
                            }

                        }
                    }
                    transaction.Commit();
                }
            }

            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }
            
            stopwatch.Stop();
            TaskDialog.Show("Время работы программы", "Время работы программы с кэшированием и многопоточностью: " + stopwatch.ElapsedMilliseconds.ToString() + " мс");

            return Result.Succeeded;
        }
    }

}



