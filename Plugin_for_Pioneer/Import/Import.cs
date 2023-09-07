using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks; // Библиотека для работы параллельными задачами
using System.Collections.Concurrent; //Библиотека, содержащая потокобезопасные коллекции

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

            if (message == "Import3d")
            {
                Import3D(uiapp, uidoc, doc);
            }

            return Result.Succeeded;
        }

        public Result Import3D(UIApplication uiapp, UIDocument uidoc, Document doc)
        {
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
                var categoryList = new List<BuiltInCategory>();

                //List<List<ElementId>> groupElements = new List<List<ElementId>>();

                List<Data> listDataElement = new List<Data>();

                Transaction t = new Transaction(doc, "UnGroup");
                t.Start();
                foreach (var seleсtedElement in selectedRef)
                {
                    Element element = doc.GetElement(seleсtedElement);

                    /*//Разгруппировка
                    if ((BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_IOSModelGroups)
                    {
                        Group group = (Group)element;
                        groupElements.Add(group.UngroupMembers().ToList());

                        foreach (var groupUp in groupElements)
                        {
                            foreach (var groupDown in groupUp)
                            {
                                if ((BuiltInCategory)groupDown.IntegerValue == BuiltInCategory.OST_IOSModelGroups)
                                {
                                    Group group1 = (Group)element;
                                    groupElements.Add(group1.UngroupMembers().ToList());
                                }
                            }
                        }
                        continue;
                    }*/

                    if ((BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_IOSModelGroups ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Sections ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Views ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Levels)
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

                    //Добавление категорий
                    Category category = element.Category;
                    BuiltInCategory enumCategory = (BuiltInCategory)category.Id.IntegerValue;
                    categoryList.Add(enumCategory);
                }
                t.Commit();

                var categorySet = new CategorySet();
                foreach (var category in categoryList.Distinct())
                {
                    categorySet.Insert(Category.GetCategory(doc, category));
                }

                using (Transaction ts = new Transaction(doc, "Add parameter"))
                {
                    ts.Start();
                    foreach (var element in elementList)
                    {
                        if (element.LookupParameter("PNR_Код по классификатору") == null
                            || element.LookupParameter("PNR_Описание по классификатору") == null)
                        {
                            CreateShared createShared_pnr_1 = new CreateShared();
                            createShared_pnr_1.CreateSharedParameter(uiapp.Application,
                                                               doc,
                                                               "PNR_Код по классификатору",
                                                               categorySet,
                                                               BuiltInParameterGroup.PG_IDENTITY_DATA,
                                                               true);
                            CreateShared createShared_pnr_2 = new CreateShared();
                            createShared_pnr_2.CreateSharedParameter(uiapp.Application,
                                                               doc,
                                                               "PNR_Описание по классификатору",
                                                               categorySet,
                                                               BuiltInParameterGroup.PG_IDENTITY_DATA,
                                                               true);
                        }
                    }
                    ts.Commit();
                }

                DateTime end; //Далее проверим как будет работать наша вычисления в многопоточном режиме                
                DateTime start = DateTime.Now; //Засекаем время

                //Многопоточность
                /*                Parallel.ForEach(listDataExcel, excel =>
                                {
                                    var desieredElement = listDataElement.FirstOrDefault(r => r.guid == excel.guid);
                                    if (desieredElement != null)
                                    {
                                        //Если парамтеры равны, ничего не делаем
                *//*                        if (desieredElement.pnr_1 == excel.pnr_1 || desieredElement.pnr_2 == excel.pnr_2)
                                            return;*//*

                                        //Если парамтеры не равны, включаем флаг изменений
                                        if (desieredElement.pnr_1 != excel.pnr_1 || desieredElement.pnr_1 != excel.pnr_2)
                                            desieredElement.flag = true;
                                    }
                                });*/

                //Без многопоточности
                foreach (var excel in listDataExcel)
                {
                    var desieredElement = listDataElement.FirstOrDefault(r => r.guid == excel.guid);
                    if (desieredElement != null)
                    {
                        //Если парамтеры равны, ничего не делаем
                        if (desieredElement.pnr_1 == excel.pnr_1)
                            continue;
                        else if (desieredElement.pnr_2 == excel.pnr_2)
                            continue;

                        //Если парамтеры не равны, включаем флаг изменений
                        if (desieredElement.pnr_1 != excel.pnr_1 || desieredElement.pnr_1 != excel.pnr_2)
                            desieredElement.flag = true;
                    }
                }

                //Заносим значения в параметр у элементов с флагом
                Transaction transaction = new Transaction(doc, "Заносим значения в параметр");
                transaction.Start();
                foreach (var desieredElement in listDataElement)
                {
                    var excelElement = listDataExcel.FirstOrDefault(r => r.guid == desieredElement.guid);
                    if (desieredElement.flag)
                    {
                        desieredElement.element.LookupParameter("PNR_Код по классификатору").Set(excelElement.pnr_1);
                        desieredElement.element.LookupParameter("PNR_Описание по классификатору").Set(excelElement.pnr_2);
                    }
                }
                transaction.Commit();

                end = DateTime.Now; // Записываем текущее время
                TimeSpan taim = (end - start);
                TaskDialog.Show("Время выполнения", $"{taim.TotalMilliseconds} миллисекунд");

                /*                Transaction tr = new Transaction(doc, "NewGroup");
                                tr.Start();
                                foreach (var group in groupElements)
                                {
                                    Group groupNew = doc.Create.NewGroup(group);
                                }
                                tr.Commit();*/
            }

            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }

            return Result.Succeeded;
        }

        public Result ImportCategory(List<BuiltInCategory> SelectedCategoryList, UIApplication uiapp, UIDocument uidoc, Document doc)
        {
            try
            {
                var selectedRef = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Выберите элементы");
                var elementList = new List<Element>();

                foreach (var seleсtedElement in selectedRef)
                {
                    Element element = doc.GetElement(seleсtedElement);
                    elementList.Add(element);
                }

                var categorySet = new CategorySet();
                foreach (var category in SelectedCategoryList.Distinct())
                {
                    categorySet.Insert(Category.GetCategory(doc, category));
                }

                using (Transaction ts = new Transaction(doc, "Add parameter"))
                {
                    ts.Start();
                    foreach (var element in elementList)
                    {
                        if (element.LookupParameter("PNR_Код по классификатору") == null || element.LookupParameter("PNR_Описание по классификатору") == null)
                        {
                            CreateShared createShared_pnr_1 = new CreateShared();
                            if (createShared_pnr_1.CreateSharedParameter(uiapp.Application,
                                                               doc,
                                                               "PNR_Код по классификатору",
                                                               categorySet,
                                                               BuiltInParameterGroup.PG_IDENTITY_DATA,
                                                               true) == 1)
                            {
                                return Result.Succeeded;
                            }

                            CreateShared createShared_pnr_2 = new CreateShared();
                            if (createShared_pnr_2.CreateSharedParameter(uiapp.Application,
                                                               doc,
                                                               "PNR_Описание по классификатору",
                                                               categorySet,
                                                               BuiltInParameterGroup.PG_IDENTITY_DATA,
                                                               true) == 1)
                            {
                                return Result.Succeeded;
                            }
                        }
                    }
                    ts.Commit();
                }

                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                openFileDialog.Filter = "All files(*.*)|*.*";

                string filePath = string.Empty;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                    filePath = openFileDialog.FileName;

                if (string.IsNullOrEmpty(filePath))
                    return Result.Cancelled;

                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(filePath);
                    ISheet sheet = workbook.GetSheetAt(index: 0);

                    int rowIndex = 0;

                    using (Transaction ts = new Transaction(doc, "Set parameters"))
                    {
                        ts.Start();
                        while (sheet.GetRow(rowIndex) != null)
                        {
                            if (sheet.GetRow(rowIndex).GetCell(0) == null ||
                                sheet.GetRow(rowIndex).GetCell(1) == null ||
                                sheet.GetRow(rowIndex).GetCell(2) == null)
                            {
                                rowIndex++;
                                continue;
                            }

                            string pnr_1 = sheet.GetRow(rowIndex).GetCell(0).StringCellValue;
                            string pnr_2 = sheet.GetRow(rowIndex).GetCell(1).StringCellValue;
                            string guid = sheet.GetRow(rowIndex).GetCell(2).StringCellValue;

                            var element = elementList.FirstOrDefault(r => r.UniqueId == guid);

                            if (element == null)
                            {
                                rowIndex++;
                                continue;
                            }

                            if (element.LookupParameter("PNR_Код по классификатору") != null)
                                element.LookupParameter("PNR_Код по классификатору").Set(pnr_1);
                            if (element.LookupParameter("PNR_Описание по классификатору") != null)
                                element.LookupParameter("PNR_Описание по классификатору").Set(pnr_2);

                            rowIndex++;
                        }
                        ts.Commit();
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }

            return Result.Succeeded;
        }
    }
}
