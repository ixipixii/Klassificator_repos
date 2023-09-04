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
using System.Threading.Tasks;
using System.Windows.Forms;

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
                //new
                List<Data> listDataExcel = new List<Data>();
                Data excelData = new Data();

                List<String> listExcelPnr_1 = new List<String>();
                List<String> listExcelPnr_2 = new List<String>();
                List<String> listExcelGuid = new List<String>();
                //new

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

                    //new
                    while (sheet.GetRow(rowIndex) != null)
                    {
                        if (sheet.GetRow(rowIndex).GetCell(0) == null ||
                            sheet.GetRow(rowIndex).GetCell(1) == null ||
                            sheet.GetRow(rowIndex).GetCell(2) == null)
                        {
                            rowIndex++;
                            continue;
                        }

                        excelData.pnr_1 = sheet.GetRow(rowIndex).GetCell(0).StringCellValue;
                        excelData.pnr_2 = sheet.GetRow(rowIndex).GetCell(1).StringCellValue;
                        excelData.guid = sheet.GetRow(rowIndex).GetCell(2).StringCellValue;                        
                        listDataExcel.Add(excelData);

                        listExcelPnr_1.Add(sheet.GetRow(rowIndex).GetCell(0).StringCellValue);
                        listExcelPnr_2.Add(sheet.GetRow(rowIndex).GetCell(1).StringCellValue);
                        listExcelGuid.Add(sheet.GetRow(rowIndex).GetCell(2).StringCellValue);

                    }
                    //new

                    /*                    using (Transaction ts = new Transaction(doc, "Set parameters"))
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
                                                    element.LookupParameter("PNR_Код по классификатору").Set(pnr_1);
                                                    element.LookupParameter("PNR_Описание по классификатору").Set(pnr_2);
                                                    rowIndex++;
                                            }
                                            ts.Commit();
                                        }*/
                }

                //Чтение элементов модели
                var selectedRef = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Выберите элементы");
                var elementList = new List<Element>();
                var categoryList = new List<BuiltInCategory>();

                List<List<ElementId>> groupElements = new List<List<ElementId>>();

                List<Data> listDataElement = new List<Data>();
                Data elementData = new Data();

                List<String> listPnr_1 = new List<String>();
                List<String> listPnr_2 = new List<String>();
                List<String> listGuid = new List<String>();

                Transaction t = new Transaction(doc, "UnGroup");
                t.Start();
                foreach (var seleсtedElement in selectedRef)
                {
                    Element element = doc.GetElement(seleсtedElement);

                    //Разгруппировка
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
                    }

                    if ((BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_IOSModelGroups ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Sections ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Views ||
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Levels)
                        continue;

                    elementList.Add(element);
                    
                    elementData.pnr_1 = element.LookupParameter("PNR_Код по классификатору").AsString();
                    elementData.pnr_2 = element.LookupParameter("PNR_Описание по классификатору").AsString();
                    elementData.guid = element.UniqueId;
                    listDataElement.Add(elementData);

                    listPnr_1.Add(element.LookupParameter("PNR_Код по классификатору").AsString());
                    listPnr_2.Add(element.LookupParameter("PNR_Описание по классификатору").AsString());
                    listGuid.Add(element.UniqueId);

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
                        if (element.LookupParameter("PNR_Код по классификатору") == null || element.LookupParameter("PNR_Описание по классификатору") == null)
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



                /*int i = 0;
                foreach (var element in listData)
                {
                    
                }*/

                int i = 0;

                foreach (var guid in listExcelGuid)
                {
                    var desieredElement = elementList.FirstOrDefault(r => r.UniqueId == guid);
                    if (desieredElement != null)
                    {
                        
                    }
                }

                Transaction tr = new Transaction(doc, "NewGroup");
                tr.Start();
                foreach (var group in groupElements)
                {
                    Group groupNew = doc.Create.NewGroup(group);
                }
                tr.Commit();
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
