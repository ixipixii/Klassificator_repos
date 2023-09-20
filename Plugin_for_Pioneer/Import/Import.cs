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
using System.Security.Policy;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Menu;
using System.Collections;

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
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_GridChains
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

                    //Добавление категорий
                    Category category = element.Category;
                    BuiltInCategory enumCategory = (BuiltInCategory)category.Id.IntegerValue;
                    categoryList.Add(enumCategory);
                }

                var categorySet = new CategorySet();
                IEnumerable<BuiltInCategory> categoryListDistinct = categoryList.Distinct();
                foreach (var category in categoryListDistinct)
                {
                    categorySet.Insert(Category.GetCategory(doc, category));
                }

                //IEnumerable<Element> elementListDistinct = elementList.Distinct();
                using (Transaction ts = new Transaction(doc, "Add parameter"))
                {
                    ts.Start();
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
                    ts.Commit();
                }

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

                List<List<ElementId>> groupElements = new List<List<ElementId>>();
                List<String> groupNames = new List<String>();
                List<Element> groupId = new List<Element>();
                List<GroupType> groupTypes = new List<GroupType>();

                var listDataElementTrueSorted = listDataElementTrue.OrderBy(pr => pr.guid).ToList();

                List<Group> groups = new List<Group>();

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
                            {
                                Group group = (Group)doc.GetElement(desieredElementTrue.element.GroupId);

                                if (group.GroupId.IntegerValue != -1)
                                    continue;

                                /*                                groupId.Add(doc.GetElement(desieredElementTrue.element.GroupId));
                                                                groupNames.Add(group.Name);
                                                                groupTypes.Add(group.GroupType);
                                                                group.GroupType.Name = "DELETE";
                                                                List<ElementId> elements = group.UngroupMembers().ToList();
                                                                doc.Delete(group.Id);
                                                                groupElements.Add(elements);*/


                                if(groups.Contains(group) == false)
                                {
                                    groups.Add(group);
                                }
                                continue;
                            }
                            //Заносим значение в параметр
                            desieredElementTrue.element.LookupParameter("PNR_Код по классификатору").Set(excelElement.pnr_1);
                            desieredElementTrue.element.LookupParameter("PNR_Описание по классификатору").Set(excelElement.pnr_2);
                        }
                    }
                    //PurgeDocument.Purge(doc);
                    transaction.Commit();
                }

                /*                Transaction tr = new Transaction(doc, "NewGroup");
                                tr.Start();
                                int i = 0;
                                foreach (var groupNewInOld in groupElements)
                                {
                                    Group groupNew = doc.Create.NewGroup(groupNewInOld);
                                    groupNew.GroupType = groupTypes[i];
                                    //groupNew.GroupType.Name = groupNames[i];
                                    i++;
                                }
                                //PurgeDocument.Purge(doc);
                                tr.Commit();*/

                TaskDialog.Show("x", $"{groups.Count()}");
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

    internal static class PurgeDocument
    {
        internal static List<ElementId> GetPurgeableElements(Document doc, List<PerformanceAdviserRuleId> performanceAdviserRuleIds)
        {
            List<FailureMessage> failureMessages = PerformanceAdviser.GetPerformanceAdviser().ExecuteRules(doc, performanceAdviserRuleIds).ToList();
            if (failureMessages.Count > 0)
            {
                List<ElementId> purgeableElementIds = failureMessages[0].GetFailingElements().ToList();
                return purgeableElementIds;
            }
            return null;
        }

        public static void Purge(Document doc)
        {
            //The internal GUID of the Performance Adviser Rule 
            const string PurgeGuid = "e8c63650-70b7-435a-9010-ec97660c1bda";

            List<PerformanceAdviserRuleId> performanceAdviserRuleIds = new List<PerformanceAdviserRuleId>();

            //Iterating through all PerformanceAdviser rules looking to find that which matches PURGE_GUID
            foreach (PerformanceAdviserRuleId performanceAdviserRuleId in PerformanceAdviser.GetPerformanceAdviser().GetAllRuleIds())
            {
                if (performanceAdviserRuleId.Guid.ToString() == PurgeGuid)
                {
                    performanceAdviserRuleIds.Add(performanceAdviserRuleId);
                    break;
                }
            }

            //Attempting to recover all purgeable elements and delete them from the document
            List<ElementId> purgeableElementIds = GetPurgeableElements(doc, performanceAdviserRuleIds);

            if (purgeableElementIds.Count > 0)
            {
                foreach (var purgeableElementId in purgeableElementIds)
                {
                    if (purgeableElementId != null)
                    {
                        if (doc.GetElement(purgeableElementId).Name == "DELETE")
                            doc.Delete(purgeableElementId);
                    }
                }
            }
        }
    }

}



