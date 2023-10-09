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

namespace Plugin_for_Pioneer
{
    [Transaction(TransactionMode.Manual)]
    internal class Create : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            CreateParameter(uiapp, uidoc, doc);
            return Result.Succeeded;
        }

        public Result CreateParameter(UIApplication uiapp, UIDocument uidoc, Document doc)
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
                var categoryList = new List<BuiltInCategory>();

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
                        (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_WindowsOpeningCut
                        )
                        continue;

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
            }

            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }

            return Result.Succeeded;
        }       
    }

}



