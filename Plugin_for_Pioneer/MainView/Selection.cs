using Autodesk.Revit.Creation;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Windows.Forms;

namespace Plugin_for_Pioneer
{
    public class Selection
    {              
        private ExternalCommandData _commandData;

        public DelegateCommand SelectCommand3D { get; }

        public DelegateCommand SelectCommandCategory { get; }

        static public List<BuiltInCategory> builtInCategories = new List<BuiltInCategory>();

        static public List<String> CategoryNameRUSList = new List<String>();

        public Selection(ExternalCommandData commandData)
        {
            _commandData = commandData;
            SelectCommand3D = new DelegateCommand(OnSelectCommand3D);
            SelectCommandCategory = new DelegateCommand(OnSelectCommandCategory);
        }

        public event EventHandler CloseRequest;
        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }

        private void OnSelectCommand3D()
        {
            RaiseCloseRequest();

            ElementSet elementSet = new ElementSet();
            Import import = new Import();
            String str = "Import3d";
            import.Execute(_commandData, ref str, elementSet);

        }

        private void OnSelectCommandCategory()
        {          
            RaiseCloseRequest();

            List<String> builtInNameList = new List<String>();



            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog.Filter = "All files(*.*)|*.*";

            string filePath = string.Empty;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
                filePath = openFileDialog.FileName;

            if (string.IsNullOrEmpty(filePath))
                TaskDialog.Show("Ошибка", "Файла нет");

            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(filePath);
                ISheet sheet = workbook.GetSheetAt(index: 0);

                int rowIndex = 0;
                while (sheet.GetRow(rowIndex) != null)
                {
                    if (sheet.GetRow(rowIndex).GetCell(0) == null ||
                        sheet.GetRow(rowIndex).GetCell(1) == null)
                    {
                        rowIndex++;
                        continue;
                    }

                    string builtInName = sheet.GetRow(rowIndex).GetCell(0).StringCellValue;
                    string categoryNameRUS = sheet.GetRow(rowIndex).GetCell(1).StringCellValue;

                    builtInNameList.Add(builtInName);
                    CategoryNameRUSList.Add(categoryNameRUS);

                    rowIndex++;
                }
            }
            
            foreach (var c in Enum.GetValues(typeof(Autodesk.Revit.DB.BuiltInCategory)))
            {
                var b = builtInNameList.FirstOrDefault(x => x.Equals(c.ToString()));

                if (b != null)
                {
                    builtInCategories.Add((BuiltInCategory)c);
                }
            }

            var window = new CategoryView.CategoryView(_commandData);
            window.ShowDialog();
        }

    }
}
