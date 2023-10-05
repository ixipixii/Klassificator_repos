using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Forms;

namespace Plugin_for_Pioneer.CategoryView
{
    public class SelectionCategory
    {
        //                 ItemsSource = "{Binding CategoryList}"
        private ExternalCommandData _commandData;

        public DelegateCommand SaveCommand { get; }

        //static public List<BuiltInCategory> CategoryList { get; set; } = new List<BuiltInCategory>();//Лист категорий

        //public List<BuiltInCategory> SelectedCategoryList { get; set; } = new List<BuiltInCategory>();//Выбранные категории

        public SelectionCategory(ExternalCommandData commandData) 
        {
            _commandData = commandData;
            SaveCommand = new DelegateCommand(OnSaveCommand);
            //CategoryList = Selection.builtInCategories;
        }

        private void OnSaveCommand()
        {
            RaiseCloseRequest();
            var uiapp = _commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Document doc = _commandData.Application.ActiveUIDocument.Document;

            if(CategoryView.SelectedCategoryList.Count == 0)
            {
                return;
            }

            Import import = new Import();
            //import.ImportCategory(CategoryView.SelectedCategoryList, uiapp, uidoc, doc);
        }

        public event EventHandler CloseRequest;

        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }
    }
}
