using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Plugin_for_Pioneer.CategoryView
{
    /// <summary>
    /// Логика взаимодействия для CategoryView.xaml
    /// </summary>
    public partial class CategoryView : Window
    {
        static public List<BuiltInCategory> SelectedCategoryList = new List<BuiltInCategory>();

        static public List<String> RusNameList { get; set; } = Selection.CategoryNameRUSList;

        public String text = String.Empty;
        public CategoryView(ExternalCommandData commandData)
        {       
            InitializeComponent();

            foreach(var item in Selection.builtInCategories)
            {
                LB.Items.Add(item);
                
            }

            SelectionCategory selectionCategory = new SelectionCategory(commandData);
            selectionCategory.CloseRequest += (s, e) => this.Close();
            DataContext = selectionCategory;
        }

        public void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selected = LB.SelectedItems.Cast<BuiltInCategory>();
            SelectedCategoryList = selected.ToList();
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            text = TB.Text;
            LB.Items.Clear();
            foreach (var item in Selection.builtInCategories)
            {
                if(item.ToString().StartsWith(text))
                {
                    LB.Items.Add(item);
                }
            }
            if(text == "")
            {
                LB.Items.Clear();
                foreach (var item in Selection.builtInCategories)
                    LB.Items.Add(item);
            }
        }
    }
}
