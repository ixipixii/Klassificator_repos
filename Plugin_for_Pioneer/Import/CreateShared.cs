using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;


namespace Plugin_for_Pioneer
{
    internal class CreateShared
    {
        public int CreateSharedParameter(Application application, Document doc, string parameterName,
                                   CategorySet categorySet, BuiltInParameterGroup builtInParameterGroup,
                                   bool inInstance)
        {
            DefinitionFile definitionFile = application.OpenSharedParameterFile();
            if (definitionFile == null)
                TaskDialog.Show("Ошибка", "Не найден ФОП");

            Definition definition = definitionFile.Groups.SelectMany(group => group.Definitions).
                FirstOrDefault(def => def.Name.Equals(parameterName));

            if (definition == null)
            {
                TaskDialog.Show("Ошибка", "Не найден указанный параметр");
                return 1;
            }

            Binding binding = application.Create.NewTypeBinding(categorySet);
            if (inInstance)
                binding = application.Create.NewInstanceBinding(categorySet);

            BindingMap map = doc.ParameterBindings;

            try
            {
                //map.Remove(definition);
                map.Insert(definition, binding, builtInParameterGroup);
                //map.ReInsert(definition, binding, builtInParameterGroup);
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                TaskDialog.Show("Ошибка", "Выбранной категории нет в модели");
                return 1;
            }

            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
                return 0;
            }
            return 0;
            
            /*map.ReInsert(definition, binding, builtInParameterGroup);*/

        }

    }
}
