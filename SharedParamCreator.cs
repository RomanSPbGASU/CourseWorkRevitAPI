using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;

namespace CourseWorkRevitAPI
{
    [Transaction(TransactionMode.Manual)]
    class SharedParamCreator : IExternalCommand
    {
        const string _filename = "SharedParams.txt";
        const string _groupname = "Coursework Parameters";
        const string _defname = "SP";
        public static List<Category> categories;
        ParameterType _deftype = ParameterType.Number;

        BuiltInCategory[] targets = new BuiltInCategory[] {
            BuiltInCategory.OST_Doors,
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_Roofs,
        };

        Category GetCategory(
          Document doc,
          BuiltInCategory target)
        {
            Category cat = null;

            if (target.Equals(BuiltInCategory.OST_IOSModelGroups))
            {
                // determine model group category:

                FilteredElementCollector collector
                  = Util.GetElementsOfType(doc, typeof(Group), // GroupType works as well
                    BuiltInCategory.OST_IOSModelGroups);

                IList<Element> modelGroups = collector.ToElements();

                if (0 == modelGroups.Count)
                {
                    Util.ErrorMsg("Please insert a model group.");
                    return cat;
                }
                else
                {
                    cat = modelGroups[0].Category;
                }
            }
            else
            {
                try
                {
                    cat = doc.Settings.Categories.get_Item(target);
                }
                catch (Exception ex)
                {
                    Util.ErrorMsg(string.Format(
                      "Error obtaining document {0} category: {1}",
                      target.ToString(), ex.Message));
                    return cat;
                }
            }
            if (null == cat)
            {
                Util.ErrorMsg(string.Format(
                  "Unable to obtain the document {0} category.",
                  target.ToString()));
            }
            return cat;
        }

        /// <summary>
        /// Create a new shared parameter
        /// </summary>
        /// <param name="doc">Document</param>
        /// <param name="cat">Category to bind the parameter definition</param>
        /// <param name="nameSuffix">Parameter name suffix</param>
        /// <param name="typeParameter">Create a type parameter? If not, it is an instance parameter.</param>
        /// <returns></returns>
        bool CreateSharedParameter(
          Document doc,
          Category cat,
          int? nameSuffix,
          bool typeParameter,
          string parameterName = null)
        {
            Application app = doc.Application;

            Autodesk.Revit.Creation.Application ca = app.Create;

            // get or set the current shared params filename:

            string filename = app.SharedParametersFilename;

            if (0 == filename.Length)
            {
                string path = _filename;
                StreamWriter stream;
                stream = new StreamWriter(File.Open(path, System.IO.FileMode.Append));
                stream.Close();
                app.SharedParametersFilename = path;
                filename = app.SharedParametersFilename;
            }

            // get the current shared params file object:

            DefinitionFile file
              = app.OpenSharedParameterFile();

            if (null == file)
            {
                Util.ErrorMsg(
                  "Error getting the shared params file.");

                return false;
            }

            // get or create the shared params group:

            DefinitionGroup group
              = file.Groups.get_Item(_groupname);

            if (null == group)
            {
                group = file.Groups.Create(_groupname);
            }

            if (null == group)
            {
                Util.ErrorMsg(
                  "Error getting the shared params group.");

                return false;
            }

            // set visibility of the new parameter:

            // Category.AllowsBoundParameters property
            // indicates if a category can have user-visible
            // shared or project parameters. If it is false,
            // it may not be bound to visible shared params
            // using the BindingMap. Please note that
            // non-user-visible parameters can still be
            // bound to these categories.

            bool visible = cat.AllowsBoundParameters;

            // get or create the shared params definition:

            string defname = parameterName ?? _defname + nameSuffix?.ToString();
            Debug.Print($"Creating parameter name: {defname}");

            Definition definition = group.Definitions.get_Item(
              defname);

            if (null == definition)
            {
                //definition = group.Definitions.Create( defname, _deftype, visible ); // 2014

                ExternalDefinitionCreationOptions opt = new ExternalDefinitionCreationOptions(defname, _deftype);

                opt.Visible = visible;

                definition = group.Definitions.Create(opt); // 2015
            }
            if (null == definition)
            {
                Util.ErrorMsg(
                  "Error creating shared parameter.");

                return false;
            }

            // create the category set containing our category for binding:

            if (SharedParamCreator.categories == null)
            {
                categories = new List<Category> { cat };
            }
            else
            {
                SharedParamCreator.categories.Add(cat);
            }

            var catSet = ca.NewCategorySet();
            foreach(var category in SharedParamCreator.categories)
            {
                catSet.Insert(category);
            }

            // bind the param:

            try
            {
                Binding binding = typeParameter
                  ? ca.NewTypeBinding(catSet) as Binding
                  : ca.NewInstanceBinding(catSet) as Binding;

                // we could check if it is already bound,
                // but it looks like insert will just ignore
                // it in that case:

                doc.ParameterBindings.Insert(definition, binding);

                // we can also specify the parameter group here:

                //doc.ParameterBindings.Insert( definition, binding,
                //  BuiltInParameterGroup.PG_GEOMETRY );

                Debug.Print(
                  "Created a shared {0} parameter '{1}' for the {2} category.",
                  (typeParameter ? "type" : "instance"),
                  defname, cat.Name);
            }
            catch (Exception ex)
            {
                Util.ErrorMsg(string.Format(
                  "Error binding shared parameter to category {0}: {1}",
                  cat.Name, ex.Message));
                return false;
            }
            return true;
        }

        /// <summary>
        /// This is dirty hack for processing List of categories
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="cat"></param>
        /// <param name="nameSuffix"></param>
        /// <param name="typeParameter"></param>
        /// <param name="parameterName"></param>
        /// <returns></returns>
        bool CreateSharedParameter(
            Document doc,
            List<BuiltInCategory> cat,
            int? nameSuffix,
            bool typeParameter,
            string parameterName = null)
        {
            bool res = true;
            foreach (var category in cat)
            {
                res &= this.CreateSharedParameter(doc, Category.GetCategory(doc, category), nameSuffix, typeParameter, parameterName);
            }
            return res;
        }

            public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication app = commandData.Application;
            Document doc = app.ActiveUIDocument.Document;

            string parameterName = "Огнестойкость";
            List<BuiltInCategory> parameterCategories = new List<BuiltInCategory> {
                BuiltInCategory.OST_Doors,
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Windows,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_Roofs,
            };

            using (Transaction t = new Transaction(doc))
            {
                // put variable in string
                t.Start($"Check Parameter {parameterName}");
                Category cat;
                int i = 0;
                //List<Element> elementsList;

                var fec = new FilteredElementCollector(doc);
                List<Element> sharedElementsList, elementList;


                try
                {
                    //doors = new List<Element>(new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_Doors).ToElements());
                    sharedElementsList = (List<Element>)(new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(SharedParameterElement)).ToElements());
                    elementList = (List<Element>)(new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(ParameterElement)).ToElements());
                }
                catch (ArgumentNullException e)
                {
                    Console.WriteLine(e.Message);
                    return Result.Failed;
                }

                List<Element> AllElements = sharedElementsList.Concat(elementList).ToList();

                foreach (var element in AllElements)
                {
                    var parameters = element.GetParameters(parameterName);
                    if (parameters.Count != 0)
                    {
                        string GoodMessage = $"Parameter {parameterName} was found. Plugin works well!";
                        Debug.Print(GoodMessage);
                        Util.InfoMsg(GoodMessage);
                        return Result.Succeeded;
                    }
                }
                
                // create instance parameters:

                foreach (BuiltInCategory target in targets)
                {
                    cat = GetCategory(doc, target);
                    if (null != cat)
                    {
                        CreateSharedParameter(doc, cat, ++i, false);
                    }
                }

                string NotSoGoodMessage = $"Parameter {parameterName} wasn't found in document... Don't panic, we can create it automatically!";
                Debug.Print(NotSoGoodMessage);
                Util.InfoMsg(NotSoGoodMessage);

                // create a type parameter:

                try
                {
                    CreateSharedParameter(doc, parameterCategories, null, false, parameterName);
                    t.Commit();
                }
                catch
                {
                    string WrongMessage = "Something went wrong. Sad but true.";
                    Debug.Print(WrongMessage);
                    Util.InfoMsg(WrongMessage);
                    t.RollBack();
                    return Result.Failed;
                }

                string ExitMessage = "Well done. Good work";
                Debug.Print(ExitMessage);
                Util.InfoMsg(ExitMessage);
            }

            return Result.Succeeded;
        }

        #region REINSERT
#if REINSERT
    public static BindSharedParamResult BindSharedParam(
      Document doc,
      Category cat,
      string paramName,
      string grpName,
      ParameterType paramType,
      bool visible,
      bool instanceBinding )
    {
      try // generic
      {
        Application app = doc.Application;

        // This is needed already here to 
        // store old ones for re-inserting

        CategorySet catSet = app.Create.NewCategorySet();

        // Loop all Binding Definitions
        // IMPORTANT NOTE: Categories.Size is ALWAYS 1 !?
        // For multiple categories, there is really one 
        // pair per each category, even though the 
        // Definitions are the same...

        DefinitionBindingMapIterator iter
          = doc.ParameterBindings.ForwardIterator();

        while( iter.MoveNext() )
        {
          Definition def = iter.Key;
          ElementBinding elemBind
            = (ElementBinding) iter.Current;

          // Got param name match

          if( paramName.Equals( def.Name,
            StringComparison.CurrentCultureIgnoreCase ) )
          {
            // Check for category match - Size is always 1!

            if( elemBind.Categories.Contains( cat ) )
            {
              // Check Param Type

              if( paramType != def.ParameterType )
                return BindSharedParamResult.eWrongParamType;

              // Check Binding Type

              if( instanceBinding )
              {
                if( elemBind.GetType() != typeof( InstanceBinding ) )
                  return BindSharedParamResult.eWrongBindingType;
              }
              else
              {
                if( elemBind.GetType() != typeof( TypeBinding ) )
                  return BindSharedParamResult.eWrongBindingType;
              }

              // Check Visibility - cannot (not exposed)
              // If here, everything is fine, 
              // ie already defined correctly

              return BindSharedParamResult.eAlreadyBound;
            }

            // If here, no category match, hence must 
            // store "other" cats for re-inserting

            else
            {
              foreach( Category catOld
                in elemBind.Categories )
                catSet.Insert( catOld ); // 1 only, but no index...
            }
          }
        }

        // If here, there is no Binding Definition for 
        // it, so make sure Param defined and then bind it!

        DefinitionFile defFile
          = GetOrCreateSharedParamsFile( app );

        DefinitionGroup defGrp
          = GetOrCreateSharedParamsGroup(
            defFile, grpName );

        Definition definition
          = GetOrCreateSharedParamDefinition(
            defGrp, paramType, paramName, visible );

        catSet.Insert( cat );

        InstanceBinding bind = null;

        if( instanceBinding )
        {
          bind = app.Create.NewInstanceBinding(
            catSet );
        }
        else
        {
          bind = app.Create.NewTypeBinding( catSet );
        }

        // There is another strange API "feature". 
        // If param has EVER been bound in a project 
        // (in above iter pairs or even if not there 
        // but once deleted), Insert always fails!? 
        // Must use .ReInsert in that case.
        // See also similar findings on this topic in: 
        // http://thebuildingcoder.typepad.com/blog/2009/09/adding-a-category-to-a-parameter-binding.html 
        // - the code-idiom below may be more generic:

        if( doc.ParameterBindings.Insert(
          definition, bind ) )
        {
          return BindSharedParamResult.eSuccessfullyBound;
        }
        else
        {
          if( doc.ParameterBindings.ReInsert(
            definition, bind ) )
          {
            return BindSharedParamResult.eSuccessfullyBound;
          }
          else
          {
            return BindSharedParamResult.eFailed;
          }
        }
      }
      catch( Exception ex )
      {
        MessageBox.Show( string.Format(
          "Error in Binding Shared Param: {0}",
          ex.Message ) );

        return BindSharedParamResult.eFailed;
      }
    }
#endif // REINSERT
        #endregion // REINSERT

        #region SetAllowVaryBetweenGroups
        /// <summary>
        /// Helper method to control `SetAllowVaryBetweenGroups` 
        /// option for instance binding param
        /// </summary>
        static void SetInstanceParamVaryBetweenGroupsBehaviour(
          Document doc,
          Guid guid,
          bool allowVaryBetweenGroups = true)
        {
            try // last resort
            {
                SharedParameterElement sp = SharedParameterElement.Lookup(doc, guid);

                // Should never happen as we will call 
                // this only for *existing* shared param.

                if (null == sp) return;

                InternalDefinition def = sp.GetDefinition();

                if (def.VariesAcrossGroups != allowVaryBetweenGroups)
                {
                    // Must be within an outer transaction!

                    def.SetAllowVaryBetweenGroups(doc, allowVaryBetweenGroups);
                }
            }
            catch { } // ideally, should report something to log...
        }

#if SetInstanceParamVaryBetweenGroupsBehaviour_SAMPLE_CALL
    // Assumes outer transaction
    public static Parameter GetOrCreateElemSharedParam( 
      Element elem,
      string paramName,
      string grpName,
      ParameterType paramType,
      bool visible,
      bool instanceBinding,
      bool userModifiable,
      Guid guid,
      bool useTempSharedParamFile,
      string tooltip = "",
      BuiltInParameterGroup uiGrp = BuiltInParameterGroup.INVALID,
      bool allowVaryBetweenGroups = true )
    {
      try
      {
        // Check if existing
        Parameter param = elem.LookupParameter( paramName );
        if( null != param )
        {
          // NOTE: If you don't want forcefully setting 
          // the "old" instance params to 
          // allowVaryBetweenGroups =true,
          // just comment the next 3 lines.
          if( instanceBinding && allowVaryBetweenGroups )
          {
            SetInstanceParamVaryBetweenGroupsBehaviour( 
              elem.Document, guid, allowVaryBetweenGroups );
          }
          return param;
        }

        // If here, need to create it (my custom 
        // implementation and classes
)

        BindSharedParamResult res = BindSharedParam( 
          elem.Document, elem.Category, paramName, grpName,
          paramType, visible, instanceBinding, userModifiable,
          guid, useTempSharedParamFile, tooltip, uiGrp );

        if( res != BindSharedParamResult.eSuccessfullyBound
          && res != BindSharedParamResult.eAlreadyBound )
        {
          return null;
        }

        // Set AllowVaryBetweenGroups for NEW Instance 
        // Binding Shared Param

        if( instanceBinding )
        {
          SetInstanceParamVaryBetweenGroupsBehaviour( 
            elem.Document, guid, allowVaryBetweenGroups );
        }

        // If here, binding is OK and param seems to be
        // IMMEDIATELY available from the very same command

        return elem.LookupParameter( paramName );
      }
      catch( Exception ex )
      {
        System.Windows.Forms.MessageBox.Show( 
          string.Format( 
            "Error in getting or creating Element Param: {0}", 
            ex.Message ) );

        return null;
      }
    }
#endif // SetInstanceParamVaryBetweenGroupsBehaviour_SAMPLE_CALL
        #endregion // SetAllowVaryBetweenGroups
    }
}
