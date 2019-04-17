using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace RevitViewAndSheetManager
{
    public class RevitManager
    {
        /////////////
        //variables//
        /////////////

        private Document doc = null;
        private UIDocument uiDoc = null;

        private TransactionGroup transactionList;

        public bool debugMode = false;

        //TODO: error handling, have a list of errors and a way to get last error

        ////////////////
        //constructors//
        ////////////////

        public RevitManager(ExternalCommandData commandData)
        {
            //check if we have been given command data
            if (commandData == null)
                throw new ArgumentNullException("commandData");

            //preparing 
            uiDoc = commandData.Application.ActiveUIDocument;
            doc = uiDoc.Document;
            transactionList = new TransactionGroup(doc);
        }

        ////////////////////
        //public functions//
        ////////////////////

        //creates a new sheet giving it a default sheet number and no title block
        public bool CreateSheet(string sheetName)
        {
            return CreateSheet(sheetName, null, null);
        }

        //creates a new sheet giving it a default sheet number and the specified title block
        public bool CreateSheet(string sheetName, string titleBlock)
        {
            return CreateSheet(sheetName, null, titleBlock);
        }

        //creates a new sheet giving it the specified sheet number and title block
        public bool CreateSheet(string sheetName, string sheetNumber, string titleBlock)
        {
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Creating new ViewSheet.");

                try
                {
                    ElementId tempTitleBlock = ElementId.InvalidElementId;

                    if (titleBlock != null)
                    {
                        tempTitleBlock = GetTitleBlockId(titleBlock);
                        if (tempTitleBlock == null)
                            throw new Exception("Unable to find specified Title Block: " + titleBlock);
                    }

                    // Create a sheet view
                    ViewSheet viewSheet = ViewSheet.Create(doc, tempTitleBlock);
                    if (viewSheet == null)
                        throw new Exception("Failed to create new ViewSheet.");

                    viewSheet.Name = sheetName;

                    //if given a sheet number attempt to change it, other wise leave it at default
                    if (sheetNumber != null)
                        viewSheet.SheetNumber = sheetNumber;

                    t.Commit();

                    return true;
                }
                catch
                {
                    if (debugMode)
                        ShowMessageBox("Warning!",
                                "ID_TaskDialog_Warning",
                                "Unable to create a new ViewSheet by the name of " + sheetName + ", of number " + sheetNumber + ", of titleblock type " + titleBlock,
                                TaskDialogIcon.TaskDialogIconWarning,
                                TaskDialogCommonButtons.Close);
                
                t.RollBack();
                    return false;
                }
            }
        }

        //wipe the first view with the given name
        public bool DeleteView(string viewName)
        {
            ElementId v = GetViewId(viewName);

            if (v != null)
                return Delete(v);
            else
                return false;
        }

        //wipe the first sheet with the given name
        public bool DeleteSheet(string sheetName)
        {
            if (sheetName != null)
                return Delete(GetSheetId(sheetName));
            else
                return false;
        }

        //moves the given view into a new family tree
        public bool MoveView(string viewName, string newLocation)
        {
            View v = GetView(viewName);
            ElementId id = GetViewFamilyType(newLocation);

            if (v != null && id != null)//makeu sure there is a view and a family to move it to
            {
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Changing view id.");
                    v.ChangeTypeId(id);
                    t.Commit();

                    return true;
                }
            }
            else
            {   //did not find a view or an element, show user an error        
                if (debugMode)
                {

                    string msg = string.Empty;

                    //preparing the error message
                    int i = 0;
                    if (v == null)
                        i++;

                    if (id == null)
                        i = +2;

                    switch (i)
                    {
                        case 1://unable to find view
                            msg = "Unable to move view '" + viewName + "', could not find view.";
                            break;

                        case 2://unable to find family
                            msg = "Unable to move view '" + viewName + "', could not find family id.";
                            break;

                        case 3://unable to find view AND template
                            msg = "Unable to move view '" + viewName + "', could not find view AND family id.";
                            break;
                    }

                    //show the message
                    ShowMessageBox("Warning!",
                                "ID_TaskDialog_Warning",
                                msg,
                                TaskDialogIcon.TaskDialogIconWarning,
                                TaskDialogCommonButtons.Close);
                }

                return false;
            }

        }

        //rotates all entities on a given view in 90 degree angles
        //TODO: fix this function, views are not currently rotating as desired, or at all. may require duplication of view BEFORE attempting to rotate? needs testing
        public void RotateView(string viewName, RotationAngle rotation)
        {
            //get the id of the view we need to rotate
            View v = GetView(viewName);

            if (v == null)
                return;//no view to work with

            //attempt to get the crop box directly
            Element cropBox = doc.GetElement(GetCropBox(v));
            
            BoundingBoxXYZ bbox = v.CropBox;

            XYZ center = 0.5 * (bbox.Max + bbox.Min);

            Line axis = Line.CreateBound(
                center, center + XYZ.BasisZ);

            RotateElement(cropBox.Id, axis, (Math.PI / 2) * (int)rotation);            
        }
        
        //rotate an element by a specific angle
        public bool RotateElement(ElementId id, Line axis, double angle)
        {
            try
            {
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Attempting to rotate element " + id.IntegerValue.ToString() + " by " + angle.ToString());
                    ElementTransformUtils.RotateElement(doc, id, axis, angle);
                    t.Commit();
                }
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        //rotate an element by 90 degree angles
        public bool RotateElement(ElementId id, Line axis, RotationAngle angle)
        {
            try
            {
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Attempting to rotate element " + id.IntegerValue.ToString() + " to " + angle.ToString());
                    ElementTransformUtils.RotateElement(doc, id, axis, (Math.PI / 2) * (int)angle);
                    t.Commit();
                }
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        //finds and renames a view
        public bool RenameView(string viewName, string newViewName)
        {
            View v = GetView(viewName);//find the view

            if (v != null)//make sure we have a view to work with
            {
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Renaming view.");
                    v.Name = newViewName;
                    t.Commit();
                }

                return true;
            }
            else
            {   //did not find aview, give the user an error
                if (debugMode)                
                    ShowMessageBox("Warning!",
                            "ID_TaskDialog_Warning",
                            "Unable to rename view '" + viewName + "', view may not exist!",
                            TaskDialogIcon.TaskDialogIconWarning,
                            TaskDialogCommonButtons.Close);
                
                return false;                
            }
        }

        //duplicates the view naming it as given, duplicate will be in same location
        public bool DuplicateViewByName(string viewName, string newViewName, ViewDuplicateOption duplicateOption)
        {
            //get the view we will duplicate first
            View v = GetView(viewName);

            if (v != null)//make sure there is indeed a view
            {
                using (Transaction t = new Transaction(doc))//duplicate the view
                {
                    t.Start("Duplicating view.");
                    ElementId id = v.Duplicate(duplicateOption);

                    View tempView = doc.GetElement(id) as View;
                    tempView.Name = newViewName;//rename the view to the new name

                    t.Commit();
                }

                return true;
            }
            else
            {
                if (debugMode)
                    ShowMessageBox("Warning!",
                                "ID_TaskDialog_Warning",
                                "Unable to duplicate view '" + viewName + "'.",
                                TaskDialogIcon.TaskDialogIconWarning,
                                TaskDialogCommonButtons.Close);

                return false;
            }
        }

        //changes the template of the given view
        public bool ChangeViewTemplate(string viewName, string templateName)
        {
            View v = GetView(viewName);
            ElementId id = GetViewTemplateId(templateName);

            if (v != null && id != null)//check that we have a view and a template to work with
            {
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Changing view template.");
                    v.ViewTemplateId = id;
                    t.Commit();
                }

                return true;
            }
            else
            {   //there was no view or template, warn the user
                //create error box                
                if (debugMode)
                {                    
                    string msg = string.Empty;

                    //preparing the error message
                    int i = 0;
                    if (v == null)
                        i++;

                    if (id == null)
                        i = +2;

                    switch (i)
                    {
                        case 1://unable to find view
                            msg = "Unable to change template for view '" + viewName + "', the view may not exist.";
                            break;

                        case 2://unable to find template
                            msg = "Unable to change template for view '" + viewName + "', the template may not exist.";
                            break;

                        case 3://unable to find view AND template
                            msg = "Unable to change template for view '" + viewName + "', the view AND template may not exist.";
                            break;
                    }

                    ShowMessageBox("Warning!",
                                "ID_TaskDialog_Warning",
                                msg,
                                TaskDialogIcon.TaskDialogIconWarning,
                                TaskDialogCommonButtons.Close);                    
                }

                return false;
            }
        }

        //removes all independent tags from the view
        public void RemoveAllIndependentTags(string viewName)
        {
            ElementId v = GetViewId(viewName);//get the view we will be working with

            if (v == null)//check if we have a view to work with
            {
                if (debugMode)
                    ShowMessageBox("Warning!",
                                "ID_TaskDialog_Warning",
                                "Unable to remove all independant tags from " + viewName + ", it may not exist.",
                                TaskDialogIcon.TaskDialogIconWarning,
                                TaskDialogCommonButtons.Close);

                return;
            }
            

            //grab all independent tags owned by the view we are worknig with
            IEnumerable<IndependentTag> coll = new FilteredElementCollector(doc)
                 .OfClass(typeof(IndependentTag))
                 .Cast<IndependentTag>()
                 .Where(i => i.OwnerViewId.Equals(v));

            //wipe all found tags
            List<ElementId> coll2 = new List<ElementId>();

            foreach (IndependentTag i in coll)
                coll2.Add(i.Id);

            Delete(coll2);
        }

        //removes ALL dimensions from given view
        public void RemoveAllDimensions(string viewName)
        {
            ElementId v = GetViewId(viewName);//get the view to work with

            if (v == null)//check if we have a view to work with
            {
                if (debugMode)
                    ShowMessageBox("Warning!",
                                "ID_TaskDialog_Warning",
                                "Unable to remove all dimensions from " + viewName + ", it may not exist.",
                                TaskDialogIcon.TaskDialogIconWarning,
                                TaskDialogCommonButtons.Close);
                return;
            }
            

            //get all the dimensions
            IEnumerable<Dimension> coll = new FilteredElementCollector(doc)
                 .OfClass(typeof(Dimension))
                 .Cast<Dimension>()
                 .Where(i => i.OwnerViewId.Equals(v));

            //wipe all found dimensions
            List<ElementId> coll2 = new List<ElementId>();

            foreach (Dimension i in coll)
                coll2.Add(i.Id);

            Delete(coll2);
        }

        //removes all dimensions of a specific type fro ma given view
        public void RemoveAllDimensionsOfType(string viewName, DimensionStyleType dimType)
        {
            ElementId v = GetViewId(viewName);//get the view to work with

            if (v == null)//check if we have a view to work with
            {
                if (debugMode)
                    ShowMessageBox("Warning!",
                                "ID_TaskDialog_Warning",
                                "Unable to remove all independant tags of type " + dimType.ToString() + " from " + viewName + ", it may not exist.",
                                TaskDialogIcon.TaskDialogIconWarning,
                                TaskDialogCommonButtons.Close);

                return;
            }

            //get all the dimensions we need to wipe
            IEnumerable<Dimension> coll = new FilteredElementCollector(doc)
                 .OfClass(typeof(Dimension))
                 .Cast<Dimension>()
                 .Where(i => i.OwnerViewId.Equals(v)
                 && i.DimensionType.StyleType == dimType);

            //wipe all found dimensions
            List<ElementId> coll2 = new List<ElementId>();

            foreach (Dimension i in coll)
                coll2.Add(i.Id);

            Delete(coll2);
        }

        //removes ALL text notes from a given view
        public void RemoveAllTextNotes(string viewName)
        {
            ElementId v = GetViewId(viewName);//get the view to work with

            if (v == null)//check if we have a view to work with    
            {
                if (debugMode)
                    ShowMessageBox("Warning!",
                                "ID_TaskDialog_Warning",
                                "Unable to remove all text notes from " + viewName + ", it may not exist.",
                                TaskDialogIcon.TaskDialogIconWarning,
                                TaskDialogCommonButtons.Close);
                return;
            }

            //get all text notes
            IEnumerable<TextNote> coll = new FilteredElementCollector(doc)
                 .OfClass(typeof(TextNote))
                 .Cast<TextNote>()
                 .Where(b => b.OwnerViewId == v);

            //wipe all found text notes
            List<ElementId> coll2 = new List<ElementId>();

            foreach (TextNote i in coll)
                coll2.Add(i.Id);

            Delete(coll2);
        }

        //wipe all generic annotation familys from the given view
        public void RemoveAllAnnotations(string viewName)
        {
            //make sure we have a view to work with
            ElementId v = GetViewId(viewName);

            if (v == null)
            {
                if (debugMode)
                    ShowMessageBox("Warning!",
                                "ID_TaskDialog_Warning",
                                "Unable to remove all generic annotations from " + viewName + ", it may not exist.",
                                TaskDialogIcon.TaskDialogIconWarning,
                                TaskDialogCommonButtons.Close);
                return;
            }

            //find all annotation symbols in the given view
            IEnumerable<FamilyInstance> coll = new FilteredElementCollector(doc, v)
                  .OfClass(typeof(FamilyInstance))
                  .Cast<FamilyInstance>()
                  .Where(b => b.Symbol.GetType().Name.Equals("AnnotationSymbolType"));

            //wipe all found symbols
            List<ElementId> coll2 = new List<ElementId>();

            foreach (FamilyInstance i in coll)
                coll2.Add(i.Id);

            Delete(coll2);
        }

        //removes all groups from the given view
        public void RemoveAllGroups(string viewName)
        {
            //make sure we have a view to work with
            ElementId v = GetViewId(viewName);

            if (v == null)
            {
                if (debugMode)
                    ShowMessageBox("Warning!",
                                "ID_TaskDialog_Warning",
                                "Unable to remove all groups from " + viewName + ", it may not exist.",
                                TaskDialogIcon.TaskDialogIconWarning,
                                TaskDialogCommonButtons.Close);
                return;
            }

            //find all annotation symbols in the given view
            IEnumerable<GroupType> coll = new FilteredElementCollector(doc, v)
                  .OfClass(typeof(GroupType))
                  .Cast<GroupType>();

            //wipe all found symbols
            List<ElementId> coll2 = new List<ElementId>();

            foreach (Element i in coll)
                coll2.Add(i.Id);

            Delete(coll2);
        }

        //adds the given view to the given sheet then returns the ElementId, places the view directly in the middle of the sheet
        public ElementId AddViewToSheet(string sheetName, string viewName)
        {
            return AddViewToSheet(sheetName, viewName, null);
        }

        //adds the given view to the given sheet then returns the ElementId, places the view at the given coordiantes on the sheet
        public ElementId AddViewToSheet(string sheetName, string viewName, XYZ xyz)
        {
            //get the sheet and the view
            ElementId s = GetSheetId(sheetName);
            ElementId v = GetViewId(viewName);

            //check if we have elements to work with
            if (s == null || v == null)
                return null;

            //check if we have been given a specific location to place the viewport
            if (xyz == null)
                xyz = new XYZ(0, 0, 0);

            //create the viewport
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Creating ViewPort.");
                Viewport vp = Viewport.Create(doc, s, v, xyz);
                t.Commit();

                return vp.Id;
            }
        }

        //returns the ID of the crop box of a given view
        //With thanks to Jeremy Tammik
        //https://thebuildingcoder.typepad.com/blog/2018/02/efficiently-retrieve-crop-box-for-given-view.html
        public ElementId GetCropBox(View view)
        {        
            if (view == null)
                return null;

            return new FilteredElementCollector(doc)
               .WherePasses(new ElementParameterFilter(
                                new FilterElementIdRule(
                                    new ParameterValueProvider(new ElementId((int)BuiltInParameter.ID_PARAM)), new FilterNumericEquals(), view.Id)))
               .ToElementIds()
               .Where<ElementId>(b => b.IntegerValue != view.Id.IntegerValue)
               .FirstOrDefault<ElementId>();
        }

        public ElementId GetCropBox(string viewName)
        {
            View v = GetView(viewName);

            return GetCropBox(v);
        }



        //load all views and return them in a viewset. skips all templates, use GetAllViewTempaltes for this
        public ViewSet GetAllViews()
        {
            //find all views
            IEnumerable<View> views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(b => !b.IsTemplate
                && b.ViewType != ViewType.DrawingSheet);

            //add all teh views into a ViewSet
            ViewSet allViews = new ViewSet();
            foreach (View v in views)
                allViews.Insert(v);

            //return the ViewSet
            return allViews;
        }

        //Specifically checks for the sheet number
        public XYZ GetSheetDimensions(string sheetName)
        {
            //find a specific title block
            View sheet = GetSheet(sheetName);


            // retrieve the title block instances            
            Parameter p;
            FilteredElementCollector a = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_TitleBlocks)
            .OfClass(typeof(FamilyInstance));

            List<string> sheight = new List<string>();
            List<string> swidth = new List<string>();

            foreach (FamilyInstance e in a)
            {
                p = e.get_Parameter(
                  BuiltInParameter.SHEET_NUMBER);

                string sheet_number = p.AsString();

                //check if we are on the correct sheet
                if (sheet_number != sheetName)
                    continue;

                //we are on the correct sheet, get the width and height and return it

                p = e.get_Parameter(
                  BuiltInParameter.SHEET_WIDTH);

                double width = p.AsDouble();

                p = e.get_Parameter(
                  BuiltInParameter.SHEET_HEIGHT);

                double height = p.AsDouble();

                XYZ xyz = new XYZ(width, height, 0);

                return xyz;
            }

            //didnt find the sheet, return null
            return null;
        }

        public void ChangeViewPortType(string sheetName, ElementId viewportId, string typeName)
        {
            ElementId i = GetSheetId(sheetName);

            if (i == null)
                return;

            IEnumerable<Viewport> viewports = new FilteredElementCollector(doc, i)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .Where(v => v.Id.Equals(viewportId));

            Viewport vp = viewports.FirstOrDefault();

            if (vp == null)
                return;

            if (!vp.Name.Equals(typeName))
                foreach (ElementId id in vp.GetValidTypes())
                {
                    ElementType type = doc.GetElement(id) as ElementType;
                    if (type.Name.Equals(typeName))
                    {
                        using (Transaction t = new Transaction(doc))
                        {
                            t.Start("Changing viewport type of sheet: '" + sheetName + "' of id: '" + viewportId.ToString() + "'");
                            vp.ChangeTypeId(type.Id);
                            t.Commit();
                        }

                        return;
                    }
                }
        }

        //checks if a view with the specified name is currently open
        public bool ViewIsOpen(string name)
        {
            View v = GetView(name);

            if (v == null)
                return false;

            ElementId id = v.Id;

            //view does not exist and so therefore can not be open, return false
            if (id == null)
                return false;

            IList<UIView> openViews = uiDoc.GetOpenUIViews();

            foreach (UIView uiV in openViews)
            {
                if (uiV.ViewId.Equals(id))
                    return true;//we found the view and it is open
            }

            return false;
        }

        //checks if a sheet with the specified name is open
        public bool SheetIsOpen(string name)
        {
            View v = GetSheet(name);

            if (v == null)
                return false;

            ElementId id = v.Id;

            //sheet does not exist to begin with, so therefore can not be open, return false
            if (id == null)
                return false;

            IList<UIView> openViews = uiDoc.GetOpenUIViews();

            foreach (UIView uiV in openViews)
            {
                if (uiV.ViewId.Equals(id))
                    return true;//we found the sheet and it is open
            }

            return false;
        }

        //shows a message box using the given parameters
        public void ShowMessageBox(string dialogName, string id, string message, TaskDialogIcon icon, TaskDialogCommonButtons buttons)
        {
            //give an error message if we do not
            TaskDialog td = new TaskDialog(dialogName);
            td.Id = id;
            td.MainIcon = icon;
            td.CommonButtons = buttons;
            td.MainContent = message;
            td.Show();
        }

        //directly moves an element around a sheet, DOES NOT MOVE IT BETWEEN SHEETS
        public void MoveElement(ElementId id, double x, double y)
        {
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Moving element with id: '" + id + "' - X: '" + x + "' Y: '" + y);
                ElementTransformUtils.MoveElement(doc, id, new XYZ(x, y, 0));
                t.Commit();
            }
        }

        public void StartTransactions(string transactionName)
        {
            if (!transactionList.HasStarted())
            {
                //prepare the transaction group and start it
                transactionList.Start(transactionName);
            }
        }

        //commits ALL transactions since the class was created        
        public void CommitTransactions()
        {
            if (transactionList.HasStarted())
                transactionList.Assimilate();
        }

        //reverts ALL transactions since the class was created
        public void RevertTransactions()
        {
            if (transactionList.HasStarted())
                transactionList.RollBack();
        }

        //exports a given sheet as a DWG file usign the given options
        //if multiple sheets of same name found exports all
        public bool ExportSheetAsDWG(string sheetName, string filePath, DWGExportOptions options)
        {
            //make sure we have data to work with
            if (options == null || filePath == string.Empty)
                return false;

            //make sure we got sheets to work with
            ICollection<ElementId> coll = GetSheetIdList(sheetName);
            if (coll == null)
                return false;

            //count how many sheets we have gone through
            int count = 1;

            foreach (ElementId id in coll)
            {
                using (Transaction t = new Transaction(doc))
                {
                    //get the sheet name and check if we need to add a nubmer to the end
                    string name = sheetName;

                    if (count > 1)
                        name += "-" + count.ToString();

                    t.Start("exporting '" + name + "' as a DWG file.");

                    ICollection<ElementId> ele = new List<ElementId>();
                    ele.Add(id);//add the file to a list as required

                    //attempt to export the sheet and rollback if it fails
                    try
                    {
                        doc.Export(filePath, name, ele, options);
                    }
                    catch (Exception e)
                    {
                        t.RollBack();
                        return false;
                    }

                    t.Commit();
                }

                count++;
            }

            return true;
        }

        //returns the location that the project exists in a string
        public string GetProjectFileLocation()
        {
            try
            {
                string str = System.IO.Path.GetDirectoryName(doc.PathName);
                return str;
            }
            catch(Exception e)
            {
                return string.Empty;
            }
        }

        //returns the file name of the project
        public string GetProjectFileName()
        {
            try
            {
                string str = (System.IO.Path.GetFileNameWithoutExtension(doc.PathName));
                return str;
            }
            catch (Exception e)
            {
                return string.Empty;
            }
        }

        /////////////////////
        //private functions//
        /////////////////////

        //searches for a view with the given name and returns its id
        private ElementId GetViewId(string viewName)
        {
            //find the view with the given name that IS NOT a template or a drawing sheet
            IEnumerable<View> views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.Name.Equals(viewName)
                && !v.IsTemplate
                && v.ViewType != ViewType.DrawingSheet);

            //check if we foudn the view
            if (views != null)
            {
                View view = views.FirstOrDefault();

                if (view != null)
                    return view.Id;//found the view
            }

            //didnt find the id, return null
            return null;
        }

        //searches for a view with the given name and returns the view itself
        private View GetView(string viewName)
        {
            //find the view with the given name that IS NOT a template or a drawing sheet
            IEnumerable<View> views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.Name.Equals(viewName)
                && !v.IsTemplate
                && v.ViewType != ViewType.DrawingSheet);

            //check if we found the view
            if (views != null)
            {
                View view = views.FirstOrDefault();

                if (view != null)
                    return view;//found the view
            }

            //didnt find the id, return null
            return null;
        }

        //searches for and returns the id of a view template
        private ElementId GetViewTemplateId(string templateName)
        {
            //find a specified view template that is NOT a drawing sheet and IS a template
            IEnumerable<View> views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.Name.Equals(templateName)
                && v.IsTemplate
                && v.ViewType != ViewType.DrawingSheet);

            //check if we found the view template
            if (views != null)
            {
                View template = views.FirstOrDefault();

                if (template != null)
                    return template.Id;//found the template
            }
            //did not find specified template
            return null;
        }

        private ElementId GetViewFamilyType(string viewTypeName)
        {
            IEnumerable<ViewFamilyType> coll = new FilteredElementCollector(doc)
                 .OfClass(typeof(ViewFamilyType))
                 .Cast<ViewFamilyType>()
                 .Where(v => v.Name.Equals(viewTypeName));

            //check if we found the view family type
            if (coll != null)
            {
                ViewFamilyType vt = coll.FirstOrDefault();

                if (vt != null)
                    return vt.Id;//found the view family type
            }

            //did not find specified view family type
            return null;
        }

        private View GetSheet(string sheetName)
        {
            //find a specific drawing sheet
            IEnumerable<View> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.Name.Equals(sheetName)
                && v.ViewType == ViewType.DrawingSheet);

            //check if we found the sheet
            if (sheets != null)
            {
                View sheet = sheets.FirstOrDefault();

                if (sheet != null)
                    return sheet;//found the sheet
            }

            //didnt find the id, return null
            return null;
        }

        private string GetSheetNumber(string sheetName)
        {
            //find a specific drawing sheet
            IEnumerable<ViewSheet> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(v => v.Name.Equals(sheetName)
                && v.ViewType == ViewType.DrawingSheet);

            //check if we found the sheet
            if (sheets != null)
            {
                ViewSheet sheet = sheets.FirstOrDefault();

                if (sheet != null)
                    return sheet.SheetNumber;//found the sheet
            }

            //didnt find the id, return null
            return null;
        }

        private ViewSheet GetSheetAsViewSheet(string sheetName)
        {
            //find a specific drawing sheet
            IEnumerable<ViewSheet> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(v => v.Name.Equals(sheetName)
                && v.ViewType == ViewType.DrawingSheet);

            //check if we found the sheet
            if (sheets != null)
            {
                ViewSheet sheet = sheets.FirstOrDefault();

                if (sheet != null)
                    return sheet;//found the sheet
            }

            //didnt find the id, return null
            return null;
        }

        //finds and returns the id of the first found sheet with the given name
        private ElementId GetSheetId(string sheetName)
        {
            //find a specific drawing sheet
            IEnumerable<View> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.Name.Equals(sheetName)
                && v.ViewType == ViewType.DrawingSheet);

            //check if we found the sheet
            if (sheets != null)
            {
                View sheet = sheets.FirstOrDefault();

                if (sheet != null)
                    return sheet.Id;//found the sheet
            }

            //didnt find the id, return null
            return null;
        }

        //finds and returns the id of ALL sheets with the given name
        private ICollection<ElementId> GetSheetIdList(string sheetName)
        {
            //find a specific drawing sheet
            IEnumerable<ViewSheet> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .OrderBy(v => v.SheetNumber)//attempt to order this list by the sheet number
                .Where(v => v.Name.Equals(sheetName)
                && v.ViewType == ViewType.DrawingSheet);

            ICollection<ElementId> coll = new List<ElementId>();

            //check if we found the sheet
            if (sheets != null)
                foreach (ViewSheet v in sheets)
                    coll.Add(v.Id);//found a sheet with name    


            //didnt find the id, return null
            if (coll.Count >= 1)
                return coll;
            else
                return null;
        }

        private ElementId GetTitleBlockId(string titleBlock)
        {
            //find a specific title block id
            IEnumerable<FamilySymbol> tb = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .Cast<FamilySymbol>()
                .Where(v => v.FamilyName.Equals(titleBlock));

            //check if we found the title block
            if (tb != null)
                return tb.FirstOrDefault().Id;
            else
                return null;
        }

        //wipes all elements with ids from the list of given
        private bool Delete(List<ElementId> eID)
        {
            if (eID != null)//make sure we have been given elements to delete
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Deleting given views.");
                    doc.Delete(eID);
                    t.Commit();
                    return true;
                }
            else
                return false;
        }

        //wipes a specific element with the given id
        private bool Delete(ElementId eID)
        {
            if (eID != null)//make sure an element was given
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Deleting given view.");
                    doc.Delete(eID);
                    t.Commit();
                    return true;
                }
            else
                return false;
        }
    }

    //enum used with rotating views
    public enum RotationAngle
    {
        Left = 1,
        Down = 2,
        Right = 3
    }
}