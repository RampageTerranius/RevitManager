using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace RevitViewAndSheetManager
{
    public class RevitManager
    {
        private Document doc = null;
        private UIDocument uiDoc = null;
        private UIApplication uiApp = null;

        private TransactionGroup transactionList;

        public bool debugMode = false;

        public Document Doc { get => doc;}
        public UIDocument UiDoc { get => uiDoc; }
        public UIApplication UiApp { get => uiApp; }

        // TODO: error handling, have a list of errors and a way to get last error.

        /// <summary>
        /// Default Constructor.
        /// Requires the user to give the current ExternalCommandData so that it may prepare uiApp, uiDoc and doc.
        /// </summary>
        public RevitManager(ExternalCommandData commandData)
        {
            // Check if we have been given command data.
            if (commandData == null)
                throw new ArgumentNullException("commandData");

            // Preparing necessary variables.
            uiApp = commandData.Application;
            uiDoc = uiApp.ActiveUIDocument;
            doc = uiDoc.Document;

            transactionList = new TransactionGroup(doc);
        }

        /// <summary>
        /// creates a new sheet giving it a default sheet number and no title block.
        /// </summary>
        public bool CreateSheet(string sheetName)
        {
            return CreateSheet(sheetName, null, null);
        }

        /// <summary>
        /// Creates a new sheet giving it a default sheet number and the specified title block.
        /// </summary>
        public bool CreateSheet(string sheetName, string titleBlock)
        {
            return CreateSheet(sheetName, null, titleBlock);
        }

        /// <summary>
        /// Creates a new sheet giving it the specified sheet number and title block.
        /// </summary>
        public bool CreateSheet(string sheetName, string sheetNumber, string titleBlock)
        {
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Creating new ViewSheet.");

                try
                {
                    // Attempt to find the title blocks id.
                    ElementId tempTitleBlock = ElementId.InvalidElementId;

                    if (titleBlock != null)
                    {
                        tempTitleBlock = GetTitleBlockId(titleBlock);
                        if (tempTitleBlock == null)
                            throw new Exception("Unable to find specified Title Block: " + titleBlock);
                    }

                    // Create a sheet view.
                    ViewSheet viewSheet = ViewSheet.Create(doc, tempTitleBlock);
                    if (viewSheet == null)
                        throw new Exception("Failed to create new ViewSheet.");

                    // Name the sheet as desired.
                    viewSheet.Name = sheetName;

                    // If given a sheet number attempt to change it, other wise leave it at default.
                    if (sheetNumber != null)
                        viewSheet.SheetNumber = sheetNumber;

                    t.Commit();

                    return true;
                }
                catch
                {
                    // Something went wrong, let the user know
                    if (debugMode)
                        ShowMessageBox("Warning!",
                                       "ID_TaskDialog_Warning",
                                       "Unable to create a new ViewSheet by the name of " + sheetName + ", of number " + sheetNumber + ", of titleblock type " + titleBlock,
                                       TaskDialogIcon.TaskDialogIconWarning,
                                       TaskDialogCommonButtons.Close);
                
                    // As something has gone wrong we want to rollback to reverse any pending changes.
                    t.RollBack();

                    return false;
                }
            }
        }

        /// <summary>
        /// Wipe the first view with the given name.
        /// </summary>
        public bool DeleteView(string viewName)
        {
            ElementId v = GetViewId(viewName);

            if (v != null)
                return Delete(v);
            else
                return false;
        }

        /// <summary>
        /// Wipe the first sheet with the given name.
        /// </summary>
        public bool DeleteSheet(string sheetName)
        {
            if (sheetName != null)
                return Delete(GetSheetId(sheetName));
            else
                return false;
        }

        /// <summary>
        /// Moves the given view into a new family tree.
        /// </summary>
        public bool MoveView(string viewName, string newLocation)
        {
            View v = GetView(viewName);
            ElementId id = GetViewFamilyType(newLocation);

            // Make sure there is a view and a family to move it to.
            if (v != null && id != null)
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
            {   
                // Did not find a view or an element, show user an error .       
                if (debugMode)
                {
                    string msg = string.Empty;

                    // Preparing the error message.
                    int i = 0;
                    if (v == null)
                        i++;

                    if (id == null)
                        i = +2;

                    switch (i)
                    {
                        case 1:
                            // Unable to find view.
                            msg = "Unable to move view '" + viewName + "', could not find view.";
                            break;

                        case 2:
                            // Unable to find family.
                            msg = "Unable to move view '" + viewName + "', could not find family id.";
                            break;

                        case 3:
                            // Unable to find view AND template.
                            msg = "Unable to move view '" + viewName + "', could not find view AND family id.";
                            break;
                    }

                    // Show the message.
                    ShowMessageBox("Warning!",
                                   "ID_TaskDialog_Warning",
                                   msg,
                                   TaskDialogIcon.TaskDialogIconWarning,
                                   TaskDialogCommonButtons.Close);
                }

                return false;
            }
        }

        /// <summary>
        /// Rotates all entities on a given view in 90 degree angles.
        /// Accepts ONLY RotationAngle.
        /// </summary>
        public void RotateView(string viewName, RotationAngle rotation)
        {
            // Get the id of the view we need to rotate.
            View v = GetView(viewName);

            // Stop here if there is no view to work with.
            if (v == null)
                return;

            // Attempt to get the crop box directly.
            Element cropBox = doc.GetElement(GetCropBox(v));

            // Push the rotation to the main element rotation function.
            RotateElement(cropBox.Id, rotation);
        }

        /// <summary>
        /// Rotate an element by a specific angle.
        /// </summary>
        public bool RotateElement(ElementId id, double angle)
        {
            try
            {
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Attempting to rotate element " + id.IntegerValue.ToString() + " by " + angle.ToString());
                    // Find the element we are working with.
                    Element e = doc.GetElement(id);

                    // Get bounding box.
                    BoundingBoxXYZ bbox = e.get_BoundingBox(doc.ActiveView);

                    // Calculate the center of the bounding box.
                    XYZ center = 0.5 * (bbox.Max + bbox.Min);

                    // Create a axis to use from the center to the left.
                    Line axis = Line.CreateBound(center, center + XYZ.BasisZ);

                    // Attempt to rotate the element using the given parameters.
                    // As our axis is from center to left we need to rotate as negative to get the correct rotation direction.
                    ElementTransformUtils.RotateElement(doc, id, axis, -angle);

                    t.Commit();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Rotate an element by 90 degree angles.
        /// </summary>
        public bool RotateElement(ElementId id, RotationAngle angle)
        {
             return RotateElement(id, (DegreesToRadians(90) * (int)angle));
        }

        /// <summary>
        /// Finds and renames a view.
        /// </summary>
        public bool RenameView(string viewName, string newViewName)
        {
            // Get the view first
            View v = GetView(viewName);

            // Make sure we have a view to work with first.
            if (v != null)
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
            {   
                // Did not find a view, give the user an error.
                if (debugMode)                
                    ShowMessageBox("Warning!",
                                   "ID_TaskDialog_Warning",
                                   "Unable to rename view '" + viewName + "', view may not exist!",
                                   TaskDialogIcon.TaskDialogIconWarning,
                                   TaskDialogCommonButtons.Close);
                
                return false;                
            }
        }

        /// <summary>
        /// Duplicates the view naming it as given, duplicate will be in same location.
        /// </summary>
        public bool DuplicateViewByName(string viewName, string newViewName, ViewDuplicateOption duplicateOption)
        {
            // Get the view we will duplicate first.
            View v = GetView(viewName);

            // Make sure there is indeed a view.
            if (v != null)
            {
                // Duplicate the view.
                using (Transaction t = new Transaction(doc))
                {
                    // Duplicate the view first.
                    t.Start("Duplicating view.");
                    ElementId id = v.Duplicate(duplicateOption);

                    // Rename the view to the new name.
                    View tempView = doc.GetElement(id) as View;
                    tempView.Name = newViewName;

                    t.Commit();
                }

                return true;
            }
            else
            {
                // There was no view, give the user an error screen.
                if (debugMode)
                    ShowMessageBox("Warning!",
                                   "ID_TaskDialog_Warning",
                                   "Unable to duplicate view '" + viewName + "'.",
                                   TaskDialogIcon.TaskDialogIconWarning,
                                   TaskDialogCommonButtons.Close);

                return false;
            }
        }

        /// <summary>
        /// Changes the template of the given view.
        /// </summary>
        public bool ChangeViewTemplate(string viewName, string templateName)
        {
            View v = GetView(viewName);
            ElementId id = GetViewTemplateId(templateName);

            // Check that we have a view and a template to work with first.
            if (v != null && id != null)
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
            {   
                // There was no view or template, warn the user.
                // Create error box.
                if (debugMode)
                {                    
                    string msg = string.Empty;

                    // Preparing the error message.
                    int i = 0;
                    if (v == null)
                        i++;

                    if (id == null)
                        i = +2;

                    switch (i)
                    {
                        case 1:
                            // Unable to find view.
                            msg = "Unable to change template for view '" + viewName + "', the view may not exist.";
                            break;

                        case 2:
                            // Unable to find template.
                            msg = "Unable to change template for view '" + viewName + "', the template may not exist.";
                            break;

                        case 3:
                            // Unable to find view AND template.
                            msg = "Unable to change template for view '" + viewName + "', the view AND template may not exist.";
                            break;
                    }

                    // Show the message box.
                    ShowMessageBox("Warning!",
                                   "ID_TaskDialog_Warning",
                                   msg,
                                   TaskDialogIcon.TaskDialogIconWarning,
                                   TaskDialogCommonButtons.Close);                    
                }

                return false;
            }
        }

        /// <summary>
        /// Removes all independent tags from the view.
        /// </summary>
        public void RemoveAllIndependentTags(string viewName)
        {
            // Get the view we will be working with.
            ElementId v = GetViewId(viewName);

            // Check if we have a view to work with.
            if (v == null)
            {
                if (debugMode)
                    ShowMessageBox("Warning!",
                                   "ID_TaskDialog_Warning",
                                   "Unable to remove all independant tags from " + viewName + ", it may not exist.",
                                   TaskDialogIcon.TaskDialogIconWarning,
                                   TaskDialogCommonButtons.Close);

                return;
            }
            

            // Grab all independent tags owned by the view we are worknig with.
            IEnumerable<IndependentTag> coll = new FilteredElementCollector(doc)
                .OfClass(typeof(IndependentTag))
                .Cast<IndependentTag>()
                .Where(i => i.OwnerViewId.Equals(v));

            // Wipe all found tags.
            List<ElementId> coll2 = new List<ElementId>();

            foreach (IndependentTag i in coll)
                coll2.Add(i.Id);

            Delete(coll2);
        }

        /// <summary>
        /// Removes ALL dimensions from given view.
        /// </summary>
        public void RemoveAllDimensions(string viewName)
        {
            // Get the view to work with.
            ElementId v = GetViewId(viewName);

            // Check if we have a view to work with.
            if (v == null)
            {
                if (debugMode)
                    ShowMessageBox("Warning!",
                                   "ID_TaskDialog_Warning",
                                   "Unable to remove all dimensions from " + viewName + ", it may not exist.",
                                   TaskDialogIcon.TaskDialogIconWarning,
                                   TaskDialogCommonButtons.Close);
                return;
            }            

            // Get all the dimensions.
            IEnumerable<Dimension> coll = new FilteredElementCollector(doc)
                .OfClass(typeof(Dimension))
                .Cast<Dimension>()
                .Where(i => i.OwnerViewId.Equals(v));

            // Wipe all found dimensions.
            List<ElementId> coll2 = new List<ElementId>();

            foreach (Dimension i in coll)
                coll2.Add(i.Id);

            Delete(coll2);
        }

        /// <summary>
        /// Removes all dimensions of a specific type from a given view.
        /// </summary>
        public void RemoveAllDimensionsOfType(string viewName, DimensionStyleType dimType)
        {
            // Get the view to work with.
            ElementId v = GetViewId(viewName);

            // Check if we have a view to work with.
            if (v == null)
            {
                if (debugMode)
                    ShowMessageBox("Warning!",
                                   "ID_TaskDialog_Warning",
                                   "Unable to remove all independant tags of type " + dimType.ToString() + " from " + viewName + ", it may not exist.",
                                   TaskDialogIcon.TaskDialogIconWarning,
                                   TaskDialogCommonButtons.Close);

                return;
            }

            // Get all the dimensions we need to wipe.
            IEnumerable<Dimension> coll = new FilteredElementCollector(doc)
                .OfClass(typeof(Dimension))
                .Cast<Dimension>()
                .Where(i => i.OwnerViewId.Equals(v)
                && i.DimensionType.StyleType == dimType);

            // Wipe all found dimensions.
            List<ElementId> coll2 = new List<ElementId>();

            foreach (Dimension i in coll)
                coll2.Add(i.Id);

            Delete(coll2);
        }

        /// <summary>
        /// Removes ALL text notes from a given view.
        /// </summary>
        public void RemoveAllTextNotes(string viewName)
        {
            // Get the view to work with.
            ElementId v = GetViewId(viewName);

            // Check if we have a view to work with.
            if (v == null)  
            {
                if (debugMode)
                    ShowMessageBox("Warning!",
                                   "ID_TaskDialog_Warning",
                                   "Unable to remove all text notes from " + viewName + ", it may not exist.",
                                   TaskDialogIcon.TaskDialogIconWarning,
                                   TaskDialogCommonButtons.Close);
                return;
            }

            // Get all text notes.
            IEnumerable<TextNote> coll = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNote))
                .Cast<TextNote>()
                .Where(b => b.OwnerViewId == v);

            // Wipe all found text notes.
            List<ElementId> coll2 = new List<ElementId>();

            foreach (TextNote i in coll)
                coll2.Add(i.Id);

            Delete(coll2);
        }

        /// <summary>
        /// Wipe all generic annotation familys from the given view.
        /// </summary>
        public void RemoveAllAnnotations(string viewName)
        {
            // Make sure we have a view to work with.
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

            // Find all annotation symbols in the given view.
            IEnumerable<FamilyInstance> coll = new FilteredElementCollector(doc, v)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(b => b.Symbol.GetType().Name.Equals("AnnotationSymbolType"));

            // Wipe all found symbols.
            List<ElementId> coll2 = new List<ElementId>();

            foreach (FamilyInstance i in coll)
                coll2.Add(i.Id);

            Delete(coll2);
        }

        /// <summary>
        /// Removes all groups from the given view.
        /// </summary>
        public void RemoveAllGroups(string viewName)
        {
            // Make sure we have a view to work with.
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

            // Find all annotation symbols in the given view.
            IEnumerable<GroupType> coll = new FilteredElementCollector(doc, v)
                .OfClass(typeof(GroupType))
                .Cast<GroupType>();

            // Wipe all found symbols.
            List<ElementId> coll2 = new List<ElementId>();

            foreach (Element i in coll)
                coll2.Add(i.Id);

            Delete(coll2);
        }

        /// <summary>
        /// Adds the given view to the given sheet then returns the ElementId, places the view directly in the middle of the sheet.
        /// </summary>
        public ElementId AddViewToSheet(string sheetName, string viewName)
        {
            return AddViewToSheet(sheetName, viewName, null);
        }

        /// <summary>
        /// Adds the given view to the given sheet then returns the ElementId, places the view at the given coordiantes on the sheet.
        /// </summary>
        public ElementId AddViewToSheet(string sheetName, string viewName, XYZ xyz)
        {
            // Get the sheet and the view.
            ElementId s = GetSheetId(sheetName);
            ElementId v = GetViewId(viewName);

            // Check if we have elements to work with.
            if (s == null || v == null)
                return null;

            // Check if we have been given a specific location to place the viewport.
            if (xyz == null)
                xyz = new XYZ(0, 0, 0);

            // Create the viewport.
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Creating ViewPort.");
                Viewport vp = Viewport.Create(doc, s, v, xyz);
                t.Commit();

                return vp.Id;
            }
        }

        /// <summary>
        /// Returns the ID of the crop box of a given view.
        /// </summary>
        //With thanks to Jeremy Tammik/Konrads Samulis
        //https://thebuildingcoder.typepad.com/blog/2018/02/efficiently-retrieve-crop-box-for-given-view.html
        public ElementId GetCropBox(View view)
        {        
            // Check that we actually have a view to work with.
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

        /// <summary>
        /// Gets crop box using the name of the view instead of the view its self.
        /// </summary>
        public ElementId GetCropBox(string viewName)
        {
            // Does not need to check if we have a view as GetCropBox(View view) will check for us.
            View v = GetView(viewName);

            return GetCropBox(v);
        }

        /// <summary>
        /// Load all views and return them in a viewset. skips all templates, use GetAllViewTemplates for this.
        /// </summary>
        public ViewSet GetAllViews()
        {
            // Find all views.
            IEnumerable<View> views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(b => !b.IsTemplate
                && b.ViewType != ViewType.DrawingSheet);

            // Add all the views into a ViewSet.
            ViewSet allViews = new ViewSet();

            foreach (View v in views)
                allViews.Insert(v);

            // Return the ViewSet.
            return allViews;
        }

        /// <summary>
        /// Specifically checks for the sheet number
        /// </summary>
        public XYZ GetSheetDimensions(string sheetName)
        {
            // Find a specific title block.
            View sheet = GetSheet(sheetName);


            // Retrieve the title block instances. 
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

                // Check if we are on the correct sheet.
                if (sheet_number != sheetName)
                    continue;

                // We are on the correct sheet, get the width and height and return it.

                p = e.get_Parameter(
                  BuiltInParameter.SHEET_WIDTH);

                double width = p.AsDouble();

                p = e.get_Parameter(
                  BuiltInParameter.SHEET_HEIGHT);

                double height = p.AsDouble();

                XYZ xyz = new XYZ(width, height, 0);

                return xyz;
            }

            // Didnt find the sheet, return null.
            return null;
        }

        /// <summary>
        /// Changes a viewport on the given sheets viewporttype into another type.
        /// </summary>
        public void ChangeViewPortType(string sheetName, ElementId viewportId, string typeName)
        {
            ElementId i = GetSheetId(sheetName);

            // Make sure we have a sheet to work with.
            if (i == null)
                return;

            // Get the viewport we are after.
            IEnumerable<Viewport> viewports = new FilteredElementCollector(doc, i)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .Where(v => v.Id.Equals(viewportId));

            Viewport vp = viewports.FirstOrDefault();

            // Make sure we have a viewport to work with.
            if (vp == null)
                return;

            // Look for the view we need to change.
            if (!vp.Name.Equals(typeName))
                foreach (ElementId id in vp.GetValidTypes())
                {
                    ElementType type = doc.GetElement(id) as ElementType;
                    if (type.Name.Equals(typeName))
                    {
                        // Found the view. Change its type and stop the method.
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

        /// <summary>
        /// Checks if a view with the specified name is currently open
        /// </summary>
        public bool ViewIsOpen(string name)
        {
            View v = GetView(name);

            // Make sure we have a view to work with, if we dont it doesnt exist and therefore can not be open.
            if (v == null)
                return false;

            // Get a list of all open views and check if the oen we are checking for is open.
            IList<UIView> openViews = uiDoc.GetOpenUIViews();

            // Check if the view is open.
            foreach (UIView uiV in openViews)            
                if (uiV.ViewId.Equals(v.Id))
                    return true;

            return false;
        }

        /// <summary>
        /// Checks if a sheet with the specified name is open.
        /// </summary>
        public bool SheetIsOpen(string name)
        {
            View v = GetSheet(name);

            // Make sure we have a sheet to work with, if we dont it doesnt exist and therefore cane not be open.
            if (v == null)
                return false;

            IList<UIView> openViews = uiDoc.GetOpenUIViews();

            // Check if the sheet is open.
            foreach (UIView uiV in openViews)            
                if (uiV.ViewId.Equals(v.Id))
                    return true;         

            return false;
        }

        /// <summary>
        /// Shows a message box using the given parameters
        /// </summary>
        public TaskDialogResult ShowMessageBox(string dialogName, string id, string message, TaskDialogIcon icon, TaskDialogCommonButtons buttons)
        {
            TaskDialog td = new TaskDialog(dialogName);
            td.Id = id;
            td.MainIcon = icon;
            td.CommonButtons = buttons;
            td.MainContent = message;
            return td.Show();
        }

        /// <summary>
        /// Shows a defaulted message box using the given strings.
        /// The box will have no icon and will be a single ok button.
        /// </summary>
        public TaskDialogResult ShowMessageBox(string dialogName, string message)
        {
            TaskDialog td = new TaskDialog(dialogName);
            td.Id = dialogName;
            td.MainIcon = TaskDialogIcon.TaskDialogIconNone;
            td.CommonButtons = TaskDialogCommonButtons.Ok;
            td.MainContent = message;
            return td.Show();
        }

        /// <summary>
        /// Directly moves an element around a sheet, DOES NOT MOVE IT BETWEEN SHEETS.
        /// </summary>

        public void MoveElement(ElementId id, double x, double y)
        {
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Moving element with id: '" + id + "' - X: '" + x + "' Y: '" + y);
                ElementTransformUtils.MoveElement(doc, id, new XYZ(x, y, 0));
                t.Commit();
            }
        }

        /// <summary>
        /// Starts the main transaction group.
        /// </summary>
        public bool StartTransactions(string transactionName)
        {
            if (!transactionList.HasStarted())
                if (transactionList.Start(transactionName) == TransactionStatus.Started)
                    return true;

            return false;
        }

        /// <summary>
        /// Commits ALL transactions since StartTransactions was run under a single transaction
        /// </summary>
        public bool CommitTransactions()
        {
            if (transactionList.HasStarted())
                if (transactionList.Assimilate() == TransactionStatus.Committed)
                    return true;

            return false;
        }

        /// <summary>
        /// Reverts ALL transactions since StartTransactions was run.
        /// </summary>
        public bool RevertTransactions()
        {
            if (transactionList.HasStarted())
                if (transactionList.RollBack() == TransactionStatus.RolledBack)
                    return true;

            return false;
        }

        /// <summary>
        /// Exports a given sheet as a DWG file using the given options.
        /// If multiple sheets of same name found exports each one with a number at the end donating its position.
        /// </summary>
        public bool ExportSheetAsDWG(string sheetName, string filePath, DWGExportOptions options)
        {
            // Make sure we have data to work with.
            if (options == null || filePath == string.Empty)
                return false;

            // Make sure we got sheets to work with.
            ICollection<ElementId> coll = GetSheetIdList(sheetName);
            if (coll == null)
                return false;

            // Count how many sheets we have gone through.
            int count = 1;

            foreach (ElementId id in coll)
            {
                using (Transaction t = new Transaction(doc))
                {
                    // Get the sheet name and check if we need to add a number to the end.
                    string name = sheetName;

                    if (count > 1)
                        name += "-" + count.ToString();

                    t.Start("exporting '" + name + "' as a DWG file.");

                    // Add the file to a list as required.
                    ICollection<ElementId> ele = new List<ElementId>();
                    ele.Add(id);

                    // Attempt to export the sheet and rollback if it fails.
                    try
                    {
                        doc.Export(filePath, name, ele, options);
                    }
                    catch
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

        /// <summary>
        /// Returns the location that the project exists as a string.
        /// </summary>
        public string GetProjectFileLocation()
        {
            try
            {
                return System.IO.Path.GetDirectoryName(doc.PathName);
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Returns the file name of the project.
        /// </summary>
        public string GetProjectFileName()
        {
            try
            {
                return System.IO.Path.GetFileNameWithoutExtension(doc.PathName);
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Returns the location of the local user Addins folder for the in-use version of revit.
        /// Addins in this folder are user specific.
        /// </summary>
        public string GetAddinsAppDataLocation()
        {
            try
            {
                string appDataloc = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

                return appDataloc + @"\Autodesk\Revit\Addins\" + UiApp.Application.VersionNumber;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Returns the location of the common Addins folder for the in-use version of revit.
        /// Addins in this folder are non-user specific.
        /// </summary>
        public string GetAddinsProgramDataLocation()
        {
            try
            {
                string appDataloc = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);

                return appDataloc + @"\Autodesk\Revit\Addins\" + UiApp.Application.VersionNumber;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Returns the location of the local user Revit settings folder for the in-use version of revit.
        /// </summary>
        public string GetRevitAppDataLocation()
        {
            try
            {
                string appDataloc = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

                return appDataloc + @"\Autodesk\Revit";
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Returns the location of the common Revit folder for the in-use version of revit.
        /// </summary
        public string GetRevitProgramDataLocation()
        {
            try
            {
                string appDataloc = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);

                return appDataloc + @"\Autodesk\Revit";
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Ask the user to select a point and return the selected point.
        /// </summary>
        public XYZ PickPoint(string message)
        {
            try
            {
                return uiDoc.Selection.PickPoint(message);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Ask the user to select a point and return the selected point.
        /// </summary>
        public XYZ PickPoint(ObjectSnapTypes type, string message)
        {
            try
            {
                return uiDoc.Selection.PickPoint(type, message);
            }
            catch
            {
                return null;
            }            
        }

        /// <summary>
        /// Ask the user to select a object and return the selected object.
        /// </summary>
        public Reference PickObject(ObjectType ot, SelectionFilter sf, string message)
        {
            ISelectionFilter filter = null;

            switch(sf)
            {
                case SelectionFilter.Building:
                    filter = new BuildingSelectionFilter();
                    break;

                case SelectionFilter.TextNote:
                    filter = new TextNoteSelectionFilter();
                    break;
            }

            return PickObject(ot, filter, message);
        }

        //// <summary>
        /// Ask the user to select a object and return the selected object.
        /// </summary>
        public Reference PickObject(ObjectType ot, string message)
        {
            return PickObject(ot, (ISelectionFilter)null, message);
        }

        /// <summary>
        /// Ask the user to select a object and return the selected object.
        /// </summary>
        public Reference PickObject(ObjectType ot, ISelectionFilter sf, string message)
        {
            if (ot == ObjectType.Nothing || message == string.Empty)
                return null;

            try
            {
                if (sf == null)
                    return uiDoc.Selection.PickObject(ot, message);
                else
                    return uiDoc.Selection.PickObject(ot, sf, message);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Asks the user to pick objects and returns all picked objects as IList<Reference>.
        /// </summary>
        public IList<Reference> PickObjects(ObjectType ot, SelectionFilter sf, string message)
        {
            ISelectionFilter filter = null;

            switch (sf)
            {
                case SelectionFilter.Building:
                    filter = new BuildingSelectionFilter();
                    break;

                case SelectionFilter.TextNote:
                    filter = new TextNoteSelectionFilter();
                    break;
            }

            return PickObjects(ot, filter, message);
        }

        /// <summary>
        /// Asks the user to pick objects and returns all picked objects as IList<Reference>.
        /// </summary>
        public IList<Reference> PickObjects(ObjectType ot, string message)
        {
            return PickObjects(ot, (ISelectionFilter)null, message);
        }

        /// <summary>
        /// Asks the user to pick objects and returns all picked objects as IList<Reference>.
        /// </summary>
        public IList<Reference> PickObjects(ObjectType ot, ISelectionFilter sf, string message)
        {
            if (ot == ObjectType.Nothing || message == string.Empty)
                return null;

            try
            {
                if (sf == null)
                    return uiDoc.Selection.PickObjects(ot, message);
                else
                    return uiDoc.Selection.PickObjects(ot, sf, message);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Return a list of every element between two given points.
        /// </summary>
        public IEnumerable<Element> GetAllElementsBetweenTwoPoints(XYZ first, XYZ last)
        {
            // Make sure we have data to work with.
            if (first == null || last == null)
                return null;

            IEnumerable<Element> eList = new List<Element>();

            double minX, minY, minZ, maxX, maxY, maxZ;

            if (first.X <= last.X)
            {
                minX = first.X;
                maxX = last.X;
            } else
            {
                maxX = first.X;
                minX = last.X;
            }

            if (first.Y <= last.Y)
            {
                minY = first.Y;
                maxY = last.Y;
            }
            else
            {
                maxY = first.Y;
                minY = last.Y;
            }                          

            if (first.Z <= last.Z)
            {
                minZ = first.Z;
                maxZ = last.Z;
            }
            else
            {
                maxZ = first.Z;
                minZ = last.Z;
            }

            XYZ min = new XYZ(minX, minY, minZ);
            XYZ max = new XYZ(maxX, maxY, maxZ);

            Outline ol = new Outline(min, max);

            BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(ol);

            FilteredElementCollector coll = new FilteredElementCollector(doc, doc.ActiveView.Id);
            eList = coll.WherePasses(filter).
                            Where(e => e.get_Geometry(new Options { ComputeReferences = true }) != null);

            return eList;
        }

        /// <summary>
        /// Opens a view in the current project and makes it active in the ui.
        /// </summary>
        public void OpenView(string viewName)
        {
            // Get the sheet.
            View view = GetView(viewName);

            // Make sure we have data to work with.
            if (view == null)
                return;

            // Switch the active view and make sure to refresh the screen.
            uiDoc.ActiveView = view;
            uiDoc.RefreshActiveView();
        }

        /// <summary>
        /// Opens a sheet in the current project and makes it active in the ui.
        /// </summary>
        public void OpenSheet(string sheetName)
        {
            // Get the sheet.
            View view = GetSheet(sheetName);

            // Make sure we have data to work with.
            if (view == null)
                return;

            // Switch the active view and make sure to refresh the screen.
            uiDoc.ActiveView = view;
            uiDoc.RefreshActiveView();
        }

        /// <summary>
        /// Closes an open view with the given name, will not close if it is the ONLY open doc.
        /// </summary>
        public void CloseActiveTab(string tabName)
        {
            // Get a list of all active ui views.
            IList<UIView> viewList = uiDoc.GetOpenUIViews();

            // Make sure we have data to work with.
            if (viewList == null)
                return;

            // MUST have at least 2 tabs open as we can not close the last tab.
            if (viewList.Count <= 1)
                return;

            // Iterate through all uiviews and attempt to close the first one found with the given name.
            foreach(UIView uiView in viewList)
            {
                try
                {
                    // Convert the tab into a view so we can check its name.
                    View view = doc.GetElement(uiView.ViewId) as View;

                    //check if the tab has the name we are looking for
                    if (view.Name == tabName)
                    {
                        // Close it and quit the foreach statement.
                        uiView.Close();
                        break;
                    }
                }
                catch
                {
                    // Exception caught, try next tab.
                    continue;
                }
            }
        }

        /// <summary>
        /// Creates a WinForm asking the user to enter data into the supplied text box.
        /// This function does not check the returning string and returns it as was exactly added, it is up to the coder to error check this data once it is returned.
        /// </summary>
        public string GetUserInput(string windowName, string labelText)
        {
            string tempstr = "";

            // Prepare a new form.
            System.Windows.Forms.Form form = new System.Windows.Forms.Form();

            // Set some basic settings.
            form.Size = new System.Drawing.Size(400, 162);
            form.Text = windowName;
            form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            form.MinimizeBox = false;
            form.MaximizeBox = false;

            // Create the buttons.
            System.Windows.Forms.Button bOk = new System.Windows.Forms.Button();
            System.Windows.Forms.Button bCancel = new System.Windows.Forms.Button();

            // Create the label.
            System.Windows.Forms.Label lLabel = new System.Windows.Forms.Label();

            // Create the text box.
            System.Windows.Forms.TextBox tbTextBox = new System.Windows.Forms.TextBox();

            // Add our controls.
            form.Controls.Add(bOk);
            form.Controls.Add(bCancel);
            form.Controls.Add(lLabel);
            form.Controls.Add(tbTextBox);

            // Setup the controls as needed.
            bOk.Text = "Ok";
            bOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            bOk.Location = new System.Drawing.Point(10, 90);

            bCancel.Text = "Cancel";
            bCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            bCancel.Location = new System.Drawing.Point(95, 90);

            lLabel.Text = labelText;
            lLabel.Location = new System.Drawing.Point(10, 10);
            lLabel.Size = new System.Drawing.Size(380, 46);

            tbTextBox.Text = "";
            tbTextBox.Location = new System.Drawing.Point(10, 56);
            tbTextBox.Size = new System.Drawing.Size(370, 23);

            // Setup the enter and exc button hotkeysfor the form.
            form.AcceptButton = bOk;
            form.CancelButton = bCancel;

            // Check if we should grab the data from the textbox.
            if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)            
                tempstr = tbTextBox.Text;            

            return tempstr;
        }

        /// <summary>
        /// Returns the element in the document of the given reference.
        /// </summary>
        public Element GetElementViaReference(Reference r)
        {
            return doc.GetElement(r.ElementId);
        }

        /// <summary>
        /// Returns the given degrees as radians.
        /// </summary>
        public double DegreesToRadians(double degrees)
        {
            return (degrees * Math.PI) / 180;
        }

        /// <summary>
        /// Returns the given radians as degrees.
        /// </summary>
        public double RadiansToDegrees(double radians)
        {
            return (radians * 180) / Math.PI;
        }

        /// <summary>
        /// Returns the state of the given room.
        /// </summary>
        //with thanks to jeremy tammik
        //https://thebuildingcoder.typepad.com/blog/2016/04/how-to-distinguish-redundant-rooms.html
        public RoomState DistinguishRoom(Autodesk.Revit.DB.Architecture.Room room)
        {
            // Check if the room is Placed.
            if (room.Area > 0)
                return RoomState.Placed;
            // If not check if its NotPlaced.
            else if (room.Location == null)
                return RoomState.NotPlaced;
            // Other wise it must be Redundant or NotEnclosed.
            else 
            {

                SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions();

                IList<IList<BoundarySegment>> segs
                  = room.GetBoundarySegments(opt);

                if (segs == null)
                    return RoomState.NotEnclosed;
                else if (segs.Count == 0)
                    return RoomState.Redundant;
            }

            // We didnt find the state.
            return RoomState.Unknown;
        }

        /// <summary>
        /// Retuns a parameter from the given familyinstance with the given name.
        /// </summary>
        public Parameter GetInstanceParameter(FamilyInstance familyInstance, string parameterName)
        {
            // Attempt to get the parameter data.
            if (familyInstance != null)
                return GetElementParameter(doc.GetElement(familyInstance.Id), parameterName); ;

            // Did not find a parameter with the given name.
            return null;
        }

        /// <summary>
        /// Retuns the value of a parameter from the given familyinstance with the given name.
        /// </summary>
        public string GetInstanceParameterValue(FamilyInstance familyInstance, string parameterName)
        {
            // Attempt to get the parameter data.
            if (familyInstance != null)
                foreach (Parameter p in familyInstance.Parameters)
                    if (p.Definition.Name == parameterName)
                        if (p.HasValue)
                            return p.AsValueString();
                        else
                            return string.Empty;

            // Did not find a parameter with the given name.
            return string.Empty;
        }

        /// <summary>
        /// Changes the type on a given family instance.
        /// Searches for the new type, slower then directly giving it the FamilySymbol.
        /// </summary>
        public bool ChangeInstanceType(FamilyInstance familyInstance, string newFamily, string newType)
        {
            // Get the id of the type we will be changing to first.
            IEnumerable<FamilySymbol> coll = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(s => s.FamilyName.Equals(newFamily)
                && s.Name.Equals(newType));

            if (coll.Count() == 0)
                return false;

            ElementId type = coll.FirstOrDefault().Id;

            if (!familyInstance.IsValidType(type))
                return false;

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Change instance Type of " + familyInstance.Name + " to " + newFamily + ":" + newType);
                familyInstance.ChangeTypeId(type);
                t.Commit();
            }            

            return true;
        }

        /// <summary>
        /// Changes the type on a given family instance.
        /// </summary>
        public bool ChangeInstanceType(FamilyInstance familyInstance, FamilySymbol newType)
        {
            if (familyInstance == null || newType == null)
                return false;

            if (!familyInstance.IsValidType(newType.Id))
                return false;

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Change instance Type of " + familyInstance.Name + " to " + newType.Name + ":" + newType);
                familyInstance.ChangeTypeId(newType.Id);
                t.Commit();
            }

            return true;
        }

        /// <summary>
        /// Changes the type on a given family instance.
        /// </summary>
        public bool ChangeElementType(ElementId eId, FamilySymbol newType)
        {
            if (eId == null || newType == null)
                return false;

            Element e = doc.GetElement(eId);

            if (!e.IsValidType(newType.Id))
                return false;

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Change instance Type of " + e.Name + " to " + newType.Name + ":" + newType);
                e.ChangeTypeId(newType.Id);
                t.Commit();
            }

            return true;
        }

        /// <summary>
        /// Changes the family of a given instance.
        /// Searches for the new family by name, slower then directly giving it the FamilySymbol.
        /// </summary>
        public bool ChangeFamily(FamilyInstance familyInstance, string newFamily)
        {
            // Get the id of the type we will be changing to first.
            IEnumerable<FamilySymbol> coll = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(s => s.FamilyName.Equals(newFamily));

            if (coll.Count() == 0)
                return false;

            ElementId type = coll.FirstOrDefault().Id;

            // Make sure the new type is valid for this instance.     
            if (!familyInstance.IsValidType(type))
                return false;

            // Update family instance to the new type   
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Change family Type of " + familyInstance + " to " + newFamily);
                familyInstance.ChangeTypeId(type);

                t.Commit();
            }            

            return true;
        }

        /// <summary>
        /// Changes the family of a given instance.
        /// </summary>
        public bool ChangeFamily(FamilyInstance familyInstance, FamilySymbol newFamily)
        {
            if (familyInstance == null || newFamily == null)
                return false;

            // Make sure the new type is valid for this instance.     
            if (!familyInstance.IsValidType(newFamily.Id))
                return false;

            // Update family instance to the new type   
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Change family Type of " + familyInstance + " to " + newFamily);
                familyInstance.ChangeTypeId(newFamily.Id);

                t.Commit();
            }

            return true;
        }

        /// <summary>
        /// Changes the family of a given element.
        /// </summary>
        public bool ChangeFamily(ElementId instance, string newFamily)
        {
            // Get the id of the type we will be changing to first.
            IEnumerable<FamilySymbol> coll = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(s => s.FamilyName.Equals(newFamily));

            if (coll.Count() == 0)
                return false;

            ElementId type = coll.FirstOrDefault().Id;

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Change family of element " + instance.ToString() + " to " + newFamily);

                try
                {                    
                    FamilyInstance temp = doc.GetElement(instance) as FamilyInstance;

                    temp.ChangeTypeId(type);
                }
                catch
                {
                    t.RollBack();
                    return false;
                }                

                t.Commit();
            }

            return true;
        }

        /// <summary>
        /// Changes the parameters on a given family symbol. 
        /// </summary>
        public bool ChangeFamilyTypeParameter(FamilySymbol familySymbol, string parameterName, string newParameter)
        {
            foreach (Parameter p in familySymbol.Parameters)
                if (p.Definition.Name.ToLower().Equals(parameterName.ToLower()))
                    using (Transaction t = new Transaction(doc))
                    {
                        bool result = false;

                        t.Start("Edit Parameter");

                        switch (p.StorageType)
                        {
                            case StorageType.Integer:
                                int tempi;
                                int.TryParse(newParameter, out tempi);
                                result = p.Set(tempi);
                                break;

                            case StorageType.Double:
                                double tempd;
                                double.TryParse(newParameter, out tempd);
                                result = p.Set(tempd);
                                break;

                            case StorageType.ElementId:
                                Element tempele = doc.GetElement(newParameter);

                                if (tempele != null)
                                    result = p.Set(tempele.Id);
                                break;

                            case StorageType.String:
                                result = p.Set(newParameter);
                                break;
                        }

                        if (result)
                            t.Commit();
                        else
                            t.RollBack();

                        return result;
                    }            

            return false;
        }

        /// <summary>
        /// Changes the parameters on a given family.
        /// </summary>
        public bool ChangeFamilyInstanceParameter(FamilyInstance familyInstance, string parameterName, string newParameter)
        {
            foreach (Parameter p in familyInstance.Parameters)
                if (p.Definition.Name.ToLower().Equals(parameterName.ToLower()))
                    using (Transaction t = new Transaction(doc))
                    {
                        bool result = false;

                        t.Start("Edit Parameter");

                        switch (p.StorageType)
                        {
                            case StorageType.Integer:
                                int tempi;
                                int.TryParse(newParameter, out tempi);
                                result = p.Set(tempi);
                                break;

                            case StorageType.Double:
                                double tempd;
                                double.TryParse(newParameter, out tempd);
                                result = p.Set(tempd);
                                break;

                            case StorageType.ElementId:
                                Element tempele = doc.GetElement(newParameter);

                                if (tempele != null)
                                    result = p.Set(tempele.Id);
                                break;

                            case StorageType.String:
                                result = p.Set(newParameter);
                                break;
                        }

                        if (result)
                            t.Commit();
                        else
                            t.RollBack();

                        return result;
                    }
            

            return false;
        }

        /// <summary>
        /// Changes the parameters on a given element.
        /// </summary>
        public bool ChangeElementParameter(ElementId eId, string parameterName, string newParameter)
        {            
            FamilyInstance familyInstance;

            //attempt to turn the given elementid into a family instance
            try
            {
                familyInstance = doc.GetElement(eId) as FamilyInstance;
            }
            catch
            {
                return false;
            }

            foreach (Parameter p in familyInstance.Parameters)
                if (p.Definition.Name.ToLower().Equals(parameterName.ToLower()))
                    using (Transaction t = new Transaction(doc))
                    {
                        bool result = false;

                        t.Start("Edit Parameter");

                        switch (p.StorageType)
                        {
                            case StorageType.Integer:
                                int tempi;
                                int.TryParse(newParameter, out tempi);
                                result = p.Set(tempi);
                                break;

                            case StorageType.Double:
                                double tempd;
                                double.TryParse(newParameter, out tempd);
                                result = p.Set(tempd);
                                break;
                                
                            case StorageType.ElementId:
                                Element tempele = doc.GetElement(newParameter);

                                if (tempele != null)
                                    result = p.Set(tempele.Id);
                                break;

                            case StorageType.String:
                                result = p.Set(newParameter);
                                break;
                        }

                        if (result)
                            t.Commit();
                        else
                            t.RollBack();

                        return result;
                    }
            

            return false;
        }

        /// <summary>
        /// Returns an array of FamilyInstance where every member is a part of the given family.
        /// </summary>
        public FamilyInstance[] GetAllFamilyInstancesOfFamily(string familyName)
        {
            IEnumerable<FamilyInstance> coll = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(s => s.Symbol.FamilyName.ToLower().Equals(familyName.ToLower()));

            if (coll.Count() == 0)
                return null;

            return coll.ToArray();
        }

        /// <summary>
        /// Returns the FamilySymbol of a family type in the given family.
        /// </summary>
        public FamilySymbol GetFamilyType(string familyName, string familyType)
        {
            // Get the id of the type we will be changing to first.
            IEnumerable<FamilySymbol> coll = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(s => s.FamilyName.ToLower().Equals(familyName.ToLower())
                && s.Name.ToLower().Equals(familyType.ToLower()));

            if (coll.Count() == 0)
                return null;

            return coll.FirstOrDefault();
        }

        /// <summary>
        /// Finds a material by the given name and returns the ElementId
        /// </summary>
        public ElementId GetMaterialElementID(string materialName)
        {
            IEnumerable<Material> coll = new FilteredElementCollector(doc)
            .OfClass(typeof(Material))
            .Cast<Material>()
            .Where(s => s.Name == materialName);

            if (coll != null && coll.Count() > 0)
                return coll.FirstOrDefault().Id;

            else return null;
        }

        /// <summary>
        /// Retuns a parameter from the given FamilySymbol with the given name.
        /// </summary>
        public Parameter GetFamilyTypeParameter(FamilySymbol familySymbol, string parameterName)
        {
            // Attempt to get the parameter data.
            if (familySymbol != null)
                return GetElementParameter(doc.GetElement(familySymbol.Id), parameterName);

            // Did not find a parameter with the given name.
            return null;
        }

        /// <summary>
        /// Retuns a parameter from the given element with the given name.
        /// </summary>
        public Parameter GetElementParameter(Element ele, string parameterName)
        {
            // Attempt to get the parameter data.
            if (ele != null)
                foreach (Parameter p in ele.Parameters)
                    if (p.Definition.Name.ToLower().Equals(parameterName.ToLower()))
                        return p;

            // Did not find a parameter with the given name.
            return null;
        }

        /// <summary>
        /// Finds and returns the Family of the given name.
        /// </summary>
        public Family GetFamily(string familyName)
        {
            IEnumerable<Family> coll = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Where(s => s.Name.Equals(familyName));

            if (coll.Count() == 0)
                return null;

            return coll.FirstOrDefault();
        }

        /// <summary>
        /// Wipes all elements from the document with ids from the given list.
        /// </summary>
        public bool Delete(List<ElementId> eID)
        {
            // Make sure we have been given elements to delete.
            if (eID != null)
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Deleting given element(s).");
                    doc.Delete(eID);
                    t.Commit();

                    return true;
                }
            else
                return false;
        }

        /// <summary>
        /// Wipes a specific element with the given id.
        /// </summary>
        public bool Delete(ElementId eID)
        {
            // Make sure an element was given.
            if (eID != null)
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Deleting given element.");
                    doc.Delete(eID);
                    t.Commit();

                    return true;
                }
            else
                return false;
        }

        /// <summary>
        /// Wrapper class for converting IntPtr to IWin32Window.        
        /// </summary>
        // code thanks to Jeremy Tammik
        // https://thebuildingcoder.typepad.com/blog/2012/05/the-schedule-api-and-access-to-schedule-data.html
        public class JtWindowHandle : System.Windows.Forms.IWin32Window
        {
            IntPtr _hwnd;

            public JtWindowHandle(IntPtr h)
            {
                System.Diagnostics.Debug.Assert(IntPtr.Zero != h,
                  "expected non-null window handle");

                _hwnd = h;
            }

            public IntPtr Handle
            {
                get
                {
                    return _hwnd;
                }
            }
        }

        /// <summary>
        /// Searches for a view with the given name and returns its id.
        /// </summary>
        private ElementId GetViewId(string viewName)
        {
            // Find the view with the given name that IS NOT a template or a drawing sheet.
            IEnumerable<View> views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.Name.Equals(viewName)
                && !v.IsTemplate
                && v.ViewType != ViewType.DrawingSheet);

            // Check if we found the view.
            if (views != null)
            {
                View view = views.FirstOrDefault();

                if (view != null)
                    return view.Id;
            }

            // Didnt find the id.
            return null;
        }

        /// <summary>
        /// Searches for a view with the given name and returns the view itself.
        /// </summary>
        private View GetView(string viewName)
        {
            // Find the view with the given name that IS NOT a template or a drawing sheet.
            IEnumerable<View> views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.Name.Equals(viewName)
                && !v.IsTemplate
                && v.ViewType != ViewType.DrawingSheet);

            // Check if we found the view.
            if (views != null)
            {
                View view = views.FirstOrDefault();

                if (view != null)
                    return view;
            }

            // Didnt find the id, return null.
            return null;
        }

        /// <summary>
        /// Searches for and returns the id of a view template.
        /// </summary>
        private ElementId GetViewTemplateId(string templateName)
        {
            // Find a specified view template that is NOT a drawing sheet and IS a template.
            IEnumerable<View> views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.Name.Equals(templateName)
                && v.IsTemplate
                && v.ViewType != ViewType.DrawingSheet);

            // Check if we found the view template.
            if (views != null)
            {
                View template = views.FirstOrDefault();

                if (template != null)
                    return template.Id;
            }

            // Did not find specified template.
            return null;
        }

        /// <summary>
        /// Gets the ID of the ViewFamiliyType of a view.
        /// </summary>
        private ElementId GetViewFamilyType(string viewTypeName)
        {
            IEnumerable<ViewFamilyType> coll = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .Where(v => v.Name.Equals(viewTypeName));

            // Check if we found the view family type.
            if (coll != null)
            {
                ViewFamilyType vt = coll.FirstOrDefault();

                if (vt != null)
                    return vt.Id;
            }

            // Did not find specified view family type.
            return null;
        }

        /// <summary>
        /// Gets the first view of the given name.
        /// </summary>
        private View GetSheet(string sheetName)
        {
            // Find a specific drawing sheet.
            IEnumerable<View> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.Name.Equals(sheetName)
                && v.ViewType == ViewType.DrawingSheet);

            // Check if we found the sheet.
            if (sheets != null)
            {
                View sheet = sheets.FirstOrDefault();

                if (sheet != null)
                    return sheet;
            }

            // Didnt find the id, return null.
            return null;
        }

        /// <summary>
        /// Gets the sheet number of the first view with the given name.
        /// </summary>
        private string GetSheetNumber(string sheetName)
        {
            // Find a specific drawing sheet.
            IEnumerable<ViewSheet> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(v => v.Name.Equals(sheetName)
                && v.ViewType == ViewType.DrawingSheet);

            // Check if we found the sheet.
            if (sheets != null)
            {
                ViewSheet sheet = sheets.FirstOrDefault();

                if (sheet != null)
                    return sheet.SheetNumber;
            }

            // Didnt find the id, return null.
            return null;
        }

        /// <summary>
        /// Gets the view of the first view with the given name and returns it as a ViewSheet.
        /// </summary>
        private ViewSheet GetSheetAsViewSheet(string sheetName)
        {
            // Find a specific drawing sheet.
            IEnumerable<ViewSheet> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(v => v.Name.Equals(sheetName)
                && v.ViewType == ViewType.DrawingSheet);

            // Check if we found the sheet.
            if (sheets != null)
            {
                ViewSheet sheet = sheets.FirstOrDefault();

                if (sheet != null)
                    return sheet;
            }

            // Didnt find the id, return null.
            return null;
        }

        /// <summary>
        /// Finds and returns the id of the first found sheet with the given name.
        /// </summary>
        private ElementId GetSheetId(string sheetName)
        {
            // Find a specific drawing sheet.
            IEnumerable<View> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.Name.Equals(sheetName)
                && v.ViewType == ViewType.DrawingSheet);

            // Check if we found the sheet.
            if (sheets != null)
            {
                View sheet = sheets.FirstOrDefault();

                if (sheet != null)
                    return sheet.Id;
            }

            // Didnt find the id, return null.
            return null;
        }

        /// <summary>
        /// Finds and returns the id of ALL sheets with the given name.
        /// </summary>
        private ICollection<ElementId> GetSheetIdList(string sheetName)
        {
            // Find a specific drawing sheet.
            IEnumerable<ViewSheet> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .OrderBy(v => v.SheetNumber)
                .Where(v => v.Name.Equals(sheetName)
                && v.ViewType == ViewType.DrawingSheet);

            ICollection<ElementId> coll = new List<ElementId>();

            // Check if we found the sheet.
            if (sheets != null)
                foreach (ViewSheet v in sheets)
                    coll.Add(v.Id); 


            // Didnt find the id, return null.
            if (coll.Count >= 1)
                return coll;
            else
                return null;
        }

        /// <summary>
        /// Gets the id of the first TitleBlock wit hthe given name.
        /// </summary>
        private ElementId GetTitleBlockId(string titleBlock)
        {
            // Find a specific title block id.
            IEnumerable<FamilySymbol> tb = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .Cast<FamilySymbol>()
                .Where(v => v.FamilyName.Equals(titleBlock));

            // Check if we found the title block.
            if (tb != null)
                return tb.FirstOrDefault().Id;
            else
                return null;
        }             

        // Classes for use with selection filters.
        private class BuildingSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element e)
            {
                if (e is Wall)
                    return true;
                else if (e is RoofBase)
                    return true;
                else if (e is Floor)
                    return true;
                else if (e is SlabEdge)
                    return true;
                else if (e is FamilyInstance)
                    return true;

                return false;
            }

            public bool AllowReference(Reference r, XYZ p)
            {
                return true;
            }
        }
        private class TextNoteSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element e)
            {
                if (e is TextNote)
                    return true;

                return false;
            }

            public bool AllowReference(Reference r, XYZ p)
            {
                return true;
            }
        }
    }    

    // Enum used with rotating views.
    public enum RotationAngle
    {
        Left = 1,
        Down = 2,
        Right = 3
    }

    // Enum used with selecting specific objects via PickObject or PickObjects.
    public enum SelectionFilter
    {
        Building,
        TextNote
    }

    // Used via DistinguishRoom to determine the state of the room.
    public enum RoomState
    {
        Placed,
        NotPlaced,
        NotEnclosed,
        Redundant,
        Unknown
    }
}