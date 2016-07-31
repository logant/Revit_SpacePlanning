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

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using System.Data;

namespace LMN.Revit.SpacePlanning
{
    /// <summary>
    /// Interaction logic for ImportProgramForm.xaml
    /// </summary>
    public partial class ImportProgramForm : Window
    {
        // Brushes for the button fills
        LinearGradientBrush eBrush = new LinearGradientBrush(
            System.Windows.Media.Color.FromArgb(255, 245, 245, 245), 
            System.Windows.Media.Color.FromArgb(255, 195, 195, 195), 
            new System.Windows.Point(0, 0), 
            new System.Windows.Point(0, 1));
        SolidColorBrush lBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(0, 0, 0, 0));

        List<Level> levels = new List<Level>();
        List<FamilySymbol> massFamilies = new List<FamilySymbol>();
        Level selectedLevel = null;
        FamilySymbol selectedSymbol = null;
        string excelFilePath;
        List<string> worksheetNames;
        List<MassObject> masses;
        List<string> headers;
        List<ParameterObj> instanceParameters;
        List<ParameterObj> typeParameters;
        List<Parameter> paramsOut = new List<Parameter>();
        List<string> massGroupingParams;
        UIDocument uiDoc;
        List<Material> projMaterials;
        List<int> columnIndecies;
        List<string> groupingNames;
        List<List<MassObject>> organizedMasses;

        // Creation Settings
        public ParameterObj DeptParam { get; set; }
        public ParameterObj RoomTypeParam { get; set; }
        public ParameterObj RoomNameParam { get; set; }
        public ParameterObj RoomNumberParam { get; set; }
        public ParameterObj ProgramAreaParam { get; set; }
        public ParameterObj WidthParam { get; set; }
        public ParameterObj DepthParam { get; set; }

        // Schema information
        string excelFileName;
        string excelWorksheetName;
        List<string> paramNames;
        string creationDateTime;
        // massGroupingParams, see above
        // symbolId, see selectedSymbol above
        // levelId, see selectedLevel above
         




        public ImportProgramForm(UIDocument uidoc, string excelPath)
        {
            excelFilePath = excelPath;
            uiDoc = uidoc;
            InitializeComponent();

            importDataGrid.CanUserReorderColumns = false;
            importDataGrid.CanUserAddRows = false;
            importDataGrid.CanUserDeleteRows = false;
            importDataGrid.CanUserSortColumns = false;

            // Read in the worksheets and assign to the appropriate control
            worksheetNames = RevitCommon.Excel.GetWorksheetNames(excelPath);
            worksheetComboBox.ItemsSource = worksheetNames;
            worksheetComboBox.SelectedIndex = 0;

            // Get a list of the project levesl
            FilteredElementCollector levelCollector = new FilteredElementCollector(uidoc.Document).OfCategory(BuiltInCategory.OST_Levels).OfClass(typeof(Level));
            foreach(Level lvl in levelCollector)
            {
                levels.Add(lvl);
            }
            levels.Sort((x, y) => x.Elevation.CompareTo(y.Elevation));
            levelComboBox.ItemsSource = levels;
            levelComboBox.DisplayMemberPath = "Name";
            levelComboBox.SelectedIndex = 0;

            // Get a list of the mass families in the document
            FilteredElementCollector massCollector = new FilteredElementCollector(uidoc.Document).OfCategory(BuiltInCategory.OST_Mass).OfClass(typeof(FamilySymbol));
            foreach(FamilySymbol fs in massCollector)
            {
                massFamilies.Add(fs);
            }
            massFamilies.Sort((x, y) => (x.FamilyName + " - " + x.Name).CompareTo(y.FamilyName + " - " + y.Name));
            familyComboBox.ItemsSource = massFamilies;
            familyComboBox.SelectedIndex = 0;

            // Get the project materials
            projMaterials = new List<Material>();
            FilteredElementCollector matCol = new FilteredElementCollector(uiDoc.Document).OfCategory(BuiltInCategory.OST_Materials).OfClass(typeof(Material));
            foreach(Material m in matCol)
            {
                projMaterials.Add(m);
            }
            projMaterials.Sort((x, y) => x.Name.CompareTo(y.Name));
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void createButton_Click(object sender, RoutedEventArgs e)
        {
            // Get the currently set materials
            List<Material> mats = new List<Material>();
            foreach(UIElement uiElem in controlPanel.Children)
            {
                MaterialSelectorControl matCtrl = uiElem as MaterialSelectorControl;
                mats.Add(matCtrl.Material);
            }
            
            System.Diagnostics.Process proc = System.Diagnostics.Process.GetCurrentProcess();
            IntPtr handle = proc.MainWindowHandle;
            SettingsForm form = new SettingsForm(organizedMasses, mats, paramsOut, uiDoc);
            System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(form);
            helper.Owner = handle;
            Hide();
            form.ShowDialog();
            List<Parameter> selectedParameters = form.SelectedParameters;

            // Get the DisplayUnitTypes for length and area.
            Units units = uiDoc.Document.GetUnits();
            FormatOptions areaFO = units.GetFormatOptions(UnitType.UT_Area);
            FormatOptions lengthFO = units.GetFormatOptions(UnitType.UT_Length);
            DisplayUnitType lengthDUT = lengthFO.DisplayUnits;
            DisplayUnitType areaDUT = areaFO.DisplayUnits;

            double defaultHeight = 10;
            if (lengthDUT == DisplayUnitType.DUT_CENTIMETERS || lengthDUT == DisplayUnitType.DUT_DECIMETERS || lengthDUT == DisplayUnitType.DUT_METERS ||
                 lengthDUT == DisplayUnitType.DUT_METERS_CENTIMETERS || lengthDUT == DisplayUnitType.DUT_MILLIMETERS)
                defaultHeight = 9.84252;
            string temp = "start";
            try
            {
                // Time to start creating masses, gather information for the schema
                excelFileName = new System.IO.FileInfo(excelFilePath).Name;
                temp = "excel file";
                excelWorksheetName = worksheetComboBox.SelectedItem.ToString();
                temp = "worksheet";
                paramNames = new List<string>();
                foreach (Parameter p in selectedParameters)
                {
                    if (p != null)
                        paramNames.Add(p.Definition.Name);
                    else
                        paramNames.Add(string.Empty);
                }
                temp = "params";
                creationDateTime = DateTime.Now.ToString("yyyymmdd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                temp = "creation date";
                // Other data, massGroupingParams, symbolId, and levelId are already defined.
            }
            catch (Exception ex)
            {
                TaskDialog.Show(temp, ex.Message);
            }
            // Create the masses
            using (Transaction t = new Transaction(uiDoc.Document, "Create Masses"))
            {
                t.Start();

                // Make sure the Symbol is active
                if (!selectedSymbol.IsActive)
                {
                    selectedSymbol.Activate();
                    uiDoc.Document.Regenerate();
                }

                // Iterate through the masses
                double currentX = 0;

                for (int i = 0; i < organizedMasses.Count; i++)
                {
                    
                    try
                    {
                        List<MassObject> massGrp = organizedMasses[i];
                        
                        double maxX = 0;
                        double currentY = 0;
                        string groupId = groupingNames[i];
                        foreach (MassObject mass in massGrp)
                        {
                            for (int j = 0; j < mass.Quantity; j++)
                            {
                                // Location
                                XYZ loc = new XYZ(currentX, currentY, selectedLevel.ProjectElevation);

                                //FamilyInstance
                                FamilyInstance fi = uiDoc.Document.Create.NewFamilyInstance(loc, selectedSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                
                                // Calculate width and depth
                                double sideLength = 10;
                                if (mass.ProgramArea > 0)
                                    sideLength = Math.Sqrt(mass.ProgramArea);
                                if (sideLength > maxX)
                                    maxX = sideLength;

                                // Set the parameters as necessary
                                if (selectedParameters[0] != null && selectedParameters[0].Definition.ParameterType == ParameterType.Text)
                                    fi.get_Parameter(selectedParameters[0].Definition).Set(mass.Department);
                                
                                if (selectedParameters[1] != null && selectedParameters[1].Definition.ParameterType == ParameterType.Text)
                                    fi.get_Parameter(selectedParameters[1].Definition).Set(mass.RoomType);
                                
                                if (selectedParameters[2] != null && selectedParameters[2].Definition.ParameterType == ParameterType.Text)
                                    fi.get_Parameter(selectedParameters[2].Definition).Set(mass.RoomName);
                                
                                if (selectedParameters[3] != null && selectedParameters[3].Definition.ParameterType == ParameterType.Text)
                                    fi.get_Parameter(selectedParameters[3].Definition).Set(mass.RoomNumber);
                                
                                if (selectedParameters[4] != null && selectedParameters[4].Definition.ParameterType == ParameterType.Area)
                                    fi.get_Parameter(selectedParameters[4].Definition).Set(UnitUtils.ConvertToInternalUnits(mass.ProgramArea, areaDUT));
                                
                                if (selectedParameters[5] != null && selectedParameters[5].Definition.ParameterType == ParameterType.Material)
                                    fi.get_Parameter(selectedParameters[5].Definition).Set(mats[i].Id);
                                
                                if (selectedParameters[6] != null && selectedParameters[6].Definition.ParameterType == ParameterType.Length)
                                    fi.get_Parameter(selectedParameters[6].Definition).Set(UnitUtils.ConvertToInternalUnits(sideLength, lengthDUT));
                                
                                if (selectedParameters[7] != null && selectedParameters[7].Definition.ParameterType == ParameterType.Length)
                                    fi.get_Parameter(selectedParameters[7].Definition).Set(UnitUtils.ConvertToInternalUnits(sideLength, lengthDUT));
                                
                                if (selectedParameters[8] != null && selectedParameters[8].Definition.ParameterType == ParameterType.Length)
                                    fi.get_Parameter(selectedParameters[8].Definition).Set(UnitUtils.ConvertToInternalUnits(defaultHeight, lengthDUT));
                                
                                for (int k = 9; k < mass.Parameters.Count + 9; k++)
                                {
                                    if (k - 9 < mass.Parameters.Count)
                                    {
                                        ParameterObj po = mass.Parameters[k - 9];
                                        if (selectedParameters[k] != null)
                                        {
                                            try
                                            {
                                                switch (selectedParameters[k].StorageType)
                                                {
                                                    case StorageType.Double:
                                                        double dbl = 0;
                                                        if (double.TryParse(po.Value, out dbl))
                                                        {
                                                            Parameter p = fi.get_Parameter(selectedParameters[k].Definition);
                                                            if (p.Definition.ParameterType == ParameterType.Area)
                                                                p.Set(UnitUtils.ConvertToInternalUnits(dbl, areaDUT));
                                                            else if (p.Definition.ParameterType == ParameterType.Length)
                                                                p.Set(UnitUtils.ConvertToInternalUnits(dbl, lengthDUT));
                                                            else
                                                                p.Set(dbl);
                                                        }
                                                        break;
                                                    case StorageType.Integer:
                                                        int integer = 0;
                                                        if (int.TryParse(po.Value, out integer))
                                                            fi.get_Parameter(selectedParameters[k].Definition).Set(integer);
                                                        break;
                                                    case StorageType.ElementId:
                                                        int elemIdInt = 0;
                                                        if (int.TryParse(po.Value, out integer))
                                                            fi.get_Parameter(selectedParameters[k].Definition).Set(new ElementId(elemIdInt));
                                                        break;
                                                    default:
                                                        fi.get_Parameter(selectedParameters[k].Definition).Set(po.Value);
                                                        break;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }

                                // Store the local Schema data, ie the creation datetime.
                                try
                                { 
                                    StoreElemDateTime(fi, groupId);
                                }
                                catch (Exception ex)
                                {
                                    TaskDialog.Show("Schema Error", ex.Message);
                                }

                                // Adjust the y Position
                                currentY = currentY + sideLength + 5;
                            }
                        }
                        // Adjust the X position
                        currentX = currentX + maxX + 5;
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Error", ex.Message);
                    }
                }

                // Store the schema data for the entire creation set.
                //try
                //{
                //    if (StoreImportData())
                //        TaskDialog.Show("Success", "Should have successfully created the collection schema.");
                //    else
                //        TaskDialog.Show("Failure", "Failed to create the collection schema.");
                        
                //}
                //catch (Exception ex)
                //{
                //    TaskDialog.Show("Schema Error", ex.Message);
                //}
                t.Commit();
            }

            Close();
        }

        private void createButton_MouseEnter(object sender, MouseEventArgs e)
        {
            createRect.Fill = eBrush;
        }

        private void createButton_MouseLeave(object sender, MouseEventArgs e)
        {
            createRect.Fill = lBrush;
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void cancelButton_MouseEnter(object sender, MouseEventArgs e)
        {
            cancelRect.Fill = eBrush;
        }

        private void cancelButton_MouseLeave(object sender, MouseEventArgs e)
        {
            cancelRect.Fill = lBrush;
        }

        private void familyComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            selectedSymbol = massFamilies[familyComboBox.SelectedIndex];
                                                   
            if (selectedLevel != null && selectedSymbol != null)
            {
                try
                {
                    
                    // Get the instance parameters
                    using (Transaction t = new Transaction(uiDoc.Document, "Temporarily Create Mass"))
                    {
                        t.Start();

                        // Make sure the Symbol is active
                        if(!selectedSymbol.IsActive)
                        {
                            selectedSymbol.Activate();
                            uiDoc.Document.Regenerate();
                        }

                        instanceParameters = new List<ParameterObj>();
                        paramsOut = new List<Parameter>();
                        FamilyInstance tempFI = uiDoc.Document.Create.NewFamilyInstance(XYZ.Zero, selectedSymbol, selectedLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        //TaskDialog.Show("Test", "FamilySymbol: " + selectedSymbol.Id.IntegerValue.ToString());
                        foreach (Parameter p in tempFI.Parameters)
                        {
                            if (!p.IsReadOnly)
                            {
                                paramsOut.Add(p);
                                ParameterObj po = new ParameterObj();
                                po.IsInstance = true;
                                po.Name = p.Definition.Name;
                                po.ParameterType = p.Definition.ParameterType;
                                po.Value = null;
                                instanceParameters.Add(po);
                            }
                        }
                        // Rollback the family creation
                        t.RollBack();
                    }
                    
                    // Get the type parameters
                    typeParameters = new List<ParameterObj>();
                    foreach (Parameter p in selectedSymbol.Parameters)
                    {
                        if (!p.IsReadOnly)
                        {
                            ParameterObj po = new ParameterObj();
                            po.IsInstance = false;
                            po.Name = p.Definition.Name;
                            po.ParameterType = p.Definition.ParameterType;
                            po.Value = null;
                            typeParameters.Add(po);
                        }
                    }

                    // sort the parameters
                    instanceParameters.Sort((x, y) => x.Name.CompareTo(y.Name));
                    typeParameters.Sort((x, y) => x.Name.CompareTo(y.Name));
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", ex.Message);
                }
            }
        }

        private void levelComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            selectedLevel = levels[levelComboBox.SelectedIndex];
        }

        private void worksheetComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string worksheet = worksheetComboBox.SelectedItem.ToString();
            List<List<string>> readData = RevitCommon.Excel.Read(excelFilePath, worksheet);

            BuildTable(readData);
            if (groupingNames != null && groupingNames.Count > 0 && importDataGrid.Columns.Count > 1)
                ResetMaterialControls();
            else
                controlPanel.Children.Clear();
        }

        private void BuildTable(List<List<string>> excelData)
        {
            
            DataTable table = new DataTable();

            headers = excelData[0];

            bool accurateHeaders = true;
            for (int i = 0; i < headers.Count; i++)
            {
                if (!accurateHeaders)
                    break;

                switch(i)
                {
                    case 0:
                        if (headers[i].ToUpper() == "DEPARTMENT" || headers[i].ToUpper() == "DEPT")
                            accurateHeaders = true;
                        else
                            accurateHeaders = false;
                        break;
                    case 1:
                        if (headers[i].ToUpper() == "ROOM TYPE" || headers[i].ToUpper() == "ROOMTYPE")
                            accurateHeaders = true;
                        else
                            accurateHeaders = false;
                        break;
                    case 2:
                        if (headers[i].ToUpper() == "ROOM NAME" || headers[i].ToUpper() == "ROOMNAME")
                            accurateHeaders = true;
                        else
                            accurateHeaders = false;
                        break;
                    case 3:
                        if (headers[i].ToUpper() == "ROOM NUMBER" || headers[i].ToUpper() == "ROOMNUMBER")
                            accurateHeaders = true;
                        else
                            accurateHeaders = false;
                        break;
                    case 4:
                        if (headers[i].ToUpper() == "PROGRAM AREA" || headers[i].ToUpper() == "PROGRAMAREA")
                            accurateHeaders = true;
                        else
                            accurateHeaders = false;
                        break;
                    case 5:
                        if (headers[i].ToUpper() == "QTY" || headers[i].ToUpper() == "QUANTITY")
                            accurateHeaders = true;
                        else
                            accurateHeaders = false;
                        break;
                    default:
                        break;
                }
                DataColumn dc = new DataColumn();
                dc.ColumnName = headers[i];
                table.Columns.Add(dc);
            }

            if (accurateHeaders)
            {
                // Create the mass objects and fill out the datagrid table
                masses = new List<MassObject>();
                string currentDept = null;
                string currentType = null;
                for (int i = 1; i < excelData.Count; i++)
                {
                    
                    MassObject mass = new MassObject();
                    List<ParameterObj> massParams = new List<ParameterObj>();
                    bool valid = true;
                    for (int j = 0; j < excelData[i].Count; j++)
                    {
                        try
                        {
                            if (j == 0)
                            {
                                if (excelData[i][j] != string.Empty && excelData[i][j] != null && excelData[i][j].Length > 0)
                                {
                                    currentDept = excelData[i][j];
                                    mass.Department = excelData[i][j];
                                }
                                else if (currentDept != null)
                                    mass.Department = currentDept;
                                else
                                    mass.Department = "NULL";
                            }
                            else if (j == 1)
                            {
                                if (excelData[i][j] != string.Empty && excelData[i][j] != null && excelData[i][j].Length > 0)
                                {
                                    currentType = excelData[i][j];
                                    mass.RoomType = excelData[i][j];
                                }
                                else if (currentType != null)
                                    mass.RoomType = currentType;
                                else
                                    mass.RoomType = "NULL";
                            }
                            else if (j == 2)
                            {
                                if (excelData[i][j] != string.Empty && excelData[i][j] != null && excelData[i][j].Length > 0)
                                {
                                    mass.RoomName = excelData[i][j];
                                }
                                else
                                {
                                    valid = false;
                                    break;
                                }
                            }
                            else if (j == 3)
                            {
                                if (excelData[i][j] != string.Empty && excelData[i][j] != null && excelData[i][j].Length > 0)
                                {
                                    mass.RoomNumber = excelData[i][j];
                                }
                                else
                                {
                                    valid = false;
                                    break;
                                }
                            }
                            else if (j == 4)
                            {
                                if (excelData[i][j] != string.Empty && excelData[i][j] != null && excelData[i][j].Length > 0)
                                {
                                    double area = double.NaN;
                                    double.TryParse(excelData[i][j], out area);
                                    mass.ProgramArea = area;
                                }
                                else
                                    continue;
                            }
                            else if (j == 5)
                            {
                                if (excelData[i][j] != string.Empty && excelData[i][j] != null && excelData[i][j].Length > 0)
                                {
                                    int qty = 1;
                                    int.TryParse(excelData[i][j], out qty);
                                    mass.Quantity = qty;
                                }
                                else
                                    continue;
                            }
                            else if(j > 5)
                            {
                                ParameterObj param = new ParameterObj();
                                param.Name = headers[j];
                                param.Value = excelData[i][j];
                                massParams.Add(param);
                            }
                        }
                        catch { }
                    }

                    if(valid)
                    {
                        
                            DataRow dr = table.NewRow();
                            dr[0] = mass.Department;
                            dr[1] = mass.RoomType;
                            dr[2] = mass.RoomName;
                            dr[3] = mass.RoomNumber;
                            dr[4] = mass.ProgramArea.ToString();
                            dr[5] = mass.Quantity.ToString();
                        try
                        {
                            for (int j = 0; j < excelData[0].Count - 6; j++)
                            {
                                dr[j + 6] = massParams[j].Value;
                            }
                            mass.Parameters = massParams;
                        }
                        catch { }
                        masses.Add(mass);
                        table.Rows.Add(dr);
                        
                    }
                    
                }
                DataView view = new DataView(table);
                importDataGrid.ItemsSource = view;
                importDataGrid.RowBackground = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 255,255));
            }
            else
            {
                table = new DataTable();
                DataColumn col = new DataColumn();
                col.ColumnName = "WARNING";
                table.Columns.Add(col);
                DataRow row = table.NewRow();
                row[0] = "Excel worksheet is not in the expected format.  The first six column headers should be formatted as:\n\nDEPARTMENT | ROOM TYPE | ROOM NAME | ROOM NUMBER | PROGRAM AREA | QUANTITY\n\nAdditional data can be appended as columns after these first six";
                table.Rows.Add(row);

                DataView view = new DataView(table);
                importDataGrid.ItemsSource = view;
                importDataGrid.RowBackground = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 203, 153));
            }

            // Size the columns and set them to be read only 
            double columnSize = 1.0 / importDataGrid.Columns.Count;
            for (int i = 0; i < importDataGrid.Columns.Count; i++)
            {
                importDataGrid.Columns[i].Width = new DataGridLength(columnSize, DataGridLengthUnitType.Star);
                importDataGrid.Columns[i].IsReadOnly = true;
            }
        }

        private void settingsButton_Click(object sender, RoutedEventArgs e)
        {
            // Launch the settings form.  This will be to specify the parameter(s)
            // used to assign material and provide grouping
            if (selectedLevel != null && selectedSymbol != null && instanceParameters != null)
            {
                // Launch the settings form
                System.Diagnostics.Process proc = System.Diagnostics.Process.GetCurrentProcess();
                IntPtr handle = proc.MainWindowHandle;
                GroupingForm form = new GroupingForm(headers, instanceParameters);
                System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(form);
                helper.Owner = handle;
                form.ShowDialog();
                
                if(form.DialogResult.HasValue && form.DialogResult == true)
                {
                    massGroupingParams = form.SelectedParams.ToList();

                    ResetMaterialControls();
                }
            }
        }

        private void settingsButton_MouseEnter(object sender, MouseEventArgs e)
        {
            settingsRect.Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 195, 195, 195));
        }

        private void settingsButton_MouseLeave(object sender, MouseEventArgs e)
        {
            settingsRect.Fill = lBrush;
        }

        public void ResetMaterialControls()
        {
            // Identify the indices of the 
            columnIndecies = new List<int>();
            foreach (string groupParam in massGroupingParams)
            {
                for (int i = 0; i < headers.Count; i++)
                {
                    string val = headers[i];
                    if (val == groupParam)
                    {
                        columnIndecies.Add(i);
                        break;
                    }
                }
            }

            // Group figure out how many groupings we need
            groupingNames = new List<string>();
            foreach (object obj in importDataGrid.Items)
            {
                try
                {
                    DataRowView drv = obj as DataRowView;
                    string gName = string.Empty;
                    foreach (int i in columnIndecies)
                    {
                        if (gName == string.Empty)
                            gName = drv.Row.ItemArray[i].ToString();
                        else
                            gName += " - " + drv.Row.ItemArray[i].ToString();
                    }
                    if (!groupingNames.Contains(gName))
                        groupingNames.Add(gName);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error: ", ex.Message);
                }
            }

            // instantiate the controls onto the material control area
            controlPanel.Children.Clear();
            organizedMasses = new List<List<MassObject>>();
            foreach (string gName in groupingNames)
            {
                MaterialSelectorControl ctrl = new MaterialSelectorControl(gName, projMaterials);
                ctrl.Margin = new Thickness(0, 0, 0, 0);
                controlPanel.Children.Add(ctrl);
                organizedMasses.Add(new List<MassObject>());
            }

            foreach (object obj in importDataGrid.Items)
            {
                try
                {
                    DataRowView drv = obj as DataRowView;
                    string curName = string.Empty;
                    foreach (int i in columnIndecies)
                    {
                        if (curName == string.Empty)
                            curName = drv.Row.ItemArray[i].ToString();
                        else
                            curName += " - " + drv.Row.ItemArray[i].ToString();
                    }
                    int index = groupingNames.IndexOf(curName);
                    
                    
                    string rName = drv.Row.ItemArray[2].ToString();
                    string rNumb = drv.Row.ItemArray[3].ToString();
                    if (rName == string.Empty && rName == null && rNumb == string.Empty && rNumb == null)
                        continue;
                    string dept = drv.Row.ItemArray[0].ToString();
                    string rType = drv.Row.ItemArray[1].ToString();
                    string progAreaStr = drv.Row.ItemArray[4].ToString();
                    string qtyStr = drv.Row.ItemArray[5].ToString();
                    List<ParameterObj> paramObjs = new List<ParameterObj>();
                    for(int i = 6; i < drv.Row.ItemArray.Length; i++)
                    {
                        ParameterObj po = new ParameterObj();
                        po.Name = "#" + headers[i] + "#";
                        po.Value = drv.Row.ItemArray[i].ToString();
                        paramObjs.Add(po);
                    }
                    MassObject mo = new MassObject();
                    mo.RoomName = rName;
                    mo.RoomNumber = rNumb;

                    if (dept != null && dept != string.Empty)
                        mo.Department = dept;
                    else
                        mo.Department = string.Empty;

                    if (rType != null && rType != string.Empty)
                        mo.RoomType = rType;
                    else
                        mo.RoomType = string.Empty;

                    double area = 0;
                    if (double.TryParse(progAreaStr, out area))
                        mo.ProgramArea = area;
                    else
                        mo.ProgramArea = 0;

                    int qty = 1;
                    if (int.TryParse(qtyStr, out qty))
                        mo.Quantity = qty;
                    else
                        mo.Quantity = 1;

                    //if (paramObjs != null && paramObjs.Count > 0)
                    mo.Parameters = paramObjs;
                    organizedMasses[index].Add(mo);
                }
                catch { }
            }
        }

        private bool StoreElemDateTime(Element elem, string groupId)
        {
            string message = "start";
            try
            {
                // Add schema data
                SchemaBuilder sBuilder = new SchemaBuilder(Properties.Settings.Default.SchemaIdElement);
                sBuilder.SetReadAccessLevel(AccessLevel.Public);
                sBuilder.SetWriteAccessLevel(AccessLevel.Public);
                sBuilder.SetVendorId("LMNA");
                sBuilder.SetSchemaName("ProgramElem");

                FieldBuilder fileFB = sBuilder.AddSimpleField("FileName", typeof(string));
                fileFB.SetDocumentation("File path to the excel file.");
                message = "file name set";
                FieldBuilder worksheetFB = sBuilder.AddSimpleField("Worksheet", typeof(string));
                worksheetFB.SetDocumentation("Worksheet names for the imports");
                message = "worksheet set";
                FieldBuilder paramsFB = sBuilder.AddArrayField("ParameterNames", typeof(string));
                paramsFB.SetDocumentation("Parameter names selected for this set of masses.");
                message = "parameternames set";
                FieldBuilder groupingsFB = sBuilder.AddArrayField("GroupingParameters", typeof(string));
                groupingsFB.SetDocumentation("Mass grouping parameters.");
                message = "grouping set";
                FieldBuilder symbolFB = sBuilder.AddSimpleField("SymbolId", typeof(ElementId));
                symbolFB.SetDocumentation("Symbol for mass family creation.");
                message = "symbols et";
                FieldBuilder levelFB = sBuilder.AddSimpleField("LevelId", typeof(ElementId));
                levelFB.SetDocumentation("Level for mass family creation.");
                message = "level set";
                FieldBuilder dtFB = sBuilder.AddSimpleField("CreationTime", typeof(string));
                dtFB.SetDocumentation("Creation time of this mass element per import.");
                message = "datetime set";
                FieldBuilder gidFB = sBuilder.AddSimpleField("GroupId", typeof(string));
                gidFB.SetDocumentation("GroupID based on GroupingParameters");
                Schema schema = sBuilder.Finish();
                
                Entity entity = new Entity(schema);
                entity.Set<string>("FileName", excelFileName);
                entity.Set<string>("Worksheet", excelWorksheetName);
                entity.Set<IList<string>>(schema.GetField("ParameterNames"), paramNames);
                entity.Set<IList<string>>(schema.GetField("GroupingParameters"), massGroupingParams);
                entity.Set<ElementId>("SymbolId", selectedSymbol.Id);
                entity.Set<ElementId>("LevelId", selectedLevel.Id);
                entity.Set<string>("CreationTime", creationDateTime);
                entity.Set<string>("GroupId", groupId);
                elem.SetEntity(entity);

                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show(message, ex.Message);
                return false;
            }
        }
    }
}
