using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;

namespace LMN.Revit.SpacePlanning
{
    [Transaction(TransactionMode.Manual)]
    public class UpdateMassesCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;

                string fileName = null;
                using (System.Windows.Forms.OpenFileDialog dlg = new System.Windows.Forms.OpenFileDialog())
                {
                    dlg.Title = "Select an Excel File";
                    dlg.RestoreDirectory = true;
                    dlg.Filter = "Excel (*.xlsx; *.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";

                    System.Windows.Forms.DialogResult result = dlg.ShowDialog();
                    if(result == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = dlg.FileName;
                    }
                }

                if(fileName != null)
                {
                    if(fileName.Split(new char[] { '.' }).Last().ToLower() == "xls" || fileName.Split(new char[] { '.' }).Last().ToLower() == "xlsx")
                    {
                        string fileOnly = new System.IO.FileInfo(fileName).Name;
                        // Launch the workset selector
                        System.Diagnostics.Process proc = System.Diagnostics.Process.GetCurrentProcess();
                        IntPtr handle = proc.MainWindowHandle;
                        WorksheetSelectWindow selectWindow = new WorksheetSelectWindow(RevitCommon.Excel.GetWorksheetNames(fileName));

                        System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(selectWindow);
                        helper.Owner = handle;
                        selectWindow.ShowDialog();

                        if(selectWindow.WorksheetName != null)
                        {
                            Schema schema = Schema.Lookup(Properties.Settings.Default.SchemaIdElement);
                            
                            // Get the masses from the project that mach the file name and worksheet.
                            IList<Element> massElements = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Mass)
                                .WhereElementIsNotElementType()
                                .Where(x => x.GetEntity(schema) != null && x.GetEntity(schema).Schema != null)
                                .ToList();
                            
                            List<Element> filteredMasses = new List<Element>();
                            List<string> parameterNames = new List<string>();
                            List<string> groupingParams = new List<string>();
                            ElementId symbolId = null;
                            ElementId levelId = null;
                            Level selectedLevel = null;
                            FamilySymbol selectedSymbol = null;
                            string creationTime = null;

                            foreach(Element elem in massElements)
                            {
                                Entity entity = elem.GetEntity(schema);
                                if (entity != null && entity.IsValid())
                                {
                                    string elemFile = entity.Get<string>(schema.GetField("FileName"));
                                    string worksheet = entity.Get<string>(schema.GetField("Worksheet"));
                                    if (fileOnly.ToLower() == elemFile.ToLower() && worksheet.ToLower() == selectWindow.WorksheetName.ToLower())
                                    {
                                        filteredMasses.Add(elem);
                                        if (parameterNames.Count == 0)
                                            parameterNames = entity.Get<IList<string>>(schema.GetField("ParameterNames")).ToList();
                                        if (groupingParams.Count == 0)
                                            groupingParams = entity.Get<IList<string>>(schema.GetField("GroupingParameters")).ToList();
                                        if (symbolId == null)
                                            selectedSymbol = doc.GetElement(entity.Get<ElementId>(schema.GetField("SymbolId"))) as FamilySymbol;
                                        if (levelId == null)
                                            selectedLevel = doc.GetElement(entity.Get<ElementId>(schema.GetField("LevelId"))) as Level;
                                        if (creationTime == null)
                                            creationTime = entity.Get<string>(schema.GetField("CreationTime"));
                                    }
                                }
                            }
                            
                            if (filteredMasses.Count > 0)
                            {
                                string headerStream = string.Empty;
                                // Excel data
                                List<List<string>> data = RevitCommon.Excel.Read(fileName, selectWindow.WorksheetName);
                                List<MassObject> masses = ExcelToMassData(data);
                                
                                List<string> headers = data[0];
                                List<int> indices = new List<int>();
                                foreach (string gn in groupingParams)
                                {
                                    for (int i = 0; i < headers.Count; i++)
                                    {
                                        if(headers[i] == gn)
                                        {
                                            indices.Add(i);
                                            break;
                                        }
                                    }
                                }

                                // find all of the unique groupings
                                List<string> uniqueGroups = new List<string>();
                                for(int i = 1; i < masses.Count; i++)
                                {
                                    MassObject mo = masses[i];
                                    string groupName = GetGroupID(mo, indices);
                                    if (!uniqueGroups.Contains(groupName) )
                                        uniqueGroups.Add(groupName);
                                }

                                // Organize the masses into groups
                                List<List<MassObject>> organizedMasses = OrganizeMasses(masses, uniqueGroups, indices);

                                // Get the DisplayUnitTypes for length and area.
                                Units units = doc.GetUnits();
                                FormatOptions areaFO = units.GetFormatOptions(UnitType.UT_Area);
                                FormatOptions lengthFO = units.GetFormatOptions(UnitType.UT_Length);
                                DisplayUnitType lengthDUT = lengthFO.DisplayUnits;
                                DisplayUnitType areaDUT = areaFO.DisplayUnits;

                                double defaultHeight = 10;
                                if (lengthDUT == DisplayUnitType.DUT_CENTIMETERS || lengthDUT == DisplayUnitType.DUT_DECIMETERS || lengthDUT == DisplayUnitType.DUT_METERS ||
                                     lengthDUT == DisplayUnitType.DUT_METERS_CENTIMETERS || lengthDUT == DisplayUnitType.DUT_MILLIMETERS)
                                    defaultHeight = 9.84252;

                                Transaction trans = new Transaction(doc, "Update Space Planning Masses");
                                trans.Start();
                                
                                // Iterate through the organized masses and either update an existing one or create a new one.
                                double currentX = 0;
                                foreach (List<MassObject> massGroup in organizedMasses)
                                {
                                    double maxX = 0;
                                    double currentY = 0;
                                    foreach (MassObject mass in massGroup)
                                    {
                                        string rName = mass.RoomName;
                                        string rNumber = mass.RoomNumber;
                                        string groupName = mass.GroupId;
                                        //TaskDialog.Show("Mass", "GroupName: " + groupName + "\nrName: " + rName + "\nrNumber: " + rNumber);
                                        ElementId matId = null;
                                        List<Element> massesToupdate = new List<Element>();
                                        foreach(Element elem in massElements)
                                        {
                                            Entity entity = null;
                                            try
                                            {
                                                entity = elem.GetEntity(schema);
                                            }
                                            catch { }
                                            if (entity != null)
                                            {
                                                string groupId = entity.Get<string>(schema.GetField("GroupId"));
                                                List<string> paramNames = entity.Get<IList<string>>(schema.GetField("ParameterNames")).ToList();
                                                string roomname = elem.LookupParameter(paramNames[2]).AsString();
                                                string roomnumber = elem.LookupParameter(paramNames[3]).AsString();
                                                if (roomname == rName && roomnumber == rNumber && groupId == groupName)
                                                {
                                                    massesToupdate.Add(elem);
                                                }


                                                if (matId == null && groupName == groupId && paramNames[5] != null)
                                                {
                                                    matId = elem.LookupParameter(paramNames[5]).AsElementId();
                                                }
                                            }
                                            
                                        }
                                        //TaskDialog.Show("Testing", "Group: " + groupIndex.ToString() + "\nObject: " + innerIndex.ToString());

                                        // Calculate width and depth
                                        double sideLength = 10;
                                        if (mass.ProgramArea > 0)
                                            sideLength = Math.Sqrt(mass.ProgramArea);
                                        if (sideLength > maxX)
                                            maxX = sideLength;
                                        //TaskDialog.Show("Test", "MassesToUpdate: " + massesToupdate.Count.ToString());
                                        if (massesToupdate.Count > 0)
                                        {
                                            // Update mass element
                                            foreach (Element updateElem in massesToupdate)
                                            {
                                                try
                                                {
                                                    // update the program area and extra parameters only
                                                    FamilyInstance fi = updateElem as FamilyInstance;

                                                    Parameter areaParam = null;
                                                    if (parameterNames[4] != null)
                                                        areaParam = fi.LookupParameter(parameterNames[4]);
                                                    if (areaParam != null)
                                                        areaParam.Set(UnitUtils.ConvertToInternalUnits(mass.ProgramArea, areaDUT));

                                                    for (int k = 9; k < parameterNames.Count; k++)
                                                    {
                                                        ParameterObj po = mass.Parameters[k - 9];
                                                        Parameter miscParam = null;
                                                        miscParam = updateElem.LookupParameter(parameterNames[k]);
                                                        if (miscParam != null)
                                                        {
                                                            switch (miscParam.StorageType)
                                                            {
                                                                case StorageType.Double:
                                                                    double dbl = 0;
                                                                    if (double.TryParse(po.Value, out dbl))
                                                                    {
                                                                        if (miscParam.Definition.ParameterType == ParameterType.Area)
                                                                            miscParam.Set(UnitUtils.ConvertToInternalUnits(dbl, areaDUT));
                                                                        else if (miscParam.Definition.ParameterType == ParameterType.Length)
                                                                            miscParam.Set(UnitUtils.ConvertToInternalUnits(dbl, lengthDUT));
                                                                        else
                                                                            miscParam.Set(dbl);
                                                                    }
                                                                    break;
                                                                case StorageType.Integer:
                                                                    int integer = 0;
                                                                    if (int.TryParse(po.Value, out integer))
                                                                        miscParam.Set(integer);
                                                                    break;
                                                                case StorageType.ElementId:
                                                                    int elemIdInt = 0;
                                                                    if (int.TryParse(po.Value, out integer))
                                                                        miscParam.Set(new ElementId(elemIdInt));
                                                                    break;
                                                                default:
                                                                    miscParam.Set(po.Value);
                                                                    break;
                                                            }
                                                        }
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    TaskDialog.Show("error", ex.Message);
                                                }

                                            }
                                        }
                                        if (mass.Quantity - massesToupdate.Count > 0)
                                        {
                                            // Create a new mass element
                                            // need to track X and Y position for creating masses
                                            
                                            for (int j = 0; j < mass.Quantity - massesToupdate.Count; j++)
                                            {
                                                // Location
                                                XYZ loc = new XYZ(currentX, currentY, selectedLevel.ProjectElevation);

                                                //FamilyInstance
                                                FamilyInstance fi = doc.Create.NewFamilyInstance(loc, selectedSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);



                                                // Set the parameters as necessary
                                                if (parameterNames[0] != null)
                                                    fi.LookupParameter(parameterNames[0]).Set(mass.Department);

                                                if (parameterNames[1] != null)
                                                    fi.LookupParameter(parameterNames[1]).Set(mass.RoomType);

                                                if (parameterNames[2] != null)
                                                    fi.LookupParameter(parameterNames[2]).Set(mass.RoomName);

                                                if (parameterNames[3] != null)
                                                    fi.LookupParameter(parameterNames[3]).Set(mass.RoomNumber);

                                                if (parameterNames[4] != null)
                                                    fi.LookupParameter(parameterNames[4]).Set(UnitUtils.ConvertToInternalUnits(mass.ProgramArea, areaDUT));

                                                if (parameterNames[5] != null && matId != null)
                                                    fi.LookupParameter(parameterNames[5]).Set(matId);

                                                if (parameterNames[6] != null)
                                                    fi.LookupParameter(parameterNames[6]).Set(UnitUtils.ConvertToInternalUnits(sideLength, lengthDUT));

                                                if (parameterNames[7] != null)
                                                    fi.LookupParameter(parameterNames[7]).Set(UnitUtils.ConvertToInternalUnits(sideLength, lengthDUT));

                                                if (parameterNames[8] != null)
                                                    fi.LookupParameter(parameterNames[8]).Set(UnitUtils.ConvertToInternalUnits(defaultHeight, lengthDUT));

                                                for (int k = 9; k < parameterNames.Count; k++)
                                                {

                                                    ParameterObj po = mass.Parameters[k - 9];

                                                    if (parameterNames[k] != null)
                                                    {
                                                        try
                                                        {
                                                            Parameter p = fi.LookupParameter(parameterNames[k]);
                                                            switch (p.StorageType)
                                                            {
                                                                case StorageType.Double:
                                                                    double dbl = 0;
                                                                    if (double.TryParse(po.Value, out dbl))
                                                                    {
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
                                                                        p.Set(integer);
                                                                    break;
                                                                case StorageType.ElementId:
                                                                    int elemIdInt = 0;
                                                                    if (int.TryParse(po.Value, out integer))
                                                                        p.Set(new ElementId(elemIdInt));
                                                                    break;
                                                                default:
                                                                    p.Set(po.Value);
                                                                    break;
                                                            }
                                                        }
                                                        catch { }
                                                    }
                                                }

                                                // Store the local Schema data, ie the creation datetime.
                                                try
                                                {
                                                    StoreElemDateTime(fi, fileOnly, selectWindow.WorksheetName, parameterNames, groupingParams, selectedSymbol, selectedLevel, creationTime, groupName);
                                                }
                                                catch (Exception ex)
                                                {
                                                    TaskDialog.Show("Schema Error", ex.Message);
                                                }
                                                currentY = currentY + sideLength + 5;
                                            }
                                        }
                                    }
                                    // Adjust the X position
                                    currentX = currentX + maxX + 5;
                                }
                                trans.Commit();
                            }
                        }
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

        }

        public List<MassObject> ExcelToMassData(List<List<string>> excelData)
        {
            List<string> headers = excelData[0];

            bool accurateHeaders = true;
            for (int i = 0; i < headers.Count; i++)
            {
                if (!accurateHeaders)
                    break;

                switch (i)
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
            }

            if (accurateHeaders)
            {
                // Create the mass objects and fill out the datagrid table
                List<MassObject> masses = new List<MassObject>();
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
                            else if (j > 5)
                            {
                                ParameterObj param = new ParameterObj();
                                param.Name = headers[j];
                                param.Value = excelData[i][j];
                                massParams.Add(param);
                            }
                        }
                        catch { }
                    }

                    if (valid)
                    {
                        mass.Parameters = massParams;
                        masses.Add(mass);
                    }
                }
                return masses;
            }
            else
            {
                TaskDialog.Show("Warning", "The selected Excel worksheet is not in a valid format.");
                return null;
            }
        }

        public List<List<MassObject>> OrganizeMasses(List<MassObject> masses, List<string> groupNames, List<int> indices)
        {
            List<List<MassObject>> orgMasses = new List<List<MassObject>>();

            foreach(string s in groupNames)
            {
                orgMasses.Add(new List<MassObject>());
            }

            foreach(MassObject mo in masses)
            {
                string groupName = GetGroupID(mo, indices);

                if(groupName != string.Empty)
                {
                    mo.GroupId = groupName;
                    int index = groupNames.IndexOf(groupName);
                    if(index != -1)
                    {
                        orgMasses[index].Add(mo);
                    }
                }
            }

            return orgMasses;
        }

        private string GetGroupID(MassObject mo, List<int> indices)
        {
            try
            {
                string groupName = string.Empty;
                foreach (int index in indices)
                {
                    switch (index)
                    {
                        case 0:
                            if (groupName == string.Empty)
                                groupName = mo.Department;
                            else
                                groupName += " - " + mo.Department;
                            break;
                        case 1:
                            if (groupName == string.Empty)
                                groupName = mo.RoomType;
                            else
                                groupName += " - " + mo.RoomType;
                            break;
                        case 2:
                            if (groupName == string.Empty)
                                groupName = mo.RoomName;
                            else
                                groupName += " - " + mo.RoomName;
                            break;
                        case 3:
                            if (groupName == string.Empty)
                                groupName = mo.RoomNumber;
                            else
                                groupName += " - " + mo.RoomNumber;
                            break;
                        case 4:
                            if (groupName == string.Empty)
                                groupName = mo.ProgramArea.ToString();
                            else
                                groupName += " - " + mo.ProgramArea.ToString();
                            break;
                        case 5:
                            if (groupName == string.Empty)
                                groupName = mo.Quantity.ToString();
                            else
                                groupName += " - " + mo.Quantity.ToString();
                            break;
                    }
                }
                return groupName;
            }
            catch
            {
                return null;
            }
        }

        private bool StoreElemDateTime(Element elem, string excelFileName, string excelWorksheetName, List<string> paramNames, List<string> massGroupingParams, FamilySymbol selectedSymbol, Level selectedLevel, string creationDateTime, string groupId)
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
