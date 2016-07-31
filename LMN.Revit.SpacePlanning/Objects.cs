using System.Collections.Generic;

namespace LMN.Revit.SpacePlanning
{
    public class MassObject
    {
        public string GroupId { get; set; }
        public string Department { get; set; }
        public string RoomType { get; set; }
        public string RoomName { get; set; }
        public string RoomNumber { get; set; }
        public double ProgramArea { get; set; }
        public int Quantity { get; set; }
        public List<ParameterObj> Parameters { get; set; }
    }

    public class ParameterObj
    {
        public string Name { get; set; }
        public string Value { get; set; }
        public bool IsInstance { get; set; }
        public Autodesk.Revit.DB.ParameterType ParameterType { get; set; }

        public static ParameterObj Invalid
        {
            get
            {
                ParameterObj invalid = new ParameterObj();
                invalid.IsInstance = false;
                invalid.Name = "No Parameter Selected";
                invalid.ParameterType = Autodesk.Revit.DB.ParameterType.Invalid;
                invalid.Value = null;
                return invalid;
            }
        }
    }
}
