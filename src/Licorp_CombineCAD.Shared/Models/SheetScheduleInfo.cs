namespace Licorp_CombineCAD.Models
{
    public class SheetScheduleInfo
    {
        public string ElementIdValue { get; set; }
        public string Name { get; set; }

        public override string ToString()
        {
            return Name ?? "";
        }
    }
}
