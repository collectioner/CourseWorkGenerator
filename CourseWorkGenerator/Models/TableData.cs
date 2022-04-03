namespace CourseWorkGenerator.Models
{
    public class TableData
    {
        public string ErrorValueTextFormat { get; set; }
        public IReadOnlyList<string> HeaderCells { get; set; }
        public IReadOnlyList<EntityData> Entities { get; set; }
    }
}
