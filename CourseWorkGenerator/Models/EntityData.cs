namespace CourseWorkGenerator.Models
{
    public class EntityData
    {
        public string Title { get; set; }
        public IReadOnlyList<EntityValue> Values { get; set; }
    }

    public class EntityValue
    {
        public int ExperimentNumber { get; set; }
        public string Value { get; set; }
        public string Error { get; set; }
    }
}
