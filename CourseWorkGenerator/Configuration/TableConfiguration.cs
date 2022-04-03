namespace CourseWorkGenerator.Configuration
{
    public class TableConfiguration
    {
        public int NumberOfExperiments { get; set; }
        public string HeaderFormat { get; set; }
        public string ErrorValueTextFormat { get; set; }
        public IReadOnlyList<EntityConfiguration> Entities { get; set; }
    }
}
