namespace CourseWorkGenerator.Configuration
{
    public class TableConfiguration
    {
        public int NumberOfExperiments { get; set; }
        public string HeaderFormat { get; set; }
        public IReadOnlyList<DataConfiguration> Cells { get; set; }
    }
}
