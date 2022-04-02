namespace CourseWorkGenerator.Configuration
{
    public class AppConfiguration
    {
        public string SourceFileName { get; set; }
        public IReadOnlyList<TableConfiguration> Tables { get; set; }
    }
}
