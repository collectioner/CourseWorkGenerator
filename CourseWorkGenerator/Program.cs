using CourseWorkGenerator;
using CourseWorkGenerator.Configuration;
using Newtonsoft.Json;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
string configurationData = File.ReadAllText("configuration.json");

var configuration = JsonConvert.DeserializeObject<AppConfiguration>(configurationData);
if (configuration is null)
{
    throw new Exception("App configuration is null.");
}

IReadOnlyList<IReadOnlyList<IReadOnlyList<string>>> tablesData = DataParser.Parse(configuration);

TableGenerator.Generate(tablesData);

Console.WriteLine("Data processed.");