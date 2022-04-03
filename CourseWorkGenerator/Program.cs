using CourseWorkGenerator;
using CourseWorkGenerator.Configuration;
using CourseWorkGenerator.Models;
using Newtonsoft.Json;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
string configurationData = File.ReadAllText("configuration.json");

var configuration = JsonConvert.DeserializeObject<AppConfiguration>(configurationData);
if (configuration is null)
{
    throw new Exception("App configuration is null.");
}

IReadOnlyList<TableData> tablesData = DataParser.Parse(configuration);

DocumentGenerator.Generate(tablesData);

Console.WriteLine("EntityData processed.");