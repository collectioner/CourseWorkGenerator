using CourseWorkGenerator.Configuration;
using CourseWorkGenerator.Models;
using OfficeOpenXml;

namespace CourseWorkGenerator
{
    public static class DataParser
    {
        public static IReadOnlyList<TableData> Parse(AppConfiguration configuration)
        {
            using var package = new ExcelPackage(new FileInfo(configuration.SourceFileName));
            ExcelWorksheet firstWorksheet = package.Workbook.Worksheets[0];

            return configuration.Tables
                .Select(t => ParseTable(firstWorksheet, t))
                .ToList();
        }

        private static TableData ParseTable(ExcelWorksheet worksheet, TableConfiguration tableConfiguration)
        {
            List<string> header = tableConfiguration.Entities
                .Select(c => string.Format(tableConfiguration.HeaderFormat, c.DataHeader))
                .ToList();

            return new TableData
            {
                ErrorValueTextFormat = tableConfiguration.ErrorValueTextFormat,
                HeaderCells = header,
                Entities = tableConfiguration.Entities
                    .Select(e => ParseEntity(worksheet, e, tableConfiguration.NumberOfExperiments))
                    .ToList()
            };
        }

        private static EntityData ParseEntity(
            ExcelWorksheet worksheet, 
            EntityConfiguration entityConfiguration, 
            int experimentsCount)
        {
            var data = new List<EntityValue>();

            for (int i = 0; i < experimentsCount; i++)
            {
                var entityValue = new EntityValue
                {
                    ExperimentNumber = i + 1
                };

                object value = worksheet.Cells[$"{entityConfiguration.ValueCell}{entityConfiguration.StartRowNumber + i}"].Value;
                if (value is null)
                {
                    data.Add(entityValue);
                }
                else
                {
                    object error = worksheet.Cells[$"{entityConfiguration.ErrorCell}{entityConfiguration.StartRowNumber + i}"].Value;

                    entityValue.Value = value.ToString(); // ToString() returns string?, but we assume that string exists
                    entityValue.Error = error.ToString();

                    data.Add(entityValue);
                }
            }

            return new EntityData
            {
                Title = entityConfiguration.DataHeader,
                Values = data
            };
        }
    }
}
