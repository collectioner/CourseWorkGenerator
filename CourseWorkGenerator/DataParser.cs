using CourseWorkGenerator.Configuration;
using OfficeOpenXml;

namespace CourseWorkGenerator
{
    public static class DataParser
    {
        public static IReadOnlyList<IReadOnlyList<IReadOnlyList<string>>> Parse(AppConfiguration configuration)
        {
            using var package = new ExcelPackage(new FileInfo(configuration.SourceFileName));
            ExcelWorksheet firstWorksheet = package.Workbook.Worksheets[0];

            return configuration.Tables
                .Select(t => ParseTable(firstWorksheet, t))
                .ToList();
        }

        private static IReadOnlyList<IReadOnlyList<string>> ParseTable(ExcelWorksheet worksheet, TableConfiguration tableConfiguration)
        {
            List<string> header = tableConfiguration.Cells
                .Select(c => string.Format(tableConfiguration.HeaderFormat, c.DataHeader))
                .ToList();

            header.Insert(0, "#");

            var tableData = new List<List<string>> { header };

            for (int i = 0; i < tableConfiguration.NumberOfExperiments; i++)
            {
                var data = new List<string> { (i + 1).ToString() };

                foreach (DataConfiguration cell in tableConfiguration.Cells)
                {
                    object value = worksheet.Cells[$"{cell.ValueCell}{cell.StartRowNumber + i}"].Value;
                    if (value is null)
                    {
                        data.Add(string.Empty);
                    }
                    else
                    {
                        object error = worksheet.Cells[$"{cell.ErrorCell}{cell.StartRowNumber + i}"].Value;

                        data.Add($"{value} ± {error}".Replace(".", ","));
                    }
                }

                tableData.Add(data);
            }

            return tableData;
        }
    }
}
