using System.Reflection;
using System.Text;
using CourseWorkGenerator.Models;
using Word = Microsoft.Office.Interop.Word;

namespace CourseWorkGenerator
{
    public static class DocumentGenerator
    {
        private static readonly object EndOfDoc = "\\endofdoc"; // constant value

        public static void Generate(IReadOnlyList<TableData> tablesData)
        {
            string? currentDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (string.IsNullOrEmpty(currentDirectory))
            {
                throw new Exception("Current directory is null.");
            }

            var wordApplication = new Word.Application();
            Word.Document wordDoc = wordApplication.Documents.Add(Missing.Value, Missing.Value);

            try
            {
                foreach (TableData tableData in tablesData)
                {
                    GenerateTable(tableData, ref wordDoc);
                }

                wordDoc.SaveAs(Path.Combine(currentDirectory, "result.docx"));
            }
            finally
            {
                wordDoc.Close(false);
                wordApplication.Quit(false);
            }
        }

        private static void GenerateTable(TableData tableData, ref Word.Document wordDoc)
        {
            int rowCount = tableData.Entities.First().Values.Count + 1; // +1 for header
            int columnCount = tableData.HeaderCells.Count + 1; // +1 for iteration column

            var range = wordDoc.Bookmarks.get_Item(EndOfDoc).Range;
            var table = wordDoc.Tables.Add(range, rowCount, columnCount);

            table.Cell(1, 1).Range.Text = "#";

            for (int index = 0; index < tableData.HeaderCells.Count; index++)
            {
                table.Cell(1, index + 2).Range.Text = tableData.HeaderCells[index];
            }

            for (int rowIndex = 1; rowIndex < rowCount; rowIndex++)
            {
                table.Cell(rowIndex + 1, 1).Range.Text = rowIndex.ToString();

                for (int columnIndex = 1; columnIndex < columnCount; columnIndex++)
                {

                    table.Cell(rowIndex + 1, columnIndex + 1).Range.Text =
                        tableData.GetEntityValue(columnIndex - 1, rowIndex - 1);
                }
            }

            // insert an empty paragraph after table
            range = wordDoc.Bookmarks.get_Item(EndOfDoc).Range;
            wordDoc.Content.Paragraphs.Add(range);

            WriteTableDescription(tableData, ref wordDoc);
        }

        private static void WriteTableDescription(TableData tableData, ref Word.Document wordDoc)
        {
            var range = wordDoc.Bookmarks.get_Item(EndOfDoc).Range;
            var paragraph = wordDoc.Content.Paragraphs.Add(range);

            var sb = new StringBuilder();

            foreach (EntityData entityData in tableData.Entities)
            {
                sb.Append($"{entityData.Title}:{Environment.NewLine}");

                List<string> values = entityData.Values
                    .OrderBy(v => v.Value)
                    .ToList()
                    .Select(s => BuildStageString(s, tableData.ErrorValueTextFormat))
                    .ToList();

                string value = string.Join(", ", values);

                sb.Append($"{value}{Environment.NewLine}{Environment.NewLine}");
            }

            paragraph.Range.Text = sb.ToString();

            // insert an empty paragraph
            range = wordDoc.Bookmarks.get_Item(EndOfDoc).Range;
            wordDoc.Content.Paragraphs.Add(range);
        }

        private static string BuildStageString(EntityValue entityValue, string errorValueFormat)
        {
            string stageTranslation = DocumentHelper.GetStageNumberTranslation(entityValue.ExperimentNumber);

            return $"стадия {stageTranslation} ({entityValue.GetFormattedEntityValue(errorValueFormat)})";
        }
    }
}
