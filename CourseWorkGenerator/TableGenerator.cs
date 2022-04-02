using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

namespace CourseWorkGenerator
{
    public static class TableGenerator
    {
        public static void Generate(IReadOnlyList<IReadOnlyList<IReadOnlyList<string>>> tablesData)
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
                foreach (IReadOnlyList<IReadOnlyList<string>> tableData in tablesData)
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

        private static void GenerateTable(IReadOnlyList<IReadOnlyList<string>> tableData, ref Word.Document wordDoc)
        {
            object oEndOfDoc = "\\endofdoc"; // constant value

            int rowCount = tableData.Count;
            int columnCount = tableData.First().Count;

            var range = wordDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            var table = wordDoc.Tables.Add(range, rowCount, columnCount);

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                table.Cell(rowIndex + 1, columnIndex + 1).Range.Text = tableData[rowIndex][columnIndex];
            }

            // insert an empty paragraph after table
            range = wordDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            wordDoc.Content.Paragraphs.Add(range);
        }
    }
}
