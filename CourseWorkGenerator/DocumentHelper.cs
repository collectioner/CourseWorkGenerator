namespace CourseWorkGenerator
{
    public static class DocumentHelper
    {
        private static readonly Dictionary<int, string> Map = new()
        {
            {1, "первая"},
            {2, "вторая"},
            {3, "третья"},
            {4, "четвертая"},
            {5, "пятая"},
            {6, "шестая"},
            {7, "седьмая"},
            {8, "восьмая"},
            {9, "девятая"},
            {10, "десятая"}
        };

        public static string GetStageNumberTranslation(int stage)
        {
            if (!Map.ContainsKey(stage))
            {
                throw new ArgumentOutOfRangeException(nameof(stage));
            }

            return Map[stage];
        }
    }
}
