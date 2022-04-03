using CourseWorkGenerator.Models;

namespace CourseWorkGenerator
{
    public static class Extensions
    {
        public static string GetEntityValue(this TableData tableData, int entityNumber, int experimentNumber)
        {
            EntityValue entityValue = tableData.Entities[entityNumber].Values[experimentNumber];

            return entityValue.GetFormattedEntityValue("{0}");
        }

        public static string GetFormattedEntityValue(this EntityValue entityValue, string errorValueFormat)
        {
            if (!string.IsNullOrEmpty(entityValue.Value) && !string.IsNullOrEmpty(entityValue.Value))
            {
                string errorValue = string.Format(errorValueFormat, entityValue.Error);

                return $"{entityValue.Value} ± {errorValue}".Replace(".", ",");
            }

            return string.Empty;
        }
    }
}
