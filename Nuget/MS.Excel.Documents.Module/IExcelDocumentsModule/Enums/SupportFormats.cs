namespace IExcelDocumentsModule.Enums
{
    public static class SupportFormats
    {
        #region Constants

        public const string csv = ".csv";

        public const string xls = ".xls";

        public const string xlsx = ".xlsx";

        #endregion

        public static bool IsSupportedExcelFile(this string fileFormat)
        {
            switch (fileFormat.ToLower())
            {
                case csv:
                case xls:
                case xlsx:
                    return true;
                default:
                    return false;
            }
        }
    }
}
