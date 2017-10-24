namespace excelio.Readers
{
    using GemBox.Spreadsheet;
    using JetBrains.Annotations;

    internal class ExcelReaderFactory
    {
        [CanBeNull]
        public static ExcelReader Create (int fileFormat, ExcelFile workbook)
        {
            var reader = new InputExcelReader(workbook);
            return reader;
        } 
    }
}