namespace excelio.Readers
{
    using System.Collections.Generic;
    using GemBox.Spreadsheet;
    using JetBrains.Annotations;

    internal abstract class ExcelReader
    {
        protected readonly ExcelFile Workbook;

        public List<Participant> Participants { get; } = new List<Participant>();

        internal ExcelReader ([NotNull] ExcelFile workbook)
        {
            Workbook = workbook;
        }

        public abstract void Read ();
    }
}