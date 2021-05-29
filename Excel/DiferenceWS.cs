namespace DiffExcel.Excel
{
    public class DiferenceWs
    {
        public DiferenceWs(int numberRow, int numberCol, string valueSource, string valueTarget, string table)
        {
            NumberRow = numberRow;
            NumberCol = numberCol;
            ValueSource = valueSource;
            ValueTarget = valueTarget;
            Table = table;
        }

        private int NumberRow { get; set; }
        private int NumberCol { get; set; }

        private string ValueSource { get; set; }
        private string ValueTarget { get; set; }
        private string Table { get; set; }

        public override string ToString()
        {
            return $"[TABLE {Table}] -> Row: {NumberRow}, Col: {NumberCol}, ValueSource: {ValueSource}, ValueTarget: {ValueTarget}";
        }
    }
}
