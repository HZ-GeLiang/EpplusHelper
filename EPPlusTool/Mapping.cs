namespace EPPlusTool
{
    internal class ExcelInfo
    {
        public string R1C1 { get; set; }
        public int ColumnIndex { get; set; }
        public string Value { get; set; }
    }

    internal class ModelProp
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public List<string> Attribute { get; set; }
    }
}
