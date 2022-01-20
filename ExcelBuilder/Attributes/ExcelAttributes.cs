namespace ExcelBuilder.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelDisplayName : Attribute
    {
        public string Title { get; init; }
        public ExcelDisplayName(string title) => Title = title;
    }


    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelIgnore : Attribute { }


    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelTotal : Attribute { }
}
