using ExcelBuilder.Attributes;

namespace ExcelBuilder.Model
{
    public class People
    {
        [ExcelIgnore]
        public int Id { get; set; }

        [ExcelDisplayName("نام")]
        public string Name { get; set; }

        [ExcelDisplayName("فامیلی")]
        public string Family { get; set; }

        [ExcelTotal]
        [ExcelDisplayName("سن")]
        public decimal Age { get; set; }
    }
}
