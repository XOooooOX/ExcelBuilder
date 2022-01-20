using ExcelBuilder.Helper;
using ExcelBuilder.Model;

List<People> peoples = new();

peoples.Add(new People()
{
    Id = 1,
    Name = "Mohammad Hosein",
    Family = "Ghelich Khani",
    Age = 35
});

peoples.Add(new People()
{
    Id = 2,
    Name = "Mahsa",
    Family = "Mohammad Zade",
    Age = 32
});

var excelFile = peoples.ExportExcel();

var fileName = $"Excel{excelFile.Extension}";

File.WriteAllBytes(fileName, excelFile.File);