using ForburySolution.Helpers;
using ForburySolution.Utilities;
using OfficeOpenXml;


namespace ForburySolution.ExcelTests;

public class ExcelColumnValidationTests
{
    private ExcelWorksheet worksheet;
    
    [SetUp]
    public void Setup()
    {
        worksheet = ExcelTestHelper.LoadTestExcel();
    }
    
    
    [Test]
    [TestCase("F12","Base")] // Base Sales
    [TestCase("F13","Percentage Rent Tier 1")] // Percentage Rent Tier 1
    [TestCase("F14","Percentage Rent Tier 2")] // Percentage Rent Tier 2
    [TestCase("F15","Percentage Rent Tier 3")] // Percentage Rent Tier 3
    public void VerifyColumnExists(string columnCell, string columnName)
    {
        TestLogger.LogResult("VerifyColumnExists", columnCell,worksheet.Cells[columnCell].ToString());
        Assert.That(worksheet.Cells[columnCell], Is.Not.Null, $"Required column {columnCell} is missing for column {columnName}.");
    }

   
    [Test]
    [TestCase("F12", 10,"Base")] // Base Sales, 10 rows
    [TestCase("F13", 10,"Percentage Rent Tier 1")] // Percentage Rent Tier 1, 10 rows
    [TestCase("F14", 10,"Percentage Rent Tier 1")] // Percentage Rent Tier 2, 10 rows
    [TestCase("F15", 10,"Percentage Rent Tier 1")] // Percentage Rent Tier 3, 10 rows
    public void VerifyColumnValuesAreNotNull(string column, int rowCount, string columnName)
    {
        for (var i = 12; i <= rowCount; i++)
        {
            var cellAddress = $"{column.Substring(12,1)}{i + 9}"; 
            var cellValue = worksheet.Cells[cellAddress].Value;
            TestLogger.LogResult("VerifyColumnValuesAreNotNull", cellValue, cellAddress);
            Assert.That(cellValue, Is.Not.Null, $"Column {column} has a missing value at {cellAddress} for column {columnName}.");
        }
    }
    
    [Test]
    public void VerifyColumnValuesMatchExpected()
    {
        
        var expectedValues = new Dictionary<string, double>
        {
            { "G10", 1 },
            { "G11", 50000000 },
            { "G12", 750000 },
            { "G13", 468750 },
            { "G14", 562500 },
            { "G15", 125000 }
        };

        foreach (var cell in expectedValues)
        {
            var actualValue = worksheet.Cells[cell.Key].Value;
            TestLogger.LogResult("VerifyColumnValuesMatchExpected", cell.Value, actualValue);
            Assert.That(Convert.ToDouble(actualValue), Is.EqualTo(cell.Value), $"Mismatch in {cell.Key}");
        }
    }
    
     
    [TearDown]
    public void TearDown()
    {
        worksheet?.Dispose();
        TestLogger.PrintSummary();
    }
}