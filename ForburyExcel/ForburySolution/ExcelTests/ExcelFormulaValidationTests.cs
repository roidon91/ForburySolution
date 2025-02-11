using ForburySolution.Helpers;
using ForburySolution.Utilities;
using OfficeOpenXml;

namespace ForburySolution.ExcelTests;

public class ExcelFormulaValidationTests
{
    private ExcelWorksheet worksheet;
    
    [SetUp]
    public void Setup()
    {
        worksheet = ExcelTestHelper.LoadTestExcel();
        TestLogger.ClearLog();
    }
    
    [Test]
    [TestCase("Year 1","G12","IF(G10<=5,$C$16,$C$17)")]
    [TestCase("Year 2","H12","IF(H10<=5,$C$16,$C$17)")]
    [TestCase("Year 3","I12","IF(I10<=5,$C$16,$C$17)")]
    [TestCase("Year 4","J12","IF(J10<=5,$C$16,$C$17)")]
    [TestCase("Year 5","K12","IF(K10<=5,$C$16,$C$17)")]
    [TestCase("Year 6","L12","IF(L10<=5,$C$16,$C$17)")]
    [TestCase("Year 7","M12","IF(M10<=5,$C$16,$C$17)")]
    [TestCase("Year 8","N12","IF(N10<=5,$C$16,$C$17)")]
    [TestCase("Year 9","O12","IF(O10<=5,$C$16,$C$17)")]
    [TestCase("Year 10","P12","IF(P10<=5,$C$16,$C$17)")]
    public void VerifyExcelModelIntegrityForBaseFormula(string year,string cellValue, string formula)
    {
        var formulaCells = new Dictionary<string, string>
        {
            {cellValue, formula}
        };

        foreach (var cell in formulaCells)
        {
            var actualFormula = worksheet.Cells[cell.Key].Formula;
            Console.WriteLine($"Checking formula in {cell.Key}: {actualFormula}");
            
            TestLogger.LogResult("VerifyExcelModelIntegrityForBaseFormula", formula,actualFormula);
            Assert.That(actualFormula, Is.Not.Null, $"Base Formula is missing in {cell.Key} for {year}");
            Assert.That(actualFormula, Is.EqualTo(cell.Value), $"Base Formula tampered in {cell.Key} for {year}");
        }
    }
    
    [Test]
    [TestCase("Year 1","G13","(G12/$C$10-G12/$C$9)*$C$18")]
    [TestCase("Year 2","H13","(H12/$C$10-H12/$C$9)*$C$18")]
    [TestCase("Year 3","I13","(I12/$C$10-I12/$C$9)*$C$18")]
    [TestCase("Year 4","J13","(J12/$C$10-J12/$C$9)*$C$18")]
    [TestCase("Year 5","K13","(K12/$C$10-K12/$C$9)*$C$18")]
    [TestCase("Year 6","L13","(L12/$C$10-L12/$C$9)*$C$18")]
    [TestCase("Year 7","M13","(M12/$C$10-M12/$C$9)*$C$18")]
    [TestCase("Year 8","N13","(N12/$C$10-N12/$C$9)*$C$18")]
    [TestCase("Year 9","O13", "(O12/$C$10-O12/$C$9)*$C$18")]
    [TestCase("Year 10","P13","(P12/$C$10-P12/$C$9)*$C$18")]
    public void VerifyExcelModelIntegrityForPercentageRentTire1Formula(string year,string cellValue, string formula)
    {
        var formulaCells = new Dictionary<string, string>
        {
            {cellValue, formula}
        };

        foreach (var cell in formulaCells)
        {
            var actualFormula = worksheet.Cells[cell.Key].Formula;
            Console.WriteLine($"Checking formula in {cell.Key}: {actualFormula}");
            
            TestLogger.LogResult("VerifyExcelModelIntegrityForPercentageRentTire1Formula", formula,actualFormula);
            Assert.That(actualFormula, Is.Not.Null, $"Formula for PercentageRentTire1 is missing in {cell.Key} for {year}");
            Assert.That(actualFormula, Is.EqualTo(cell.Value), $"Formula for PercentageRentTire1tampered in {cell.Key} for {year}");
        }
    }
    
    [Test]
    [TestCase("Year 1","G14","(G12/$C$11-G12/$C$10)*$C$19")]
    [TestCase("Year 2","H14","(H12/$C$11-H12/$C$10)*$C$19")]
    [TestCase("Year 3","I14","(I12/$C$11-I12/$C$10)*$C$19")]
    [TestCase("Year 4","J14","(J12/$C$11-J12/$C$10)*$C$19")]
    [TestCase("Year 5","K14","(K12/$C$11-K12/$C$10)*$C$19")]
    [TestCase("Year 6","L14","(L12/$C$11-L12/$C$10)*$C$19")]
    [TestCase("Year 7","M14","(M12/$C$11-M12/$C$10)*$C$19")]
    [TestCase("Year 8","N14","(N12/$C$11-N12/$C$10)*$C$19")]
    [TestCase("Year 9","O14", "(O12/$C$11-O12/$C$10)*$C$19")]
    [TestCase("Year 10","P14","(P12/$C$11-P12/$C$10)*$C$19")]
    public void VerifyExcelModelIntegrityForPercentageRentTire2Formula(string year,string cellValue, string formula)
    {
        var formulaCells = new Dictionary<string, string>
        {
            {cellValue, formula}
        };

        foreach (var cell in formulaCells)
        {
            var actualFormula = worksheet.Cells[cell.Key].Formula;
            Console.WriteLine($"Checking formula in {cell.Key}: {actualFormula}");
            TestLogger.LogResult("VerifyExcelModelIntegrityForPercentageRentTire2Formula", formula,actualFormula);
            Assert.That(actualFormula, Is.Not.Null, $"Formula for PercentageRentTire2 is missing in {cell.Key} for {year}");
            Assert.That(actualFormula, Is.EqualTo(cell.Value), $"Formula for PercentageRentTire2 tampered in {cell.Key} for {year}");
        }
    }

    [Test]
    [TestCase("Year 1","G15","(G11-G12/$C$11)*$C$20")]
    [TestCase("Year 2","H15","(H11-H12/$C$11)*$C$20")]
    [TestCase("Year 3","I15","(I11-I12/$C$11)*$C$20")]
    [TestCase("Year 4","J15","(J11-J12/$C$11)*$C$20")]
    [TestCase("Year 5","K15","(K11-K12/$C$11)*$C$20")]
    [TestCase("Year 6","L15","(L11-L12/$C$11)*$C$20")]
    [TestCase("Year 7","M15","(M11-M12/$C$11)*$C$20")]
    [TestCase("Year 8","N15","(N11-N12/$C$11)*$C$20")]
    [TestCase("Year 9","O15", "(O11-O12/$C$11)*$C$20")]
    [TestCase("Year 10","P15","(P11-P12/$C$11)*$C$20")]
    public void VerifyExcelModelIntegrityForPercentageRentTire3Formula(string year,string cellValue, string formula)
    {
        var formulaCells = new Dictionary<string, string>
        {
            {cellValue, formula}
        };

        foreach (var cell in formulaCells)
        {
            var actualFormula = worksheet.Cells[cell.Key].Formula;
            Console.WriteLine($"Checking formula in {cell.Key}: {actualFormula}");
            TestLogger.LogResult("VerifyExcelModelIntegrityForPercentageRentTire3Formula", formula,actualFormula);
            Assert.That(actualFormula, Is.Not.Null, $"Formula for PercentageRentTire3 is missing in {cell.Key} for {year}");
            Assert.That(actualFormula, Is.EqualTo(cell.Value), $"Formula for PercentageRentTire3 tampered in {cell.Key} for {year}");
        }
    }
    
    
    [Test]
    [TestCase("Year 1","G11","C15")]
    [TestCase("Year 2","H11","$G$11*(1+$C$21)")]
    [TestCase("Year 3","I11","$G$11*(1+$C$21)")]
    [TestCase("Year 4","J11","$G$11*(1+$C$21)")]
    [TestCase("Year 5","K11","$G$11*(1+$C$21)")]
    [TestCase("Year 6","L11","$G$11*(1+$C$21)")]
    [TestCase("Year 7","M11","$G$11*(1+$C$21)")]
    [TestCase("Year 8","N11","$G$11*(1+$C$21)")]
    [TestCase("Year 9","O11", "$G$11*(1+$C$21)")]
    [TestCase("Year 10","P11","$G$11*(1+$C$21)")]
    public void VerifyExcelModelIntegrityForYearSalesFormula(string year,string cellValue, string formula)
    {
        var formulaCells = new Dictionary<string, string>
        {
            {cellValue, formula}
        };

        foreach (var cell in formulaCells)
        {
            var actualFormula = worksheet.Cells[cell.Key].Formula;
            Console.WriteLine($"Checking formula in {cell.Key}: {actualFormula}");
            TestLogger.LogResult("VerifyExcelModelIntegrityForYearSalesFormula", formula,actualFormula);
            Assert.That(actualFormula, Is.Not.Null, $"Formula for Year Sales is missing in {cell.Key} for {year}");
            Assert.That(actualFormula, Is.EqualTo(cell.Value), $"Formula for Year Sales tampered in {cell.Key} for {year}");
        }
    }
    
    
    [TearDown]
    public void TearDown()
    {
        worksheet?.Dispose();
        TestLogger.PrintSummary();
    }
}