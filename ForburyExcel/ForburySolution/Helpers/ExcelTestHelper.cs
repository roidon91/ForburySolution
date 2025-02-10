using ClosedXML.Excel;
using ForburySolution.Utilities;

namespace ForburySolution.Helpers;


    public static class ExcelTestHelper
    {
        private const string BaseDirectory = "TestData/";

        public static string GetFile(string version = "latest")
        {
            var directoryPath = version == "latest"
                ? Path.Combine(BaseDirectory, "Latest")
                : Path.Combine(BaseDirectory, "Archives");

            return FileUtils.GetLatestFile(directoryPath);
        }

        public static string GetInvalidFile(string type)
        {
            var invalidDirectory = Path.Combine(BaseDirectory, "Invalid");
            var filePath = Path.Combine(invalidDirectory, $"test_data_{type}.xlsx");

            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Invalid test file not found: {filePath}");

            return filePath;
        }

        public static IXLWorksheet LoadTestExcel(string version = "latest")
        {
            var filePath = GetFile(version);
            return ExcelUtils.LoadExcelWorksheet(filePath);
        }
    }