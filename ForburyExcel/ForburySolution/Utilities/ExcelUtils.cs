using ClosedXML.Excel;
using OfficeOpenXml;

namespace ForburySolution.Utilities;

    public static class ExcelUtils
    {
        static ExcelUtils()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Required for EPPlus
        }

        /// <summary>
        /// Loads an Excel worksheet from a file.
        /// </summary>
        public static ExcelWorksheet LoadWorksheet(string filePath)
        {
            if (!FileUtils.FileExists(filePath))
                throw new FileNotFoundException($"Excel file not found: {filePath}");

            var package = new ExcelPackage(new FileInfo(filePath));
            return package.Workbook.Worksheets[0]; // Load first sheet
        }
    }