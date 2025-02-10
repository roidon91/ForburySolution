using ClosedXML.Excel;

namespace ForburySolution.Utilities;


    public static class ExcelUtils
    {
        public static IXLWorksheet LoadExcelWorksheet(string filePath)
        {
            if (!FileUtils.FileExists(filePath))
                throw new FileNotFoundException($"Excel file not found: {filePath}");

            var workbook = new XLWorkbook(filePath);
            return workbook.Worksheet(1);
        }

        public static List<string>? GetHeaders(IXLWorksheet worksheet)
        {
            return worksheet.FirstRowUsed()?.CellsUsed().Select(c => c.GetString()).ToList();
        }

        public static List<string> GetColumnValues(IXLWorksheet worksheet, int columnNumber)
        {
            return worksheet.Column(columnNumber).CellsUsed().Skip(1) // Skip header row
                .Select(c => c.GetString()).ToList();
        }
    }