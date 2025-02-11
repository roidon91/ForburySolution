namespace ForburySolution.Utilities;

public static class TestLogger
{
    private const string LogFilePath = "TestResults.log";
    private static readonly List<string> Summary = [];

    /// <summary>
    /// Logs the test result for a given test case.
    /// </summary>
    public static void LogResult(string testName, object expected, object actual)
    {
        var status = expected.Equals(actual) ? "PASS" : "FAIL";
        
        var logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TestLoggerFile.log");
        var logEntry = $"{DateTime.Now}: {testName} - {status} | Expected: {expected} | Actual: {actual}";

       
        File.AppendAllText(LogFilePath, logEntry + Environment.NewLine);

        
        if (status == "FAIL")
        {
            Summary.Add(logEntry);
        }
        
        Console.WriteLine($"Log saved: {logFilePath}");
    }

    /// <summary>
    /// Prints a summary of all test cases, highlighting failures.
    /// </summary>
    public static void PrintSummary()
    {
        Console.WriteLine("\n=== Test Summary ===");
        if (Summary.Count == 0)
        {
            Console.WriteLine("All tests passed successfully!");
        }
        else
        {
            Console.WriteLine("Some tests failed:");
            foreach (var entry in Summary)
            {
                Console.WriteLine(entry);
            }
        }
    }

    /// <summary>
    /// Clears the previous log file before running new tests.
    /// </summary>
    public static void ClearLog()
    {
        if (File.Exists(LogFilePath))
        {
            File.Delete(LogFilePath);
        }
    }
}