namespace ForburySolution.Utilities;

    public static class FileUtils
    {
        public static bool FileExists(string filePath)
        {
            return File.Exists(filePath);
        }

        public static string GetLatestFile(string directoryPath, string extension = "*.xlsx")
        {
            var files = Directory.GetFiles(directoryPath, extension)
                .OrderByDescending(f => new FileInfo(f).CreationTime)
                .ToList();

            return files.Count > 0 ? files.First() : throw new FileNotFoundException("No Excel files found.");
        }
    }
