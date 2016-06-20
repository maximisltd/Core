using System.IO;

namespace Maximis.Toolkit.IO
{
    public enum PathType { File, Directory };

    public static class FileHelper
    {
        public static void AppendToFile(string filePath, string content)
        {
            if (string.IsNullOrEmpty(filePath)) return;
            EnsureDirectoryExists(filePath, PathType.File);
            File.AppendAllText(filePath, content);
        }

        public static DirectoryInfo EnsureDirectoryExists(string fileOrFolderPath, PathType pathType)
        {
            DirectoryInfo result = null;
            if (pathType == PathType.File)
            {
                FileInfo fi = new FileInfo(fileOrFolderPath);
                if (!Directory.Exists(fi.DirectoryName)) result = Directory.CreateDirectory(fi.DirectoryName);
            }
            else
            {
                if (!Directory.Exists(fileOrFolderPath)) result = Directory.CreateDirectory(fileOrFolderPath);
            }
            return result;
        }

        public static string ReadFromFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return null;
            if (!File.Exists(filePath)) return null;
            return File.ReadAllText(filePath);
        }

        public static void WriteToFile(string filePath, string content)
        {
            if (string.IsNullOrEmpty(filePath)) return;
            EnsureDirectoryExists(filePath, PathType.File);
            if (File.Exists(filePath)) File.Delete(filePath);
            File.WriteAllText(filePath, content);
        }
    }
}