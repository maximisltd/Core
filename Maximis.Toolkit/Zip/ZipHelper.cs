using ICSharpCode.SharpZipLib.Zip;
using System.Collections;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;

namespace Maximis.Toolkit.Zip
{
    public static class ZipHelper
    {
        public static Stream GetZipEntryContent(string zipFilePath, Regex zipEntryRegex)
        {
            using (FileStream fileStream = File.OpenRead(zipFilePath))
            {
                ZipFile zipFile = new ZipFile(fileStream);
                IEnumerator enumerator = zipFile.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    ZipEntry zipEntry = (ZipEntry)enumerator.Current;
                    if (zipEntryRegex.IsMatch(zipEntry.Name))
                    {
                        return zipFile.GetInputStream(zipEntry);
                    }
                }
            }
            return null;
        }

        public static byte[] GZipCompress(byte[] bytes)
        {
            using (MemoryStream input = new MemoryStream(bytes))
            using (MemoryStream output = new MemoryStream())
            {
                using (GZipStream zip = new GZipStream(output, CompressionMode.Compress))
                {
                    input.CopyTo(zip);
                }

                return output.ToArray();
            }
        }

        public static byte[] GZipDecompress(byte[] bytes)
        {
            using (MemoryStream input = new MemoryStream(bytes))
            using (MemoryStream output = new MemoryStream())
            {
                using (GZipStream zip = new GZipStream(input, CompressionMode.Decompress))
                {
                    zip.CopyTo(output);
                }

                return output.ToArray();
            }
        }
    }
}