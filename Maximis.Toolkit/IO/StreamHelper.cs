using System.IO;

namespace Maximis.Toolkit.IO
{
    public static class StreamHelper
    {
        public static string ReadStringFromStream(Stream stream)
        {
            using (StreamReader streamReader = new StreamReader(stream))
            {
                return streamReader.ReadToEnd();
            }
        }
    }
}