using System.IO;

//using System.IO.Compression;

namespace Maximis.Toolkit.Office
{
    public class WordDocumentModifier45 : WordDocumentModifierBase
    {
        public override void PerformOperation(Stream input, Stream output)
        {
            //// Open the source and destination Word documents as Zip Archives
            //using (ZipArchive inZip = new ZipArchive(input, ZipArchiveMode.Read, true))
            //using (ZipArchive outZip = new ZipArchive(output, ZipArchiveMode.Create, true))
            //{
            //    // Loop through entries (files inside archive)
            //    foreach (ZipArchiveEntry inZipEntry in inZip.Entries)
            //    {
            //        ZipArchiveEntry outZipEntry = outZip.CreateEntry(inZipEntry.FullName);

            //        if (inZipEntry.FullName == "word/document.xml")
            //        {
            //            // Entry "word/document.xml" contains the document content

            //            // Read content into a string
            //            string documentXml = null;
            //            using (StreamReader sr = new StreamReader(inZipEntry.Open())) { documentXml = sr.ReadToEnd(); }
            //            documentXml = base.PerformOperationWorker(documentXml);

            //            // Write updated XML to output zip
            //            using (StreamWriter sw = new StreamWriter(outZipEntry.Open())) { sw.Write(documentXml); }
            //        }
            //        else
            //        {
            //            // For all other entries, copy directly
            //            using (Stream inStr = inZipEntry.Open())
            //            using (Stream outStr = outZipEntry.Open())
            //            {
            //                inStr.CopyTo(outStr);
            //            }
            //        }
            //    }
            //}
        }
    }
}