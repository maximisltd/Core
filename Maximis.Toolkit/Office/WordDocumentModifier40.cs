using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections;
using System.IO;

namespace Maximis.Toolkit.Office
{
    public class WordDocumentModifier40 : WordDocumentModifierBase
    {
        public override void PerformOperation(Stream input, Stream output)
        {
            ZipFile inZip = new ZipFile(input);
            ZipOutputStream outZip = new ZipOutputStream(output);
            IEnumerator enumerator = inZip.GetEnumerator();
            while (enumerator.MoveNext())
            {
                ZipEntry inZipEntry = (ZipEntry)enumerator.Current;
                ZipEntry outZipEntry = new ZipEntry(inZipEntry.Name) { DateTime = DateTime.Now };

                outZip.PutNextEntry(outZipEntry);

                if (inZipEntry.Name == "word/document.xml")
                {
                    // Entry "word/document.xml" contains the document content

                    // Read content into a string
                    string documentXml = null;
                    using (StreamReader sr = new StreamReader(inZip.GetInputStream(inZipEntry))) { documentXml = sr.ReadToEnd(); }
                    documentXml = base.PerformOperationWorker(documentXml);

                    // Write updated XML to output zip - note StreamWriter not disposed on purpose
                    StreamWriter sw = new StreamWriter(outZip);
                    sw.Write(documentXml);
                    sw.Flush();
                }
                else
                {
                    // For all other entries, copy directly
                    using (Stream inStr = inZip.GetInputStream(inZipEntry))
                    {
                        inStr.CopyTo(outZip);
                    }
                }
                outZip.CloseEntry();
            }
            outZip.Finish();
        }
    }
}