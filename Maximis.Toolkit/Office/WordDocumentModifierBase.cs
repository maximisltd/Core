using System.Collections.Generic;
using System.IO;

namespace Maximis.Toolkit.Office
{
    public enum CarriageReturnBehaviour
    {
        Ignore,
        LineBreak,
        ParagraphBreak
    }

    public abstract class WordDocumentModifierBase
    {
        private static readonly string CARRIAGE_RETURN = ((char)13).ToString();

        public WordDocumentModifierBase()
        {
        }

        public List<WordDocumentModifierReplacement> Replacements { get; set; }

        public void AddReplacement(string originalText, string replacement)
        {
            if (this.Replacements == null) this.Replacements = new List<WordDocumentModifierReplacement>();
            this.Replacements.Add(new WordDocumentModifierReplacement { OriginalText = originalText, Replacement = replacement });
        }

        public abstract void PerformOperation(Stream input, Stream output);

        public byte[] PerformOperation(byte[] input)
        {
            using (MemoryStream inStr = new MemoryStream(input))
            using (MemoryStream outStr = new MemoryStream())
            {
                PerformOperation(inStr, outStr);
                return outStr.ToArray();
            }
        }

        protected string PerformOperationWorker(string documentXml)
        {
            string result = documentXml;

            // Perform replacements on string
            if (this.Replacements != null)
            {
                foreach (WordDocumentModifierReplacement replacement in this.Replacements)
                {
                    string repText = replacement.Replacement.Replace("&", "&amp;");
                    if (replacement.CarriageReturnBehaviour != CarriageReturnBehaviour.Ignore)
                    {
                        repText = repText.Replace(CARRIAGE_RETURN, string.Empty);
                        switch (replacement.CarriageReturnBehaviour)
                        {
                            case CarriageReturnBehaviour.LineBreak:
                                repText = repText.Replace(((char)10).ToString(), "<w:br/>");
                                break;

                            case CarriageReturnBehaviour.ParagraphBreak:
                                repText = string.Format("<w:p>{0}</w:p>", repText.Replace(((char)10).ToString(), "</w:p><w:p>"));
                                break;
                        }
                    }
                    result = result.Replace(replacement.OriginalText, repText);
                }
            }

            return result;
        }
    }

    public class WordDocumentModifierReplacement
    {
        public CarriageReturnBehaviour CarriageReturnBehaviour { get; set; }

        public string OriginalText { get; set; }

        public string Replacement { get; set; }
    }
}