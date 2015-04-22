using System;
using System.Text.RegularExpressions;

namespace Maximis.Toolkit.Html
{
    public static class HtmlHelper
    {
        private static readonly Regex htmlTags = new Regex("<.+?>");
        private static readonly Regex newLineMarker = new Regex("¬+");
        private static readonly Regex paraClose = new Regex("</p>", RegexOptions.IgnoreCase);
        private static readonly Regex whiteSpace = new Regex("\\s+");

        public static string StripHtml(string html, bool includeLineBreaks = true)
        {
            // Convert linebreaks and closing paragraph tags to a "¬"
            html = html.Replace(Environment.NewLine, "¬");
            html = paraClose.Replace(html, "¬");

            // Replace all HTML tags with a space
            html = htmlTags.Replace(html, " ");

            // Reduce multiple spaces to a single space
            html = whiteSpace.Replace(html, " ");

            // Catch a space after a line break
            html = html.Replace("¬ ", "¬");

            // Now replace one or more line break markers with two line breaks or a space
            if (includeLineBreaks)
            {
                html = newLineMarker.Replace(html, Environment.NewLine + Environment.NewLine);
            }
            else
            {
                html = newLineMarker.Replace(html, " ");
            }

            // Trim and return
            return html.Trim();
        }
    }
}