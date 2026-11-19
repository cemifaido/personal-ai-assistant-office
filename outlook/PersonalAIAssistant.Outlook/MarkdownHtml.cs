using System;
using System.Net;
using System.Text;

namespace PersonalAIAssistant
{
    /// <summary>
    /// Markdown essenziale -> HTML per le email. Tutto il testo viene prima escapato:
    /// l'unico markup prodotto e' quello generato qui (paragrafi, grassetto, elenchi, titoli).
    /// </summary>
    internal static class MarkdownHtml
    {
        public static string Convert(string md)
        {
            var lines = (md ?? "").Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var sb = new StringBuilder();
            bool inList = false;
            foreach (var raw in lines)
            {
                var t = raw.Trim();
                bool bullet = t.StartsWith("- ") || t.StartsWith("* ") || t.StartsWith("• ");
                if (bullet)
                {
                    if (!inList) { sb.Append("<ul style=\"margin:4px 0 8px 18px;padding:0\">"); inList = true; }
                    sb.Append("<li>").Append(Inline(t.Substring(2))).Append("</li>");
                    continue;
                }
                if (inList) { sb.Append("</ul>"); inList = false; }
                if (t.Length == 0) sb.Append("<br>");
                else if (t.StartsWith("#")) sb.Append("<p style=\"margin:8px 0 4px\"><b>").Append(Inline(t.TrimStart('#').Trim())).Append("</b></p>");
                else sb.Append("<p style=\"margin:0\">").Append(Inline(t)).Append("</p>");
            }
            if (inList) sb.Append("</ul>");
            return sb.ToString();
        }

        private static string Inline(string s)
        {
            var parts = WebUtility.HtmlEncode(s.Replace("`", "")).Split(new[] { "**" }, StringSplitOptions.None);
            var sb = new StringBuilder();
            for (int i = 0; i < parts.Length; i++)
            {
                if (i > 0) sb.Append(i % 2 == 1 ? "<b>" : "</b>");
                sb.Append(parts[i]);
            }
            if (parts.Length % 2 == 0) sb.Append("</b>"); // ** non chiuso
            return sb.ToString();
        }

        /// <summary>Inserisce l'HTML subito dopo il tag &lt;body ...&gt;, mantenendo firma e messaggio citato.</summary>
        public static string InsertAfterBody(string html, string block)
        {
            html = html ?? "";
            int p = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
            int q = p >= 0 ? html.IndexOf('>', p) : -1;
            return q >= 0 ? html.Substring(0, q + 1) + block + html.Substring(q + 1) : block + html;
        }
    }
}
