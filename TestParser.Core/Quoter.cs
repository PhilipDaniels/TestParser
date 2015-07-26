using System;
using System.Linq;

namespace TestParser.Core
{
    public static class Quoter
    {
        public static string CSVQuote(string text)
        {
            if (HasCSVSpecialChars(text))
                text = Quote(text);

            return text;
        }

        public static string KVPQuote(string text)
        {
            if (HasKVPSpecialChars(text))
                text = Quote(text);

            return text;
        }

        static bool HasCSVSpecialChars(string text)
        {
            return text.Contains('"') || text.Contains(',') || text.Contains('\r') || text.Contains('\n') || text.Contains('\t');
        }

        static bool HasKVPSpecialChars(string text)
        {
            return HasCSVSpecialChars(text) || text.Contains(' ') || text.Contains('=');
        }

        static string Quote(string text)
        {
            return String.Concat('"', text.Replace("\"", "\"\""), '"');
        }
    }
}
