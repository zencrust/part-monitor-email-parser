using System.Linq;
using System.Collections.Generic;

namespace MailParser
{
    internal static class Extensions
    {
        public static string GetValue(this IList<string> inputArray, string titleString)
        {
            var result = inputArray.First(x => x.Contains(titleString))
                    .Substring(titleString.Length);
            return result.Trim();
        }

        public static string GetTimeStamp(this string timeStamp)
        {
            var timestampIndex = timeStamp.IndexOf("AM", System.StringComparison.InvariantCulture);
            if (timestampIndex == -1)
            {
                timestampIndex = timeStamp.IndexOf("PM", System.StringComparison.InvariantCulture);
            }
            return timeStamp.Substring(0, timestampIndex + 2);
        }
    }
}
