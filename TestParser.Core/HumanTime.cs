using System;

namespace TestParser.Core
{
    /// <summary>
    /// Provides functions to format a duration in seconds into a human readable string
    /// of the form "hh:mm:ss.ff".
    /// </summary>
    public static class HumanTime
    {
        /// <summary>
        /// Formats a time in seconds as a human readable string of the form "hh:mm:ss.ff".
        /// </summary>
        /// <param name="durationInSeconds">The duration in seconds.</param>
        /// <returns>Formatted duration string.</returns>
        public static string ToHumanString(int durationInSeconds)
        {
            return ToHumanString((decimal)durationInSeconds);
        }

        /// <summary>
        /// Formats a time in seconds as a human readable string of the form "hh:mm:ss.ff".
        /// </summary>
        /// <param name="durationInSeconds">The duration in seconds.</param>
        /// <returns>Formatted duration string.</returns>
        public static string ToHumanString(double durationInSeconds)
        {
            return ToHumanString((decimal)durationInSeconds);
        }

        /// <summary>
        /// Formats a time in seconds as a human readable string of the form "hh:mm:ss.ff".
        /// </summary>
        /// <param name="durationInSeconds">The duration in seconds.</param>
        /// <returns>Formatted duration string.</returns>
        public static string ToHumanString(decimal durationInSeconds)
        {
            int hours = (int)(durationInSeconds / 3600);
            durationInSeconds -= hours * 3600;

            int minutes = (int)(durationInSeconds / 60);
            durationInSeconds -= minutes * 60;

            string result = String.Format("{0:00}:{1:00}:{2:00.00}", hours, minutes, durationInSeconds);
            return result;
        }
    }
}
