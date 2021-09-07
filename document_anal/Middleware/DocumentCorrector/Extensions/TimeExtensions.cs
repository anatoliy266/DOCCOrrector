using System;
using System.Collections.Generic;
using System.Text;

namespace document_anal.Middleware.WordCorrector.Extensions
{
    public static class TimeExtensions
    {
        public static int ToInt(this TimeZoneInfo timeZone)
        {
            string localTimeZone = TimeZoneInfo.Local.ToString();
            char[] timeStr = new char[6];
            localTimeZone.CopyTo(5, timeStr, 0, 2);

            int time = int.Parse(String.Join("", timeStr));

            return time;
        }

    }
}
