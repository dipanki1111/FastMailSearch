using System;
using System.Globalization;

namespace FastMailSearch.Core.Search
{
    public static class DateParser
    {
        public static (DateTime? fromDate, DateTime? toDate) ParseRelativeDate(string value)
        {
            DateTime now = DateTime.Now;
            DateTime today = now.Date;

            switch (value.ToLowerInvariant().Trim())
            {
                case "today":
                    return (today, today.AddDays(1).AddTicks(-1));

                case "yesterday":
                    DateTime yesterday = today.AddDays(-1);
                    return (yesterday, yesterday.AddDays(1).AddTicks(-1));

                case "thisweek":
                    // Monday as start of week
                    int diff = (7 + (today.DayOfWeek - DayOfWeek.Monday)) % 7;
                    DateTime startOfWeek = today.AddDays(-1 * diff);
                    return (startOfWeek, startOfWeek.AddDays(7).AddTicks(-1));

                case "thismonth":
                    DateTime startOfMonth = new DateTime(today.Year, today.Month, 1);
                    DateTime endOfMonth = startOfMonth.AddMonths(1).AddTicks(-1);
                    return (startOfMonth, endOfMonth);

                default:
                    if (DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out DateTime parsed))
                    {
                        return (parsed.Date, parsed.Date.AddDays(1).AddTicks(-1));
                    }
                    if (DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTime exactParsed))
                    {
                        return (exactParsed.Date, exactParsed.Date.AddDays(1).AddTicks(-1));
                    }
                    return (null, null);
            }
        }

        public static DateTime? ParseSingleDate(string value, bool isEnd = false)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;

            value = value.Trim().ToLowerInvariant();
            var relative = ParseRelativeDate(value);
            if (relative.fromDate.HasValue)
            {
                return isEnd ? relative.toDate : relative.fromDate;
            }

            string[] formats = new string[] { "yyyy-MM-dd", "yyyy/MM/dd", "MM/dd/yyyy", "dd/MM/yyyy", "yyyy-M-d" };
            if (DateTime.TryParseExact(value, formats, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTime dt))
            {
                return isEnd ? dt.Date.AddDays(1).AddTicks(-1) : dt.Date;
            }
            if (DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out dt))
            {
                return isEnd ? dt.Date.AddDays(1).AddTicks(-1) : dt.Date;
            }

            return null;
        }
    }
}
