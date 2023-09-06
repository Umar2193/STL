using System.Globalization;

namespace STL.Helpers {
    public static class CommonHelper {

        public const string DateFormat = "yyyy-MM-dd";

        public static DateTime StartOfWeek(this DateTime dt, DayOfWeek startOfWeek = DayOfWeek.Monday) {
            int diff = (7 + (dt.DayOfWeek - startOfWeek)) % 7;
            var date = dt.AddDays(-1 * diff).Date;
            return date;
        }

        public static DateTime EndOfWeek(this DateTime dt, DayOfWeek startOfWeek = DayOfWeek.Monday) {
            var date = StartOfWeek(dt, startOfWeek).AddDays(4);
            return date;
        }

        public static List<string> WeekList(DateTime startDate, DateTime endDate) {
            List<DateTime> weeks = new List<DateTime>();
            TimeSpan span = endDate - startDate;
            for(int i = 0; i <= span.Days; i++) {
                DateTime date = startDate.AddDays(i);
                if(date.DayOfWeek == DayOfWeek.Monday) {
                    weeks.Add(date);
                }
            }

            return weeks.Select(c => c.ToString($"yyyy - Week {DateTimeFormatInfo.CurrentInfo.Calendar.GetWeekOfYear(c, CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday)}")).ToList();
        }

        public static DateTime GetStartDate(string columnName) {
            int weekNumber = int.Parse(columnName.Split(' ')[3], CultureInfo.InvariantCulture);
            int year = int.Parse(columnName.Split(' ')[0], CultureInfo.InvariantCulture);

            DateTime jan1 = new DateTime(year, 1, 1);
            int daysOffset = (int)jan1.DayOfWeek - 1;
            DateTime firstMondayOfYear = jan1.AddDays(-daysOffset);
            return firstMondayOfYear.AddDays((weekNumber - 1) * 7);
        }

        public static DateTime GetEndDate(string columnName) {
            return GetStartDate(columnName).AddDays(4);
        }

        public static string GetRootPath() {
//#if DEBUG
           // return Environment.GetEnvironmentVariable("OneDriveConsumer");
//#endif

            if(!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OneDrive"))) {
                return Environment.GetEnvironmentVariable("OneDrive");
            } else if(!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OneDriveCommercial"))) {
                return Environment.GetEnvironmentVariable("OneDriveCommercial");
            } else if(!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OneDriveConsumer"))) {
                return Environment.GetEnvironmentVariable("OneDriveConsumer");
            }
            return string.Empty;
        }
    }
}
