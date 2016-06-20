using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;

namespace Maximis.Toolkit.Xrm
{
    public static class BusinessClosureHelper
    {
        public static readonly List<DayOfWeek> WeekendDays = new List<DayOfWeek> { DayOfWeek.Saturday, DayOfWeek.Sunday };

        public static DateTime AddWorkingDays(DateTime date, int daysToAdd, IList<DateTime> closureDates)
        {
            int addedDays = 0;
            DateTime result = date;
            if (daysToAdd > 0)
            {
                while (addedDays < daysToAdd)
                {
                    result = result.AddDays(1);
                    if (IsWorkingDay(result, closureDates)) addedDays++;
                }
            }
            else
            {
                while (addedDays > daysToAdd)
                {
                    result = result.AddDays(-1);
                    if (IsWorkingDay(result, closureDates)) addedDays--;
                }
            }
            return result;
        }

        public static Entity GetBusinessClosureCalendar(IOrganizationService orgService)
        {
            QueryExpression calendarRulesQuery = new QueryExpression("calendar") { ColumnSet = new ColumnSet(true) };
            calendarRulesQuery.Criteria.AddCondition("name", ConditionOperator.Equal, "Business Closure Calendar");
            return QueryHelper.RetrieveSingleEntity(orgService, calendarRulesQuery);
        }

        public static List<DateTime> GetBusinessClosureDates(Entity businessClosureCalendar)
        {
            List<DateTime> closureDates = new List<DateTime>();
            foreach (Entity calendarRule in businessClosureCalendar.GetAttributeValue<EntityCollection>("calendarrules").Entities)
            {
                DateTime ruleStartDate = calendarRule.GetLocalDateTimeValue("effectiveintervalstart");
                int ruleDays = calendarRule.GetAttributeValue<int>("duration") / 1440;
                for (int i = 0; i < ruleDays; i++)
                {
                    closureDates.Add(ruleStartDate);
                    ruleStartDate = ruleStartDate.AddDays(1);
                }
            }
            return closureDates;
        }

        public static bool IsWorkingDay(DateTime date, IList<DateTime> closureDates, IList<DayOfWeek> nonWorkingDays)
        {
            if (closureDates.Contains(date)) return false;
            if (nonWorkingDays.Contains(date.DayOfWeek)) return false;
            return true;
        }

        public static bool IsWorkingDay(DateTime date, IList<DateTime> closureDates)
        {
            return IsWorkingDay(date, closureDates, WeekendDays);
        }
    }
}