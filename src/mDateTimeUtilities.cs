/*
 *  This is part of the KAParser4DotNet Library
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * This software is distributed without any warranty.
 *
 * @author Domenico Mammola (mimmo71@gmail.com - www.mammola.net)
 *
 */
using System;
using System.Globalization;

namespace Mammola.KAParser.mDateTimeUtilities
{
  public static class DateTimeUtilities
  {
    // http://stackoverflow.com/questions/11154673/get-the-correct-week-number-of-a-given-date
    public static int WeekOfYear_ISO8601(DateTime fromDate)
    {
        // Get jan 1st of the year
        DateTime startOfYear = fromDate.AddDays(- fromDate.Day + 1).AddMonths(- fromDate.Month +1);
        // Get dec 31st of the year
        DateTime endOfYear = startOfYear.AddYears(1).AddDays(-1);
        // ISO 8601 weeks start with Monday 
        // The first week of a year includes the first Thursday 
        // DayOfWeek returns 0 for sunday up to 6 for saterday
        int[] iso8601Correction = {6,7,8,9,10,4,5};
        int nds = fromDate.Subtract(startOfYear).Days  + iso8601Correction[(int)startOfYear.DayOfWeek];
        int wk = nds / 7;
        switch(wk)
        {
            case 0 : 
                // Return weeknumber of dec 31st of the previous year
                return WeekOfYear_ISO8601(startOfYear.AddDays(-1));
            case 53 : 
                // If dec 31st falls before thursday it is week 01 of next year
                if (endOfYear.DayOfWeek < DayOfWeek.Thursday)
                    return 1;
                else
                    return wk;
            default : return wk;
        }
    }

    /*public static DateTime StartOfTheWeek (DateTime fromDate)
    {
      System.Globalization.CultureInfo ci = 
          System.Threading.Thread.CurrentThread.CurrentCulture;
      DayOfWeek fdow = ci.DateTimeFormat.FirstDayOfWeek;
      DayOfWeek today = DateTime.Now.DayOfWeek;
      DateTime sow = DateTime.Now.AddDays(-(today - fdow)).Date;      
    }*/
  
    // http://stackoverflow.com/questions/38039/how-can-i-get-the-datetime-for-the-start-of-the-week
    public static DateTime StartOfWeek(DateTime fromDate)
    {      
      int diff = fromDate.DayOfWeek - CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek;
      if (diff < 0)
      {
        diff += 7;
      }

      return fromDate.AddDays(-1 * diff).Date;
    }

    // http://stackoverflow.com/questions/24245523/getting-the-first-and-last-day-of-a-month-using-a-given-datetime-object
    public static DateTime FirstDayOfMonth(DateTime fromDate)
    {
      return new DateTime(fromDate.Year, fromDate.Month, 1);
    }

    // http://stackoverflow.com/questions/24245523/getting-the-first-and-last-day-of-a-month-using-a-given-datetime-object
    public static int DaysInMonth(DateTime fromDate)
    {
      return DateTime.DaysInMonth(fromDate.Year, fromDate.Month);
    }

    // http://stackoverflow.com/questions/24245523/getting-the-first-and-last-day-of-a-month-using-a-given-datetime-object
    public static DateTime LastDayOfMonth(DateTime fromDate)
    {
      return new DateTime(fromDate.Year, fromDate.Month, DaysInMonth(fromDate));
    }

  }
}