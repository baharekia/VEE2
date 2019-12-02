using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VEE2
{
    public class CreateTempTime
    {
        public int tm;
        public int ty;
        public int td;

        public DateTime CreateTime(DateTime convertedDate)
        {
            if (convertedDate.Day < 20 & convertedDate.Day > 1)
            {
                tm = convertedDate.Month - 1;
                if (tm < 1)
                {
                    ty = convertedDate.Year - 1;
                    tm = 12;
                }
                ty = convertedDate.Year;
            }

            else
            {
                tm = convertedDate.Month;
                ty = convertedDate.Year;
            }
            return new DateTime(ty, tm, 19, 23, 59, 59);
        }
    }
}
