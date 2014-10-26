using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace csv_интерпритация
{
    public class its_time_to_begin
    {
        public int
            hour = 0,
            minute = 0,
            second = 0;
        public its_time_to_begin()
        {

        }
        public its_time_to_begin(int hour, int minute, int second)
        {
            this.hour = hour;
            this.minute = minute;
            this.second = second;
        }
    }
}