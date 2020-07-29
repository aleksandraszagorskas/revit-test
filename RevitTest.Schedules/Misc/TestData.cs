using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitTest.Schedules.Misc
{
    public class TestData
    {
        public object Val1 { get; set; }
        public object Val2 { get; set; }
        public object Val3 { get; set; }


        public override string ToString()
        {
            List<string> vals = new List<string>();
            if (Val1 != null)
            {
                vals.Add(Val1.ToString());
            }
            if (Val2 != null)
            {
                vals.Add(Val2.ToString());
            }
            if (Val3 != null)
            {
                vals.Add(Val3.ToString());
            }
            return String.Join(",", vals);
        }
    }
}
