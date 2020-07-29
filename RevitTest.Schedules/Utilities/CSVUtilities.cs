using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitTest.Schedules.Utilities
{
    public class CSVUtilities
    {
        public static void ExportList<T>(string filePath, IEnumerable<T> sequence)
        {
            StringBuilder builder = new StringBuilder();
            foreach (var item in sequence)
            {
                builder.AppendLine(item.ToString());
            }

            Console.WriteLine(builder.ToString());
            File.WriteAllText(
                Path.Combine(filePath),
                builder.ToString());
            Console.ReadLine();
        }
    }
}
