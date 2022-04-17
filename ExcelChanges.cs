using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoParser
{
    class ExcelChanges : OpenExcel
    {
        public static string content1 = "БЫЛО:\n\n";
        public static string content2 = "СТАЛО:\n\n";
        public static string contentBefore = "\nУдалённые Угрозы:\n";
        public static string contentAfter = "\nНовые Угрозы:\n";
        public static int countUpdates = 0;
        public static HashSet<string> NamesAfter = new HashSet<string>();

        public void Compare()
        {
            
            foreach (var item in Maxiresult3)
            {
                NamesAfter.Add(item.NameUBI);
            }

            foreach (var item3 in result3)
            {
                if (NamesBefore.Contains(item3.NameUBI))
                {
                    int g = NamesBefore.ToList().LastIndexOf(item3.NameUBI);

                    if (item3.ToString().Equals(result[g].ToString()))
                    {
                    }
                    else
                    {
                        countUpdates++;                                 // число измененных записей
                        content1 += $"{countUpdates}.{result[g]}\n";    // было
                        content2 += $"{countUpdates}.{item3}\n";        // стало  
                    }
                }
            }

            NamesBefore.ExceptWith(NamesAfter);     // в этом списке имена тех угроз, которые были раньше, но отсутствуют сейчас
            NamesAfter.ExceptWith(NamesBefore2);    // в этом списке имена тех угроз, которые появились сейчас, но отсутствовали раньше

            if (NamesBefore.Count > NamesAfter.Count) { countUpdates += NamesBefore.Count; }
            else { countUpdates += NamesAfter.Count; }

            foreach (var item in result)
            {
                if (NamesBefore.Contains(item.NameUBI))
                {
                    contentBefore += item;
                }
            }
            foreach (var item in result3)
            {
                if (NamesAfter.Contains(item.NameUBI))
                {
                    contentAfter += item;
                }
            }

        }
    }
}
