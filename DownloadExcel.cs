using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace AutoParser
{
    public class DownloadExcel
    {
        public static void Download()
        {
            WebClient wc = new WebClient();

            wc.DownloadFileAsync(new Uri("https://bdu.fstec.ru/files/documents/thrlist.xlsx"), $@"{Environment.CurrentDirectory}\thrlist.xlsx");
           
        }
    }
}
