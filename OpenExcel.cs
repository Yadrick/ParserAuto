using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.IO;
using ExcelDataReader;

namespace AutoParser
{
    //КЛАСС УМЕЕТ: ОТКРЫВАТЬ ФАЙЛ, ЧИТАТЬ ДАННЫЕ ИЗ ТАБЛИЦЫ И ВОЗВРАЩАТЬ ИХ ЗНАЧЕНИЕ
    //ПЕРЕПИСЫВАТЬ ЗНАЧЕНИЕ ТАБДИЦЫ, СОХРАНЯТЬ ФАЙЛ ПОСЛЕ РЕДАКТИРОВАНИЯ, СОХРАНЯТЬ ФАЙЛ В ОДЕЛЬНОМ ФАЙЛЕ, ЗАКРЫВАТЬ ФАЙЛ
    public class OpenExcel
    {
        string path = "";                                                                       // путь Excel
        Application excel = new Application();                                                  //     
        Workbook wb;                                                                            // Эксель файл?!)
        public Worksheet ws;                                                                           // номер страницы в Excell
        public string message = "";

        public double MaxCountSheet { get; set; } = 0;
        public int CountSheet { get; set; } = 0;
        public int k = 0;

        public List<DataFromExcelMINI> Miniresult2 = new List<DataFromExcelMINI>();                // данные result2 в кратком формате

        public List<DataFromExcel> result = new List<DataFromExcel>();                     //сюда добавляю все данные таблицы
        public List<DataFromExcelMINI> Maxiresult = new List<DataFromExcelMINI>();                // данные result в кратком формате

        public List<DataFromExcel> result3 = new List<DataFromExcel>();
        public List<DataFromExcelMINI> Maxiresult3 = new List<DataFromExcelMINI>();

        public HashSet<string> NamesBefore = new HashSet<string>();
        public HashSet<string> NamesBefore2 = new HashSet<string>();

        public string sheetContent = "";


        public OpenExcel()
        {
            this.path = $@"{Environment.CurrentDirectory}\thrlist.xlsx";
            try
            {
                wb = excel.Workbooks.Open(path);
                ws = wb.Worksheets[1];
                
            }
            catch (Exception)
            {
                message = "Файла с локальной базой не существует или он находится не в папке с приложением! Проведите первичную загрузку данных или переместите в папку с приложением.";
            }

            
        }



        public void gridUploaded()
        {
            int j = 3;
            DeleteUnneedColumns();
            while (ReadCell(j, 0) != "")
            {
                result3.Add(new DataFromExcel(ReadCell(j, 0), ReadCell(j, 1), ReadCell(j, 2), ReadCell(j, 3), ReadCell(j, 4), ReadCell(j, 5), ReadCell(j, 6), ReadCell(j, 7)));
                Maxiresult3.Add(new DataFromExcelMINI("УБИ." + ReadCell(j, 0), ReadCell(j, 1)));
                j++;
            }

            result = null;
            Maxiresult = null;
            Maxiresult = Maxiresult3;
            result = result3;
            CountSheet = 0;
        }


        public void gridLoaded()
        {
            Miniresult2.Clear();
            DeleteUnneedColumns();
            int j = 3;
            // заполняет в List result все строки из Excel, начиная с j=2
            if (result.Count < 5)// такое условие, чтобы загрузилось в первый раз, а больше не загружалось повторно
            {
                while (ReadCell(j, 0) != "")
                {
                    result.Add(new DataFromExcel(ReadCell(j, 0), ReadCell(j, 1), ReadCell(j, 2), ReadCell(j, 3), ReadCell(j, 4), ReadCell(j, 5), ReadCell(j, 6), ReadCell(j, 7)));
                    Maxiresult.Add(new DataFromExcelMINI("УБИ." + result[j - 3].Id, result[j - 3].NameUBI));
                    NamesBefore.Add($"{result[j - 3].NameUBI}");
                    NamesBefore2.Add($"{result[j - 3].NameUBI}");
                    j++;
                }
            }
            MaxCountSheet = result.Count / 20;                                      // максимальное число страниц, где хранится 20 строк



            //для отображения 20 строк на одном листе(если есть 20 в БД)
            for (int i = CountSheet * 20; i < 20 + (CountSheet * 20); i++)
            {
                if (i < result.Count)
                {
                    //result2.Add(result[i]);                                       //здесь заполняется лист, в котором 20 полей с полной инфой о каждом
                    Miniresult2.Add(Maxiresult[i]);                                 // здесь краткая инфа 20 полей
                }
                else break;
            }

            // цикл для корректного переключения страниц
            if (Miniresult2.Count > 0)
            {
                string prosto = Miniresult2[Miniresult2.Count - 1].Id.Substring(4);
                k = Convert.ToInt32(prosto);

                sheetContent = $"{Miniresult2[0].Id.Substring(4)} - {k}";
            }

            //gridSource = Miniresult2;
        }



        public void DeleteUnneedColumns()
        {
            int i = 2;
            if (ws != null)
            {
                if (ws.Cells[i, 9] != null)
                {
                    while (ws.Cells[i, 9].Value2 != null && ws.Cells[i, 10].Value2 != null)
                    {
                        ws.Cells[i, 9].Clear();// = null;
                        ws.Cells[i, 9].Clear();// = null;
                        ws.Cells[i, 10].Clear();
                        i++;
                    }
                }
                
            }
            
        }


        // Функция для читки данных из таблицы Excell, i j - номера строк и столбцов
        public string ReadCell(int i, int j)
        {
            j++;
            if (ws.Cells[i, j].Value2 != null)
            {
                return Convert.ToString(ws.Cells[i, j].Text);
            }
            else { return ""; }
        }


        // функция сохраняет файл, после редактирования например
        public void Save()
        {
            if (wb != null) { wb.Save(); }

        }
        // функция сохраняет файл в отдельном месте path (полный путь + название файла) Если просто имя файла, не знаю, куда он сохраняется))) 
        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }

        // Функция закрыает открытый файл Excel, который открывается в конструкторе класса
        public void Close()
        {
            if (wb != null) { wb.Close(); }
        }
    }
}
