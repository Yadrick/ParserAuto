using ExcelDataReader;
using Microsoft.Win32;
//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace AutoParser
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

       
        OpenExcel excel = new OpenExcel();                                                   // переменная для открытого Файла Эксель
        int countOfCopy = 1;                                                                // число сохраненных копий файла Эксель

        public MainWindow()
        {
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            InitializeComponent();
            if (excel.ws != null) { excel.gridLoaded(); grid.ItemsSource = excel.Miniresult2; }
            else { MessageBox.Show(excel.message); }
        }


        private void nextSheet_Click(object sender, RoutedEventArgs e)
        {
            if (excel.CountSheet <= excel.MaxCountSheet -1)                  // потому что начинаем с 0 страницы в данном условии
            {
                excel.CountSheet++;
                excel.gridLoaded();
                grid.ItemsSource = null;
                grid.ItemsSource = excel.Miniresult2;
            }
        }

        private void backSheet_Click(object sender, RoutedEventArgs e)
        {
            if (excel.CountSheet == 0) { }
            else 
            {
                excel.CountSheet--; excel.gridLoaded(); grid.ItemsSource = null;
                grid.ItemsSource = excel.Miniresult2;
            }
        }

        private void SaveeASS_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (excel.ws != null)
                {
                    excel.SaveAs($@"{Environment.CurrentDirectory}\thrlist{ countOfCopy}.xlsx");
                    MessageBox.Show($@"Данные успешно сохранены в {Environment.CurrentDirectory}\thrlist{ countOfCopy}.xlsx");
                    countOfCopy++;
                }
                else { MessageBox.Show($@"А сохранять то нечего..."); }
            }
            catch (Exception)
            {
                countOfCopy++;
                SaveeASS_Click(sender, e);
            }
            
        }


        

        private void End_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (excel != null) { excel.Close(); }
                Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Не выходит корректно закрыть приложение. Excel может быть открыт в Процессах компьютера.");
                Close();
            }
        }

        private void Download_Click(object sender, RoutedEventArgs e)
        {
            if (excel.ws == null)
            {
                DownloadExcel.Download();
                MessageBox.Show("Данные успешно загружены! Парсим...");
                excel = new OpenExcel();
                excel.gridLoaded();
                grid.ItemsSource = excel.Miniresult2;
                MessageBox.Show("Запарсено!");
            }
            else { MessageBox.Show("Данные уже загружены!"); }
        }


        private void Update_Click(object sender, RoutedEventArgs e)
        {
            textAboutUBI.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;

            MessageBox.Show("Обновляем информацию...\nПожалуйста, подождите.");
            excel.Close();
            excel.ws = null;
            try
            {
                File.Delete($@"{Environment.CurrentDirectory}\thrlist.xlsx");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // пытаемся скачать новую версию экселя
            try
            { 
                DownloadExcel.Download();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            // нужно открыть новый эксель
            try
            {
                excel = new OpenExcel();
                excel.gridUploaded();
            }
            catch (Exception)
            {
                MessageBox.Show(excel.message);
            }

            // Сравнение!
            ExcelChanges LetsChanges = new ExcelChanges();
            LetsChanges.Compare();
       

            textBox2.Text = ExcelChanges.content1+ ExcelChanges.contentBefore;
            textBox3.Text = ExcelChanges.content2 + ExcelChanges.contentAfter;


            excel.gridLoaded();
            grid.ItemsSource = null;
            grid.ItemsSource = excel.Miniresult2;

            MessageBox.Show($"Загружено!\nКоличество изменённых записей: {ExcelChanges.countUpdates}");
        }


        private void Read_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int h = Convert.ToInt32(searchUBI.Text);
                textAboutUBI.Text = excel.result[h - 1].ToString();

            }
            catch (Exception)
            {
                MessageBox.Show("Введено некорректное значение!");
            }
            //if (int.Parse(searchUBI.Text) > 0 && int.Parse(searchUBI.Text) < result.Count) { informationWithUpdate.Content = result[Convert.ToInt32(searchUBI.Text) - 1]; }
            //else { MessageBox.Show("Введено некорректное значение!"); }
        }     
    }
}
