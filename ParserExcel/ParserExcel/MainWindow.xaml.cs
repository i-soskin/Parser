using System.Windows;
using Microsoft.Win32;
using ExcelDataReader;
using System.Data;
using System.IO;
using MySql.Data.MySqlClient;
using System;
using System.Text.RegularExpressions;

namespace WPFExcelView
{
    class DataBase // Класс для работы с базой данных SQL
    {
        MySqlConnection connection = new MySqlConnection("server = localhost; port = 3306; user = root; password = 1234; database = ubibase");

        public void openConnection()
        {
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
        }

        public void closeConnection()
        {
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Close();
        }

        public MySqlConnection getConnection()
        {
            return connection;
        }
    }

    public partial class MainWindow : Window 
    {
        IExcelDataReader edr; // Объекты для считывания и вывода таблицы

        private DataSet dataSet;
        private DataView dataView;
        private DataTable dataTable;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "EXCEL Files (*.xlsx)|*.xlsx|EXCEL Files 2003 (*.xls)|*.xls|All files (*.*)|*.*"; // Условие для отображения только нужного формата файлов
            if (openFileDialog.ShowDialog() != true)
                return;
            dataView = readFile(openFileDialog.FileName);
        }

        private DataView readFile(string fileNames)
        {

            var extension = fileNames.Substring(fileNames.LastIndexOf('.'));
            // Создание потока чтения файла.
            FileStream stream = File.Open(fileNames, FileMode.Open, FileAccess.Read);
  
            // Читатель если файл с расширением *.xlsx.
            if (extension == ".xlsx")
                edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
            // Читатель если файл с расширением *.xls.
            else if (extension == ".xls")
                edr = ExcelReaderFactory.CreateBinaryReader(stream);

            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };
            // Получаем DataView и начинаем обработку.
            dataSet = edr.AsDataSet(conf);
            DataView dtView = dataSet.Tables[0].AsDataView();
            DataTable distinctValues = dtView.ToTable(true);
            dataTable = dtView.ToTable(true);
            
            // Импортирует данные полученные из таблицы для выполнения задания 6
            DataBase dataBase = new DataBase(); 
            MySqlCommand command = new MySqlCommand("INSERT INFO 'ubi2' ('Id','Name', 'Description', ' Source', Object', Confidential', 'Integrity', 'Avalibilty', 'DataAdd', 'DataChange') VALUES (NULL,@Name,@Description,@Source,@Object,@Confidential,@Integrity,@Avalibilty,@DataAdd,@DataChange)", dataBase.getConnection());
           
            command.Parameters.Add("@Name", MySqlDbType.VarChar).Value = distinctValues.Rows[1][0];
            command.Parameters.Add("@Description", MySqlDbType.VarChar).Value = distinctValues.Rows[1][1];
            command.Parameters.Add("@Source", MySqlDbType.VarChar).Value = distinctValues.Rows[1][2];
            command.Parameters.Add("@Object", MySqlDbType.VarChar).Value = distinctValues.Rows[1][3];
            command.Parameters.Add("@Confidential", MySqlDbType.Int32).Value = distinctValues.Rows[1][4];
            command.Parameters.Add("@Integrity", MySqlDbType.Int32).Value = distinctValues.Rows[1][5];
            command.Parameters.Add("@Avalibilty", MySqlDbType.Int32).Value = distinctValues.Rows[1][6];
            command.Parameters.Add("@DataAdd", MySqlDbType.VarChar).Value = distinctValues.Rows[1][7];
            command.Parameters.Add("@DataChange", MySqlDbType.VarChar).Value = distinctValues.Rows[1][8];
                
         
            //dataBase.openConnection(); // Если нет сервера с SQl базой данных, то будет ошибка соединения поэтому тут стоит комент и в базу не записывается!

            dataBase.closeConnection();
            
            // Освобождение ресурсов после чтения.

            MessageBox.Show("Успешно загружена база данных");
            edr.Close();
            dtView = new DataView(distinctValues);
            return dtView;
        }
            // Функции обработки кнопок
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string input = InputBox.Text;
            string pattern = @"^([0-9])+$"; // Регулярное выражение для адекватности введёного значения строки.
            if(dataTable != null)
            {
                if (Regex.IsMatch(input, pattern, RegexOptions.IgnoreCase))
                {
                    int number = Convert.ToInt32(input);
                    if (number > 0 && number < dataTable.Rows.Count)
                    {
                        DbGrig.ItemsSource = "";
                        string Confidential = "Нет"; // По умолчанию ставим Нет
                        string Integrity = "Нет";
                        string Avalibilty = "Нет";
                        if (dataTable.Rows[number][5].ToString() == "1") // Замена 0 и 1 на Да и Нет
                            Confidential = "Да";
                        if (dataTable.Rows[number][6].ToString() == "1") // Замена 0 и 1 на Да и Нет
                            Integrity = "Да";
                        if (dataTable.Rows[number][7].ToString() == "1") // Замена 0 и 1 на Да и Нет
                            Avalibilty = "Да";
                        MessageBox.Show("Индификатор УБИ: " + dataTable.Rows[number][0].ToString() + "\n" // Вывод задачи 5 пункта. Вся информация по нужной нам угрозе выводится отдельно
                             + "Наименование: " + dataTable.Rows[number][1].ToString() + "\n"
                             + "Описание: " + dataTable.Rows[number][2].ToString() + "\n"
                             + "Источник угрозы: " + dataTable.Rows[number][3].ToString() + "\n"
                             + "Объект действия: " + dataTable.Rows[number][4].ToString() + "\n"
                             + "Нарушение конфиденциальности: " + Confidential + "\n"
                             + "Нарушение целостности: " + Integrity + "\n"
                             + "Нарушение доступности: " + Avalibilty + "\n"
                             + "Дата включения: " + dataTable.Rows[number][8].ToString() + "\n"
                             + "Дата изменения: " + dataTable.Rows[number][9].ToString());
                    }    
                    else
                        MessageBox.Show("Введите корректное значение номера УБИ"); // Проверяем корректность значения и что оно входит в диапазон
                }
                else
                    MessageBox.Show("Введите корректное значение номера УБИ");
            }
            else
                MessageBox.Show("Сначала выберите файл Excel");
        }


        private void View2columns_Click(object sender, RoutedEventArgs e)
        {
            if (dataView != null)
            {
                dataTable = dataView.ToTable(true, "Общая информация", "Column1"); // Вывод общего перечня угроз по заданию 3
                DataView dtView = new DataView(dataTable);
                DbGrig.ItemsSource = dtView;
            }
            else
                MessageBox.Show("Сначала выберите файл Excel"); // Проверка на наличее выбранного файла
        }

        private void AllInformation_Click(object sender, RoutedEventArgs e) // Вывод всей таблицы с информацией по заданию 4
        {
            if (dataView != null)
                DbGrig.ItemsSource = dataView;
            else
                MessageBox.Show("Сначала выберите файл Excel"); // Проверка на наличее выбранного файла
        }
        private void DbGrig_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

        }

        private void TextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }
    }
}
