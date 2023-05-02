using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Xceed.Document.NET;
using Xceed.Words.NET;
using System.IO;

namespace Template_4335.Windows
{
    /// <summary>
    /// Логика взаимодействия для Klevtsov_4335.xaml
    /// </summary>
    public partial class Klevtsov_4335 : System.Windows.Window
    {
        public Klevtsov_4335()
        {
            InitializeComponent();
        }

        private void Import_Click_1(object sender, RoutedEventArgs e)
        {
            string filePath = "C:/Users/2004y/OneDrive/Рабочий стол/isr.xslx";
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.ActiveSheet;
            Range range = worksheet.UsedRange;
            System.Data.DataTable dt = new System.Data.DataTable();
            for (int i = 1; i <= range.Rows.Count; i++)
            {
                DataRow row = dt.NewRow();
                for (int j = 1; j <= range.Columns.Count; j++)
                {
                    if (i == 1)
                    {
                        dt.Columns.Add(range.Cells[i, j].Value2.ToString());
                    }
                    else
                    {
                        row[j - 1] = range.Cells[i, j].Value2;
                    }
                }
                if (i > 1)
                {
                    dt.Rows.Add(row);
                }
            }
            workbook.Close(false);
            excelApp.Quit();

            // Save data to database table
            using (SqlConnection conn = new SqlConnection(@"data source=DESKTOP-3QE7AG6\MSSQLLocalDB;initial catalog=isrpo1;integrated security=True"))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("INSERT INTO V VALUES (@column1, @column2, @column3)", conn);
                cmd.Parameters.Add("@column1", SqlDbType.Int);
                cmd.Parameters.Add("@column2", SqlDbType.NVarChar);
                cmd.Parameters.Add("@column3", SqlDbType.NVarChar);
                foreach (DataRow row in dt.Rows)
                {
                    cmd.Parameters["@column1"].Value = row["Код"].ToString();
                    cmd.Parameters["@column2"].Value = row["Должность"].ToString();
                    cmd.Parameters["@column3"].Value = row["Логин"].ToString();
                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            }
        }

        private DataRow CreateRowFromReader(SqlDataReader reader, System.Data.DataTable table)
        {
            DataRow row = table.NewRow();
            row["Код"] = reader["Код"].ToString();
            row["Должность"] = reader["Должность"].ToString();
            row["Логин"] = reader["Логин"].ToString();
            return row;
        }

        private void Export_Click_1(object sender, RoutedEventArgs e)
        {
            string newFilePath = @"C:/Users/2004y/OneDrive/Рабочий стол/isr.xslx";
            // Group data by input type
            Dictionary<string, System.Data.DataTable> dataByInputType = new Dictionary<string, System.Data.DataTable>();
            using (SqlConnection conn = new SqlConnection(@"data source=DESKTOP-3QE7AG6\MSSQLLocalDB;initial catalog=isrpo1;integrated security=True"))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("SELECT Код,Должность,Логин FROM М", conn);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string inputType = reader["Код"].ToString();
                    if (!dataByInputType.ContainsKey(inputType))
                    {
                        dataByInputType[inputType] = new System.Data.DataTable();
                        dataByInputType[inputType].Columns.Add("Код");
                        dataByInputType[inputType].Columns.Add("Должность");
                        dataByInputType[inputType].Columns.Add("Логин");
                    }
                    DataRow row = CreateRowFromReader(reader, dataByInputType[inputType]);
                    dataByInputType[inputType].Rows.Add(row);
                }
            }

            // Save data to new xlsx-file
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Add();
            foreach (string inputType in dataByInputType.Keys)
            {
                System.Data.DataTable data = dataByInputType[inputType];
                Worksheet worksheet = (Worksheet)workbook.Worksheets.Add();
                worksheet.Name = inputType;
                Range headerRange = worksheet.Range["A1:C1"];
                headerRange.Merge();
                headerRange.Value2 = inputType;
                headerRange.Font.Bold = true;
                headerRange.Font.Size = 16;
                Range dataRange = worksheet.Range["A2:C" + (data.Rows.Count + 1)];
                dataRange.Value2 = ConvertDataTableToArray(data);
                dataRange.Columns.AutoFit();
            }
            workbook.SaveAs(newFilePath);
            workbook.Close(false);
            excelApp.Quit();
        }

        private object[,] ConvertDataTableToArray(System.Data.DataTable dt)
        {
            object[,] array = new object[dt.Rows.Count, dt.Columns.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow row = dt.Rows[i];
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    array[i, j] = row[j];
                }
            }
            return array;
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            var connectionString = @"data source=DESKTOP-3QE7AG6\MSSQLLocalDB;initial catalog=isrpo1;integrated security=True";
            var loader = new JsonToDbLoader(connectionString);

            // Load JSON file to database
            var jsonFilePath = "C:/Users/2004y/source/repos/Klevtsov_4335_/Template_4335/isrpo.json";
            loader.LoadJsonToDb(jsonFilePath);
        }
        public class JsonToDbLoader
        {
            private readonly string _connectionString;

            public JsonToDbLoader(string connectionString)
            {
                _connectionString = connectionString;
            }

            public void LoadJsonToDb(string jsonFilePath)
            {
                var json = File.ReadAllText(jsonFilePath);
                var data = JsonConvert.DeserializeObject<List<JsonData>>(json);

                foreach (var item in data)
                {
                    using (var connection = new SqlConnection(_connectionString))
                    {
                        connection.Open();

                        var sql = "INSERT INTO V (Код, Должность, Логин) VALUES (@Code, @Position, @Login)";
                        var command = new SqlCommand(sql, connection);
                        command.Parameters.AddWithValue("@Code", item.Код);
                        command.Parameters.AddWithValue("@Position", item.Должность);
                        command.Parameters.AddWithValue("@Login", item.Логин);
                        command.ExecuteNonQuery();
                    }
                }
            }

            public void SaveDataToDb(int code, string position, string login)
            {
                using (var connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    var sql = "INSERT INTO V (Код, Должность, Логин) VALUES (@Code, @Position, @Login)";
                    var command = new SqlCommand(sql, connection);
                    command.Parameters.AddWithValue("@Code", code);
                    command.Parameters.AddWithValue("@Position", position);
                    command.Parameters.AddWithValue("@Login", login);
                    command.ExecuteNonQuery();
                }
            }
        }
        public class JsonToDocxConverter
        {
            private readonly string _connectionString;

            public JsonToDocxConverter(string connectionString)
            {
                _connectionString = connectionString;
            }

            public void ConvertJsonToDocx(string jsonFilePath, string outputFilePath)
            {
                var json = File.ReadAllText(jsonFilePath);
                var data = JsonConvert.DeserializeObject<List<JsonData>>(json);

                // Group data by input type
                var groupedData = data.GroupBy(d => d.Должность);

                // Create new docx document
                using (var document = DocX.Create(outputFilePath))
                {
                    foreach (var group in groupedData)
                    {
                        // Add new page for each category
                        document.InsertSectionPageBreak();

                        // Add category title with employee count for Директор position
                        var title = group.Key;
                        if (title == "Директор")
                        {
                            var employeeCount = group.Count(g => g.Должность == "Директор");
                            title += $" ({employeeCount} директоров)";
                        }
                        var heading = document.InsertParagraph(title);
                        heading.Bold();
                        heading.FontSize(16);
                        heading.SpacingAfter(20d);

                        // Add data table for category
                        var table = document.AddTable(group.Count(), 3);
                        table.Alignment = Alignment.center;
                        table.AutoFit = AutoFit.Window;
                        int i = 0;
                        foreach (var item in group)
                        {
                            table.Rows[i].Cells[0].Paragraphs.First().Append(item.Код.ToString());
                            table.Rows[i].Cells[1].Paragraphs.First().Append(item.Должность);
                            table.Rows[i].Cells[2].Paragraphs.First().Append(item.Логин);
                            i++;
                        }
                        document.InsertTable(table);
                    }
                    document.Save();
                }
            }
        }
        private void Export_Click(object sender, RoutedEventArgs e)
        {
            var connectionString = @"data source=DESKTOP-3QE7AG6\MSSQLLocalDB;initial catalog=isrpo1;integrated security=True";
            var converter = new JsonToDocxConverter(connectionString);

            // Convert JSON to docx
            var jsonFilePath = "C:/Users/2004y/source/repos/Klevtsov_4335_/Template_4335/isrpo.json";
            var outputFilePath = "B:/desctop2/1.docx";
            converter.ConvertJsonToDocx(jsonFilePath, outputFilePath);
        }
        public class JsonData
        {
            public int Код { get; set; }
            public string Должность { get; set; }
            public string Логин { get; set; }
        }
    }
}
