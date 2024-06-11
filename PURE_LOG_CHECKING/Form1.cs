using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDataReader;
using System.Text; // 确保包含这个命名空间

namespace PURE_LOG_CHECKING
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // 清空文本框内容
                Result.Clear();

                // Open file dialog to select Excel file
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
                };
                if (openFileDialog.ShowDialog() != DialogResult.OK)
                    throw new Exception("No Excel file selected");

                string excelFilePath = openFileDialog.FileName;

                // Read the Excel file and get the values from the Guidelines sheet
                var guidelines = ReadGuidelinesSheet(excelFilePath);

                // Create a StringBuilder to accumulate the results
                StringBuilder resultBuilder = new StringBuilder();

                // Process the guidelines table
                ProcessGuidelines(guidelines, resultBuilder);

                // Set the accumulated results to the Result TextBox
                Result.Text = resultBuilder.ToString();

                MessageBox.Show("Processing completed successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private DataTable ReadGuidelinesSheet(string excelFilePath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    var guidelinesSheet = result.Tables["Guidelines"];
                    return guidelinesSheet;
                }
            }
        }

        private void ProcessGuidelines(DataTable guidelines, StringBuilder resultBuilder)
        {
            // 从第二行开始遍历
            for (int i = 1; i < guidelines.Rows.Count; i++)
            {
                var row = guidelines.Rows[i];

                // 检查是否为空白行
                if (row.ItemArray.All(field => string.IsNullOrEmpty(field?.ToString())))
                {
                    break;
                }

                string key = row[0].ToString().Replace(" ", "");
                string instruction = row[4].ToString().Replace(" ", "");

                string transactionAmount = "000000001000";
                string transactionType = "00";
                string transactionTime = "240604";
                string currencycode = "0978";

                string arc = "3030"; // 默认值
                string iad = "0102030405060708"; // 固定值

                if (!string.IsNullOrEmpty(instruction))
                {
                    var amountMatch = Regex.Match(instruction, @"Pleaseentertransactionamountas(\d+\.\d+)");
                    if (amountMatch.Success)
                    {
                        decimal amount = decimal.Parse(amountMatch.Groups[1].Value);
                        transactionAmount = ((int)(amount * 100)).ToString("D12");
                    }

                    var typeMatch = Regex.Match(instruction, @"Transactiontypeas(\d+)");
                    if (typeMatch.Success)
                    {
                        transactionType = typeMatch.Groups[1].Value;
                    }

                    var arcMatch = Regex.Match(instruction, @"PleaseconfigurehosttosendARC");
                    if (arcMatch.Success)
                    {
                        var arcValueMatch = Regex.Match(instruction, @"PleaseconfigurehosttosendARC=3030");
                        if (arcValueMatch.Success)
                        {
                            arc = "3030";
                        }
                        else
                        {
                            arc = "3035";
                        }
                    }
                }

                // 构建输出字符串
                string output = $"{key}: {transactionAmount}, {transactionType}, {currencycode}, {transactionTime}, ARC: {arc}, IAD: {iad}";

                // 将结果追加到 StringBuilder
                resultBuilder.AppendLine(output);
            }
        }
    }
}

