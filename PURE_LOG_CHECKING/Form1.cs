using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDataReader;
using System.Text;
using System.Collections.Generic;

namespace PURE_LOG_CHECKING
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private static Dictionary<string, string> tagNames = new Dictionary<string, string>
    {
        { "50", "Application Label" },
        { "57", "Track 2 Equivalent Data" },
        { "5A", "PAN" },
        { "5F20", "Cardholder Name" },
        { "5F24", "Application Expiration Date" },
        { "5F25", "Application Effective Date" },
        { "5F28", "Issuer Country Code" },
        { "5F2A", "currencycode" },
        { "5F2D", "Language Preference" },
        { "5F34", "Application PAN Sequence Number" },
        { "6F", "SELECT Response Message Data Field (FCI)" },
        { "71", "Critical Issuer Script Command" },
        { "72", "Non Critical Issuer Script Command" },
        { "82", "Application Interchange Profile (AIP)" },
        { "84", "Dedicated File Name (AID)" },
        { "85", "Memory Slot Update Entry Setting" },
        { "8A", "arc" },
        { "8C", "Card Risk Management Data Object List 1 (CDOL1)" },
        { "8D", "Card Risk Management Data Object List 2(CDOL2)" },
        { "8E", "Cardholder Verification Method (CVM) List" },
        { "8F", "Certification Authority Public Key Index" },
        { "90", "Issuer Public Key Certificate" },
        { "91", "iad" },
        { "92", "Issuer Public Key Remainder" },
        { "93", "Signed Static Application Data" },
        { "94", "Application File Locator (AFL)" },
        { "95", "Terminal Verification Result (TVR)" },
        { "9A", "transactionTime" },
        { "9C", "transactionType" },
        { "9F01", "Acquirer Identifier" },
        { "9F02", "transaction amount" },
        { "9F03", "Amount, Other (Numeric)" },
        { "9F07", "Application Usage Control" },
        { "9F08", "Application Version Number" },
        { "9F09", "Terminal Application Version Number" },
        { "9F0D", "Issuer Action Code - Default" },
        { "9F0E", "Issuer Action Code - Denial" },
        { "9F0F", "Issuer Action Code - Online" },
        { "9F10", "Issuer Application Data (IAD)" },
        { "9F11", "Issuer Code Table Index" },
        { "9F12", "Application Preferred Name" },
        { "9F13", "Last Online Application Transaction Counter (ATC) Register" },
        { "9F14", "Lower Consecutive Offline Limit" },
        { "9F15", "Merchant Category Code" },
        { "9F16", "Merchant Identifier" },
        { "9F1A", "Terminal Country Code" },
        { "9F1E", "Interface Device(IFD) Serial Number" },
        { "9F1F", "Track 1 Discretionary Data" },
        { "9F21", "Transaction Time" },
        { "9F23", "Upper Consecutive Offline Limit" },
        { "9F26", "Application Cryptogram (AC)" },
        { "9F27", "Cryptogram Information Data (CID)" },
        { "9F2A", "Kernel Identifier" },
        { "9F32", "Issuer Public Key Exponent" },
        { "9F33", "Terminal Capabilities" },
        { "9F34", "CVM Result" },
        { "9F35", "Terminal Type" },
        { "9F36", "Application Transaction Counter (ATC)" },
        { "9F37", "Unpredictable Number" },
        { "9F38", "Processing Options Data Object List (PDOL)" },
        { "9F42", "Application Currency Code" },
        { "9F45", "Data Authentication Code (DAC)" },
        { "9F46", "ICC Public Key Certificate" },
        { "9F47", "ICC Public Key Exponent" },
        { "9F48", "ICC Public Key Remainder" },
        { "9F49", "Dynamic Data Authentication Data Object List (DDOL)" },
        { "9F4A", "Static Data Authentication Tag List" },
        { "9F4B", "Signed Dynamic Application Data" },
        { "9F4D", "Log Entry" },
        { "9F4E", "Merchant Name and Location" },
        { "9F4F", "Log Format" },
        { "9F50", "COTA Offline Balance" },
        { "9F70", "GDDOL - List of data elements to retrieve using GET DATA" },
        { "9F71", "GDDOL Resulting Buffer" },
        { "9F72", "Memory Slot Identifier" },
        { "9F73", "Issuer Script Results" },
        { "9F74", "Data Elements Update Result" },
        { "9F75", "Echo Card Identifier" },
        { "9F76", "Terminal Transaction Data" },
        { "9F77", "Terminal Dedicated Data" },
        { "A2", "Memory Slot Update Entry Type 1" },
        { "A3", "Memory Slot Update Entry Type 2" },
        { "A4", "Memory Slot Update Entry Type 3" },
        { "A5", "File Control Information (FCI) Proprietary Template" },
        { "BF0C", "File Control Information (FCI) Issuer Discretionary Data" },
        { "BF70", "Memory Slot Update template" },
        { "BF71", "Memory Slot Read Template" },
        { "C5", "Contactless Cryptogram Information Data (CCID)" },
        { "C7", "Terminal Transaction Processing Information (TTPI)" },
        { "CD", "CRM Currency Code" },
        { "9F66", "Terminal Transaction Qualifiers (TTQ)" },
        { "0F", "Unknown Tag" }
    };

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
                    Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                    Title = "Please select PURE_Contactless_Reader_Guidelines_vX.X.xlsx"
                };
                if (openFileDialog.ShowDialog() != DialogResult.OK)
                    throw new Exception("No Excel file selected");

                string excelFilePath = openFileDialog.FileName;

                // Read the Excel file and get the values from the Guidelines sheet
                var guidelines = ReadGuidelinesSheet(excelFilePath);

                // Create a StringBuilder to accumulate the results
                StringBuilder resultBuilder = new StringBuilder();

                //// Process the guidelines table
                //ProcessGuidelines(guidelines, resultBuilder);

                // 选择文件夹
                using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
                {
                    folderBrowserDialog.Description = "Select a folder for storing PURE logs in txt format";
                    if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                    {
                        string selectedFolderPath = folderBrowserDialog.SelectedPath;

                        // 遍历文件夹和子文件夹中的TXT文件
                        var txtFiles = Directory.GetFiles(selectedFolderPath, "*.txt", SearchOption.AllDirectories);

                        // Dictionary to hold errors for each key
                        var errorDictionary = new Dictionary<string, List<string>>();
                        // 处理每个TXT文件
                        foreach (var txtFile in txtFiles)
                        {
                            try
                            {
                                string fileName = Path.GetFileNameWithoutExtension(txtFile);
                                string[] fileNameParts = fileName.Split('_');

                                // 初始化key为空字符串
                                string key = string.Empty;

                                for (int i = 0; i < fileNameParts.Length; i++)
                                {
                                    // 如果遇到以"2024"开头的部分，就停止并清除后面的部分
                                    if (fileNameParts[i].StartsWith("2024"))
                                    {
                                        break;
                                    }

                                    // 如果key不为空，则添加下划线
                                    if (!string.IsNullOrEmpty(key))
                                    {
                                        key += "_";
                                    }

                                    // 添加当前部分到key
                                    key += fileNameParts[i];
                                }

                                // 在此处读取TXT文件内容并与Excel数据进行比较
                                CompareWithExcelData(guidelines, key, txtFile, resultBuilder);
                                

                            }
                            catch { 
                                MessageBox.Show($"Error processing file: {txtFile}");
                                }
                        }

                        // 显示结果
                        if (resultBuilder.Length == 0)
                        {
                            Result.Text = "All Test Cases do not have any issues.";
                        }
                        else
                        {
                            Result.Text = resultBuilder.ToString();
                        }
                        
                        MessageBox.Show("Processing completed successfully!");
                    }
                }
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

        //private void ProcessGuidelines(DataTable guidelines, StringBuilder resultBuilder)
        //{
        //    // 从第二行开始遍历
        //    for (int i = 1; i < guidelines.Rows.Count; i++)
        //    {
        //        var row = guidelines.Rows[i];

        //        // 检查是否为空白行
        //        if (row.ItemArray.All(field => string.IsNullOrEmpty(field?.ToString())))
        //        {
        //            break;
        //        }

        //        string key = row[0].ToString().Replace(" ", "");
        //        string instruction = row[4].ToString().Replace(" ", "");

        //        string transactionAmount = "000000001000";
        //        string transactionType = "00";
        //        string transactionTime = "240604";
        //        string currencyCode = "0978";

        //        string arc = "3030"; // 默认值
        //        string iad = "0102030405060708"; // 固定值

        //        if (!string.IsNullOrEmpty(instruction))
        //        {
        //            var amountMatch = Regex.Match(instruction, @"Pleaseentertransactionamountas(\d+\.\d+)");
        //            if (amountMatch.Success)
        //            {
        //                decimal amount = decimal.Parse(amountMatch.Groups[1].Value);
        //                transactionAmount = ((int)(amount * 100)).ToString("D12");
        //            }

        //            var typeMatch = Regex.Match(instruction, @"Transactiontypeas(\d+)");
        //            if (typeMatch.Success)
        //            {
        //                transactionType = typeMatch.Groups[1].Value;
        //            }

        //            var arcMatch = Regex.Match(instruction, @"PleaseconfigurehosttosendARC");
        //            if (arcMatch.Success)
        //            {
        //                var arcValueMatch = Regex.Match(instruction, @"PleaseconfigurehosttosendARC=3030");
        //                if (arcValueMatch.Success)
        //                {
        //                    arc = "3030";
        //                }
        //                else
        //                {
        //                    arc = "3035";
        //                }
        //            }
        //        }

        //        //// 将从Excel读取的值与TXT文件进行比较
        //        //resultBuilder.AppendLine($"Key: {key}, Amount: {transactionAmount}, Type: {transactionType}, Time: {transactionTime}, Code: {currencyCode}, ARC: {arc}, IAD: {iad}");
        //    }
        //}
        private void ParseAndCompareData(DataTable guidelines, string key, string txtContent, string pdol_data, string gpo_data, StringBuilder resultBuilder)
        {
            var pdolValues = ParsePdol(pdol_data);
            var gpoValues = ParseGpoData(pdolValues, gpo_data);

            // 验证长度一致性
            int totalLength = 0;
            foreach (var (_, length) in pdolValues)
            {
                totalLength += length;
            }
            if (totalLength != gpo_data.Length / 2)
            {
                resultBuilder.AppendLine($"{key}:The length of PDOL Value and GPO Data Value is inconsistent");
                return;
            }

            // 提取TXT文件中的值
            string transactionAmount = ExtractValueFromTxt(gpoValues, "9F02");
            string transactionType = ExtractValueFromTxt(gpoValues, "9C");
            string transactionTime = ExtractValueFromTxt(gpoValues, "9A");
            string currencyCode = ExtractValueFromTxt(gpoValues, "5F2A");
            string arc = ExtractValueFromTxt(gpoValues, "8A");
            string iad = ExtractValueFromTxt(gpoValues, "9F10");
            // Append extracted values to the resultBuilder
            //resultBuilder.AppendLine($"{key}:{transactionAmount}:{transactionType}:{transactionTime}:{currencyCode}:{arc}:{iad}");
            // 比较TXT文件中的数据与Excel中的数据
            foreach (DataRow row in guidelines.Rows)
            {
                if (row[0].ToString().Replace(" ", "") == key)
                {
                    // 在此处进行比较逻辑
                    string instruction = row[4].ToString().Replace(" ", "");
                    string expectedTransactionAmount = "000000001000";
                    string expectedTransactionType = "00";
                    string expectedTransactionTime = "240604";
                    string expectedCurrencyCode = "0978";
                    string expectedArc = "3030"; // 默认值
                    string expectedIad = "0102030405060708"; // 固定值

                    if (key == "PRE_PROC_42" || key == "PRE_PROC_46" || key == "PRE_PROC_50")
                    {
                        expectedCurrencyCode = "0840";
                    }
                    else if (key == "PRE_PROC_53" || key == "PRE_PROC_55" || key == "PRE_PROC_57" || key == "PRE_PROC_59" || key == "PRE_PROC_61" || key == "PRE_PROC_63" || key == "PRE_PROC_65")
                    {
                        expectedCurrencyCode = "0056";
                    }
                    else if (key == "PDOL_PROC_002")
                    {
                        expectedTransactionAmount = "0000001000";
                        expectedTransactionTime = "";
                    }
                    else if (key == "PDOL_PROC_004" || key == "PDOL_PROC_005")
                    {
                        expectedTransactionAmount = "00000000001000";
                    }
                    else if (key == "Online_processing_038" || key == "Online_processing_039" || key == "Online_processing_083")
                    {
                        expectedTransactionAmount = "00001000";
                    }
                    else if (key == "Online_processing_040" || key == "Online_processing_085")
                    {
                        expectedTransactionAmount = "0000000000001000";
                    }
                    else if (key == "Processing_restrictions_037")
                    {
                        expectedTransactionTime = "991231";
                    }
                    else if (key == "Processing_restrictions_040")
                    {
                        expectedTransactionTime = "500101";
                    }
                    else if (transactionType == "90" && transactionAmount == "000000000000" && transactionTime == "000000" && currencyCode == "0000")
                    {
                        expectedTransactionTime = "000000";
                        expectedTransactionAmount = "000000000000";
                        expectedCurrencyCode = "0000";
                    }
                    else if (transactionType == "90" && transactionAmount == "000000000000" && transactionTime != "000000" && currencyCode != "0000")
                    {
                        //expectedTransactionTime = "000000";
                        expectedTransactionAmount = "000000000000";
                        //expectedCurrencyCode = "0000";
                    }
                    if (!string.IsNullOrEmpty(instruction))
                    {
                        var amountMatch = Regex.Match(instruction, @"Pleaseentertransactionamountas(\d+\.\d+)");
                        if (amountMatch.Success)
                        {
                            decimal amount = decimal.Parse(amountMatch.Groups[1].Value);
                            expectedTransactionAmount = ((int)(amount * 100)).ToString("D12");
                            if (key == "PDOL_PROC_002")
                            {
                                expectedTransactionAmount = "0000001000";
                                expectedTransactionTime = "";
                            }
                            else if (key == "PDOL_PROC_004" || key == "PDOL_PROC_005")
                            {
                                expectedTransactionAmount = "00000000001000";
                            }
                            else if (key == "Online_processing_038" || key == "Online_processing_039" || key == "Online_processing_083")
                            {
                                expectedTransactionAmount = "00001000";
                            }
                            else if (key == "Online_processing_040" || key == "Online_processing_085")
                            {
                                expectedTransactionAmount = "0000000000001000";
                            }
                            else if (transactionType == "90" && transactionAmount == "000000000000" && transactionTime == "000000" && currencyCode == "0000")
                            {
                                expectedTransactionTime = "000000";
                                expectedTransactionAmount = "000000000000";
                                expectedCurrencyCode = "0000";
                            }
                            else if (transactionType == "90" && transactionAmount == "000000000000" && transactionTime != "000000" && currencyCode != "0000")
                            {
                                //expectedTransactionTime = "000000";
                                expectedTransactionAmount = "000000000000";
                                //expectedCurrencyCode = "0000";
                            }
                        }

                        var typeMatch = Regex.Match(instruction, @"Transactiontypeas(\d+)");
                        if (typeMatch.Success)
                        {
                            expectedTransactionType = typeMatch.Groups[1].Value;
                        }

                        var arcMatch = Regex.Match(instruction, @"ARC");
                        if (arcMatch.Success)
                        {
                            var arcValueMatch = Regex.Match(instruction, @"ARC=3030");
                            if (arcValueMatch.Success)
                            {
                                expectedArc = "3030";
                            }
                            else
                            {
                                expectedArc = "3035";
                            }
                        }
                    }
                    // 比较逻辑
                    CompareAndAppendResult(resultBuilder, "Transaction Amount (9F02)", expectedTransactionAmount, transactionAmount, key);
                    CompareAndAppendResult(resultBuilder, "Transaction Type (9C)", expectedTransactionType, transactionType, key);
                    if (key == "Processing_restrictions_037" || key == "Processing_restrictions_040" || transactionType == "90" && transactionTime == "000000" || key == "PDOL_PROC_002")
                    {
                        CompareAndAppendResult(resultBuilder, "Transaction Date (9A)", expectedTransactionTime, transactionTime, key);
                    }
                    else
                    {
                        ValidateExtractedValue(key, "Transaction Time", transactionTime, expectedTransactionTime, resultBuilder, checkFirstTwoChars: true);
                    }                   
                    CompareAndAppendResult(resultBuilder, "Currency Code (5F2A)", expectedCurrencyCode, currencyCode, key);
                    CompareAndAppendResult(resultBuilder, "ARC (8A)", expectedArc, arc, key);
                    CompareAndAppendResult(resultBuilder, "IAD (9F10)", expectedIad, iad, key);
                }
            }
        }
        private void ValidateExtractedValue(string key, string fieldName, string extractedValue, string expectedValue, StringBuilder resultBuilder, bool checkFirstTwoChars = false)
        {
            if (checkFirstTwoChars)
            {
                if (!extractedValue.StartsWith(expectedValue.Substring(0, 2)))
                {
                    resultBuilder.AppendLine($"{key}: {fieldName} mismatch, expected value is {expectedValue.Substring(0, 2)}xxxx, but the actual value is {extractedValue}.");
                }
            }
            else
            {
                if (extractedValue != expectedValue)
                {
                    resultBuilder.AppendLine($"{key}: {fieldName} mismatch, expected {expectedValue}, but got {extractedValue}");
                }
            }
        }
        private string ExtractValueFromTxt(List<(string tag, string value)> gpoValues, string tag)
        {
            var result = gpoValues.FirstOrDefault(v => v.tag == tag);
            return Regex.Replace(result.value ?? "", @"\s|\(.*?\)", ""); // 去掉空格和括号内容
        }

        private void CompareAndAppendResult(StringBuilder resultBuilder, string fieldName, string expectedValue, string actualValue, string key)
        {
            if (!string.IsNullOrEmpty(actualValue) && expectedValue.Replace(" ", "") != actualValue)
            {
                resultBuilder.AppendLine($"{key}: Expected {fieldName} is {expectedValue}, actual value is {actualValue}.");
            }
        }
        private List<(string, int)> ParsePdol(string pdol_data)
        {
            var tags = new List<string>
        {
            "50", "57", "5A", "5F20", "5F24", "5F25", "5F28", "5F2A", "5F2D", "5F34", "6F", "71", "72", "82", "84", "85", "8A", "8C", "8D", "8E", "8F", "90", "91", "92", "93", "94", "95", "9A", "9C", "9F01", "9F02", "9F03", "9F07", "9F08", "9F09", "9F0D", "9F0E", "9F0F", "9F10", "9F11", "9F12", "9F13", "9F14", "9F15", "9F16", "9F1A", "9F1E", "9F1F", "9F21", "9F23", "9F26", "9F27", "9F2A", "9F32", "9F33", "9F34", "9F35", "9F36", "9F37", "9F38", "9F42","9F45", "9F46", "9F47", "9F48", "9F49", "9F4A", "9F4B", "9F4D", "9F4E", "9F4F", "9F50", "9F70", "9F71", "9F72", "9F73", "9F74", "9F75", "9F76", "9F77", "A2", "A3", "A4", "A5", "BF0C", "BF70", "BF71", "C5", "C7", "CD", "9F66", "0F"
        };

            var pdolValues = new List<(string, int)>();
            int i = 0;
            while (i < pdol_data.Length)
            {
                string tag = pdol_data.Substring(i, 4);
                if (!tags.Contains(tag))
                {
                    tag = pdol_data.Substring(i, 2);
                    string lengthHex = pdol_data.Substring(i + 2, 2);
                    int length = int.Parse(lengthHex, System.Globalization.NumberStyles.HexNumber);
                    pdolValues.Add((tag, length));
                    i += 4;
                }
                else
                {
                    string lengthHex = pdol_data.Substring(i + 4, 2);
                    int length = int.Parse(lengthHex, System.Globalization.NumberStyles.HexNumber);
                    pdolValues.Add((tag, length));
                    i += 6;
                }
            }

            return pdolValues;
        }

        private List<(string, string)> ParseGpoData(List<(string, int)> pdolValues, string gpo_data)
        {
            var gpoValues = new List<(string, string)>();
            int i = 0;
            foreach (var (tag, length) in pdolValues)
            {
                string value = gpo_data.Substring(i, length * 2);
                string formattedValue = string.Join(" ", Regex.Split(value, "(?<=\\G..)(?=.)"));
                int valueLength = value.Length / 2;
                string byteLabel = valueLength == 1 ? "Byte" : "Bytes";
                formattedValue += $" ({valueLength} {byteLabel})";
                gpoValues.Add((tag, formattedValue));
                i += length * 2;
            }

            return gpoValues;
        }
        private void CompareWithExcelData(DataTable guidelines, string key, string txtFilePath, StringBuilder resultBuilder)
        {
            string txtContent = File.ReadAllText(txtFilePath).Replace(" ", "").Replace("\r", "").Replace("\n", "");

            string pdol_data = "";
            string gpo_data = "";


            // 优先获取80A8到cla之间的数据
            var gpoMatch80A8 = Regex.Match(txtContent, @"80A8(.*?)cla", RegexOptions.Singleline);
            if (gpoMatch80A8.Success)
            {
                // 获取并处理匹配结果
                string gpoSegment80A8 = gpoMatch80A8.Groups[1].Value.Replace(" ", "").Replace("cla", ""); // 清除cla

                // 检查前7个字节是否包含81
                string firstSevenBytes = gpoSegment80A8.Length >= 14 ? gpoSegment80A8.Substring(0, 10) : gpoSegment80A8;
                if (firstSevenBytes.Contains("81"))
                {
                    gpoSegment80A8 = firstSevenBytes.Replace("81", "") + gpoSegment80A8.Substring(10, gpoSegment80A8.Length - 12);
                }

                if (gpoSegment80A8.Length > 16 && !firstSevenBytes.Contains("81"))
                {
                    gpo_data = gpoSegment80A8.Substring(10, gpoSegment80A8.Length - 12);
                }
                else if(gpoSegment80A8.Length > 16 && firstSevenBytes.Contains("81"))
                {
                    gpo_data = gpoSegment80A8.Substring(10, gpoSegment80A8.Length - 10);
                }

                // 获取9F38的数据
                var pdolMatch9F38 = Regex.Match(txtContent, @"<li>9F38(.*?)</li>");
                if (pdolMatch9F38.Success)
                {
                    string pdolSegment9F38 = pdolMatch9F38.Groups[1].Value.Replace(" ", "").Replace("ProcessingOptionsDataObjectList(PDOL)", "");
                    if (pdolSegment9F38.Length > 2)
                    {
                        pdol_data = pdolSegment9F38.Substring(2);
                        if (key != "PDOL_PROC_002")
                        {
                            if (!pdol_data.Contains("9F02") || !pdol_data.Contains("9A") || !pdol_data.Contains("9C") || !pdol_data.Contains("5F2A"))
                            {
                                gpo_data = "";
                                var gpoMatch = Regex.Matches(txtContent, @":80AE(.*?)cla", RegexOptions.Singleline);
                                if (gpoMatch.Count > 0)
                                {
                                    foreach (Match match in gpoMatch)
                                    {
                                        string gpoSegment = match.Groups[1].Value.Replace(" ", "").Replace(":", "");
                                        if (!gpoSegment.Contains("3030") || !gpoSegment.Contains("3035"))
                                        {
                                            gpo_data = gpoSegment;
                                            if (gpo_data.Length > 12)
                                            {
                                                gpo_data = gpo_data.Substring(6, gpo_data.Length - 8);
                                            }// 获取8C的数据
                                            var pdolMatch = Regex.Match(txtContent, @"<li>8C(.*?)</li>");
                                            if (pdolMatch.Success)
                                            {
                                                pdol_data = pdolMatch.Groups[1].Value.Replace(" ", "").Replace("CardRiskManagementDataObjectList1(CDOL1)","");
                                                if (pdol_data.Length > 4)
                                                {
                                                    pdol_data = pdol_data.Substring(2);
                                                }
                                            }
                                        }

                                    }
                                }
                            }

                        }
                    }
                }
            }
         
            var gpoMatches = Regex.Matches(txtContent, @":80AE(.*?)cla", RegexOptions.Singleline);
            if (gpoMatches.Count > 0)
            {
                foreach (Match match in gpoMatches)
                {
                    string gpoSegment = match.Groups[1].Value.Replace(" ", "").Replace(":", "");
                    
                    if (gpoSegment.Contains("3030") || gpoSegment.Contains("3035"))
                    {
                        gpo_data = gpoSegment;
                        if (gpo_data.Length > 12)
                        {
                            gpo_data = gpo_data.Substring(6, gpo_data.Length - 8);
                        }

                        // 获取8D的数据
                        var pdolMatch = Regex.Match(txtContent, @"<li>8D(.*?)</li>");
                        if (pdolMatch.Success)
                        {
                            pdol_data = pdolMatch.Groups[1].Value.Replace(" ", "").Replace("CardRiskManagementDataObjectList2(CDOL2)", "");
                            if (pdol_data.Length > 4)
                            {
                                pdol_data = pdol_data.Substring(2);
                            }
                        }
                    }
                }
                

          
               
            }  // 比较TXT文件中的数据与Excel中的数据
               // 判断 pdol_data 和 gpo_data 是否为空
            if (string.IsNullOrEmpty(pdol_data) || string.IsNullOrEmpty(gpo_data))
            {
                resultBuilder.AppendLine($"{key}: Can not check the log.");
                return;
            }
            // 解析并比较数据
            if (!string.IsNullOrEmpty(gpo_data) && !string.IsNullOrEmpty(pdol_data))
            {
                ParseAndCompareData(guidelines, key, txtContent, pdol_data, gpo_data, resultBuilder);
            }

        }
                }
            }
                   