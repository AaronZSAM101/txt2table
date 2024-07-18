using OfficeOpenXml;
using System.Text.RegularExpressions;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length > 0)
        {
            string filePath = args[0];

            if (File.Exists(filePath))
            {
                // 读取文本文件内容
                string content;
                using (StreamReader sr = new StreamReader(filePath))
                {
                    content = sr.ReadToEnd();
                }

                // 处理文本内容
                string processedContent = ProcessTextContent(content);

                // 将处理后的内容写入 Excel
                WriteToExcel(processedContent);

                Console.WriteLine($"处理完成，Excel 文件已保存。");
            }
            else
            {
                Console.WriteLine($"指定的文件不存在：{filePath}");
            }
        }
        else
        {
            Console.WriteLine("请将要处理的文本文件拖放到此可执行文件上。");
        }
    }

    static string ProcessTextContent(string content)
    {
        // 按顺序替换指定文本
        content = content.Replace("趋势信号机器人", "趋势信号机器人:");
        content = content.Replace("趋势 : ", "趋势:");

        // 使用正则表达式按顺序删除指定行
        content = Regex.Replace(content, @"RSI.*\n", "");
        content = Regex.Replace(content, @"MaSlope.*\n", "");
        content = Regex.Replace(content, @"涨跌幅.*\n", "");
        content = Regex.Replace(content, @" #.*\n", "");
        content = Regex.Replace(content, @" -.*\n", "");
        content = Regex.Replace(content, @", *.*\n", "\n");

        return content;
    }

    static void WriteToExcel(string processedContent)
    {
        // 将内容按照“趋势信号机器人:”分割
        string[] entries = Regex.Split(processedContent.Trim(), "(?=趋势信号机器人:)");

        // 初始化表头和表格数据列表
        List<string[]> tableData = new List<string[]>();
        string[] header = { "趋势信号机器人", "策略编号", "交易对象", "触发进场时间", "进场价格", "趋势" };
        tableData.Add(header);

        // 处理每个条目，并添加到表格数据中
        foreach (string entry in entries)
        {
            if (!string.IsNullOrWhiteSpace(entry))
            {
                string[] lines = entry.Trim().Split('\n');
                List<string> rowData = new List<string>();
                foreach (string line in lines)
                {
                    if (line.Contains(":"))
                    {
                        string[] parts = line.Split(new[] { ':' }, 2);
                        if (parts.Length > 1)
                        {
                            rowData.Add(parts[1].Trim());
                        }
                    }
                }
                tableData.Add(rowData.ToArray());
            }
        }

        // 创建一个新的 Workbook
        using (ExcelPackage excelPackage = new ExcelPackage())
        {
            ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add(DateTime.Now.ToString("yyyy-MM-dd"));

            // 将数据写入 Excel 表格
            int rowIndex = 1;
            foreach (string[] row in tableData)
            {
                for (int colIndex = 0; colIndex < row.Length; colIndex++)
                {
                    sheet.Cells[rowIndex, colIndex + 1].Value = row[colIndex];
                }
                rowIndex++;
            }

            // 生成当前时刻的文件名
            string currentTime = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string filename = $"output_table_{currentTime}.xlsx";

            // 保存 Excel 文件
            FileInfo excelFile = new FileInfo(filename);
            excelPackage.SaveAs(excelFile);

            Console.WriteLine($"Excel 文件 \"{filename}\" 已保存。");
        }
    }
}
