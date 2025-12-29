using OfficeOpenXml;

namespace Excel2CSharp.Tools;

public static class ExcelTools
{
    /// <summary>
    /// 获取文件夹下所有 Excel 文件
    /// </summary>
    /// <param name="basePath"></param>
    /// <param name="extension"></param>
    /// <returns></returns>
    private static List<FileInfo> GetFiles(string basePath, string extension = "xlsx")
    {
        if (!Directory.Exists(basePath)) return [];
        var directoryInfo = new DirectoryInfo(basePath);
        var files = directoryInfo.GetFiles($"*.{extension}", SearchOption.AllDirectories);
        return files.Where(fi => !fi.Name.Contains("~$")).ToList();
    }

    /// <summary>
    /// 读取 Excel 内容
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
    public static List<ExcelData> ReadExcel(string filePath)
    {
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Not Found Excel: {filePath}");
        }
        var fileInfo = new FileInfo(filePath);
        var excel = new List<ExcelData>();
        using var package = new ExcelPackage(fileInfo);
        foreach (var worksheet in package.Workbook.Worksheets)
        {
            var data = new ExcelData
            {
                sheetName = worksheet.Name,
                datas = new object[worksheet.Dimension.End.Row + 1, worksheet.Dimension.End.Column + 1]
            };
            for (int c = worksheet.Dimension.Start.Column, c1 = worksheet.Dimension.End.Column; c <= c1; c++)
            {
                for (int r = worksheet.Dimension.Start.Row, r1 = worksheet.Dimension.End.Row; r <= r1; r++)
                {
                    data.datas[r, c] = worksheet.GetValue(r, c);
                }
            }
            excel.Add(data);
        }

        return excel;
    }

    public static Dictionary<string, List<ConfigData>> GetAllExcelData(string basePath)
    {
        var excels = new Dictionary<string, List<ConfigData>>();
        var files = GetFiles(basePath);
        foreach (var file in files)
        {
            var excelDatas = ReadExcel(file.FullName);
            var list = (from data in excelDatas
                where !data.sheetName.Contains('#')
                select new ConfigData
                {
                    name = data.sheetName, excelData = data
                }).ToList();
            excels.Add(file.Name, list);
        }

        return excels;
    }
}

public struct ExcelData
{
    public string sheetName;
    public object[,] datas;
}

public struct ConfigData
{
    public string? name;
    public ExcelData excelData;
}