using Excel2CSharp.Tools;

internal abstract class Program
{
    private static readonly Dictionary<string, string> cmd_map = new ()
    {
        { "efl", "查看所有excel文件" },
        { "esl", "查看excel所有工作簿sheet 参数是excel文件名(带后缀名)" },
        { "", "" },
        { " ", " " },
        { "export", "导出配置表 \n    第一个参数:\n        all:导出全部; \n        excel文件名(带后缀名):导出对应文件名所有工作簿sheet; \n        excel文件名(带后缀名)-工作簿sheet名:导出对应文件名的对应工作簿sheet名\n    第二个参数:\n        json: 导出json格式数据\n        xml: 导出xml格式数据" },
        { "exit", "退出命令行模式" },
        { "help", "查看所有命令" },
    };

    private static readonly string[] options = [ "手动模式", "命令行模式", ];

    private const string cmd_key = "etc";
    private const string Indicator = "* "; // 前导符

    public enum OperateMode
    {
        None,
        Manual,
        Command,
    }

    private static OperateMode operate = OperateMode.None;
    
    private static int startLine = 3;
    
    private static readonly List<string> manualOptions = [];
    
    // 表示当前所选
    private static int currentIndex = -1;
    // 表示前一个选项
    private static int previousIndex = -1;
    
    private static Dictionary<string, List<ConfigData>> excels = new();

    private static void Main()
    {
        // Excel2CSharp
        var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excel");
        if (!Directory.Exists(path))
        {
            Console.WriteLine("Not Found Excel Path: " + path);
            return;
        }
        excels = ExcelTools.GetAllExcelData(path);
        
        InitMenu(operate,  options);
        
        do
        {
            if (operate == OperateMode.None)
            {
                var key = Console.ReadKey(true);
                if (key is { Key: ConsoleKey.UpArrow })
                {
                    previousIndex = currentIndex; // 保存前一个被选项索引
                    currentIndex--;
                }
                if (key is { Key: ConsoleKey.DownArrow })
                {
                    previousIndex = currentIndex;
                    currentIndex++;
                }
                var index = currentIndex - startLine + 1;
                if (key is { Key: ConsoleKey.Enter })
                {
                    if (index < 1) index = 1;
                    if (options != null && index > options.Length) index = options.Length;
                    operate = (OperateMode)index;
                    if (operate != OperateMode.None)
                    {
                        if (operate == OperateMode.Manual)
                        {
                            InitManualMenuData(0);
                            InitMenu(operate, manualOptions.ToArray());
                        }
                        if (operate == OperateMode.Command)
                        {
                            InitMenu(operate, null);
                            Console.CursorVisible = true;
                        }
                    }
                    else
                    {
                        InitMenu(operate, options);
                    }

                    continue;
                }

                // 先清除前一个选项的标记
                if (previousIndex > -1 && options != null && previousIndex < options.Length + startLine)
                {
                    Console.SetCursorPosition(0, previousIndex);
                    Console.ResetColor();
                    var s = options[previousIndex - startLine];
                    var menu = $"[{Array.IndexOf(options, s)}] {s}";
                    Console.Write("".PadLeft(Indicator.Length, ' ') + menu);
                }

                // 再看看当前项有没有超出范围
                if (currentIndex < startLine) currentIndex = startLine;
                if (options != null && currentIndex > options.Length + startLine - 1) currentIndex = options.Length + startLine - 1;
                Console.BackgroundColor = ConsoleColor.Blue;    // 背景蓝色
                // 设置当前选择项的标记
                Console.SetCursorPosition(0, currentIndex);

                if (options != null)
                {
                    var menu = $"[{Array.IndexOf(options, options[currentIndex - startLine])}] {options[currentIndex - startLine]}";
                    Console.WriteLine($"{Indicator}{menu}");
                }
            }
            if (operate == OperateMode.Command)
            {
                Command();
            }
            if (operate == OperateMode.Manual)
            {
                Manual();
            }
        } while (true);
        // ReSharper disable once FunctionNeverReturns
    }
    
    private static void WriteTitle(OperateMode mode)
    {
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("################Excel2CSharp################");
        switch (mode)
        {
            case OperateMode.None:
            case OperateMode.Manual:
                Console.WriteLine("#       ↑↓ 选择操作, Enter 确认选择        #");
                break;
            case OperateMode.Command:
                Console.WriteLine("# 1. 命令以etc开头, 后面加空格和具体命令   #");
                Console.WriteLine("# 2. 输入'exit'按回车键退出命令行模式      #");
                Console.WriteLine("# 3. 输入'help'按回车键查看所有命令        #");
                break;
            default:
                break;
        }
        Console.WriteLine("############################################");
        Console.ResetColor();
        startLine = Console.CursorTop;
    }

    private static void InitMenu(OperateMode mode, string[]? menuOptions)
    {
        Console.ResetColor();
        Console.Clear();
        Console.SetCursorPosition(0, 0);
        currentIndex = -1;
        previousIndex = -1;
        WriteTitle(mode);
        // 下面这行是隐藏光标，这样好看一些
        Console.CursorVisible = false;
        if (menuOptions == null) return;
        // 先输出选项
        foreach (var s in menuOptions)
        {
            var menu = $"[{Array.IndexOf(menuOptions, s)}] {s}";
            Console.WriteLine(menu.PadLeft(Indicator.Length + menu.Length));
        }
    }
    
    private static int manualLayer = 0;
    private static void InitManualMenuData(int layer, string key =  "")
    {
        manualLayer = layer;
        manualOptions.Clear();
        switch (layer)
        {
            case 0:
                manualOptions.Add("导出所有excel表");
                foreach (var file in excels.Keys)
                {
                    manualOptions.Add(file);
                }
                manualOptions.Add("退出");
                break;
            case 1:
                manualOptions.Add($"导出[{key}]表");
                excelFile = key;
                foreach (var data in excels[key])
                {
                    manualOptions.Add(data.name);
                }
                manualOptions.Add("返回");
                break;
            case 2:
                excelSheet = key;
                manualOptions.Add("Json");
                manualOptions.Add("Xml");
                manualOptions.Add("返回");
                break;
        }
    }
    private static string excelFile = "";
    private static string excelSheet = "";
    private static void Manual()
    {
        var key = Console.ReadKey(true);
        if (key is { Key: ConsoleKey.UpArrow })
        {
            previousIndex = currentIndex; // 保存前一个被选项索引
            currentIndex--;
        }
        if (key is { Key: ConsoleKey.DownArrow })
        {
            previousIndex = currentIndex;
            currentIndex++;
        }
        var index = currentIndex - startLine;
        if (key is { Key: ConsoleKey.Enter })
        {
            if (index < 0) index = 0;
            if (index > manualOptions.Count - 1) index = manualOptions.Count - 1;
            
            if (index == 0)
            {
                switch (manualLayer)
                {
                    case 0:
                        if (manualOptions[index].EndsWith(".xlsx"))
                        {
                            InitManualMenuData(1, manualOptions[index]);
                            InitMenu(operate, manualOptions.ToArray());
                        }
                        else
                        {
                            ExecuteCmd("export", "all");
                        }
                        return;
                    case 1:
                        InitManualMenuData(2);
                        InitMenu(operate, manualOptions.ToArray());
                        return;
                }
            }
            if (index == manualOptions.Count - 1)
            {
                switch (manualLayer)
                {
                    case 2:
                        InitManualMenuData(1, excelFile);
                        InitMenu(operate, manualOptions.ToArray());
                        return;
                    case 1:
                        InitManualMenuData(0);
                        InitMenu(operate, manualOptions.ToArray());
                        return;
                    case 0:
                        operate = OperateMode.None;
                        InitMenu(operate,  options);
                        return;
                }
                return;
            }

            switch (manualLayer)
            {
                case 2:
                    var args = string.IsNullOrEmpty(excelSheet) ? $"{excelFile} {manualOptions[index].ToLower()}" : $"{excelFile}-{excelSheet} {manualOptions[index].ToLower()}";
                    ExecuteCmd("export", args);
                    return;
                case 1:
                    InitManualMenuData(2, manualOptions[index]);
                    InitMenu(operate, manualOptions.ToArray());
                    return;
                case 0:
                    InitManualMenuData(1, manualOptions[index]);
                    InitMenu(operate, manualOptions.ToArray());
                    return;
            }
        }

        // 先清除前一个选项的标记
        if (previousIndex > -1 && previousIndex < manualOptions.Count + startLine)
        {
            Console.SetCursorPosition(0, previousIndex);
            Console.ResetColor();
            var s = manualOptions[previousIndex - startLine];
            var _menu = $"[{manualOptions.IndexOf(s)}] {s}";
            Console.Write("".PadLeft(Indicator.Length, ' ') + _menu);
        }

        // 再看看当前项有没有超出范围
        if (currentIndex < startLine) currentIndex = startLine;
        if (currentIndex > manualOptions.Count + startLine - 1) currentIndex = manualOptions.Count + startLine - 1;
        Console.BackgroundColor = ConsoleColor.Blue;    // 背景蓝色
        // 设置当前选择项的标记
        Console.SetCursorPosition(0, currentIndex);
        
        var menu = $"[{manualOptions.IndexOf(manualOptions[currentIndex - startLine])}] {manualOptions[currentIndex - startLine]}";
        Console.WriteLine($"{Indicator}{menu}");
    }

    private static void Command()
    {
        var list = cmd_map.Keys.ToArray();
        Console.Write("请输入命令: ");
        var input = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(input))
        {
            return;
        }

        var cmd_group = input.Split(' ');
        if (cmd_group[0] != cmd_key)
        {
            Console.WriteLine($"{cmd_group[0]}: 未找到的命令");
            return;
        }

        if (cmd_group.Length == 1)
        {
            ExecuteCmd($"{list[^1]}");
            return;
        }

        var cmd = cmd_group[1];
        var args = cmd_group.Skip(2).ToArray();
        if (!list.Contains(cmd))
        {
            Console.WriteLine($"etc: '{cmd}'不是etc命令,请使用'etc {list[^1]}'查看帮助");
            return;
        }

        ExecuteCmd(cmd, args);
    }

    private static void ExecuteCmd(string cmd, params string[] args)
    {
        switch (cmd)
        {
            case "":
            case " ":
            case "help":
            {
                var list_key = cmd_map.Keys.ToArray();
                var list_value = cmd_map.Values.ToArray();
                for (var i = 0; i < cmd_map.Keys.ToArray().Length; i++)
                {
                    if (!string.IsNullOrWhiteSpace(cmd_map.Keys.ToArray()[i]))
                    {
                        Console.WriteLine($" -- {list_key[i]}: {list_value[i]}");
                    }
                }
                return;
            }
            case "exit":
                operate = OperateMode.None;
                Console.Clear();
                Console.SetCursorPosition(0, 0);
                InitMenu(operate,  options);
                return;
            case "efl":
            {
                foreach (var file in excels.Keys)
                {
                    Console.WriteLine(file);
                }
                return;
            }
            case "esl":
            {
                if (args.Length == 0)
                {
                    Console.WriteLine("excel文件名(带后缀名)不能为空");
                    return;
                }
                var arg = args[0];
                if (excels.TryGetValue(arg, out var excel))
                {
                    foreach (var data in excel)
                    {
                        Console.WriteLine(data.name);
                    }
                }
                else
                {
                    Console.WriteLine($"未找到excel文件:{arg}");
                }
                return;
            }
            case "export":
            {
                // Console.WriteLine($"输入命令为{cmd}, 参数为{string.Join(" ", args)}");
                if (args.Length < 2)
                {
                    Console.WriteLine("至少包含两个参数,第一个为导表参数.第二个为数据类型");
                    return;
                }
                var arg0 = args[0];
                var arg1 = args[1];
                var mold = Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(arg1);
                if (!Enum.TryParse<BuildTools.ConfigMold>(mold, out var configMold))
                {
                    Console.WriteLine("第二个参数不是正确的数据类型");
                }
                if (!arg0.Contains('-'))
                {
                    if (arg0 == "all")
                    {
                        foreach (var data in excels.SelectMany(excel => excel.Value))
                        {
                            if (data.excelData.datas.GetLength(1) < 5) continue;
                            if (data.excelData.datas.GetLength(0) < 7) continue;
                            BuildConfig(data.excelData, configMold);
                        }
                    }
                    else if (arg0.EndsWith(".xlsx"))
                    {
                        foreach (var data in excels[arg0].Where(data => data.excelData.datas.GetLength(1) >= 5).Where(data => data.excelData.datas.GetLength(0) >= 7))
                        {
                            BuildConfig(data.excelData, configMold);
                        }
                    }
                    else
                    {
                        Console.WriteLine("第一个参数不是正确的表格名称");
                    }
                }
                else
                {
                    var arg_split = arg0.Split('-');
                    if (arg_split.Length != 2)
                    {
                        Console.WriteLine("第一个参数不正确");
                        return;
                    }
                    if (!arg_split[0].EndsWith(".xlsx"))
                    {
                        Console.WriteLine("第一个参数不是正确的表格名称");
                        return;
                    }
                    var data = excels[arg_split[0]].FirstOrDefault(data => data.name == arg_split[1]);
                    BuildConfig(data.excelData, configMold);
                }
                return;
            }
        }
    }
    
    private static void BuildConfig(ExcelData data, BuildTools.ConfigMold mold)
    {
        BuildTools.CreateCSharp(data, mold);
        switch (mold)
        {
            case BuildTools.ConfigMold.Json:
                BuildTools.CreateJsonConfig(data);
                break;
            case BuildTools.ConfigMold.Xml:
                BuildTools.CreateXmlConfig(data);
                break;
            default:
                break;
        }
    }
}