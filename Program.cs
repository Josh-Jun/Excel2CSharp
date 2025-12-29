using Excel2CSharp.Tools;

internal abstract class Program
{
    private static readonly Dictionary<string, string> cmd_map = new ()
    {
        { "efl", "查看所有excel文件" },
        { "esl", "查看excel所有工作簿sheet 参数是excel文件名(带后缀名)" },
        { "", "" },
        { "export", "导出配置表 \n    all:导出全部; \n    excel文件名(带后缀名):导出对应文件名所有工作簿sheet; \n    excel文件名(带后缀名)-工作簿sheet名:导出对应文件名的对应工作簿sheet名" },
        { "exit", "退出命令行模式" },
        { "help", "查看所有命令" },
    };

    private static readonly string[]? options = [ "手动模式", "命令行模式", ];

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
                if (key is { Key: ConsoleKey.Enter })
                {
                    var index = currentIndex - startLine + 1;
                    if (index < 0) index = 0;
                    if (options != null && index > options.Length - 1) index = options.Length - 1;
                    operate = (OperateMode)index;
                    Console.CursorVisible = true;
                    Console.ResetColor();
                    Console.Clear();
                    Console.SetCursorPosition(0, 0);
                    currentIndex = -1;
                    previousIndex = -1;
                    if (operate != OperateMode.None)
                    {
                        if (operate == OperateMode.Manual)
                        {
                            manualOptions.Clear();
                            foreach (var file in excels.Keys)
                            {
                                manualOptions.Add(file);
                            }
                            manualOptions.Add("退出");
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
                    var menu = $"({Array.IndexOf(options, s)}) {s}";
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
                    var menu = $"({Array.IndexOf(options, options?[currentIndex - startLine])}) {options?[currentIndex - startLine]}";
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
        WriteTitle(mode);
        // 下面这行是隐藏光标，这样好看一些
        Console.CursorVisible = false;
        if (menuOptions == null) return;
        // 先输出选项
        foreach (var s in menuOptions)
        {
            var menu = $"({Array.IndexOf(menuOptions, s)}) {s}";
            Console.WriteLine(menu.PadLeft(Indicator.Length + menu.Length));
        }
    }
    
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
        if (key is { Key: ConsoleKey.Enter })
        {
            var index = currentIndex - startLine + 1;
            if (index < 0) index = 0;
            if (index > manualOptions.Count) index = manualOptions.Count;
            
            if (index == manualOptions.Count)
            {
                operate = OperateMode.None;
                Console.ResetColor();
                Console.Clear();
                Console.SetCursorPosition(0, 0);
                InitMenu(operate,  options);
                currentIndex = -1;
                previousIndex = -1;
                return;
            }
            var option = manualOptions[index];
        }

        // 先清除前一个选项的标记
        if (previousIndex > -1 && previousIndex < manualOptions.Count + startLine)
        {
            Console.SetCursorPosition(0, previousIndex);
            Console.ResetColor();
            var s = manualOptions[previousIndex - startLine];
            var _menu = $"({manualOptions.IndexOf(s)}) {s}";
            Console.Write("".PadLeft(Indicator.Length, ' ') + _menu);
        }

        // 再看看当前项有没有超出范围
        if (currentIndex < startLine) currentIndex = startLine;
        if (currentIndex > manualOptions.Count + startLine - 1) currentIndex = manualOptions.Count + startLine - 1;
        Console.BackgroundColor = ConsoleColor.Blue;    // 背景蓝色
        // 设置当前选择项的标记
        Console.SetCursorPosition(0, currentIndex);
        
        var menu = $"({manualOptions.IndexOf(manualOptions[currentIndex - startLine])}) {manualOptions[currentIndex - startLine]}";
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
        var list_key = cmd_map.Keys.ToArray();
        var list_value = cmd_map.Values.ToArray();
        if (cmd == "help")
        {
            for (var i = 0; i < cmd_map.Keys.ToArray().Length; i++)
            {
                if (!string.IsNullOrEmpty(cmd_map.Keys.ToArray()[i]))
                {
                    Console.WriteLine($" -- {list_key[i]}: {list_value[i]}");
                }
            }
            return;
        }
        if (cmd == "exit")
        {
            operate = OperateMode.None;
            Console.Clear();
            Console.SetCursorPosition(0, 0);
            InitMenu(operate,  options);
            return;
        }
        Console.WriteLine($"输入命令为{cmd}, 参数为{string.Join(" ", args)}");
        if (args.Length == 0)
        {
            return;
        }
    }
}