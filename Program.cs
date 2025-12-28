internal class Program
{
    private static void Main()
    {
        // Excel2CSharp
        string[] cmd_list = ["fl", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "--help"];
        string cmd_key = "etc";
        Console.WriteLine("################Excel2CSharp################");

        do
        {
            Console.Write("请输入命令：");
            var input = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(input))
            {
                continue;
            }
            var cmd_group = input.Split();
            if(cmd_group[0] != cmd_key)
            {
                Console.WriteLine($"{cmd_group[0]}:未找到的命令");
                continue;
            }
            if(cmd_group.Length == 1)
            {
                ExcuteCmd("--help");
                continue;
            }
            for(int i = 1; i < cmd_group.Length; i++)
            {
                var cmd = cmd_group[i];
                if (cmd_list.Contains(cmd))
                {
                    ExcuteCmd(cmd);
                }
                else
                {

                    Console.WriteLine($"etc:‘{cmd_group[i]}’不是etc命令，请使用‘etc --help’查看帮助");
                }
            }
        } while (true);
    }

    private static void ExcuteCmd(string cmd)
    {
        Console.WriteLine($"输入命令为{cmd}");
    }
}