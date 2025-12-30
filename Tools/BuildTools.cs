using System.Diagnostics;
using System.Net.Mime;
using System.Text;

namespace Excel2CSharp.Tools;

public class BuildTools
{
    public enum ConfigMold
    {
        Json,
        Xml
    }
    
    private static readonly string _basePath = new DirectoryInfo("../..").FullName;
    public static void CreateCSharp(ExcelData data, ConfigMold mold)
    {
        var stringBuilder = new StringBuilder();

        stringBuilder.AppendLine("using System;");
        stringBuilder.AppendLine("using UnityEngine;");
        stringBuilder.AppendLine("using App.Core.Tools;");
        stringBuilder.AppendLine("using App.Core.Helper;");
        stringBuilder.AppendLine("using System.Collections.Generic;");
        stringBuilder.AppendLine("using System.Xml.Serialization;");
        stringBuilder.AppendLine("");
        stringBuilder.AppendLine("namespace App.Core.Master");
        stringBuilder.AppendLine("{");
        stringBuilder.AppendLine("    [Config]");
        stringBuilder.AppendLine(
            $"    public class {data.sheetName}{mold}Config : Singleton<{data.sheetName}{mold}Config>, IConfig");
        stringBuilder.AppendLine("    {");
        stringBuilder.AppendLine(
            $"        private {data.sheetName}{mold}Data _data = new {data.sheetName}{mold}Data();");
        stringBuilder.AppendLine(
            $"        private readonly Dictionary<int, {data.sheetName}> _dict = new Dictionary<int, {data.sheetName}>();");
        stringBuilder.AppendLine(
            $"        private const string location = \"Assets/Bundles/Builtin/Configs/{mold}/{data.sheetName}{mold}Data.{mold.ToString().ToLower()}\";");
        stringBuilder.AppendLine("        public void Load()");
        stringBuilder.AppendLine("        {");
        stringBuilder.AppendLine(
            $"            var textAsset = AssetsMaster.Instance.LoadAssetSync<TextAsset>(location);");
        switch (mold)
        {
            case ConfigMold.Json:
                stringBuilder.AppendLine(
                    $"            _data = JsonUtility.FromJson<{data.sheetName}{mold}Data>(textAsset.text);");
                break;
            case ConfigMold.Xml:
                stringBuilder.AppendLine(
                    $"            _data = XmlTools.ProtoDeSerialize<{data.sheetName}{mold}Data>(textAsset.bytes);");
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(mold), mold, null);
        }

        stringBuilder.AppendLine($"            foreach (var data in _data.{data.sheetName}s)");
        stringBuilder.AppendLine("            {");
        stringBuilder.AppendLine("                _dict.Add(data.Id, data);");
        stringBuilder.AppendLine("            }");
        stringBuilder.AppendLine("        }");
        stringBuilder.AppendLine($"        public {data.sheetName} Get(int id)");
        stringBuilder.AppendLine("        {");
        stringBuilder.AppendLine("            _dict.TryGetValue(id, out var value);");
        stringBuilder.AppendLine("            if (value == null)");
        stringBuilder.AppendLine("            {");
        stringBuilder.AppendLine(
            $"                throw new Exception($\"找不到config数据,表名=[{{nameof({data.sheetName}{mold}Config)}}],id=[{{id}}]\");");
        stringBuilder.AppendLine("            }");
        stringBuilder.AppendLine("            return value;");
        stringBuilder.AppendLine("        }");
        stringBuilder.AppendLine("        public bool Contains(int id)");
        stringBuilder.AppendLine("        {");
        stringBuilder.AppendLine("            return _dict.ContainsKey(id);");
        stringBuilder.AppendLine("        }");
        stringBuilder.AppendLine($"        public Dictionary<int, {data.sheetName}> GetAll()");
        stringBuilder.AppendLine("        {");
        stringBuilder.AppendLine("            return _dict;");
        stringBuilder.AppendLine("        }");
        stringBuilder.AppendLine("    }");

        switch (mold)
        {
            case ConfigMold.Json:
                CreateJsonCSharp(data, ref stringBuilder);
                break;
            case ConfigMold.Xml:
                CreateXmlCSharp(data, ref stringBuilder);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(mold), mold, null);
        }

        stringBuilder.Append('}');

        var output = $"{_basePath}/Assets/App/Scripts/Frame/Core/Master/Config/{mold}/{data.sheetName}{mold}Config.cs";
        SaveFile(output, stringBuilder);
    }

    public static void CreateJsonCSharp(ExcelData data, ref StringBuilder stringBuilder)
    {
        stringBuilder.AppendLine("    [System.Serializable]");
        stringBuilder.AppendLine($"    public class {data.sheetName}JsonData");
        stringBuilder.AppendLine("    {");

        stringBuilder.AppendLine(
            $"        public List<{data.sheetName}> {data.sheetName}s = new List<{data.sheetName}>();");

        stringBuilder.AppendLine("    }");

        stringBuilder.AppendLine("    [System.Serializable]");
        stringBuilder.AppendLine($"    public class {data.sheetName}");
        stringBuilder.AppendLine("    {");

        for (var c = 3; c < data.datas.GetLength(1); c++)
        {
            if ($"{data.datas[1, c]}".Contains("#") || $"{data.datas[2, c]}".Contains("#")) continue;
            stringBuilder.AppendLine($"        /// <summary>{data.datas[3, c]}</summary>");
            stringBuilder.AppendLine($"        public {data.datas[5, c]} {data.datas[4, c]};");
        }

        stringBuilder.AppendLine("    }");
    }

    public static void CreateXmlCSharp(ExcelData data, ref StringBuilder stringBuilder)
    {
        stringBuilder.AppendLine("    [System.Serializable]");
        stringBuilder.AppendLine($"    public class {data.sheetName}XmlData");
        stringBuilder.AppendLine("    {");

        stringBuilder.AppendLine($"        [XmlElement(\"{data.sheetName}s\")]");
        stringBuilder.AppendLine(
            $"        public List<{data.sheetName}> {data.sheetName}s = new List<{data.sheetName}>();");

        stringBuilder.AppendLine("    }");

        stringBuilder.AppendLine("    [System.Serializable]");
        stringBuilder.AppendLine($"    public class {data.sheetName}");
        stringBuilder.AppendLine("    {");

        for (var c = 3; c < data.datas.GetLength(1); c++)
        {
            if ($"{data.datas[1, c]}".Contains('#') || $"{data.datas[2, c]}".Contains('#')) continue;
            stringBuilder.AppendLine($"        /// <summary>{data.datas[3, c]}</summary>");
            stringBuilder.AppendLine($"        [XmlAttribute(\"{data.datas[4, c]}\")]");
            stringBuilder.AppendLine($"        public {data.datas[5, c]} {data.datas[4, c]};");
        }

        stringBuilder.AppendLine("    }");
    }

    public static void CreateJsonConfig(ExcelData data)
    {
        var ids = new List<string>();
        var stringBuilder = new StringBuilder();
        stringBuilder.AppendLine("{");
        stringBuilder.AppendLine($"  \"{data.sheetName}s\": [");

        for (var r = 6; r < data.datas.GetLength(0); r++)
        {
            if ($"{data.datas[r, 1]}".Contains('#') || $"{data.datas[r, 2]}".Contains('#')) continue;
            stringBuilder.AppendLine("    {");
            for (var c = 3; c < data.datas.GetLength(1); c++)
            {
                if ($"{data.datas[1, c]}".Contains('#') || $"{data.datas[2, c]}".Contains('#')) continue;
                var dataStr = data.datas[r, c].ToString();
                var typeStr = data.datas[5, c].ToString();
                if ($"{data.datas[4, c]}" == "Id")
                {
                    if (dataStr != null && !ids.Contains(dataStr))
                    {
                        ids.Add(dataStr);
                    }
                    else
                    {
                        continue;
                    }
                }

                var _str = c == data.datas.GetLength(1) - 1 ? "" : ",";
                if (dataStr == null || typeStr == null) continue;
                stringBuilder.AppendLine($"      \"{data.datas[4, c]}\": {GetJsonData(dataStr, typeStr)}{_str}");
            }

            var str = data.datas.GetLength(0) - 1 == r ? "    }" : "    },";
            stringBuilder.AppendLine(str);
        }

        stringBuilder.AppendLine("  ]");

        stringBuilder.Append('}');

        var output = $"{_basePath}/Assets/Bundles/Builtin/Configs/Json/{data.sheetName}JsonData.json";
        SaveFile(output, stringBuilder);
    }

    private static string GetJsonData(string dataStr, string typeStr)
    {
        var sb = new StringBuilder();
        var array = dataStr.Split('|');
        if (typeStr.Contains("[]"))
        {
            sb.Append('[');
        }

        var arraySplit = array.Length > 1 ? "," : "";
        foreach (var value in array)
        {
            if (typeStr.Contains("string"))
            {
                sb.Append($"\"{value}\"");
            }
            else if (typeStr.Contains("Vector3"))
            {
                if (string.IsNullOrEmpty(value))
                {
                    sb.Append("{\"x\":0,\"y\":0,\"z\":0}");
                }
                else
                {
                    var str = value.Split(',');
                    sb.Append($"{{\"x\":{str[0]},\"y\":{str[1]},\"z\":{str[2]}}}");
                }
            }
            else
            {
                var str = string.IsNullOrEmpty(value) ? "0" : $"{value}";
                sb.Append($"{str}");
            }

            sb.Append(arraySplit);
        }

        if (!string.IsNullOrEmpty(arraySplit))
        {
            sb.Remove(sb.Length - 1, 1);
        }

        if (!typeStr.Contains("[]")) return sb.ToString();
        sb.Append(']');
        return sb.ToString();
    }

    public static void CreateXmlConfig(ExcelData data)
    {
        var ids = new List<string>();
        var stringBuilder = new StringBuilder();
        stringBuilder.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        stringBuilder.AppendLine($"<{data.sheetName}XmlData>");

        for (var r = 6; r < data.datas.GetLength(0); r++)
        {
            if ($"{data.datas[r, 1]}".Contains('#') || $"{data.datas[r, 2]}".Contains('#')) continue;
            stringBuilder.Append($"    <{data.sheetName}s");
            for (var c = 3; c < data.datas.GetLength(1); c++)
            {
                if ($"{data.datas[1, c]}".Contains('#') || $"{data.datas[2, c]}".Contains('#')) continue;

                if (data.datas[5, c].ToString()!.Contains("int") || data.datas[5, c].ToString()!.Contains("long") ||
                    data.datas[5, c].ToString()!.Contains("float"))
                {
                    data.datas[r, c] ??= 0;
                }

                if ($"{data.datas[4, c]}" == "Id")
                {
                    if (!ids.Contains($"{data.datas[r, c]}"))
                    {
                        ids.Add($"{data.datas[r, c]}");
                    }
                    else
                    {
                        continue;
                    }
                }

                stringBuilder.Append($" {data.datas[4, c]}=\"{data.datas[r, c]}\"");
            }

            stringBuilder.AppendLine($"/>");
        }

        stringBuilder.Append($"</{data.sheetName}XmlData>");

        var output = $"{_basePath}/Assets/Bundles/Builtin/Configs/Xml/{data.sheetName}XmlData.xml";
        SaveFile(output, stringBuilder);
    }

    private static void SaveFile(string output, StringBuilder stringBuilder)
    {
        var fs1 = new FileStream(output, FileMode.Create, FileAccess.Write);
        var sw = new StreamWriter(fs1);
        sw.WriteLine(stringBuilder.ToString()); //开始写入值
        sw.Close();
        fs1.Close();
    }
}