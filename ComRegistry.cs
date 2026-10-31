using Microsoft.Win32;
using System.Diagnostics;
using System.Text;

namespace VSTORepair;

internal enum RegScope
{
    Merged,
    User,
    Machine
}

internal static class ComRegistry
{
    public const string PsOaInterface = "{00020424-0000-0000-C000-000000000046}";
    public const string PsDispatch = "{00020420-0000-0000-C000-000000000046}";

    public static IEnumerable<RegistryView> AllViews()
    {
        yield return RegistryView.Registry32;
        if (Environment.Is64BitOperatingSystem)
            yield return RegistryView.Registry64;
    }

    public static string ViewName(RegistryView view) => view == RegistryView.Registry64 ? "64 位" : "32 位";

    public static string Platform(RegistryView view) => view == RegistryView.Registry64 ? "win64" : "win32";

    public static string Braced(Guid guid) => "{" + guid.ToString().ToUpperInvariant() + "}";

    /// <summary>
    /// Merged 读取 HKEY_CLASSES_ROOT，也就是 COM 实际看到的“当前用户覆盖计算机”的合并结果。
    /// </summary>
    public static RegistryKey? Open(RegScope scope, RegistryView view, string subKey, bool writable = false)
    {
        var hive = scope switch
        {
            RegScope.User => RegistryHive.CurrentUser,
            RegScope.Machine => RegistryHive.LocalMachine,
            _ => RegistryHive.ClassesRoot
        };
        var prefix = scope == RegScope.Merged ? "" : @"Software\Classes\";
        try
        {
            var root = RegistryKey.OpenBaseKey(hive, view);
            return root.OpenSubKey(prefix + subKey, writable);
        }
        catch (Exception ex) when (ex is System.Security.SecurityException || ex is UnauthorizedAccessException || ex is IOException)
        {
            return null;
        }
    }

    public static string? ReadString(RegistryKey? key, string? name)
    {
        if (key == null)
            return null;
        object? value;
        try
        {
            value = key.GetValue(name ?? "");
        }
        catch (Exception ex) when (ex is System.Security.SecurityException || ex is UnauthorizedAccessException || ex is IOException)
        {
            return null;
        }

        var text = value as string ?? (value == null ? null : Convert.ToString(value));
        if (text == null)
            return null;
        text = text.Trim().Trim('"').Trim();
        return text.Length == 0 ? null : text;
    }

    public static string? ReadString(RegScope scope, RegistryView view, string subKey, string? name)
    {
        using var key = Open(scope, view, subKey);
        return ReadString(key, name);
    }

    public static bool KeyExists(RegScope scope, RegistryView view, string subKey)
    {
        using var key = Open(scope, view, subKey);
        return key != null;
    }

    public static string[] SubKeyNames(RegistryKey? key)
    {
        if (key == null)
            return Array.Empty<string>();
        try
        {
            return key.GetSubKeyNames();
        }
        catch (Exception ex) when (ex is System.Security.SecurityException || ex is UnauthorizedAccessException || ex is IOException)
        {
            return Array.Empty<string>();
        }
    }

    public static string? ReadMachineValue(RegistryHive hive, RegistryView view, string subKey, string? name)
    {
        try
        {
            var root = RegistryKey.OpenBaseKey(hive, view);
            using var key = root.OpenSubKey(subKey);
            return ReadString(key, name);
        }
        catch (Exception ex) when (ex is System.Security.SecurityException || ex is UnauthorizedAccessException || ex is IOException)
        {
            return null;
        }
    }
}

/// <summary>
/// 记录每一次对 HKCU\Software\Classes 的写入，撤销时按原样恢复旧值、删除新建的键。
/// </summary>
internal sealed class RegistryJournal
{
    private readonly List<string> _lines = new List<string>();
    private readonly HashSet<string> _recordedValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _recordedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    public int Count => _lines.Count;

    public void SetUserValue(RegistryView view, string subKey, string? name, string value)
    {
        EnsureKey(view, subKey);
        using var key = ComRegistry.Open(RegScope.User, view, subKey, true)
            ?? throw new IOException(@"无法写入 HKCU\Software\Classes\" + subKey);
        var valueName = name ?? "";
        if (_recordedValues.Add(view + "|" + subKey + "|" + valueName))
        {
            var old = key.GetValue(valueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
            var kind = old == null ? RegistryValueKind.Unknown : key.GetValueKind(valueName);
            _lines.Add(Join("V", view.ToString(), subKey, valueName, old == null ? "0" : "1", kind.ToString(),
                old == null ? "" : Convert.ToString(old) ?? ""));
        }

        key.SetValue(valueName, value, RegistryValueKind.String);
    }

    /// <summary>
    /// 删除当前用户下的子树。调用方须先把内容导出到 backupReg，撤销时再 import。
    /// </summary>
    public void DeleteUserTree(RegistryView view, string subKey, string backupReg)
    {
        if (!ComRegistry.KeyExists(RegScope.User, view, subKey))
            return;
        if (_recordedKeys.Add("DELU|" + view + "|" + subKey))
            _lines.Add(Join("DELU", view.ToString(), subKey, backupReg));
        DeleteTree(RegistryHive.CurrentUser, view, @"Software\Classes\" + subKey);
    }

    /// <summary>
    /// 删除计算机下的子树（需要管理员）。调用方须先把内容导出到 backupReg。
    /// </summary>
    public void DeleteMachineTree(RegistryView view, string subKey, string backupReg)
    {
        if (!ComRegistry.KeyExists(RegScope.Machine, view, subKey))
            return;
        if (_recordedKeys.Add("DELM|" + view + "|" + subKey))
            _lines.Add(Join("DELM", view.ToString(), subKey, backupReg));
        DeleteTree(RegistryHive.LocalMachine, view, @"Software\Classes\" + subKey);
    }

    public void Save(string file)
    {
        File.WriteAllLines(file, _lines, new UTF8Encoding(false));
    }

    public static string Undo(string file)
    {
        var log = new StringBuilder();
        var entries = File.ReadAllLines(file).Where(l => l.Length > 0).Select(Split).ToList();
        entries.Reverse();

        foreach (var f in entries.Where(f => (f[0] == "DELU" || f[0] == "DELM") && f.Length >= 4))
        {
            try
            {
                if (File.Exists(f[3]))
                    RegImport(f[3], (RegistryView)Enum.Parse(typeof(RegistryView), f[1]));
                else
                    log.AppendLine("无法恢复删除项（备份不存在）: " + f[2]);
            }
            catch (Exception ex)
            {
                log.AppendLine("恢复删除失败: " + f[2] + "  " + ex.Message);
            }
        }

        foreach (var f in entries.Where(f => f[0] == "V" && f.Length >= 7))
        {
            var view = (RegistryView)Enum.Parse(typeof(RegistryView), f[1]);
            try
            {
                using var key = ComRegistry.Open(RegScope.User, view, f[2], true);
                if (key == null)
                    continue;
                if (f[4] == "1")
                {
                    var kind = (RegistryValueKind)Enum.Parse(typeof(RegistryValueKind), f[5]);
                    if (kind != RegistryValueKind.String && kind != RegistryValueKind.ExpandString)
                        kind = RegistryValueKind.String;
                    key.SetValue(f[3], f[6], kind);
                }
                else
                {
                    key.DeleteValue(f[3], false);
                }
            }
            catch (Exception ex)
            {
                log.AppendLine("恢复失败: " + f[2] + " " + f[3] + "  " + ex.Message);
            }
        }

        foreach (var f in entries.Where(f => f[0] == "K" && f.Length >= 3))
        {
            var view = (RegistryView)Enum.Parse(typeof(RegistryView), f[1]);
            var path = f[2];
            var slash = path.LastIndexOf('\\');
            var parentPath = slash < 0 ? "" : path.Substring(0, slash);
            var leaf = slash < 0 ? path : path.Substring(slash + 1);
            try
            {
                var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
                using var parent = root.OpenSubKey((@"Software\Classes\" + parentPath).TrimEnd('\\'), true);
                parent?.DeleteSubKeyTree(leaf, false);
            }
            catch (Exception ex)
            {
                log.AppendLine("删除失败: " + path + "  " + ex.Message);
            }
        }

        log.Insert(0, "已按记录撤销 " + entries.Count + " 项修改。" + Environment.NewLine);
        return log.ToString();
    }

    private static void DeleteTree(RegistryHive hive, RegistryView view, string path)
    {
        var slash = path.LastIndexOf('\\');
        var parentPath = slash < 0 ? "" : path.Substring(0, slash);
        var leaf = slash < 0 ? path : path.Substring(slash + 1);
        using var root = RegistryKey.OpenBaseKey(hive, view);
        if (string.IsNullOrEmpty(parentPath))
        {
            root.DeleteSubKeyTree(leaf, false);
            return;
        }

        using var parent = root.OpenSubKey(parentPath, true);
        if (parent == null)
            throw new IOException("无法打开注册表父键: " + parentPath);
        parent.DeleteSubKeyTree(leaf, false);
    }

    private static void RegImport(string file, RegistryView view)
    {
        var flag = view == RegistryView.Registry64 ? "/reg:64" : "/reg:32";
        var psi = new ProcessStartInfo
        {
            FileName = Path.Combine(Environment.SystemDirectory, "reg.exe"),
            Arguments = "import \"" + file + "\" " + flag,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        using var process = Process.Start(psi);
        process?.WaitForExit(60000);
        if (process != null && process.ExitCode != 0)
            throw new IOException("reg import 失败，退出码 " + process.ExitCode);
    }

    private void EnsureKey(RegistryView view, string subKey)
    {
        var path = "";
        foreach (var part in subKey.Split('\\'))
        {
            path = path.Length == 0 ? part : path + "\\" + part;
            if (ComRegistry.KeyExists(RegScope.User, view, path))
                continue;
            var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
            using (root.CreateSubKey(@"Software\Classes\" + path))
            {
            }

            if (_recordedKeys.Add(view + "|" + path))
                _lines.Add(Join("K", view.ToString(), path));
        }
    }

    private static string Join(params string[] fields) => string.Join("\t", fields.Select(Uri.EscapeDataString));

    private static string[] Split(string line) => line.Split('\t').Select(Uri.UnescapeDataString).ToArray();
}
