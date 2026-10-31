using Microsoft.Win32;

namespace VSTORepair;

internal sealed class WordPluginInfo
{
    public string ProgId { get; set; } = "";
    public string FriendlyName { get; set; } = "";
    public string Description { get; set; } = "";
    public string? Manifest { get; set; }
    public int? LoadBehavior { get; set; }
    public bool InUser { get; set; }
    public bool InMachine { get; set; }
    public bool InWpsWhitelist { get; set; }
    public bool InWpsBlacklist { get; set; }
    public List<string> Locations { get; } = new List<string>();

    public string StatusText
    {
        get
        {
            if (LoadBehavior == null)
                return InWpsWhitelist ? "仅白名单（未注册到 Word\\Addins）" : "未注册";
            return LoadBehavior.Value switch
            {
                0 => "已禁用（不自动加载）",
                1 => "已连接（不自动启动）",
                2 => "启动加载但已断开（常为加载失败）",
                3 => "已启用（启动时加载）",
                8 => "按需加载（未连接）",
                9 => "按需加载（已连接）",
                16 => "首次启动加载后按需",
                _ => "LoadBehavior=" + LoadBehavior.Value
            };
        }
    }

    public string ScopeText =>
        InUser && InMachine ? "当前用户+计算机" :
        InMachine ? "计算机" :
        InUser ? "当前用户" : "—";

    public string WpsText
    {
        get
        {
            if (InWpsBlacklist && InWpsWhitelist) return "白名单+黑名单";
            if (InWpsBlacklist) return "黑名单";
            if (InWpsWhitelist) return "白名单";
            return "未登记";
        }
    }

    public bool IsEnabled => LoadBehavior is 3 or 9 or 16 or 1;
}

internal sealed class PluginService
{
    private const string WordAddins = @"Software\Microsoft\Office\Word\Addins";
    private const string WpsWl = @"Software\Kingsoft\Office\WPS\AddinsWL";
    private const string WpsBl = @"Software\Kingsoft\Office\WPS\AddinsBL";

    public List<WordPluginInfo> ListPlugins()
    {
        var map = new Dictionary<string, WordPluginInfo>(StringComparer.OrdinalIgnoreCase);

        foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
        {
            foreach (var view in ComRegistry.AllViews())
                ReadAddins(map, hive, view);
        }

        foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
        {
            foreach (var view in ComRegistry.AllViews())
            {
                foreach (var name in ReadValueNames(hive, view, WpsWl))
                {
                    var info = GetOrAdd(map, name);
                    info.InWpsWhitelist = true;
                    if (string.IsNullOrEmpty(info.FriendlyName))
                        info.FriendlyName = name;
                }

                foreach (var name in ReadValueNames(hive, view, WpsBl))
                {
                    var info = GetOrAdd(map, name);
                    info.InWpsBlacklist = true;
                    if (string.IsNullOrEmpty(info.FriendlyName))
                        info.FriendlyName = name;
                }
            }
        }

        return map.Values
            .OrderByDescending(p => p.IsEnabled)
            .ThenBy(p => p.FriendlyName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public string Enable(WordPluginInfo plugin)
    {
        if (!plugin.InUser && !plugin.InMachine)
        {
            AddWpsWhitelist(plugin.ProgId);
            RemoveWpsBlacklist(plugin.ProgId);
            return plugin.ProgId + " 尚未注册到 Word\\Addins（只有白名单残留）。已加入 WPS 白名单；请先重新安装该插件以完成注册，再点“启用加载”。";
        }

        SetLoadBehaviorEverywhere(plugin.ProgId, 3);
        AddWpsWhitelist(plugin.ProgId);
        RemoveWpsBlacklist(plugin.ProgId);
        return "已启用 " + plugin.ProgId + "（LoadBehavior=3，已加入 WPS 白名单并移出黑名单）。请完全退出后重新打开 WPS/Word。";
    }

    public string Disable(WordPluginInfo plugin)
    {
        if (!plugin.InUser && !plugin.InMachine)
            return plugin.ProgId + " 尚未注册到 Word\\Addins，无法改 LoadBehavior。可移出 WPS 白名单。";
        SetLoadBehaviorEverywhere(plugin.ProgId, 0);
        return "已禁用 " + plugin.ProgId + "（LoadBehavior=0）。请完全退出后重新打开 WPS/Word。";
    }

    public string AddWhitelist(WordPluginInfo plugin)
    {
        AddWpsWhitelist(plugin.ProgId);
        RemoveWpsBlacklist(plugin.ProgId);
        return "已将 " + plugin.ProgId + " 加入 WPS 白名单（并移出黑名单）。";
    }

    public string RemoveWhitelist(WordPluginInfo plugin)
    {
        RemoveWpsWhitelist(plugin.ProgId);
        return "已将 " + plugin.ProgId + " 移出 WPS 白名单。";
    }

    public string ClearOfficeDisabledItems()
    {
        var cleared = 0;
        foreach (var ver in new[] { "16.0", "15.0", "14.0", "12.0" })
        {
            var path = @"Software\Microsoft\Office\" + ver + @"\Word\Resiliency\DisabledItems";
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(path, true);
                if (key == null)
                    continue;
                foreach (var name in key.GetValueNames().ToList())
                {
                    key.DeleteValue(name, false);
                    cleared++;
                }
            }
            catch (Exception ex) when (ex is UnauthorizedAccessException || ex is System.Security.SecurityException || ex is IOException)
            {
            }
        }

        return cleared == 0
            ? "未找到 Word 的“自动禁用项”记录。"
            : "已清除 " + cleared + " 条 Word 自动禁用记录。请重新打开 Word/WPS。";
    }

    private static void ReadAddins(Dictionary<string, WordPluginInfo> map, RegistryHive hive, RegistryView view)
    {
        try
        {
            using var root = RegistryKey.OpenBaseKey(hive, view);
            using var addins = root.OpenSubKey(WordAddins);
            if (addins == null)
                return;
            foreach (var progId in addins.GetSubKeyNames())
            {
                using var one = addins.OpenSubKey(progId);
                if (one == null)
                    continue;
                var info = GetOrAdd(map, progId);
                if (hive == RegistryHive.CurrentUser)
                    info.InUser = true;
                else
                    info.InMachine = true;

                var friendly = one.GetValue("FriendlyName") as string;
                var desc = one.GetValue("Description") as string;
                var manifest = one.GetValue("Manifest") as string;
                var lb = one.GetValue("LoadBehavior");
                if (!string.IsNullOrWhiteSpace(friendly) && (string.IsNullOrEmpty(info.FriendlyName) || info.FriendlyName == progId))
                    info.FriendlyName = friendly!.Trim();
                if (!string.IsNullOrWhiteSpace(desc) && string.IsNullOrEmpty(info.Description))
                    info.Description = desc!.Trim();
                if (!string.IsNullOrWhiteSpace(manifest) && string.IsNullOrEmpty(info.Manifest))
                    info.Manifest = manifest!.Trim();
                if (lb != null)
                {
                    var n = Convert.ToInt32(lb);
                    // 用户项覆盖计算机项的“当前生效”观感：优先取用户，否则取已有/机器。
                    if (hive == RegistryHive.CurrentUser || info.LoadBehavior == null)
                        info.LoadBehavior = n;
                }

                var loc = (hive == RegistryHive.CurrentUser ? "HKCU" : "HKLM")
                    + (view == RegistryView.Registry64 ? "64" : "32")
                    + @"\…\Word\Addins\" + progId;
                if (!info.Locations.Contains(loc, StringComparer.OrdinalIgnoreCase))
                    info.Locations.Add(loc);
            }
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException || ex is System.Security.SecurityException || ex is IOException)
        {
        }
    }

    private static WordPluginInfo GetOrAdd(Dictionary<string, WordPluginInfo> map, string progId)
    {
        if (!map.TryGetValue(progId, out var info))
        {
            info = new WordPluginInfo { ProgId = progId, FriendlyName = progId };
            map[progId] = info;
        }

        return info;
    }

    private static IEnumerable<string> ReadValueNames(RegistryHive hive, RegistryView view, string path)
    {
        var list = new List<string>();
        try
        {
            if (view == RegistryView.Registry64 && !Environment.Is64BitOperatingSystem)
                return list;
            using var root = RegistryKey.OpenBaseKey(hive, view);
            using var key = root.OpenSubKey(path);
            if (key == null)
                return list;
            foreach (var name in key.GetValueNames())
            {
                if (!string.IsNullOrWhiteSpace(name))
                    list.Add(name);
            }
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException || ex is System.Security.SecurityException || ex is IOException)
        {
        }

        return list;
    }

    private static void SetLoadBehaviorEverywhere(string progId, int value)
    {
        foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
        {
            foreach (var view in ComRegistry.AllViews())
            {
                try
                {
                    using var root = RegistryKey.OpenBaseKey(hive, view);
                    using var key = root.OpenSubKey(WordAddins + "\\" + progId, true);
                    key?.SetValue("LoadBehavior", value, RegistryValueKind.DWord);
                }
                catch (Exception ex) when (ex is UnauthorizedAccessException || ex is System.Security.SecurityException || ex is IOException)
                {
                }
            }
        }
    }

    private static void AddWpsWhitelist(string progId)
    {
        WriteWpsListValue(WpsWl, progId, add: true);
        WriteWpsListValue(WpsBl, progId, add: false);
    }

    private static void RemoveWpsWhitelist(string progId) => WriteWpsListValue(WpsWl, progId, add: false);

    private static void RemoveWpsBlacklist(string progId) => WriteWpsListValue(WpsBl, progId, add: false);

    private static void WriteWpsListValue(string listPath, string progId, bool add)
    {
        foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
        {
            foreach (var view in ComRegistry.AllViews())
            {
                try
                {
                    using var root = RegistryKey.OpenBaseKey(hive, view);
                    if (add)
                    {
                        using var key = root.CreateSubKey(listPath);
                        key?.SetValue(progId, "", RegistryValueKind.String);
                    }
                    else
                    {
                        using var key = root.OpenSubKey(listPath, true);
                        if (key?.GetValue(progId) != null)
                            key.DeleteValue(progId, false);
                    }
                }
                catch (Exception ex) when (ex is UnauthorizedAccessException || ex is System.Security.SecurityException || ex is IOException)
                {
                }
            }
        }
    }
}
