using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace VSTORepair;

internal sealed class ScanReport
{
    public ScanResult Result { get; set; } = null!;
    public string Text { get; set; } = "";
    public string Summary { get; set; } = "";
    public bool HasProblem { get; set; }
    public bool CanRepair { get; set; }
}

internal sealed class RepairOutcome
{
    public string Text { get; set; } = "";
    public bool Success { get; set; }
    public bool Blocked { get; set; }
    public string? BackupDirectory { get; set; }
}

internal sealed class RepairService
{
    private const int DetailLimit = 30;

    public static string DataDirectory
    {
        get
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VSTORepair");
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

    public ScanReport Scan(string? errorText)
    {
        var result = new Scanner().Scan(errorText);
        var problems = result.Problems.ToList();
        var ignored = result.IgnoredPeripheral.ToList();
        var summary = BuildSummary(result, problems, ignored);
        var text = new StringBuilder();
        text.AppendLine("VSTORepair");
        text.AppendLine("时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        text.AppendLine("运行身份: " + Environment.UserDomainName + "\\" + Environment.UserName
            + "    管理员: " + (Program.IsAdmin() ? "是" : "否")
            + "    系统: " + (Environment.Is64BitOperatingSystem ? "64 位" : "32 位"));
        text.AppendLine();
        text.AppendLine("【结论】");
        text.AppendLine(summary);
        text.AppendLine();
        AppendRequested(text, result);
        AppendEnvironment(text, result);
        AppendFamilies(text, result);
        AppendChecks(text, result, problems, ignored);

        File.WriteAllText(Path.Combine(DataDirectory, "last-scan.txt"), text.ToString(), new UTF8Encoding(true));
        return new ScanReport
        {
            Result = result,
            Text = text.ToString(),
            Summary = summary,
            HasProblem = result.HasProblem,
            CanRepair = result.Plan.HasFixes
        };
    }

    public static string? RepairBlocker()
    {
        var running = new[] { "wps", "et", "wpp" }
            .SelectMany(name =>
            {
                try
                {
                    return Process.GetProcessesByName(name);
                }
                catch (InvalidOperationException)
                {
                    return Array.Empty<Process>();
                }
            })
            .Select(p => p.ProcessName + ".exe")
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (running.Count > 0)
            return "WPS 正在运行（" + string.Join("、", running) + "）。请先完全退出 WPS，包括右下角托盘图标，再修复。";

        var desktopUser = SessionUser();
        if (desktopUser != null && !string.Equals(desktopUser, Environment.UserName, StringComparison.OrdinalIgnoreCase))
            return "本程序正以 " + Environment.UserName + " 的身份运行，但登录桌面的是 " + desktopUser
                + "。修复只写入运行身份的当前用户注册表，请直接双击运行，不要用其他账号“以管理员身份运行”。";
        return null;
    }

    public RepairOutcome Repair(string? errorText)
    {
        var before = Scan(errorText);
        if (!before.CanRepair)
        {
            return new RepairOutcome
            {
                Success = !before.HasProblem,
                Text = before.Text + Environment.NewLine + "没有可以自动修复的项目。"
            };
        }

        var blocker = RepairBlocker();
        if (blocker != null)
            return new RepairOutcome { Blocked = true, Text = blocker + Environment.NewLine + Environment.NewLine + before.Text };

        var backup = Path.Combine(DataDirectory, "backup", DateTime.Now.ToString("yyyyMMdd-HHmmss"));
        Directory.CreateDirectory(backup);
        ExportBackups(backup);

        var journal = new RegistryJournal();
        var log = new StringBuilder();
        var failed = false;
        try
        {
            Apply(before.Result, journal, log, backup, ref failed);
        }
        finally
        {
            journal.Save(Path.Combine(backup, "journal.txt"));
        }

        var after = Scan(errorText);
        var stillBroken = after.Result.Problems.Count() + after.Result.OrphanVersions.Count;
        var text = new StringBuilder();
        text.AppendLine(failed
            ? "修复过程中有项目失败，详见下方。"
            : "修复已完成（" + journal.Count + " 项；含当前用户写入，以及可撤销的版本删除）。");
        text.AppendLine(stillBroken == 0 ? "复查通过。请重新打开 WPS 验证插件。" : "复查后仍有 " + stillBroken + " 个问题，详见下方复查结果。");
        text.AppendLine("撤销方法: 在本工具中点“撤销上次修复”，或运行 VSTORepair.exe /undo。");
        text.AppendLine("备份目录: " + backup);
        text.AppendLine();
        text.Append(log);
        text.AppendLine();
        text.AppendLine("======== 修复后复查 ========");
        text.Append(after.Text);
        File.WriteAllText(Path.Combine(backup, "report.txt"), text.ToString(), new UTF8Encoding(true));

        return new RepairOutcome
        {
            Success = !failed && stillBroken == 0,
            BackupDirectory = backup,
            Text = text.ToString()
        };
    }

    public static string Undo()
    {
        var root = Path.Combine(DataDirectory, "backup");
        var last = Directory.Exists(root)
            ? Directory.GetDirectories(root)
                .Where(d => File.Exists(Path.Combine(d, "journal.txt")) && !File.Exists(Path.Combine(d, "undone.txt")))
                .OrderByDescending(d => d, StringComparer.Ordinal)
                .FirstOrDefault()
            : null;
        if (last == null)
            return "没有可以撤销的修复记录。";

        var blocker = RepairBlocker();
        if (blocker != null)
            return blocker;

        var log = RegistryJournal.Undo(Path.Combine(last, "journal.txt"));
        File.WriteAllText(Path.Combine(last, "undone.txt"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + Environment.NewLine + log, new UTF8Encoding(true));
        return "已撤销 " + Path.GetFileName(last) + " 的修复。" + Environment.NewLine + log;
    }

    private static void Apply(ScanResult result, RegistryJournal journal, StringBuilder log, string backup, ref bool failed)
    {
        foreach (var fix in result.Plan.VersionRemovals)
        {
            try
            {
                var verSub = "TypeLib\\" + ComRegistry.Braced(fix.LibId) + "\\" + fix.VersionKey;
                var tag = (fix.View == RegistryView.Registry64 ? "64" : "32") + "-" + ComRegistry.Braced(fix.LibId) + "-" + fix.VersionKey;
                if (fix.DeleteUser)
                {
                    var reg = Path.Combine(backup, "del-user-" + tag + ".reg");
                    RegExport(@"HKCU\Software\Classes\" + verSub, reg, fix.View == RegistryView.Registry64 ? "/reg:64" : "/reg:32");
                    if (!File.Exists(reg))
                        throw new IOException("删除前备份失败: " + reg);
                    journal.DeleteUserTree(fix.View, verSub, reg);
                }

                if (fix.DeleteMachine)
                {
                    var reg = Path.Combine(backup, "del-machine-" + tag + ".reg");
                    RegExport(@"HKLM\Software\Classes\" + verSub, reg, fix.View == RegistryView.Registry64 ? "/reg:64" : "/reg:32");
                    if (!File.Exists(reg))
                        throw new IOException("删除前备份失败: " + reg);
                    journal.DeleteMachineTree(fix.View, verSub, reg);
                }

                log.AppendLine("已完成: " + DescribeVersionRemoval(fix));
            }
            catch (Exception ex)
            {
                failed = true;
                log.AppendLine("失败: " + DescribeVersionRemoval(fix) + "  " + ex.Message);
            }
        }

        foreach (var fix in result.Plan.Paths)
        {
            try
            {
                var verSub = "TypeLib\\" + ComRegistry.Braced(fix.LibId) + "\\" + fix.VersionKey;
                var platform = ComRegistry.Platform(fix.View);
                if (ComRegistry.ReadString(RegScope.Merged, fix.View, verSub, null) == null)
                    journal.SetUserValue(fix.View, verSub, null, fix.LibName ?? "Type Library");

                // 合并视图里已有的每个 LCID 都要覆盖，否则按语言优先命中的那个可能仍是旧路径。
                using (var verKey = ComRegistry.Open(RegScope.Merged, fix.View, verSub))
                {
                    var lcids = Scanner.OrderedLcids(ComRegistry.SubKeyNames(verKey)).ToList();
                    if (!lcids.Contains("0"))
                        lcids.Insert(0, "0");
                    foreach (var lcid in lcids)
                    {
                        if (lcid != "0" && ComRegistry.ReadString(RegScope.Merged, fix.View, verSub + "\\" + lcid + "\\" + platform, null) == null)
                            continue;
                        journal.SetUserValue(fix.View, verSub + "\\" + lcid + "\\" + platform, null, fix.Path);
                    }
                }

                if (!ComRegistry.KeyExists(RegScope.Merged, fix.View, verSub + "\\FLAGS"))
                    journal.SetUserValue(fix.View, verSub + "\\FLAGS", null, "0");
                if (!ComRegistry.KeyExists(RegScope.Merged, fix.View, verSub + "\\HELPDIR"))
                    journal.SetUserValue(fix.View, verSub + "\\HELPDIR", null, Path.GetDirectoryName(ComPath.FilePart(fix.Path)) ?? "");
                log.AppendLine("已完成: " + DescribePathFix(fix));
            }
            catch (Exception ex)
            {
                failed = true;
                log.AppendLine("失败: " + DescribePathFix(fix) + "  " + ex.Message);
            }
        }

        var interfaceFailures = 0;
        foreach (var fix in result.Plan.Interfaces)
        {
            try
            {
                var sub = "Interface\\" + ComRegistry.Braced(fix.Iid);
                if (ComRegistry.ReadString(RegScope.Merged, fix.View, sub, null) == null)
                    journal.SetUserValue(fix.View, sub, null, fix.Name);
                var proxy = ComRegistry.ReadString(RegScope.Merged, fix.View, sub + "\\ProxyStubClsid32", null);
                if (fix.Reason == LinkStatus.ProxyStubMissing || proxy == null)
                {
                    journal.SetUserValue(fix.View, sub + "\\ProxyStubClsid32", null, fix.ProxyStub);
                    if (ComRegistry.ReadString(RegScope.Merged, fix.View, sub + "\\ProxyStubClsid", null) == null)
                        journal.SetUserValue(fix.View, sub + "\\ProxyStubClsid", null, fix.ProxyStub);
                }

                if (fix.SetTypeLib)
                {
                    journal.SetUserValue(fix.View, sub + "\\TypeLib", null, ComRegistry.Braced(fix.LibId));
                    journal.SetUserValue(fix.View, sub + "\\TypeLib", "Version", fix.VersionKey);
                }
            }
            catch (Exception ex)
            {
                failed = true;
                interfaceFailures++;
                if (interfaceFailures <= DetailLimit)
                    log.AppendLine("失败: 接口 " + fix.Name + " " + ComRegistry.Braced(fix.Iid) + "  " + ex.Message);
            }
        }

        foreach (var group in GroupInterfaceFixes(result.Plan.Interfaces))
            log.AppendLine("已处理: " + group);
    }

    private static void ExportBackups(string dir)
    {
        var keys = new[] { @"HKCU\Software\Classes\TypeLib", @"HKCU\Software\Classes\Interface" };
        foreach (var view in ComRegistry.AllViews())
        {
            var flag = view == RegistryView.Registry64 ? "/reg:64" : "/reg:32";
            foreach (var key in keys)
            {
                var name = key.Substring(key.LastIndexOf('\\') + 1) + "-" + (view == RegistryView.Registry64 ? "64" : "32") + ".reg";
                RegExport(key, Path.Combine(dir, name), flag);
            }
        }
    }

    private static void RegExport(string key, string file, string viewFlag)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = Path.Combine(Environment.SystemDirectory, "reg.exe"),
                Arguments = "export \"" + key + "\" \"" + file + "\" /y " + viewFlag,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using var process = Process.Start(psi);
            process?.WaitForExit(60000);
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception || ex is InvalidOperationException)
        {
        }
    }

    private static string BuildSummary(ScanResult result, List<InterfaceCheck> problems, List<InterfaceCheck> ignored)
    {
        var sb = new StringBuilder();
        var views = string.Join("、", result.Env.Views.Select(ComRegistry.ViewName));
        var orphans = result.OrphanVersions;
        if (problems.Count == 0 && orphans.Count == 0)
        {
            sb.AppendLine("未发现问题。检查了 " + views + " 注册表里 Word 插件相关接口链路（接口 → 类型库 → 文件 → 文件里有该接口），以及 Word/Office 类型库是否混入残缺版本。");
            if (ignored.Count > 0)
            {
                sb.AppendLine("另有 " + ignored.Count
                    + " 个接口被 WPS 表格/演示等旁路类型库覆盖且未能加载，与 Word 插件无关，已忽略（不修复）。");
            }

            sb.Append("如果插件仍然报错，请把错误对话框的文字粘贴到上方再诊断，并把日志发给技术支持。");
            return sb.ToString();
        }

        if (problems.Count > 0)
        {
            sb.AppendLine("发现 " + problems.Count + " 个接口注册损坏（检查范围: " + views + " 注册表）。");
            foreach (var group in GroupProblems(problems).Take(8))
                sb.AppendLine("  - " + group);
        }

        if (orphans.Count > 0)
        {
            sb.AppendLine("发现 " + orphans.Count + " 个残缺的类型库版本（本机另有可用版本，但注册表里还留着无路径/坏路径的版本键）。");
            sb.AppendLine("这类残留常导致 0x8002801D（库没有注册）：系统若精确命中残缺版本不会回退到较低的正常版本。");
            foreach (var orphan in orphans.Take(12))
                sb.AppendLine("  - " + DescribeOrphan(orphan));
        }

        if (result.Plan.HasFixes)
        {
            sb.AppendLine("可以自动修复。修复前自动备份，可撤销：");
            foreach (var fix in result.Plan.VersionRemovals)
                sb.AppendLine("  - " + DescribeVersionRemoval(fix));
            foreach (var fix in result.Plan.Paths)
                sb.AppendLine("  - " + DescribePathFix(fix));
            foreach (var group in GroupInterfaceFixes(result.Plan.Interfaces))
                sb.AppendLine("  - " + group);
        }

        if (result.Plan.Unfixable.Count > 0)
        {
            sb.AppendLine(result.Plan.Unfixable.Count + " 个接口找不到可用的类型库文件，无法自动修复：");
            foreach (var check in result.Plan.Unfixable.Take(10))
                sb.AppendLine("  - " + ComRegistry.ViewName(check.View) + " " + check.Name + " " + ComRegistry.Braced(check.Iid) + "：" + LinkStatusText.Describe(check.Status));
            sb.Append("请先用 WPS 自带的“配置工具 → 高级 → 修复”或重新安装 WPS，然后再运行本工具。");
        }

        return sb.ToString().TrimEnd();
    }

    private static IEnumerable<string> GroupProblems(IEnumerable<InterfaceCheck> problems)
    {
        return problems
            .GroupBy(c => (c.View, c.Status, Lib: c.LibIdText ?? "", Ver: c.Version ?? "", Path: c.Resolution?.Path ?? ""))
            .OrderByDescending(g => g.Count())
            .Select(g =>
            {
                var first = g.First();
                var names = string.Join("、", g.Take(4).Select(c => c.Name)) + (g.Count() > 4 ? " 等" : "");
                var lib = first.LibIdText == null ? "" : "，类型库 " + first.LibIdText + " " + (first.Version ?? "（无版本）");
                var path = first.Resolution?.Path == null ? "" : "，路径 " + first.Resolution.Path;
                return ComRegistry.ViewName(g.Key.View) + "：" + g.Count() + " 个接口（" + names + "）" + LinkStatusText.Describe(g.Key.Status)
                    + lib + path + "。通常报 " + HResults.Describe(LinkStatusText.TypicalHr(g.Key.Status));
            });
    }

    private static IEnumerable<string> GroupInterfaceFixes(IEnumerable<InterfaceFix> fixes)
    {
        return fixes
            .GroupBy(f => (f.View, f.SetTypeLib, f.LibId, f.VersionKey, f.TargetPath))
            .Select(g =>
            {
                var names = string.Join("、", g.Take(4).Select(f => f.Name)) + (g.Count() > 4 ? " 等" : "");
                var view = ComRegistry.ViewName(g.Key.View);
                if (!g.Key.SetTypeLib)
                    return view + "：为 " + g.Count() + " 个接口（" + names + "）补登 ProxyStubClsid32";
                return view + "：把 " + g.Count() + " 个接口（" + names + "）改指向 " + (g.First().TargetName ?? "类型库")
                    + " " + ComRegistry.Braced(g.Key.LibId) + " " + g.Key.VersionKey + "（" + g.Key.TargetPath + "）";
            });
    }

    private static string DescribePathFix(PathFix fix)
    {
        var old = string.IsNullOrEmpty(fix.OldPath) ? "原来没有路径" : "原路径 " + fix.OldPath;
        var count = fix.InterfaceCount > 0 ? "，影响 " + fix.InterfaceCount + " 个接口" : "";
        return ComRegistry.ViewName(fix.View) + "：把类型库 " + (fix.LibName ?? "") + " " + ComRegistry.Braced(fix.LibId) + " " + fix.VersionKey
            + " 的 " + ComRegistry.Platform(fix.View) + " 路径改为 " + fix.Path + "（" + old + count + "）";
    }

    private static string DescribeVersionRemoval(VersionRemoval fix)
    {
        var where = fix.DeleteMachine && fix.DeleteUser ? "当前用户和计算机"
            : fix.DeleteMachine ? "计算机"
            : "当前用户";
        var keep = fix.GoodVersion == null ? "" : "，保留可用版本 " + fix.GoodVersion;
        return ComRegistry.ViewName(fix.View) + "：删除残缺类型库版本 " + (fix.LibName ?? "") + " "
            + ComRegistry.Braced(fix.LibId) + " " + fix.VersionKey + "（" + where + keep + "）";
    }

    private static string DescribeOrphan(OrphanTypeLibVersion orphan)
    {
        var where = orphan.InMachine && orphan.InUser ? "当前用户+计算机"
            : orphan.InMachine ? "计算机"
            : orphan.InUser ? "当前用户"
            : "合并视图";
        var keep = orphan.GoodVersion == null ? "" : "；可用版本 " + orphan.GoodVersion;
        return ComRegistry.ViewName(orphan.View) + " " + (orphan.LibName ?? "") + " " + ComRegistry.Braced(orphan.LibId)
            + " " + orphan.VersionKey + "：" + LinkStatusText.Describe(orphan.Status) + "（" + where + keep + "）";
    }

    private static void AppendRequested(StringBuilder text, ScanResult result)
    {
        if (result.RequestedIids.Count == 0 && result.RequestedHresults.Count == 0)
            return;
        text.AppendLine("【粘贴的错误信息】");
        foreach (var hr in result.RequestedHresults)
            text.AppendLine("错误码: " + HResults.Describe(hr));
        foreach (var iid in result.RequestedIids)
        {
            var checks = result.Checks.Where(c => c.Iid == iid).ToList();
            if (checks.Count == 0)
            {
                text.AppendLine(ComRegistry.Braced(iid) + "：这是类型库 ID，已纳入检查范围。");
                continue;
            }

            foreach (var check in checks)
                text.AppendLine(ComRegistry.Braced(iid) + " " + check.Name + "（" + ComRegistry.ViewName(check.View) + "）：" + DescribeCheck(check));
        }

        text.AppendLine();
    }

    private static void AppendEnvironment(StringBuilder text, ScanResult result)
    {
        text.AppendLine("【环境】");
        if (result.Env.Wps.Count == 0)
            text.AppendLine("WPS: 未找到 wps.exe，32 位和 64 位注册表都会检查。");
        foreach (var wps in result.Env.Wps)
            text.AppendLine("WPS: " + wps.Exe + "    " + wps.Bitness + (wps.Active ? "    当前使用" : ""));
        foreach (var dir in result.Env.Office6Dirs.Where(d => result.Env.Wps.All(w => !string.Equals(w.Office6, d, StringComparison.OrdinalIgnoreCase))))
            text.AppendLine("WPS 残留目录（没有 wps.exe）: " + dir);
        if (result.Env.MicrosoftFiles.Count == 0)
            text.AppendLine("Microsoft Office: 未找到 MSWORD.OLB / MSO.DLL");
        foreach (var file in result.Env.MicrosoftFiles)
            text.AppendLine("Microsoft Office: " + file);
        text.AppendLine("检查的注册表视图: " + string.Join("、", result.Env.Views.Select(ComRegistry.ViewName)));
        text.AppendLine();
    }

    private static void AppendFamilies(StringBuilder text, ScanResult result)
    {
        text.AppendLine("【相关类型库】");
        var orphanKeys = new HashSet<string>(
            result.OrphanVersions.Select(o => ComRegistry.ViewName(o.View) + "|" + o.LibId + "|" + o.VersionKey),
            StringComparer.OrdinalIgnoreCase);
        foreach (var libId in result.Families)
        {
            text.AppendLine(ComRegistry.Braced(libId) + "  " + result.Scanner.LibName(libId));
            foreach (var view in result.Env.Views)
            {
                foreach (var (version, r) in result.Scanner.RegisteredVersions(view, libId))
                {
                    var source = r.Path == null ? "" : (r.PathFromUser ? "  [当前用户]" : "  [计算机]");
                    var orphanMark = orphanKeys.Contains(ComRegistry.ViewName(view) + "|" + libId + "|" + version)
                        ? "  ← 残缺版本，建议删除"
                        : "";
                    text.AppendLine("  " + ComRegistry.ViewName(view) + " " + version + "：" + LinkStatusText.Describe(r.Status)
                        + (r.Path == null ? "" : "  " + r.Path) + source + orphanMark);
                }
            }
        }

        text.AppendLine();
    }

    private static void AppendChecks(StringBuilder text, ScanResult result, List<InterfaceCheck> problems, List<InterfaceCheck> ignored)
    {
        text.AppendLine("【接口检查】");
        foreach (var view in result.Env.Views)
        {
            var inView = result.Checks.Where(c => c.View == view).ToList();
            var registered = inView.Count(c => c.Status != LinkStatus.InterfaceNotRegistered);
            var ok = inView.Count(c => c.Status == LinkStatus.Ok);
            var overrides = inView.Count(c => c.UserOverride);
            var ignoredInView = ignored.Count(c => c.View == view);
            text.AppendLine(ComRegistry.ViewName(view) + "：已注册 " + registered + " 个，正常 " + ok + " 个，Word 相关损坏 "
                + problems.Count(c => c.View == view) + " 个"
                + (ignoredInView > 0 ? "，忽略旁路 " + ignoredInView + " 个" : "")
                + "；其中 " + overrides + " 个被当前用户覆盖（WPS 安装时会这样做，本身不是问题）。");
        }

        foreach (var required in result.Checks.Where(c => c.Required))
            text.AppendLine("  关键接口 " + required.Name + " " + ComRegistry.Braced(required.Iid) + "（" + ComRegistry.ViewName(required.View) + "）：" + DescribeCheck(required));

        if (problems.Count > 0)
        {
            text.AppendLine();
            text.AppendLine("损坏明细（最多列出 " + DetailLimit * 4 + " 个）:");
            foreach (var check in problems.Take(DetailLimit * 4))
                text.AppendLine("  " + ComRegistry.ViewName(check.View) + " " + ComRegistry.Braced(check.Iid) + " " + check.Name + "：" + DescribeCheck(check));
        }

        if (ignored.Count > 0)
        {
            text.AppendLine();
            text.AppendLine("已忽略的旁路接口（WPS 表格/演示等，与 Word 插件无关，最多列出 8 个）:");
            foreach (var check in ignored.Take(8))
                text.AppendLine("  " + ComRegistry.ViewName(check.View) + " " + ComRegistry.Braced(check.Iid) + " " + check.Name + "：" + DescribeCheck(check));
            if (ignored.Count > 8)
                text.AppendLine("  ……其余 " + (ignored.Count - 8) + " 个同类项已省略");
        }
    }

    private static string DescribeCheck(InterfaceCheck check)
    {
        var sb = new StringBuilder(LinkStatusText.Describe(check.Status));
        if (check.LibIdText != null)
            sb.Append("  → 类型库 " + check.LibIdText + " " + (check.Version ?? "（无版本）"));
        if (check.Resolution?.Path != null)
            sb.Append("  → " + check.Resolution.Path);
        if (check.UserOverride)
            sb.Append("  [当前用户覆盖]");
        return sb.ToString();
    }

    private static string? SessionUser()
    {
        try
        {
            if (!WTSQuerySessionInformation(IntPtr.Zero, -1, 5, out var buffer, out _))
                return null;
            try
            {
                var name = Marshal.PtrToStringUni(buffer);
                return string.IsNullOrWhiteSpace(name) ? null : name;
            }
            finally
            {
                WTSFreeMemory(buffer);
            }
        }
        catch (Exception ex) when (ex is DllNotFoundException || ex is EntryPointNotFoundException)
        {
            return null;
        }
    }

    [DllImport("wtsapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool WTSQuerySessionInformation(IntPtr server, int sessionId, int infoClass, out IntPtr buffer, out int bytes);

    [DllImport("wtsapi32.dll")]
    private static extern void WTSFreeMemory(IntPtr memory);
}
