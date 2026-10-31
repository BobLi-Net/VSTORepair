using Microsoft.Win32;
using System.Text.RegularExpressions;

namespace VSTORepair;

internal enum LinkStatus
{
    Ok,
    InterfaceNotRegistered,
    NoTypeLib,
    LibNotRegistered,
    VersionNotRegistered,
    NoPlatformPath,
    FileMissing,
    LoadFailed,
    ElementNotFound,
    ProxyStubMissing
}

internal static class LinkStatusText
{
    public static string Describe(LinkStatus status) => status switch
    {
        LinkStatus.Ok => "正常",
        LinkStatus.InterfaceNotRegistered => "接口未注册",
        LinkStatus.NoTypeLib => "接口没有登记类型库",
        LinkStatus.LibNotRegistered => "类型库未注册",
        LinkStatus.VersionNotRegistered => "类型库的这个版本未注册",
        LinkStatus.NoPlatformPath => "类型库没有登记这一位数的文件路径",
        LinkStatus.FileMissing => "类型库文件不存在",
        LinkStatus.LoadFailed => "类型库文件无法加载",
        LinkStatus.ElementNotFound => "类型库里没有这个接口",
        LinkStatus.ProxyStubMissing => "接口没有可用的 ProxyStubClsid32",
        _ => status.ToString()
    };

    public static int TypicalHr(LinkStatus status) => status switch
    {
        LinkStatus.InterfaceNotRegistered => HResults.IidNotRegistered,
        LinkStatus.ProxyStubMissing => HResults.IidNotRegistered,
        LinkStatus.NoTypeLib => HResults.LibNotRegistered,
        LinkStatus.LibNotRegistered => HResults.LibNotRegistered,
        LinkStatus.VersionNotRegistered => HResults.LibNotRegistered,
        LinkStatus.NoPlatformPath => HResults.LibNotRegistered,
        LinkStatus.FileMissing => HResults.CantLoadLibrary,
        LinkStatus.LoadFailed => HResults.CantLoadLibrary,
        LinkStatus.ElementNotFound => HResults.ElementNotFound,
        _ => 0
    };
}

internal sealed class WpsInstall
{
    public string Exe { get; set; } = "";
    public string Office6 { get; set; } = "";
    public string Bitness { get; set; } = "";
    public RegistryView? View { get; set; }
    public bool Active { get; set; }
}

internal sealed class HostEnvironment
{
    public List<WpsInstall> Wps { get; } = new List<WpsInstall>();
    public List<string> Office6Dirs { get; } = new List<string>();
    public List<string> MicrosoftFiles { get; } = new List<string>();
    public string? ActiveOffice6 { get; set; }
    public List<RegistryView> Views { get; } = new List<RegistryView>();
}

internal sealed class TypeLibResolution
{
    public Guid LibId { get; set; }
    public string? RequestedVersion { get; set; }
    public LinkStatus Status { get; set; }
    public string? VersionKey { get; set; }
    public string? Lcid { get; set; }
    public string? Platform { get; set; }
    public string? Path { get; set; }
    public bool PathFromUser { get; set; }
    public TypeLibFile? File { get; set; }
}

internal sealed class InterfaceCheck
{
    public Guid Iid { get; set; }
    public string Name { get; set; } = "";
    public RegistryView View { get; set; }
    public bool Required { get; set; }
    public bool Requested { get; set; }
    public bool UserOverride { get; set; }
    public string? LibIdText { get; set; }
    public string? Version { get; set; }
    public string? ProxyStub { get; set; }
    public TypeLibResolution? Resolution { get; set; }
    public LinkStatus Status { get; set; }

    /// <summary>
    /// 本工具只把 Word 插件相关损坏当问题：关键接口、用户粘贴的接口，
    /// 或仍挂在 Word/Office（含 WPS 文字 wpsapi / 共用 ksoapi）类型库上的接口。
    /// WPS 表格 etapi、演示 wppapi 等旁路覆盖不计入，避免把正常机器误报成可修复。
    /// </summary>
    public bool IsInScope => Required || Requested || Scanner.IsWordPluginTypeLib(LibId, Resolution?.Path);

    public bool IsProblem => Status != LinkStatus.Ok
        && (Required || Requested || Status != LinkStatus.InterfaceNotRegistered)
        && IsInScope;

    public bool IsIgnoredPeripheral => Status != LinkStatus.Ok
        && Status != LinkStatus.InterfaceNotRegistered
        && !IsInScope;

    public Guid? LibId => Guid.TryParse(LibIdText, out var g) ? g : (Guid?)null;
}

internal sealed class PathFix
{
    public RegistryView View { get; set; }
    public Guid LibId { get; set; }
    public string VersionKey { get; set; } = "";
    public string Path { get; set; } = "";
    public TypeLibFile? File { get; set; }
    public string? LibName { get; set; }
    public string? OldPath { get; set; }
    public LinkStatus Reason { get; set; }
    public int InterfaceCount { get; set; }
}

internal sealed class InterfaceFix
{
    public RegistryView View { get; set; }
    public Guid Iid { get; set; }
    public string Name { get; set; } = "";
    public bool SetTypeLib { get; set; }
    public Guid LibId { get; set; }
    public string VersionKey { get; set; } = "";
    public string? TargetName { get; set; }
    public string? TargetPath { get; set; }
    public string ProxyStub { get; set; } = ComRegistry.PsOaInterface;
    public LinkStatus Reason { get; set; }
}

/// <summary>
/// 类型库登记了残缺版本（无路径/文件丢失），且本机另有可用版本。
/// LoadRegTypeLib 若精确命中该残缺版本会直接报 0x8002801D，不会回退到较低的正常版本。
/// </summary>
internal sealed class OrphanTypeLibVersion
{
    public RegistryView View { get; set; }
    public Guid LibId { get; set; }
    public string VersionKey { get; set; } = "";
    public string? LibName { get; set; }
    public LinkStatus Status { get; set; }
    public string? GoodVersion { get; set; }
    public bool InUser { get; set; }
    public bool InMachine { get; set; }
    public string? PathHint { get; set; }
}

internal sealed class VersionRemoval
{
    public RegistryView View { get; set; }
    public Guid LibId { get; set; }
    public string VersionKey { get; set; } = "";
    public string? LibName { get; set; }
    public string? GoodVersion { get; set; }
    public LinkStatus Reason { get; set; }
    public bool DeleteUser { get; set; }
    public bool DeleteMachine { get; set; }
}

internal sealed class RepairPlan
{
    public List<PathFix> Paths { get; } = new List<PathFix>();
    public List<InterfaceFix> Interfaces { get; } = new List<InterfaceFix>();
    public List<VersionRemoval> VersionRemovals { get; } = new List<VersionRemoval>();
    public List<InterfaceCheck> Unfixable { get; } = new List<InterfaceCheck>();
    public bool NeedsAdmin { get; set; }
    public bool HasFixes => Paths.Count > 0 || Interfaces.Count > 0 || VersionRemovals.Count > 0;
}

internal sealed class ScanResult
{
    public HostEnvironment Env { get; set; } = new HostEnvironment();
    public List<Guid> RequestedIids { get; } = new List<Guid>();
    public List<int> RequestedHresults { get; } = new List<int>();
    public List<InterfaceCheck> Checks { get; } = new List<InterfaceCheck>();
    public List<Guid> Families { get; } = new List<Guid>();
    public List<OrphanTypeLibVersion> OrphanVersions { get; } = new List<OrphanTypeLibVersion>();
    public RepairPlan Plan { get; set; } = new RepairPlan();
    public Scanner Scanner { get; set; } = null!;
    public IEnumerable<InterfaceCheck> Problems => Checks.Where(c => c.IsProblem);
    public IEnumerable<InterfaceCheck> IgnoredPeripheral => Checks.Where(c => c.IsIgnoredPeripheral);
    public bool HasProblem => Problems.Any() || OrphanVersions.Count > 0;
}

internal sealed class Scanner
{
    public static readonly Guid WordLibId = new Guid("00020905-0000-0000-C000-000000000046");
    public static readonly Guid OfficeLibId = new Guid("2DF8D04C-5BFA-101B-BDE5-00AA0044DE52");
    /// <summary>WPS 文字兼容类型库（Upgrade Kingsoft WPS Object Library）。</summary>
    public static readonly Guid WpsWordLibId = new Guid("00020905-0000-4B30-A977-D214852036FF");
    public static readonly Guid WordApplication = new Guid("00020970-0000-0000-C000-000000000046");
    public static readonly Guid RibbonExtensibility = new Guid("000C0396-0000-0000-C000-000000000046");

    /// <summary>
    /// 判断类型库是否属于 Word 插件关心的范围（微软 Word/Office，或 WPS 文字/共用库）。
    /// </summary>
    public static bool IsWordPluginTypeLib(Guid? libId, string? path)
    {
        if (libId == WordLibId || libId == OfficeLibId || libId == WpsWordLibId)
            return true;

        if (string.IsNullOrWhiteSpace(path))
            return false;

        string fileName;
        try
        {
            fileName = Path.GetFileName(ComPath.FilePart(path!));
        }
        catch (ArgumentException)
        {
            return false;
        }

        if (fileName.Equals("etapi.dll", StringComparison.OrdinalIgnoreCase)
            || fileName.Equals("wppapi.dll", StringComparison.OrdinalIgnoreCase))
            return false;

        return fileName.Equals("MSWORD.OLB", StringComparison.OrdinalIgnoreCase)
            || fileName.Equals("MSO.DLL", StringComparison.OrdinalIgnoreCase)
            || fileName.Equals("wpsapi.dll", StringComparison.OrdinalIgnoreCase)
            || fileName.Equals("ksoapi.dll", StringComparison.OrdinalIgnoreCase);
    }

    private static readonly Regex GuidPattern = new Regex(@"[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}");
    private static readonly Regex HrPattern = new Regex(@"0x[0-9A-Fa-f]{8}");

    private readonly TypeLibCache _cache = new TypeLibCache();
    private readonly Dictionary<string, TypeLibResolution> _resolutions = new Dictionary<string, TypeLibResolution>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<Guid, List<TypeLibFile>> _familyFiles = new Dictionary<Guid, List<TypeLibFile>>();
    private HostEnvironment _env = new HostEnvironment();

    public TypeLibCache Cache => _cache;

    public ScanResult Scan(string? errorText)
    {
        _env = DiscoverEnvironment();
        var result = new ScanResult { Env = _env, Scanner = this };
        ParseErrorText(errorText, result);

        var required = new List<Guid> { WordApplication, RibbonExtensibility };
        var families = new List<Guid> { WordLibId, OfficeLibId };
        foreach (var guid in result.RequestedIids)
        {
            if (ComRegistry.KeyExists(RegScope.Merged, RegistryView.Registry32, "TypeLib\\" + ComRegistry.Braced(guid))
                || ComRegistry.KeyExists(RegScope.Merged, RegistryView.Registry64, "TypeLib\\" + ComRegistry.Braced(guid)))
                AddDistinct(families, guid);
            else
                AddDistinct(required, guid);
        }

        foreach (var view in _env.Views)
        {
            foreach (var iid in required)
            {
                var lib = ComRegistry.ReadString(RegScope.Merged, view, "Interface\\" + ComRegistry.Braced(iid) + "\\TypeLib", null);
                if (Guid.TryParse(lib, out var libId))
                    AddDistinct(families, libId);
            }
        }

        var universe = new Dictionary<Guid, string>();
        foreach (var iid in required)
            universe[iid] = "";
        foreach (var libId in families)
        {
            foreach (var file in FamilyFiles(libId))
            {
                foreach (var info in file.Interfaces.Values)
                {
                    if (!universe.TryGetValue(info.Iid, out var name) || string.IsNullOrEmpty(name))
                        universe[info.Iid] = info.Name;
                }
            }
        }

        result.Families.AddRange(families);
        var requiredSet = new HashSet<Guid>(required);
        var requestedSet = new HashSet<Guid>(result.RequestedIids);
        foreach (var view in _env.Views)
        {
            foreach (var pair in universe)
            {
                result.Checks.Add(Check(
                    view,
                    pair.Key,
                    pair.Value,
                    requiredSet.Contains(pair.Key),
                    requestedSet.Contains(pair.Key)));
            }
        }

        result.OrphanVersions.AddRange(FindOrphanVersions(result));
        result.Plan = Plan(result);
        return result;
    }

    /// <summary>
    /// 找出“本机有正常版本，同时又登记了残缺高/旁路版本”的类型库项。
    /// 接口链路可能仍显示正常（指向 8.5），但插件若按 8.7 加载，或 LoadRegTypeLib 精确命中残缺键，就会报 0x8002801D。
    /// </summary>
    private List<OrphanTypeLibVersion> FindOrphanVersions(ScanResult result)
    {
        var orphans = new List<OrphanTypeLibVersion>();
        // 只清理 Word/Office（含 WPS 文字）族的残缺版本；表格/演示等旁路库不碰。
        foreach (var libId in result.Families.Where(id => IsWordPluginTypeLib(id, null)))
        {
            var libName = LibName(libId);
            var installed = FamilyFiles(libId);
            foreach (var view in result.Env.Views)
            {
                var versions = RegisteredVersions(view, libId).ToList();
                var good = versions.Where(v => v.Resolution.Status == LinkStatus.Ok).ToList();
                if (good.Count == 0)
                    continue;

                foreach (var (version, resolution) in versions)
                {
                    if (resolution.Status == LinkStatus.Ok)
                        continue;
                    if (resolution.Status != LinkStatus.NoPlatformPath
                        && resolution.Status != LinkStatus.FileMissing
                        && resolution.Status != LinkStatus.LoadFailed)
                        continue;

                    // 本机若装有该精确版本的文件，应走补路径而不是当孤立项删掉。
                    if (installed.Any(f => SameVersion(f.VersionKey, version)))
                        continue;

                    if (!TypeLibVersion.TryParse(version, out var major, out var minor))
                        continue;
                    var bestGood = good
                        .Select(g => (g.Version, Ok: TypeLibVersion.TryParse(g.Version, out var ma, out var mi), Major: ma, Minor: mi))
                        .Where(g => g.Ok && g.Major == major)
                        .OrderByDescending(g => g.Minor)
                        .Select(g => g.Version)
                        .FirstOrDefault()
                        ?? good.Select(g => g.Version).FirstOrDefault();

                    var verSub = "TypeLib\\" + ComRegistry.Braced(libId) + "\\" + version;
                    orphans.Add(new OrphanTypeLibVersion
                    {
                        View = view,
                        LibId = libId,
                        VersionKey = version,
                        LibName = libName,
                        Status = resolution.Status,
                        GoodVersion = bestGood,
                        InUser = ComRegistry.KeyExists(RegScope.User, view, verSub),
                        InMachine = ComRegistry.KeyExists(RegScope.Machine, view, verSub),
                        PathHint = resolution.Path
                    });
                }
            }
        }

        return orphans;
    }

    public string NameOf(Guid iid)
    {
        foreach (var file in _cache.Loaded)
        {
            if (file.Interfaces.TryGetValue(iid, out var info) && !string.IsNullOrEmpty(info.Name))
                return info.Name;
        }

        return ComRegistry.ReadString(RegScope.Merged, RegistryView.Registry32, "Interface\\" + ComRegistry.Braced(iid), null)
            ?? ComRegistry.ReadString(RegScope.Merged, RegistryView.Registry64, "Interface\\" + ComRegistry.Braced(iid), null)
            ?? "未知接口";
    }

    public string LibName(Guid libId)
    {
        var file = FamilyFiles(libId).FirstOrDefault();
        if (file != null && !string.IsNullOrWhiteSpace(file.Name))
            return file.Name!;
        using var key = ComRegistry.Open(RegScope.Merged, RegistryView.Registry32, "TypeLib\\" + ComRegistry.Braced(libId));
        foreach (var version in ComRegistry.SubKeyNames(key))
        {
            var name = ComRegistry.ReadString(RegScope.Merged, RegistryView.Registry32, "TypeLib\\" + ComRegistry.Braced(libId) + "\\" + version, null);
            if (name != null)
                return name;
        }

        return ComRegistry.Braced(libId);
    }

    public IEnumerable<(string Version, TypeLibResolution Resolution)> RegisteredVersions(RegistryView view, Guid libId)
    {
        using var key = ComRegistry.Open(RegScope.Merged, view, "TypeLib\\" + ComRegistry.Braced(libId));
        foreach (var version in ComRegistry.SubKeyNames(key).OrderBy(v => v, StringComparer.OrdinalIgnoreCase))
        {
            if (TypeLibVersion.TryParse(version, out _, out _))
                yield return (version, Resolve(view, libId, version));
        }
    }

    private static void ParseErrorText(string? text, ScanResult result)
    {
        if (string.IsNullOrWhiteSpace(text))
            return;
        foreach (Match m in GuidPattern.Matches(text))
        {
            if (Guid.TryParse(m.Value, out var guid) && !result.RequestedIids.Contains(guid))
                result.RequestedIids.Add(guid);
        }

        foreach (Match m in HrPattern.Matches(text))
        {
            if (HResults.TryParse(m.Value, out var hr) && !result.RequestedHresults.Contains(hr))
                result.RequestedHresults.Add(hr);
        }
    }

    private InterfaceCheck Check(RegistryView view, Guid iid, string name, bool required, bool requested)
    {
        var sub = "Interface\\" + ComRegistry.Braced(iid);
        var check = new InterfaceCheck
        {
            Iid = iid,
            View = view,
            Required = required,
            Requested = requested,
            Name = string.IsNullOrEmpty(name) ? NameOf(iid) : name,
            UserOverride = ComRegistry.KeyExists(RegScope.User, view, sub)
        };

        using var key = ComRegistry.Open(RegScope.Merged, view, sub);
        if (key == null)
        {
            check.Status = LinkStatus.InterfaceNotRegistered;
            return check;
        }

        using (var proxy = key.OpenSubKey("ProxyStubClsid32"))
            check.ProxyStub = ComRegistry.ReadString(proxy, null);
        using (var typeLib = key.OpenSubKey("TypeLib"))
        {
            check.LibIdText = ComRegistry.ReadString(typeLib, null);
            check.Version = ComRegistry.ReadString(typeLib, "Version");
        }

        var proxyKnown = IsUsableProxy(view, check.ProxyStub);
        var typeLibMarshaled = check.ProxyStub == null || IsSameGuid(check.ProxyStub, ComRegistry.PsOaInterface);
        var libId = check.LibId;
        if (libId == null)
        {
            // 自定义代理或纯 IDispatch 代理不依赖类型库。
            check.Status = !typeLibMarshaled && proxyKnown ? LinkStatus.Ok : LinkStatus.NoTypeLib;
            return check;
        }

        var resolution = Resolve(view, libId.Value, check.Version);
        check.Resolution = resolution;
        if (resolution.Status != LinkStatus.Ok)
        {
            check.Status = resolution.Status;
            return check;
        }

        if (!resolution.File!.Interfaces.ContainsKey(iid))
        {
            check.Status = LinkStatus.ElementNotFound;
            return check;
        }

        check.Status = proxyKnown ? LinkStatus.Ok : LinkStatus.ProxyStubMissing;
        return check;
    }

    /// <summary>
    /// 按 LoadRegTypeLib 的规则在合并视图里解析：版本先精确匹配，否则取同主版本里次版本更高的最大者。
    /// </summary>
    public TypeLibResolution Resolve(RegistryView view, Guid libId, string? versionText)
    {
        var cacheKey = view + "|" + libId + "|" + versionText;
        if (_resolutions.TryGetValue(cacheKey, out var cached))
            return cached;

        var r = new TypeLibResolution { LibId = libId, RequestedVersion = versionText };
        _resolutions[cacheKey] = r;
        var libSub = "TypeLib\\" + ComRegistry.Braced(libId);
        using var libKey = ComRegistry.Open(RegScope.Merged, view, libSub);
        if (libKey == null)
        {
            r.Status = LinkStatus.LibNotRegistered;
            return r;
        }

        var versions = ComRegistry.SubKeyNames(libKey)
            .Select(n => (Name: n, Ok: TypeLibVersion.TryParse(n, out var ma, out var mi), Major: ma, Minor: mi))
            .Where(v => v.Ok)
            .ToList();
        if (!TypeLibVersion.TryParse(versionText, out var major, out var minor))
        {
            r.Status = LinkStatus.VersionNotRegistered;
            return r;
        }

        var match = versions.FirstOrDefault(v => v.Major == major && v.Minor == minor);
        if (match.Name == null)
            match = versions.Where(v => v.Major == major && v.Minor > minor).OrderByDescending(v => v.Minor).FirstOrDefault();
        if (match.Name == null)
        {
            r.Status = LinkStatus.VersionNotRegistered;
            return r;
        }

        r.VersionKey = match.Name;
        var verSub = libSub + "\\" + match.Name;
        using var verKey = libKey.OpenSubKey(match.Name);
        var platforms = view == RegistryView.Registry64 ? new[] { "win64", "win32" } : new[] { "win32" };
        foreach (var lcid in OrderedLcids(ComRegistry.SubKeyNames(verKey)))
        {
            foreach (var platform in platforms)
            {
                var path = ComRegistry.ReadString(RegScope.Merged, view, verSub + "\\" + lcid + "\\" + platform, null);
                if (path == null)
                    continue;
                r.Lcid = lcid;
                r.Platform = platform;
                r.Path = path;
                r.PathFromUser = ComRegistry.ReadString(RegScope.User, view, verSub + "\\" + lcid + "\\" + platform, null) != null;
                break;
            }

            if (r.Path != null)
                break;
        }

        if (r.Path == null)
        {
            r.Status = LinkStatus.NoPlatformPath;
            return r;
        }

        if (!ComPath.Exists(r.Path))
        {
            r.Status = LinkStatus.FileMissing;
            return r;
        }

        r.File = _cache.Get(r.Path);
        r.Status = r.File.Loaded ? LinkStatus.Ok : LinkStatus.LoadFailed;
        return r;
    }

    public static IEnumerable<string> OrderedLcids(IEnumerable<string> subKeys)
    {
        var lcids = subKeys
            .Where(n => !n.Equals("FLAGS", StringComparison.OrdinalIgnoreCase) && !n.Equals("HELPDIR", StringComparison.OrdinalIgnoreCase))
            .Where(n => int.TryParse(n, System.Globalization.NumberStyles.HexNumber, null, out _))
            .ToList();
        return lcids.OrderBy(n => n == "0" ? 0 : 1).ThenBy(n => n, StringComparer.OrdinalIgnoreCase);
    }

    private RepairPlan Plan(ScanResult result)
    {
        var plan = new RepairPlan();
        var pathFixes = new Dictionary<string, PathFix>(StringComparer.OrdinalIgnoreCase);
        foreach (var check in result.Problems)
        {
            if (check.Status == LinkStatus.ProxyStubMissing)
            {
                plan.Interfaces.Add(new InterfaceFix
                {
                    View = check.View,
                    Iid = check.Iid,
                    Name = check.Name,
                    ProxyStub = DefaultProxy(check.Iid),
                    Reason = check.Status
                });
                continue;
            }

            if (TryPlanSameLibrary(check, pathFixes))
                continue;

            TypeLibFile? target = null;
            TypeLibResolution? current = null;
            foreach (var candidate in FilesContaining(check.Iid, check.LibId))
            {
                current = Resolve(check.View, candidate.LibId, candidate.VersionKey);
                // 只允许改写本来就坏掉的版本路径；已经能用的版本即使指向别的文件也不能动，否则会连累其他接口。
                if (current.Status != LinkStatus.Ok || current.File!.Interfaces.ContainsKey(check.Iid))
                {
                    target = candidate;
                    break;
                }
            }

            if (target == null || current == null)
            {
                plan.Unfixable.Add(check);
                continue;
            }

            if (current.Status != LinkStatus.Ok)
                AddPathFix(pathFixes, check.View, target.LibId, target.VersionKey, target, current.Path, check.Status, true);

            var alreadyPointsThere = check.LibId == target.LibId && SameVersion(check.Version, target.VersionKey);
            if (alreadyPointsThere && check.Status != LinkStatus.InterfaceNotRegistered && check.Status != LinkStatus.NoTypeLib)
                continue;

            var usableProxy = check.ProxyStub != null && IsUsableProxy(check.View, check.ProxyStub) ? check.ProxyStub : null;
            plan.Interfaces.Add(new InterfaceFix
            {
                View = check.View,
                Iid = check.Iid,
                Name = check.Name,
                SetTypeLib = true,
                LibId = target.LibId,
                VersionKey = target.VersionKey,
                TargetName = target.Name,
                TargetPath = target.Path,
                ProxyStub = usableProxy ?? (target.Interfaces[check.Iid].DispatchOnly ? ComRegistry.PsDispatch : ComRegistry.PsOaInterface),
                Reason = check.Status
            });
        }

        PlanBrokenVersionPaths(result, pathFixes);
        PlanOrphanVersions(result, plan, pathFixes);
        plan.Paths.AddRange(pathFixes.Values);
        return plan;
    }

    /// <summary>
    /// 版本键存在且本机有同版本号文件，但注册路径失效时，主动补路径（不必等接口检查失败）。
    /// </summary>
    private void PlanBrokenVersionPaths(ScanResult result, Dictionary<string, PathFix> pathFixes)
    {
        foreach (var libId in result.Families)
        {
            var installed = FamilyFiles(libId);
            foreach (var view in result.Env.Views)
            {
                foreach (var (version, resolution) in RegisteredVersions(view, libId))
                {
                    if (resolution.Status != LinkStatus.NoPlatformPath
                        && resolution.Status != LinkStatus.FileMissing
                        && resolution.Status != LinkStatus.LoadFailed)
                        continue;
                    var matching = installed.FirstOrDefault(f => SameVersion(f.VersionKey, version));
                    if (matching == null)
                        continue;
                    AddPathFix(pathFixes, view, libId, version, matching, resolution.Path, resolution.Status, false);
                }
            }
        }
    }

    private void PlanOrphanVersions(ScanResult result, RepairPlan plan, Dictionary<string, PathFix> pathFixes)
    {
        var isAdmin = Program.IsAdmin();
        foreach (var orphan in result.OrphanVersions)
        {
            var pathKey = ComRegistry.Platform(orphan.View) + "|" + orphan.LibId + "|" + orphan.VersionKey;
            if (pathFixes.ContainsKey(pathKey))
                continue;
            if (!orphan.InUser && !orphan.InMachine)
                continue;

            // 优先删除残缺版本（与手工删 8.7 一致）。只要计算机里也有该项，就需要管理员才能从合并视图里彻底清掉。
            if (!orphan.InMachine || isAdmin)
            {
                plan.VersionRemovals.Add(new VersionRemoval
                {
                    View = orphan.View,
                    LibId = orphan.LibId,
                    VersionKey = orphan.VersionKey,
                    LibName = orphan.LibName,
                    GoodVersion = orphan.GoodVersion,
                    Reason = orphan.Status,
                    DeleteUser = orphan.InUser,
                    DeleteMachine = orphan.InMachine && isAdmin
                });
                continue;
            }

            // 无管理员时无法删 HKLM，改用当前用户补路径盖住残缺项；并提示以管理员重跑可改为删除。
            plan.NeedsAdmin = true;
            var replacement = FamilyFiles(orphan.LibId)
                .OrderBy(f => LocationRank(f.Path))
                .ThenByDescending(f => f.Major)
                .ThenByDescending(f => f.Minor)
                .FirstOrDefault();
            if (replacement != null)
                AddPathFix(pathFixes, orphan.View, orphan.LibId, orphan.VersionKey, replacement, orphan.PathHint, orphan.Status, false);
        }
    }

    /// <summary>
    /// 版本键存在但路径失效（典型：WPS 升级换了版本目录）时，把这个版本键的路径改到现有安装里的同一个类型库。
    /// 版本键本身不存在（库未注册）时不在这里处理，交给后面按真实版本注册，避免凭空造版本号。
    /// </summary>
    private bool TryPlanSameLibrary(InterfaceCheck check, Dictionary<string, PathFix> pathFixes)
    {
        var fixable = check.Status == LinkStatus.NoPlatformPath || check.Status == LinkStatus.FileMissing || check.Status == LinkStatus.LoadFailed;
        var libId = check.LibId;
        var versionKey = check.Resolution?.VersionKey;
        if (!fixable || libId == null || versionKey == null || !TypeLibVersion.TryParse(versionKey, out var major, out var minor))
            return false;

        var brokenPath = check.Resolution?.Path;
        var replacement = FamilyFiles(libId.Value)
            .Where(f => f.Major == major && f.Interfaces.ContainsKey(check.Iid))
            .Where(f => brokenPath == null || !string.Equals(f.Path, ComPath.Normalize(brokenPath), StringComparison.OrdinalIgnoreCase))
            .OrderBy(f => f.Minor == minor ? 0 : 1)
            .ThenBy(f => LocationRank(f.Path))
            .ThenByDescending(f => f.Minor)
            .FirstOrDefault();
        if (replacement == null)
            return false;

        var key = ComRegistry.Platform(check.View) + "|" + libId.Value + "|" + versionKey;
        if (pathFixes.TryGetValue(key, out var existing) && existing.File != null && !existing.File.Interfaces.ContainsKey(check.Iid))
            return false;

        AddPathFix(pathFixes, check.View, libId.Value, versionKey, replacement, brokenPath, check.Status, true);
        return true;
    }

    private void AddPathFix(Dictionary<string, PathFix> fixes, RegistryView view, Guid libId, string versionKey, TypeLibFile file, string? oldPath, LinkStatus reason, bool countInterface)
    {
        var key = ComRegistry.Platform(view) + "|" + libId + "|" + versionKey;
        if (!fixes.TryGetValue(key, out var fix))
        {
            fix = new PathFix
            {
                View = view,
                LibId = libId,
                VersionKey = versionKey,
                Path = file.Path,
                File = file,
                LibName = file.Name,
                OldPath = oldPath,
                Reason = reason
            };
            fixes[key] = fix;
        }

        if (countInterface)
            fix.InterfaceCount++;
    }

    private IEnumerable<TypeLibFile> FilesContaining(Guid iid, Guid? preferredLib)
    {
        return _cache.Loaded
            .Where(f => f.Interfaces.ContainsKey(iid) && f.Major > 0)
            .OrderBy(f => LocationRank(f.Path))
            .ThenBy(f => preferredLib != null && f.LibId == preferredLib ? 0 : 1)
            .ThenByDescending(f => f.Major)
            .ThenByDescending(f => f.Minor)
            .ToList();
    }

    private string DefaultProxy(Guid iid)
    {
        var info = _cache.Loaded.Select(f => f.Interfaces.TryGetValue(iid, out var i) ? i : null).FirstOrDefault(i => i != null);
        return info != null && info.DispatchOnly ? ComRegistry.PsDispatch : ComRegistry.PsOaInterface;
    }

    public int LocationRank(string path)
    {
        if (ComPath.IsUnder(path, _env.ActiveOffice6))
            return 0;
        if (_env.MicrosoftFiles.Any(m => string.Equals(m, path, StringComparison.OrdinalIgnoreCase))
            || path.IndexOf(@"\Microsoft Office\", StringComparison.OrdinalIgnoreCase) >= 0
            || path.IndexOf(@"\Microsoft Shared\", StringComparison.OrdinalIgnoreCase) >= 0)
            return 1;
        if (_env.Office6Dirs.Any(d => ComPath.IsUnder(path, d)))
            return 2;
        return 3;
    }

    private List<TypeLibFile> FamilyFiles(Guid libId)
    {
        if (_familyFiles.TryGetValue(libId, out var cached))
            return cached;

        var files = new List<TypeLibFile>();
        foreach (var path in RegisteredPaths(libId))
        {
            if (ComPath.Exists(path))
                files.Add(_cache.Get(path));
            else
                files.AddRange(FindReplacements(path).Select(_cache.Get));
        }

        files.AddRange(_env.MicrosoftFiles.Select(_cache.Get));
        var result = files.Where(f => f.Loaded && f.LibId == libId).Distinct().ToList();
        _familyFiles[libId] = result;
        return result;
    }

    private static IEnumerable<string> RegisteredPaths(Guid libId)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var scope in new[] { RegScope.User, RegScope.Machine })
        {
            foreach (var view in ComRegistry.AllViews())
            {
                var libSub = "TypeLib\\" + ComRegistry.Braced(libId);
                using var libKey = ComRegistry.Open(scope, view, libSub);
                foreach (var version in ComRegistry.SubKeyNames(libKey))
                {
                    using var verKey = libKey!.OpenSubKey(version);
                    foreach (var lcid in OrderedLcids(ComRegistry.SubKeyNames(verKey)))
                    {
                        using var lcidKey = verKey!.OpenSubKey(lcid);
                        foreach (var platform in ComRegistry.SubKeyNames(lcidKey))
                        {
                            using var platformKey = lcidKey!.OpenSubKey(platform);
                            var path = ComRegistry.ReadString(platformKey, null);
                            if (path != null && seen.Add(path))
                                yield return path;
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// WPS 按版本号目录安装，升级后旧目录被删，注册表仍指向旧目录。按 office6 下的相对路径到现有安装目录里找同名文件。
    /// </summary>
    private IEnumerable<string> FindReplacements(string stalePath)
    {
        var filePart = ComPath.FilePart(stalePath);
        var suffix = ComPath.Suffix(stalePath);
        string fileName;
        try
        {
            fileName = Path.GetFileName(filePart);
        }
        catch (ArgumentException)
        {
            yield break;
        }

        var marker = filePart.IndexOf(@"\office6\", StringComparison.OrdinalIgnoreCase);
        var relative = marker >= 0 ? filePart.Substring(marker + @"\office6\".Length) : fileName;
        foreach (var dir in _env.Office6Dirs)
        {
            string candidate;
            try
            {
                candidate = Path.Combine(dir, relative);
            }
            catch (ArgumentException)
            {
                continue;
            }

            if (File.Exists(candidate))
                yield return candidate + suffix;
        }

        foreach (var file in _env.MicrosoftFiles)
        {
            if (string.Equals(Path.GetFileName(file), fileName, StringComparison.OrdinalIgnoreCase))
                yield return file;
        }
    }

    private static bool IsUsableProxy(RegistryView view, string? proxy)
    {
        if (!Guid.TryParse(proxy, out var clsid))
            return false;
        if (IsSameGuid(proxy, ComRegistry.PsOaInterface) || IsSameGuid(proxy, ComRegistry.PsDispatch))
            return true;
        return ComRegistry.KeyExists(RegScope.Merged, view, "CLSID\\" + ComRegistry.Braced(clsid) + "\\InprocServer32");
    }

    private static bool IsSameGuid(string? text, string guid)
    {
        return Guid.TryParse(text, out var a) && Guid.TryParse(guid, out var b) && a == b;
    }

    private static bool SameVersion(string? a, string? b)
    {
        return TypeLibVersion.TryParse(a, out var aMajor, out var aMinor)
            && TypeLibVersion.TryParse(b, out var bMajor, out var bMinor)
            && aMajor == bMajor && aMinor == bMinor;
    }

    private static void AddDistinct(List<Guid> list, Guid guid)
    {
        if (!list.Contains(guid))
            list.Add(guid);
    }

    private static HostEnvironment DiscoverEnvironment()
    {
        var env = new HostEnvironment();
        var roots = new List<string>();
        var activeRoot = ComRegistry.ReadMachineValue(RegistryHive.CurrentUser, RegistryView.Default, @"SOFTWARE\Kingsoft\Office\6.0\Common", "InstallRoot");
        foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
        {
            foreach (var view in ComRegistry.AllViews())
            {
                AddRoot(roots, ComRegistry.ReadMachineValue(hive, view, @"SOFTWARE\Kingsoft\Office\6.0\Common", "InstallRoot"));
                foreach (var location in UninstallLocations(hive, view))
                    AddRoot(roots, location);
            }
        }

        AddRoot(roots, Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Kingsoft\WPS Office"));
        AddRoot(roots, Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Kingsoft\WPS Office"));
        AddRoot(roots, Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Kingsoft\WPS Office"));

        var office6 = new List<string>();
        foreach (var root in roots)
        {
            foreach (var dir in Office6Candidates(root))
            {
                if (!office6.Contains(dir, StringComparer.OrdinalIgnoreCase))
                    office6.Add(dir);
            }
        }

        string? active = null;
        if (!string.IsNullOrWhiteSpace(activeRoot))
            active = Office6Candidates(activeRoot!).FirstOrDefault(d => File.Exists(Path.Combine(d, "wps.exe")));
        active ??= office6.Where(d => File.Exists(Path.Combine(d, "wps.exe")))
            .OrderByDescending(d => SafeWriteTime(Path.Combine(d, "wps.exe")))
            .FirstOrDefault();
        env.ActiveOffice6 = active;

        env.Office6Dirs.AddRange(office6
            .OrderBy(d => string.Equals(d, active, StringComparison.OrdinalIgnoreCase) ? 0 : 1)
            .ThenByDescending(d => SafeWriteTime(d)));

        foreach (var dir in env.Office6Dirs)
        {
            var exe = Path.Combine(dir, "wps.exe");
            if (!File.Exists(exe))
                continue;
            var bitness = PeBitness(exe);
            env.Wps.Add(new WpsInstall
            {
                Exe = exe,
                Office6 = dir,
                Bitness = bitness,
                View = bitness == "64 位" ? RegistryView.Registry64 : bitness == "32 位" ? RegistryView.Registry32 : (RegistryView?)null,
                Active = string.Equals(dir, active, StringComparison.OrdinalIgnoreCase)
            });
        }

        foreach (var view in env.Wps.Where(w => w.View != null).Select(w => w.View!.Value).Distinct())
        {
            if (view == RegistryView.Registry64 && !Environment.Is64BitOperatingSystem)
                continue;
            env.Views.Add(view);
        }

        if (env.Views.Count == 0 || env.Wps.Any(w => w.View == null))
        {
            foreach (var view in ComRegistry.AllViews())
            {
                if (!env.Views.Contains(view))
                    env.Views.Add(view);
            }
        }

        env.Views.Sort();
        env.MicrosoftFiles.AddRange(DiscoverMicrosoftFiles());
        return env;
    }

    private static IEnumerable<string> Office6Candidates(string root)
    {
        var list = new List<string>();
        try
        {
            var trimmed = root.TrimEnd('\\');
            if (!Directory.Exists(trimmed))
                return list;
            if (Path.GetFileName(trimmed).Equals("office6", StringComparison.OrdinalIgnoreCase))
            {
                list.Add(trimmed);
                var parent = Path.GetDirectoryName(Path.GetDirectoryName(trimmed));
                if (parent != null)
                    list.AddRange(Office6Candidates(parent));
                return list;
            }

            var direct = Path.Combine(trimmed, "office6");
            if (Directory.Exists(direct))
                list.Add(direct);
            foreach (var sub in Directory.GetDirectories(trimmed))
            {
                var nested = Path.Combine(sub, "office6");
                if (Directory.Exists(nested))
                    list.Add(nested);
            }
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException)
        {
        }

        return list;
    }

    private static IEnumerable<string> UninstallLocations(RegistryHive hive, RegistryView view)
    {
        var result = new List<string>();
        try
        {
            var root = RegistryKey.OpenBaseKey(hive, view);
            using var uninstall = root.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
            foreach (var name in ComRegistry.SubKeyNames(uninstall))
            {
                using var one = uninstall!.OpenSubKey(name);
                var display = ComRegistry.ReadString(one, "DisplayName");
                if (display == null || display.IndexOf("WPS", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                var location = ComRegistry.ReadString(one, "InstallLocation");
                if (location != null)
                    result.Add(location);
            }
        }
        catch (Exception ex) when (ex is System.Security.SecurityException || ex is UnauthorizedAccessException || ex is IOException)
        {
        }

        return result;
    }

    private static IEnumerable<string> DiscoverMicrosoftFiles()
    {
        var found = new List<string>();
        void Add(string? path)
        {
            if (!string.IsNullOrWhiteSpace(path) && File.Exists(path) && !found.Contains(path!, StringComparer.OrdinalIgnoreCase))
                found.Add(path!);
        }

        var programRoots = new[]
        {
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
            Environment.GetEnvironmentVariable("ProgramW6432")
        }.Where(p => !string.IsNullOrWhiteSpace(p)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

        foreach (var root in programRoots)
        {
            foreach (var ver in new[] { "Office16", "Office15", "Office14" })
            {
                Add(Path.Combine(root!, @"Microsoft Office\root\" + ver + @"\MSWORD.OLB"));
                Add(Path.Combine(root!, @"Microsoft Office\" + ver + @"\MSWORD.OLB"));
                Add(Path.Combine(root!, @"Microsoft Office\root\VFS\ProgramFilesCommonX86\Microsoft Shared\" + ver + @"\MSO.DLL"));
                Add(Path.Combine(root!, @"Microsoft Office\root\VFS\ProgramFilesCommonX64\Microsoft Shared\" + ver + @"\MSO.DLL"));
                Add(Path.Combine(root!, @"Common Files\Microsoft Shared\" + ver + @"\MSO.DLL"));
            }
        }

        foreach (var view in ComRegistry.AllViews())
        {
            var winword = ComRegistry.ReadMachineValue(RegistryHive.LocalMachine, view, @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe", null);
            if (winword != null)
                Add(Path.Combine(Path.GetDirectoryName(winword) ?? "", "MSWORD.OLB"));
            foreach (var ver in new[] { "16.0", "15.0", "14.0" })
            {
                var root = ComRegistry.ReadMachineValue(RegistryHive.LocalMachine, view, @"SOFTWARE\Microsoft\Office\" + ver + @"\Word\InstallRoot", "Path");
                if (root != null)
                    Add(Path.Combine(root, "MSWORD.OLB"));
            }
        }

        return found;
    }

    private static void AddRoot(List<string> roots, string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;
        var trimmed = path!.Trim().Trim('"').TrimEnd('\\');
        if (!roots.Contains(trimmed, StringComparer.OrdinalIgnoreCase))
            roots.Add(trimmed);
    }

    private static DateTime SafeWriteTime(string path)
    {
        try
        {
            return File.Exists(path) ? File.GetLastWriteTimeUtc(path) : Directory.GetLastWriteTimeUtc(path);
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException)
        {
            return DateTime.MinValue;
        }
    }

    private static string PeBitness(string path)
    {
        try
        {
            using var stream = File.OpenRead(path);
            using var reader = new BinaryReader(stream);
            if (reader.ReadUInt16() != 0x5A4D)
                return "未知位数";
            stream.Position = 0x3C;
            stream.Position = reader.ReadInt32();
            if (reader.ReadUInt32() != 0x4550)
                return "未知位数";
            return reader.ReadUInt16() switch
            {
                0x14C => "32 位",
                0x8664 => "64 位",
                0xAA64 => "ARM64",
                _ => "未知位数"
            };
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is EndOfStreamException)
        {
            return "未知位数";
        }
    }
}
