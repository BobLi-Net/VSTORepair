using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using CT = System.Runtime.InteropServices.ComTypes;

namespace VSTORepair;

internal static class TypeLibNative
{
    private const int RegKindNone = 2;

    // LoadTypeLib 在某些路径形式下会顺手注册类型库，诊断阶段必须用 REGKIND_NONE 保证只读。
    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
    private static extern int LoadTypeLibEx(string szFile, int regkind, out CT.ITypeLib? pptlib);

    internal static TypeLibFile Load(string path)
    {
        var file = new TypeLibFile(path);
        if (!ComPath.Exists(path))
        {
            file.Hr = HResults.CantLoadLibrary;
            file.Error = "文件不存在";
            return file;
        }

        CT.ITypeLib? lib = null;
        try
        {
            var hr = LoadTypeLibEx(path, RegKindNone, out lib);
            if (hr != 0 || lib == null)
            {
                file.Hr = hr;
                file.Error = "无法加载类型库 " + HResults.Describe(hr);
                return file;
            }

            lib.GetLibAttr(out var attrPtr);
            try
            {
                var attr = Marshal.PtrToStructure<CT.TYPELIBATTR>(attrPtr);
                file.LibId = attr.guid;
                file.Major = (ushort)attr.wMajorVerNum;
                file.Minor = (ushort)attr.wMinorVerNum;
                file.SysKind = attr.syskind;
            }
            finally
            {
                lib.ReleaseTLibAttr(attrPtr);
            }

            try
            {
                lib.GetDocumentation(-1, out var name, out var doc, out _, out _);
                file.Name = string.IsNullOrWhiteSpace(doc) ? name : doc;
            }
            catch (COMException)
            {
            }

            ReadInterfaces(lib, file);
            file.Loaded = true;
        }
        catch (Exception ex)
        {
            file.Hr = Marshal.GetHRForException(ex);
            file.Error = ex.Message;
        }
        finally
        {
            if (lib != null)
                Marshal.ReleaseComObject(lib);
        }

        return file;
    }

    private static void ReadInterfaces(CT.ITypeLib lib, TypeLibFile file)
    {
        var count = lib.GetTypeInfoCount();
        for (var i = 0; i < count; i++)
        {
            lib.GetTypeInfoType(i, out var kind);
            if (kind != CT.TYPEKIND.TKIND_INTERFACE && kind != CT.TYPEKIND.TKIND_DISPATCH)
                continue;

            lib.GetTypeInfo(i, out var info);
            try
            {
                info.GetTypeAttr(out var attrPtr);
                try
                {
                    var attr = Marshal.PtrToStructure<CT.TYPEATTR>(attrPtr);
                    info.GetDocumentation(-1, out var typeName, out _, out _, out _);
                    var dual = (attr.wTypeFlags & CT.TYPEFLAGS.TYPEFLAG_FDUAL) != 0;
                    var dispatchOnly = kind == CT.TYPEKIND.TKIND_DISPATCH && !dual;
                    if (file.Interfaces.TryGetValue(attr.guid, out var existing))
                    {
                        if (!dispatchOnly)
                            existing.DispatchOnly = false;
                    }
                    else
                    {
                        file.Interfaces[attr.guid] = new InterfaceInfo(attr.guid, typeName ?? "", dispatchOnly);
                    }
                }
                finally
                {
                    info.ReleaseTypeAttr(attrPtr);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(info);
            }
        }
    }
}

internal sealed class TypeLibFile
{
    public TypeLibFile(string path)
    {
        Path = path;
    }

    public string Path { get; }
    public bool Loaded { get; set; }
    public int Hr { get; set; }
    public string? Error { get; set; }
    public Guid LibId { get; set; }
    public ushort Major { get; set; }
    public ushort Minor { get; set; }
    public CT.SYSKIND SysKind { get; set; }
    public string? Name { get; set; }
    public Dictionary<Guid, InterfaceInfo> Interfaces { get; } = new Dictionary<Guid, InterfaceInfo>();
    public string VersionKey => TypeLibVersion.Format(Major, Minor);
}

internal sealed class InterfaceInfo
{
    public InterfaceInfo(Guid iid, string name, bool dispatchOnly)
    {
        Iid = iid;
        Name = name;
        DispatchOnly = dispatchOnly;
    }

    public Guid Iid { get; }
    public string Name { get; }
    public bool DispatchOnly { get; set; }
}

internal sealed class TypeLibCache
{
    private readonly Dictionary<string, TypeLibFile> _files = new Dictionary<string, TypeLibFile>(StringComparer.OrdinalIgnoreCase);

    public IEnumerable<TypeLibFile> Loaded => _files.Values.Where(f => f.Loaded);

    public TypeLibFile Get(string path)
    {
        var key = ComPath.Normalize(path);
        if (!_files.TryGetValue(key, out var file))
        {
            file = TypeLibNative.Load(key);
            _files[key] = file;
        }

        return file;
    }
}

internal static class TypeLibVersion
{
    // 注册表里的类型库版本（键名和 Interface\TypeLib 的 Version 值）都是十六进制，例如 Office 的 "63.1"。
    public static bool TryParse(string? text, out ushort major, out ushort minor)
    {
        major = 0;
        minor = 0;
        if (string.IsNullOrWhiteSpace(text))
            return false;
        var parts = text!.Trim().Split('.');
        if (parts.Length != 2)
            return false;
        return ushort.TryParse(parts[0], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out major)
            && ushort.TryParse(parts[1], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out minor);
    }

    public static string Format(ushort major, ushort minor)
    {
        return major.ToString("x", CultureInfo.InvariantCulture) + "." + minor.ToString("x", CultureInfo.InvariantCulture);
    }
}

internal static class ComPath
{
    // 类型库路径可以带资源序号，例如 "wpsapi.dll\3"。
    private static readonly Regex ResourceSuffix = new Regex(@"\\\d+$");

    public static string FilePart(string path)
    {
        if (File.Exists(path))
            return path;
        var match = ResourceSuffix.Match(path);
        return match.Success ? path.Substring(0, match.Index) : path;
    }

    public static string Suffix(string path)
    {
        if (File.Exists(path))
            return "";
        var match = ResourceSuffix.Match(path);
        return match.Success ? match.Value : "";
    }

    public static bool Exists(string? path)
    {
        return !string.IsNullOrWhiteSpace(path) && File.Exists(FilePart(path!));
    }

    public static string Normalize(string path)
    {
        try
        {
            return System.IO.Path.GetFullPath(FilePart(path)) + Suffix(path);
        }
        catch (Exception)
        {
            return path;
        }
    }

    public static bool IsUnder(string path, string? directory)
    {
        if (string.IsNullOrWhiteSpace(directory))
            return false;
        var dir = directory!.TrimEnd('\\') + "\\";
        return path.StartsWith(dir, StringComparison.OrdinalIgnoreCase);
    }
}

internal static class HResults
{
    public const int ElementNotFound = unchecked((int)0x8002802B);
    public const int LibNotRegistered = unchecked((int)0x8002801D);
    public const int CantLoadLibrary = unchecked((int)0x80029C4A);
    public const int IidNotRegistered = unchecked((int)0x80040155);
    public const int NoInterface = unchecked((int)0x80004002);

    private static readonly Dictionary<int, string> Names = new Dictionary<int, string>
    {
        [ElementNotFound] = "TYPE_E_ELEMENTNOTFOUND 找不到元素",
        [LibNotRegistered] = "TYPE_E_LIBNOTREGISTERED 库没有注册",
        [CantLoadLibrary] = "TYPE_E_CANTLOADLIBRARY 加载类型库/DLL 时出错",
        [IidNotRegistered] = "REGDB_E_IIDNOTREG 接口未注册",
        [NoInterface] = "E_NOINTERFACE 不支持此接口",
        [unchecked((int)0x80070002)] = "找不到文件"
    };

    public static string Format(int hr) => "0x" + unchecked((uint)hr).ToString("X8", CultureInfo.InvariantCulture);

    public static string Describe(int hr)
    {
        return Names.TryGetValue(hr, out var name) ? Format(hr) + "（" + name + "）" : Format(hr);
    }

    public static bool TryParse(string text, out int hr)
    {
        var s = text.Trim();
        if (s.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            s = s.Substring(2);
        var ok = uint.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var value);
        hr = unchecked((int)value);
        return ok;
    }
}
