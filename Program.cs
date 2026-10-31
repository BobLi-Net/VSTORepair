using System.Security.Principal;
using System.Text;

namespace VSTORepair;

internal static class Program
{
    /// <summary>
    /// 命令行：
    ///   /scan [/out:文件] [/iid:{GUID}]   只诊断。退出码 0=正常 1=有问题但不能自动修 2=可以修复
    ///   /fix  [/out:文件] [/iid:{GUID}]   诊断并修复。退出码 0=修复后正常 1=仍有问题 3=被阻止（WPS 在运行等）
    ///   /undo                             撤销上次修复
    /// </summary>
    [STAThread]
    private static int Main(string[] args)
    {
        var errorText = string.Join(" ", args.Where(a => a.StartsWith("/iid:", StringComparison.OrdinalIgnoreCase)).Select(a => a.Substring(5)));
        var output = args.FirstOrDefault(a => a.StartsWith("/out:", StringComparison.OrdinalIgnoreCase))?.Substring(5);

        if (HasSwitch(args, "/scan"))
        {
            var report = new RepairService().Scan(errorText);
            Write(output, report.Text);
            return report.CanRepair ? 2 : report.HasProblem ? 1 : 0;
        }

        if (HasSwitch(args, "/fix"))
        {
            var outcome = new RepairService().Repair(errorText);
            Write(output, outcome.Text);
            return outcome.Blocked ? 3 : outcome.Success ? 0 : 1;
        }

        if (HasSwitch(args, "/undo"))
        {
            Write(output, RepairService.Undo());
            return 0;
        }

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
        return 0;
    }

    internal static bool IsAdmin()
    {
        using var identity = WindowsIdentity.GetCurrent();
        return new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
    }

    private static bool HasSwitch(string[] args, string name) => args.Any(a => a.Equals(name, StringComparison.OrdinalIgnoreCase));

    private static void Write(string? output, string text)
    {
        var path = string.IsNullOrWhiteSpace(output) ? Path.Combine(RepairService.DataDirectory, "last-result.txt") : output!;
        File.WriteAllText(path, text, new UTF8Encoding(true));
    }
}
