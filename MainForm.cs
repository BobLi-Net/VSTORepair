using System.Diagnostics;
using System.Text;

namespace VSTORepair;

internal sealed class MainForm : Form
{
    private readonly TabControl _tabs;

    // COM 修复
    private readonly TextBox _errorText;
    private readonly TextBox _log;
    private readonly Button _scanButton;
    private readonly Button _repairButton;
    private readonly Button _undoButton;
    private readonly Button _copyButton;
    private readonly Button _folderButton;
    private readonly Label _comStatus;
    private ScanReport? _report;

    // 插件管理
    private readonly ListView _pluginList;
    private readonly Label _pluginStatus;
    private readonly TextBox _pluginFilter;
    private readonly CheckBox _showWhitelistOnly;
    private readonly Button _refreshPluginsButton;
    private readonly Button _enablePluginButton;
    private readonly Button _disablePluginButton;
    private readonly Button _whitelistButton;
    private readonly Button _unwhitelistButton;
    private readonly Button _clearDisabledButton;
    private List<WordPluginInfo> _plugins = new List<WordPluginInfo>();

    public MainForm()
    {
        Text = "VSTORepair v1.5";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(900, 600);
        Size = new Size(1080, 720);
        Font = new Font("Microsoft YaHei UI", 9F);
        try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
        catch { /* 使用默认图标 */ }

        _tabs = new TabControl { Dock = DockStyle.Fill };
        _tabs.TabPages.Add(BuildComRepairTab(out _errorText, out _log, out _scanButton, out _repairButton, out _undoButton, out _copyButton, out _folderButton, out _comStatus));
        _tabs.TabPages.Add(BuildPluginTab(out _pluginList, out _pluginStatus, out _pluginFilter, out _showWhitelistOnly,
            out _refreshPluginsButton, out _enablePluginButton, out _disablePluginButton,
            out _whitelistButton, out _unwhitelistButton, out _clearDisabledButton));
        Controls.Add(_tabs);

        _tabs.SelectedIndexChanged += async (_, _) =>
        {
            if (_tabs.SelectedIndex == 1 && _pluginList.Items.Count == 0)
                await RefreshPlugins();
        };

        Shown += async (_, _) => await RunScan();
    }

    private TabPage BuildComRepairTab(
        out TextBox errorText,
        out TextBox log,
        out Button scanButton,
        out Button repairButton,
        out Button undoButton,
        out Button copyButton,
        out Button folderButton,
        out Label status)
    {
        var page = new TabPage("COM 组件修复");

        status = new Label
        {
            Dock = DockStyle.Top,
            Height = 48,
            Padding = new Padding(12, 8, 12, 0),
            Text = "排查 WPS 插件报的“找不到元素 0x8002802B”“库没有注册 0x8002801D”等错误。诊断只读；本工具以管理员运行，可清理残缺类型库版本，修复可撤销。"
        };

        var errorLabel = new Label
        {
            Dock = DockStyle.Top,
            Height = 22,
            Padding = new Padding(12, 4, 12, 0),
            Text = "可选：把插件弹出的错误文字粘贴到这里，工具会额外检查里面提到的接口："
        };
        errorText = new TextBox
        {
            Dock = DockStyle.Top,
            Height = 54,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };
        var errorPanel = new Panel { Dock = DockStyle.Top, Height = 62, Padding = new Padding(12, 2, 12, 6) };
        errorPanel.Controls.Add(errorText);

        log = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            Font = new Font("Consolas", 10F),
            HideSelection = false
        };

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 52,
            Padding = new Padding(10, 8, 10, 8),
            FlowDirection = FlowDirection.LeftToRight
        };
        scanButton = MakeButton("诊断", async () => await RunScan());
        repairButton = MakeButton("修复", async () => await RunRepair());
        undoButton = MakeButton("撤销上次修复", async () => await RunUndo());
        copyButton = MakeButton("复制日志", () =>
        {
            if (_log.Text.Length > 0)
            {
                Clipboard.SetText(_log.Text);
                _comStatus.Text = "日志已复制，可以直接发给技术支持。";
            }

            return Task.CompletedTask;
        });
        folderButton = MakeButton("打开日志目录", () =>
        {
            Process.Start(new ProcessStartInfo { FileName = RepairService.DataDirectory, UseShellExecute = true });
            return Task.CompletedTask;
        });
        undoButton.Width = 120;
        folderButton.Width = 120;
        buttons.Controls.AddRange(new Control[] { scanButton, repairButton, undoButton, copyButton, folderButton });

        page.Controls.Add(log);
        page.Controls.Add(buttons);
        page.Controls.Add(errorPanel);
        page.Controls.Add(errorLabel);
        page.Controls.Add(status);
        return page;
    }

    private TabPage BuildPluginTab(
        out ListView list,
        out Label status,
        out TextBox filter,
        out CheckBox showWhitelistOnly,
        out Button refresh,
        out Button enable,
        out Button disable,
        out Button whitelist,
        out Button unwhitelist,
        out Button clearDisabled)
    {
        var page = new TabPage("插件管理");

        status = new Label
        {
            Dock = DockStyle.Top,
            Height = 44,
            Padding = new Padding(12, 8, 12, 0),
            Text = "列出 Word / WPS 文字的 COM/VSTO 加载项（含已禁用、加载失败断开）。修改后需完全退出并重新打开 WPS/Word 才会生效。"
        };

        var filterPanel = new Panel { Dock = DockStyle.Top, Height = 36, Padding = new Padding(12, 2, 12, 4) };
        var filterLabel = new Label { Text = "筛选：", AutoSize = true, Left = 0, Top = 8 };
        filter = new TextBox { Left = 42, Top = 4, Width = 220 };
        filter.TextChanged += (_, _) => ApplyPluginFilter();
        showWhitelistOnly = new CheckBox
        {
            Text = "显示仅白名单残留（未在 Word\\Addins 注册）",
            AutoSize = true,
            Left = 280,
            Top = 6
        };
        showWhitelistOnly.CheckedChanged += (_, _) => ApplyPluginFilter();
        filterPanel.Controls.Add(filterLabel);
        filterPanel.Controls.Add(filter);
        filterPanel.Controls.Add(showWhitelistOnly);

        list = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            MultiSelect = false,
            HideSelection = false
        };
        list.Columns.Add("名称", 180);
        list.Columns.Add("ProgId", 200);
        list.Columns.Add("状态", 180);
        list.Columns.Add("WPS", 100);
        list.Columns.Add("注册位置", 110);
        list.Columns.Add("清单/路径", 280);
        list.SelectedIndexChanged += (_, _) => UpdatePluginButtons();

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 52,
            Padding = new Padding(10, 8, 10, 8),
            FlowDirection = FlowDirection.LeftToRight
        };
        refresh = MakeButton("刷新", async () => await RefreshPlugins());
        enable = MakeButton("启用加载", async () => await RunPluginAction(p => new PluginService().Enable(p)));
        disable = MakeButton("禁用卸载", async () => await RunPluginAction(p => new PluginService().Disable(p)));
        whitelist = MakeButton("加入白名单", async () => await RunPluginAction(p => new PluginService().AddWhitelist(p)));
        unwhitelist = MakeButton("移出白名单", async () => await RunPluginAction(p => new PluginService().RemoveWhitelist(p)));
        clearDisabled = MakeButton("清除自动禁用", async () =>
        {
            var msg = await Task.Run(() => new PluginService().ClearOfficeDisabledItems());
            _pluginStatus.Text = msg;
            MessageBox.Show(this, msg, Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        });
        enable.Width = 100;
        disable.Width = 100;
        whitelist.Width = 100;
        unwhitelist.Width = 100;
        clearDisabled.Width = 120;
        buttons.Controls.AddRange(new Control[] { refresh, enable, disable, whitelist, unwhitelist, clearDisabled });

        page.Controls.Add(list);
        page.Controls.Add(buttons);
        page.Controls.Add(filterPanel);
        page.Controls.Add(status);
        return page;
    }

    private static Button MakeButton(string text, Func<Task> action)
    {
        var button = new Button { Text = text, Width = 100, Height = 30 };
        button.Click += async (_, _) => await action();
        return button;
    }

    private async Task RunScan()
    {
        var errorText = _errorText.Text;
        await BusyCom("正在诊断……", async () =>
        {
            _report = await Task.Run(() => new RepairService().Scan(errorText));
            ShowText(_report.Text);
            _comStatus.Text = _report.CanRepair
                ? "发现问题，可以修复。修复前请完全退出 WPS。"
                : _report.HasProblem
                    ? "发现问题，但无法只靠修改注册表解决，请查看结论。"
                    : "未发现接口注册问题。";
        });
    }

    private async Task RunRepair()
    {
        if (_report == null || !_report.CanRepair)
        {
            MessageBox.Show(this, _report?.Summary ?? "请先诊断。", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var blocker = RepairService.RepairBlocker();
        if (blocker != null)
        {
            MessageBox.Show(this, blocker, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var preview = new StringBuilder();
        preview.AppendLine(_report.Summary);
        preview.AppendLine();
        preview.Append("是否继续修复？");
        if (MessageBox.Show(this, preview.ToString(), Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        var errorText = _errorText.Text;
        await BusyCom("正在修复……", async () =>
        {
            var outcome = await Task.Run(() => new RepairService().Repair(errorText));
            ShowText(outcome.Text);
            _report = null;
            _comStatus.Text = outcome.Blocked ? "修复被阻止。"
                : outcome.Success ? "修复完成，复查通过。请重新打开 WPS 验证。"
                : "修复后仍有问题，请查看日志。";
            MessageBox.Show(this, _comStatus.Text, Text, MessageBoxButtons.OK,
                outcome.Success ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        });
    }

    private async Task RunUndo()
    {
        if (MessageBox.Show(this, "撤销本工具上一次写入的所有注册表修改？", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;
        await BusyCom("正在撤销……", async () =>
        {
            var text = await Task.Run(RepairService.Undo);
            ShowText(text);
            _report = null;
            _comStatus.Text = "撤销完成。";
        });
    }

    private async Task RefreshPlugins()
    {
        await BusyPlugin("正在扫描插件……", async () =>
        {
            _plugins = await Task.Run(() => new PluginService().ListPlugins());
            ApplyPluginFilter();
            var registered = _plugins.Count(p => p.InUser || p.InMachine);
            var orphanWl = _plugins.Count - registered;
            _pluginStatus.Text = "已注册 " + registered + " 个；另有 " + orphanWl
                + " 个仅白名单残留（默认隐藏）。选中后可启用、禁用或调整 WPS 白名单。";
        });
    }

    private void ApplyPluginFilter()
    {
        var q = (_pluginFilter.Text ?? "").Trim();
        var showOrphanWl = _showWhitelistOnly.Checked;
        _pluginList.BeginUpdate();
        _pluginList.Items.Clear();
        foreach (var p in _plugins)
        {
            var registered = p.InUser || p.InMachine;
            if (!registered && !showOrphanWl)
                continue;

            if (q.Length > 0
                && p.FriendlyName.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0
                && p.ProgId.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0
                && (p.Manifest == null || p.Manifest.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0)
                && p.StatusText.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0)
                continue;

            var item = new ListViewItem(p.FriendlyName);
            item.SubItems.Add(p.ProgId);
            item.SubItems.Add(p.StatusText);
            item.SubItems.Add(p.WpsText);
            item.SubItems.Add(p.ScopeText);
            item.SubItems.Add(p.Manifest ?? "");
            item.Tag = p;
            if (!p.IsEnabled)
                item.ForeColor = Color.DimGray;
            if (p.InWpsBlacklist)
                item.ForeColor = Color.DarkRed;
            _pluginList.Items.Add(item);
        }

        _pluginList.EndUpdate();
        UpdatePluginButtons();
    }

    private WordPluginInfo? SelectedPlugin() =>
        _pluginList.SelectedItems.Count == 0 ? null : _pluginList.SelectedItems[0].Tag as WordPluginInfo;

    private void UpdatePluginButtons()
    {
        var has = SelectedPlugin() != null;
        _enablePluginButton.Enabled = has;
        _disablePluginButton.Enabled = has;
        _whitelistButton.Enabled = has;
        _unwhitelistButton.Enabled = has;
    }

    private async Task RunPluginAction(Func<WordPluginInfo, string> action)
    {
        var plugin = SelectedPlugin();
        if (plugin == null)
        {
            MessageBox.Show(this, "请先在列表中选中一个插件。", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var progId = plugin.ProgId;
        await BusyPlugin("正在处理 " + progId + "……", async () =>
        {
            var msg = await Task.Run(() => action(plugin));
            _plugins = await Task.Run(() => new PluginService().ListPlugins());
            ApplyPluginFilter();
            SelectPlugin(progId);
            _pluginStatus.Text = msg;
            MessageBox.Show(this, msg, Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        });
    }

    private void SelectPlugin(string progId)
    {
        foreach (ListViewItem item in _pluginList.Items)
        {
            if (item.Tag is WordPluginInfo p && string.Equals(p.ProgId, progId, StringComparison.OrdinalIgnoreCase))
            {
                item.Selected = true;
                item.EnsureVisible();
                break;
            }
        }
    }

    private async Task BusyCom(string message, Func<Task> work)
    {
        foreach (var b in new[] { _scanButton, _repairButton, _undoButton })
            b.Enabled = false;
        UseWaitCursor = true;
        _comStatus.Text = message;
        try
        {
            await work();
        }
        catch (Exception ex)
        {
            ShowText(ex.ToString());
            _comStatus.Text = "出错了，请把日志发给技术支持。";
        }
        finally
        {
            UseWaitCursor = false;
            foreach (var b in new[] { _scanButton, _repairButton, _undoButton })
                b.Enabled = true;
        }
    }

    private async Task BusyPlugin(string message, Func<Task> work)
    {
        foreach (var b in new[] { _refreshPluginsButton, _enablePluginButton, _disablePluginButton, _whitelistButton, _unwhitelistButton, _clearDisabledButton })
            b.Enabled = false;
        UseWaitCursor = true;
        _pluginStatus.Text = message;
        try
        {
            await work();
        }
        catch (Exception ex)
        {
            _pluginStatus.Text = "出错了：" + ex.Message;
            MessageBox.Show(this, ex.ToString(), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            UseWaitCursor = false;
            _refreshPluginsButton.Enabled = true;
            _clearDisabledButton.Enabled = true;
            UpdatePluginButtons();
        }
    }

    private void ShowText(string text)
    {
        _log.Text = text.Replace("\r\n", "\n").Replace("\n", "\r\n");
        _log.SelectionStart = 0;
        _log.ScrollToCaret();
    }
}
