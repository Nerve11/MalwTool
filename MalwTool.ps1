# https://github.com/ImMALWARE/MalwTool & https://github.com/Nerve11/MalwTool
Add-Type -AssemblyName System.Windows.Forms
Add-Type -Name Window -Namespace Console -MemberDefinition '[DllImport("Kernel32.dll")]public static extern IntPtr GetConsoleWindow();[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'
[void][Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)
[System.Windows.Forms.Application]::EnableVisualStyles()
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$wc = New-Object net.webclient
$mb = [System.Windows.Forms.MessageBox]::Show
$app = 'MalwTool'
$n = [Environment]::NewLine
$hidden = New-Object System.Windows.Forms.Form -Property @{
    TopMost = $true
    WindowState = "Minimized"
}

$strings = @{
    tabs = @{
        activation = "Activation"
        download = "Download"
        other = "Other functions"
        problems = "Troubleshooting"
        info = "Info"
    }
    act = @{
        hwid = "Windows 10 or 11 (including LTSC) activation by HWID"
        kms = "Windows 10 or 11 (including LTSC) activation via KMS"
        evaluation = "Convert evaluation version of Windows 10 LTSC to full LTSC"
        server = "Activation of Windows Server 2025, Windows Server Standard, Windows Server Datacenter, 2022, 2019, 2016, 2012, 2012 R2, 1803, 1709"
        visio = "Via KMS, will be activated as %p% 2021 (older versions will be updated)"
        office = "Office %v% activation via sppc.dll file$n" + "%info%" + "The activation will also work for Office %otherv%. Office will then be automatically converted to %v%."
        o365info = "Office 365 is always the latest version of Office, it is better to choose this option.$n"
        o2024info = "Anyway, I would recommend selecting Office 365.$n"
        o2013 = "Office 2013 activation via sppc.dll file."
        prism = "Allow creation of an offline Minecraft account in Prism Launcher without Microsoft account.$nDo not start if you have already added an account! This action will delete all accounts in the launcher!"
        tl = "Premium account in TL, you will be able to disable adding advertised servers in its settings"
        act = "Activate!"
        notfound = "%p% not found!"
    }
    mc = @{
        prismok = "Offline account in Prism Launcher unlocked!"
        tlok = "Premium account in TL is activated! Now open its settings to disable advertised servers!"
        tlnoacc = "Log in to account in TL first!"
    }
    download = @{
        windows = "ISO image of the latest version of Windows %v% from the official Microsoft website"
        oonline = "Office online installer from the official Microsoft website."
        installer = " Follow the instructions of $app after starting the installer."
        oiso = "ISO archive of Office from the official Microsoft website. Run setup.exe in it."
        outdated = " Not recommended, outdated version. "
        rufus = "tool for writing ISO images to a flash drive"
        linkfail = "Failed to get a download link!$nIf retrying did not help after several tries, you can download a reupload from Archive.org"
        btn_retry = "Retry"
        btn_reupload = "Download reupload"
        btn_cancel = "Cancel"
        recommended = "recommended"
        bypass = 'For users in Russia: To install online, you need to bypass regional restrictions. Run the .exe file, wait for the "Command not supported" error, and click "Yes" in that window. Then restart the installer file. If you are not in Russia, click "No".'
        bypass_done = "Now start the installer again!"
    }
    functions = @{
        wifi = "Get passwords from saved Wi-Fi networks"
        wifi_desc = "Before starting, make sure that Wi-Fi is currently enabled"
        install = "Install "
        fileext = "Show file extensions in Explorer"
        store_desc = "For LTSC versions of Windows without installed Microsoft Store"
        store_installing = "Microsoft Store is installing! Wait a few minutes; it should appear in Start."
        drivers = "Backup drivers"
        drivers_desc = 'Before reinstalling Windows, it is best to back up all drivers so you do not have to struggle with them afterward - just select "Restore" here.'
        drivers_restore = "Restore drivers"
        drivers_select = "Select directory for creating/restoring drivers backup"
        edge = "Completely uninstall Microsoft Edge"
        compattel = "Delete system spy programs"
        compattel_desc = "Delete CompatTelRunner.exe and wsqmcons.exe"
        compattel_title = "Deleting CompatTelRunner.exe and wsqmcons.exe"
        spicetify_desc = "Spotify mod"
        hosts = "Bypass geo-blocks via hosts"
        hosts_desc = "For Russian users only. Set hosts file from dns.malw.link$nMore info at info.dns.malw.link"
        hosts_title = "Editing hosts file"
    }
    problems = @{
        cl_o16 = "Clear Office16 licenses"
        cl_o16_desc = "Only for KMS activations. Won't clear activation from $app."
        cl_o16_err = "Office16 folder not found!"
        cl_o16_title = "Clearing Office16 licenses"
        o16_folder = "Office16 folder is missing"
        o16_exists = "Folder Office16 already exists!"
        o16_title = "Downloading Office16 folder"
        ouninstall = "Office removal tool"
        cl_win = "Clear Windows KMS activation"
        cl_win_title = "Clearing Windows activation"
        sfc = "Check system files for integrity"
        sfc_shutdown_desc = "In 60 seconds there will be a reboot to check the system disk!"
        sfc_title = "Checking system"
        sfc_note = "!!! Now press Y and Enter for checking disk!!!"
        other = "I have another problem!"
        other_desc = "Contact me, even if the problem is unrelated to $app"
        ouninstall_nogethelp = "You must install Get Help app for this function! Install it and run again."
        nostore = "You must install Get Help app for this function. But to install it, you need Microsoft Store. Install it in the Other functions tab."
    }
    info = @{
        wiki = "Page on Wiki.malw.link"
        lolz = "Lolzteam thread"
        questions = "Questions? "
        telegram = "Message in Telegram"
        lolz_write = "Write in Lolzteam thread"
        github = "Write on GitHub Issues"
    }
    links = @{
        w10ltsc = "https://archive.org/download/en-us_windows_10_iot_enterprise_ltsc_2021_x64_dvd_257ad90f_20260322/en-us_windows_10_iot_enterprise_ltsc_2021_x64_dvd_257ad90f.iso"
        w11ltsc = "https://archive.org/download/en-us_windows_11_iot_enterprise_ltsc_2024_x64_dvd_f6b14814_202603/en-us_windows_11_iot_enterprise_ltsc_2024_x64_dvd_f6b14814.iso"
        ws2025 = "https://archive.org/download/en-us_windows_server_2025_updated_march_2026_x64_dvd_8e06425a/en-us_windows_server_2025_updated_march_2026_x64_dvd_8e06425a.iso"
        ws2022 = "https://archive.org/download/en-us_windows_server_2022_updated_march_2026_x64_dvd_3f772967/en-us_windows_server_2022_updated_march_2026_x64_dvd_3f772967.iso"
        ws2019 = "https://archive.org/download/en-us_windows_server_2019_x64_dvd_f9475476_202603/en-us_windows_server_2019_x64_dvd_f9475476.iso"
        w81 = "https://archive.org/download/en_windows_8.1_with_update_x64_dvd_6051480_202603/en_windows_8.1_with_update_x64_dvd_6051480.iso"
        o2024o = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2024Retail&platform=x64&language=en-us&version=O16GA"
        o2024i = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2024Retail.img"
        o2021o = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2021Retail&platform=x64&language=en-us&version=O16GA"
        o2021i = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img"
        o2019o = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2019Retail&platform=x64&language=en-us&version=O16GA"
        o2019i = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img"
        o2016o = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=en-us&version=O16GA"
        o2016i = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlusRetail.img"
        o2013o = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=en-us&version=O15GA"
        o2013i = "https://officecdn.microsoft.com/db/39168d7e-077b-48e7-872c-b232c3e72675/media/en-us/ProfessionalRetail.img"
        vis2024 = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/VisioPro2024Retail.img"
        proj2024 = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProjectPro2024Retail.img"
    }
}
$gstrings = @("irm https://raw.githubusercontent.com/ImMALWARE/$app/main/Activators", " | iex", "$env:ProgramFiles\Microsoft Office\root\vfs\System", "${env:ProgramFiles(x86)}\Microsoft Office\root\vfs\SystemX86")
$activators = @{"Win10" = "HWID.ps1"; "WinKM" = ("KM" + "S10.ps1"); "ConvertEvaluationToFull" = "LTSCEvaluationToFull.ps1"; "WinServer" = ("ServerKM" + "S.ps1"); "OfficeVisio" = "VisioProject.ps1"; "OfficeProject" = "VisioProject.ps1"; "MobaXterm" = "MXT.ps1"; "Office 365" = "Osppcs.ps1"; "Office 2024" = "Osppcs.ps1"; "Office 2021" = "Osppcs.ps1"; "Office 2019" = "Osppcs.ps1"; "Office 2016" = "Osppcs.ps1"; "Office 2013" = ("KM" + "S2013.ps1")}
$paths = @{"Office 365" = @($gstrings[2], $gstrings[3]); "Office 2024" = @($gstrings[2], $gstrings[3]); "Office 2021" = @($gstrings[2], $gstrings[3]); "Office 2019" = @($gstrings[2], $gstrings[3]); "Office 2016" = @($gstrings[2], $gstrings[3]); "Office 2013" = @("$env:ProgramFiles\Microsoft Office 15\root\vfs\System", "${env:ProgramFiles(x86)}\Microsoft Office 15\root\vfs\System"); "Prism Launcher" = @("$env:appdata\PrismLauncher"); "TL" = @("$env:appdata\.minecraft\TlauncherProfiles.json"); "MobaXterm" = @("${env:ProgramFiles(x86)}\Mobatek\MobaXterm\version.dat")}

function check_path ($paths, $prod) {
    foreach ($path in $paths) {
        if (Test-Path $path) {
            return $true
        }
    }
    $null = $mb.Invoke($strings.act.notfound.Replace('%p%', $prod))
    return $false
}

$theme = @{
    formBack = [System.Drawing.ColorTranslator]::FromHtml('#F3F5F9')
    surface = [System.Drawing.Color]::White
    surfaceAlt = [System.Drawing.ColorTranslator]::FromHtml('#EEF2FF')
    surfaceMuted = [System.Drawing.ColorTranslator]::FromHtml('#F8FAFC')
    border = [System.Drawing.ColorTranslator]::FromHtml('#D7DFEA')
    borderStrong = [System.Drawing.ColorTranslator]::FromHtml('#C4D0E3')
    accent = [System.Drawing.ColorTranslator]::FromHtml('#2563EB')
    accentSoft = [System.Drawing.ColorTranslator]::FromHtml('#DBEAFE')
    text = [System.Drawing.ColorTranslator]::FromHtml('#0F172A')
    textMuted = [System.Drawing.ColorTranslator]::FromHtml('#475569')
    textSoft = [System.Drawing.ColorTranslator]::FromHtml('#64748B')
    tabIdle = [System.Drawing.ColorTranslator]::FromHtml('#E8EEF8')
    tabActive = [System.Drawing.Color]::White
    success = [System.Drawing.ColorTranslator]::FromHtml('#16A34A')
}

function Set-ButtonTheme {
    param(
        [System.Windows.Forms.Button]$Button,
        [switch]$Primary
    )

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.UseVisualStyleBackColor = $false
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.BackColor = if ($Primary) { $theme.accent } else { $theme.textMuted }
}

function Set-RadioTheme {
    param([System.Windows.Forms.RadioButton]$Radio)

    $Radio.ForeColor = $theme.text
    $Radio.BackColor = $theme.surface
    $Radio.Cursor = [System.Windows.Forms.Cursors]::Hand
}

function New-SectionPanel {
    param(
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [string]$Title,
        [string]$Subtitle = ''
    )

    $panel = New-Object System.Windows.Forms.Panel -Property @{
        Location = [System.Drawing.Point]::new($X, $Y)
        Size = [System.Drawing.Size]::new($Width, $Height)
        BackColor = $theme.surface
        BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    }

    $titleLabel = New-Object System.Windows.Forms.Label -Property @{
        AutoSize = $true
        Location = [System.Drawing.Point]::new(16, 14)
        Font = [System.Drawing.Font]::new('Segoe UI Semibold', 11)
        ForeColor = $theme.text
        Text = $Title
        BackColor = $theme.surface
    }
    $panel.Controls.Add($titleLabel)

    if ($Subtitle) {
        $subtitleLabel = New-Object System.Windows.Forms.Label -Property @{
            Location = [System.Drawing.Point]::new(16, 38)
            Size = [System.Drawing.Size]::new($Width - 32, 32)
            Font = [System.Drawing.Font]::new('Segoe UI', 8.5)
            ForeColor = $theme.textSoft
            Text = $Subtitle
            BackColor = $theme.surface
        }
        $panel.Controls.Add($subtitleLabel)
    }

    return $panel
}

function New-BodyLabel {
    param(
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [string]$Text,
        [float]$FontSize = 8.5,
        [System.Drawing.Color]$ForeColor = $theme.textMuted
    )

    return (New-Object System.Windows.Forms.Label -Property @{
        Location = [System.Drawing.Point]::new($X, $Y)
        Size = [System.Drawing.Size]::new($Width, $Height)
        Font = [System.Drawing.Font]::new('Segoe UI', $FontSize)
        ForeColor = $ForeColor
        Text = $Text
        BackColor = [System.Drawing.Color]::Transparent
    })
}

$form = New-Object System.Windows.Forms.Form -Property @{
    StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    MaximizeBox = $false
    Text = $app
    ClientSize = [System.Drawing.Size]::new(980, 700)
    MinimumSize = [System.Drawing.Size]::new(996, 739)
    MaximumSize = [System.Drawing.Size]::new(996, 739)
    Font = [System.Drawing.Font]::new('Segoe UI', 9)
    BackColor = $theme.formBack
}

$headerPanel = New-Object System.Windows.Forms.Panel -Property @{
    Location = [System.Drawing.Point]::new(20, 18)
    Size = [System.Drawing.Size]::new(940, 72)
    BackColor = $theme.surface
    BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
}

$headerTitle = New-Object System.Windows.Forms.Label -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(20, 16)
    Font = [System.Drawing.Font]::new('Segoe UI Semibold', 17)
    ForeColor = $theme.text
    Text = $app
    BackColor = $theme.surface
}

$headerSubtitle = New-Object System.Windows.Forms.Label -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(22, 45)
    Font = [System.Drawing.Font]::new('Segoe UI', 9)
    ForeColor = $theme.textMuted
    Text = 'Refreshed interface with the same features: activation, downloads, recovery, and utilities.'
    BackColor = $theme.surface
}

$headerBadge = New-Object System.Windows.Forms.Label -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(876, 25)
    Padding = [System.Windows.Forms.Padding]::new(10, 4, 10, 4)
    Font = [System.Drawing.Font]::new('Segoe UI Semibold', 9)
    ForeColor = $theme.accent
    BackColor = $theme.accentSoft
    Text = '2.0'
}

$headerPanel.Controls.AddRange(@($headerTitle, $headerSubtitle, $headerBadge))

$tabs = New-Object System.Windows.Forms.TabControl -Property @{
    Location = [System.Drawing.Point]::new(20, 105)
    SelectedIndex = 0
    Size = [System.Drawing.Size]::new(940, 575)
    SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
    ItemSize = [System.Drawing.Size]::new(180, 34)
    DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
    Padding = [System.Drawing.Point]::new(14, 6)
    Appearance = [System.Windows.Forms.TabAppearance]::Normal
}

$ActTab = New-Object System.Windows.Forms.TabPage -Property @{
    Padding = [System.Windows.Forms.Padding]::new(0)
    Text = $strings.tabs.activation
    BackColor = $theme.formBack
}

$DlTab = New-Object System.Windows.Forms.TabPage -Property @{
    Padding = [System.Windows.Forms.Padding]::new(0)
    Text = $strings.tabs.download
    BackColor = $theme.formBack
}

$FunctionsTab = New-Object System.Windows.Forms.TabPage -Property @{
    Text = $strings.tabs.other
    BackColor = $theme.formBack
}

$ProblemsTab = New-Object System.Windows.Forms.TabPage -Property @{
    Padding = [System.Windows.Forms.Padding]::new(0)
    Text = $strings.tabs.problems
    BackColor = $theme.formBack
}

$InfoTab = New-Object System.Windows.Forms.TabPage -Property @{
    Text = $strings.tabs.info
    BackColor = $theme.formBack
}

@($ActTab, $DlTab, $FunctionsTab, $ProblemsTab, $InfoTab) | ForEach-Object { $tabs.TabPages.Add($_) }

$tabs.Add_DrawItem({
    param($sender, $e)

    $rect = $e.Bounds
    $selected = ($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -eq [System.Windows.Forms.DrawItemState]::Selected
    $bgBrush = New-Object System.Drawing.SolidBrush($(if ($selected) { $theme.tabActive } else { $theme.tabIdle }))
    $textColor = if ($selected) { $theme.text } else { $theme.textMuted }

    $e.Graphics.FillRectangle($bgBrush, $rect)
    $bgBrush.Dispose()
    $textRect = [System.Drawing.Rectangle]::new($rect.X, $rect.Y + 1, $rect.Width, $rect.Height - 2)
    [System.Windows.Forms.TextRenderer]::DrawText(
        $e.Graphics,
        $sender.TabPages[$e.Index].Text,
        [System.Drawing.Font]::new('Segoe UI Semibold', 9),
        $textRect,
        $textColor,
        [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter
    )
})

$tooltip = New-Object System.Windows.Forms.ToolTip -Property @{
    AutoPopDelay = 5000
    InitialDelay = 5
    ReshowDelay = 100
    ShowAlways = $true
}

# Activation tab

$actSystemPanel = New-SectionPanel 20 20 280 235 'Windows and Server' 'Quick activation choices for desktop and server editions.'
$actOfficePanel = New-SectionPanel 320 20 280 235 'Office and Microsoft 365' 'Office, Visio, and Project options grouped together for easier comparison.'
$actAppsPanel = New-SectionPanel 620 20 280 235 'Apps and launchers' 'Actions for third-party apps and game launchers.'
$actSummaryPanel = New-SectionPanel 20 275 880 230 'Ready to run' 'Confirm the selected scenario before launch. The commands stay the same; only the presentation is refreshed.'

$ActWin10 = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Checked = $true
    Location = [System.Drawing.Point]::new(18, 78)
    Name = "Win10"
    Text = "Windows 10/11 (HWID)"
}
$tooltip.SetToolTip($ActWin10, $strings.act.hwid)
Set-RadioTheme $ActWin10

$ActWinKM = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 108)
    Name = "WinKM"
    Text = ("Windows 10/11 (KM" + "S)")
}
$tooltip.SetToolTip($ActWinKM, $strings.act.kms)
Set-RadioTheme $ActWinKM

$ConvertEvaluationToFull = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 138)
    Name = "ConvertEvaluationToFull"
    Text = "LTSC Evaluation -> LTSC"
}
$tooltip.SetToolTip($ConvertEvaluationToFull, $strings.act.evaluation)
Set-RadioTheme $ConvertEvaluationToFull

$ActWinServer = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 168)
    Name = "WinServer"
    Text = "Win Server 2025/2022/2019/2016"
}
$tooltip.SetToolTip($ActWinServer, $strings.act.server)
Set-RadioTheme $ActWinServer

$ActVisio = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 78)
    Name = "OfficeVisio"
    Text = "Visio 2016/2019/2021"
}
$tooltip.SetToolTip($ActVisio, $strings.act.visio.Replace("%p%", "Visio"))
Set-RadioTheme $ActVisio

$ActProject = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 108)
    Name = "OfficeProject"
    Text = "Project 2016/2019/2021"
}
$tooltip.SetToolTip($ActProject, $strings.act.visio.Replace("%p%", "Project"))
Set-RadioTheme $ActProject

$ActOffice365 = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 138)
    Name = "Office 365"
    Text = "Office 365"
}
$tooltip.SetToolTip($ActOffice365, $strings.act.office.Replace("%v%", "365").Replace("%info%", $strings.act.o365info).Replace("%otherv%", "2016, 2019, 2021, 2024"))
Set-RadioTheme $ActOffice365

$ActOffice2024 = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 168)
    Name = "Office 2024"
    Text = "Office 2024"
}
$tooltip.SetToolTip($ActOffice2024, $strings.act.office.Replace("%v%", "2024").Replace("%info%", $strings.act.o2024info).Replace("%otherv%", "2016, 2019, 2021"))
Set-RadioTheme $ActOffice2024

$ActOffice2021 = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(140, 138)
    Name = "Office 2021"
    Text = "Office 2021"
}
$tooltip.SetToolTip($ActOffice2021, $strings.act.office.Replace("%v%", "2021").Replace("%info%", "").Replace("%otherv%", "2016, 2019, 2024"))
Set-RadioTheme $ActOffice2021

$ActOffice2019 = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(140, 168)
    Name = "Office 2019"
    Text = "Office 2019"
}
$tooltip.SetToolTip($ActOffice2019, $strings.act.office.Replace("%v%", "2019").Replace("%info%", "").Replace("%otherv%", "2016, 2021, 2024"))
Set-RadioTheme $ActOffice2019

$ActOffice2016 = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 198)
    Name = "Office 2016"
    Text = "Office 2016"
}
$tooltip.SetToolTip($ActOffice2016, $strings.act.office.Replace("%v%", "2016").Replace("%info%", "").Replace("%otherv%", "2019, 2021, 2024"))
Set-RadioTheme $ActOffice2016

$ActOffice2013 = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(140, 198)
    Name = "Office 2013"
    Text = "Office 2013"
}
$tooltip.SetToolTip($ActOffice2013, $strings.act.o2013)
Set-RadioTheme $ActOffice2013

$ActPrismLauncher = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 78)
    Name = "Prism Launcher"
    Text = "Prism Launcher"
}
$tooltip.SetToolTip($ActPrismLauncher, $strings.act.prism)
Set-RadioTheme $ActPrismLauncher

$ActTL = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 108)
    Name = "TL"
    Text = "TL"
}
$tooltip.SetToolTip($ActTL, $strings.act.tl)
Set-RadioTheme $ActTL

$ActMXT = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 138)
    Name = "MobaXterm"
    Text = "MobaXterm"
}
Set-RadioTheme $ActMXT

$actSelectedLabel = New-Object System.Windows.Forms.Label -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 78)
    Font = [System.Drawing.Font]::new('Segoe UI Semibold', 12)
    ForeColor = $theme.text
    BackColor = $theme.surface
    Text = ''
}

$actSummaryLabel = New-BodyLabel 18 108 660 78 '' 9 $theme.textMuted
$actSummaryLabel.BackColor = $theme.surface

$actHintLabel = New-BodyLabel 18 188 660 24 'When needed, the script will request elevation through UAC on its own.' 8.5 $theme.textSoft
$actHintLabel.BackColor = $theme.surface

$Act = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(708, 160)
    Size = [System.Drawing.Size]::new(150, 42)
    Font = [System.Drawing.Font]::new('Segoe UI Semibold', 10)
    Text = $strings.act.act
}
Set-ButtonTheme -Button $Act -Primary

function Update-ActivationSummary {
    $selected = Get-ActivationSelection
    if ($null -eq $selected) { return }
    $actSelectedLabel.Text = $selected.Text
    $actSummaryLabel.Text = $tooltip.GetToolTip($selected)
}

function Get-ActivationSelection {
    foreach ($panel in $ActTab.Controls) {
        foreach ($control in $panel.Controls) {
            if ($control -is [System.Windows.Forms.RadioButton] -and $control.Checked) {
                return $control
            }
        }
    }
}

@($ActWin10, $ActWinKM, $ConvertEvaluationToFull, $ActWinServer) | ForEach-Object { $actSystemPanel.Controls.Add($_) }
@($ActVisio, $ActProject, $ActOffice365, $ActOffice2024, $ActOffice2021, $ActOffice2019, $ActOffice2016, $ActOffice2013) | ForEach-Object { $actOfficePanel.Controls.Add($_) }
@($ActPrismLauncher, $ActTL, $ActMXT) | ForEach-Object { $actAppsPanel.Controls.Add($_) }
@($actSelectedLabel, $actSummaryLabel, $actHintLabel, $Act) | ForEach-Object { $actSummaryPanel.Controls.Add($_) }

@($ActWin10, $ActWinKM, $ConvertEvaluationToFull, $ActWinServer, $ActVisio, $ActProject, $ActOffice365, $ActOffice2024, $ActOffice2021, $ActOffice2019, $ActOffice2016, $ActOffice2013, $ActPrismLauncher, $ActTL, $ActMXT) | ForEach-Object {
    $_.Add_CheckedChanged({
        if ($this.Checked) {
            Update-ActivationSummary
        }
    })
}

@($actSystemPanel, $actOfficePanel, $actAppsPanel, $actSummaryPanel) | ForEach-Object { $ActTab.Controls.Add($_) }
Update-ActivationSummary

$Act.Add_Click({
    $prod = (Get-ActivationSelection).Name
    switch ($prod) {
        "Prism Launcher" {
            if (!(check_path $paths[$prod] $prod)) {break}
            '{"accounts": [{"entitlement": {"canPlayMinecraft": true, "ownsMinecraft": true}, "type": "MSA"}], "formatVersion": 3}' | Out-File (($paths[$prod])[0] + "\accounts.json") -Encoding ASCII
            $mb.Invoke($strings.mc.prismok, $app, "OK", "Information")
        }
        "TL" {
            if (!(check_path $paths[$prod] $prod)) {break}
            $file = Get-Content -Path $paths[$prod] -Raw
            if ($file -match '"premiumAccount": false') {
                $file -replace '"premiumAccount": false', '"premiumAccount": true' | Set-Content -Path $paths[$prod]
                $mb.Invoke($strings.mc.tlok, $app, "OK", "Information")
            }
            else {
                $mb.Invoke($strings.mc.tlnoacc, $app, "OK", "Warning")
            }
        }
        default {
            if ($paths[$prod] -ne $null -and -not (check_path $paths[$prod] $prod)) {break}
            Start-Process powershell -ArgumentList ("`$Product = '$prod'; " + $gstrings[0] + "/" + $activators[$prod] + $gstrings[1]) -Verb RunAs
        }
    }
})

# Downloads tab

$dlWindowsPanel = New-SectionPanel 20 20 280 250 'Windows' 'Fast access to current Windows builds and LTSC editions.'
$dlServerPanel = New-SectionPanel 320 20 220 180 'Windows Server' 'Direct links to server ISO images.'
$dlOfficePanel = New-SectionPanel 560 20 340 310 'Office' 'Online installers and ISO images placed side by side for faster comparison.'
$dlExtrasPanel = New-SectionPanel 20 290 520 175 'Extras' 'Extras for media creation and standalone Visio and Project downloads.'
$dlGuidePanel = New-SectionPanel 560 350 340 115 'Tip' 'Office online installers may require a regional bypass. The script will guide you after launch.'

$DlWin10 = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 78)
    Size = [System.Drawing.Size]::new(240, 34)
    Text = "Windows 10"
}
$toolTip.SetToolTip($DlWin10, $strings.download.windows.Replace("%v%", "10"))
Set-ButtonTheme $DlWin10

$DlWin11 = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 118)
    Size = [System.Drawing.Size]::new(240, 34)
    Text = "Windows 11"
}
$toolTip.SetToolTip($DlWin11, $strings.download.windows.Replace("%v%", "11"))
Set-ButtonTheme $DlWin11

$DlWin10Ltsc = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 158)
    Size = [System.Drawing.Size]::new(240, 34)
    Text = "Win 10 LTSC 2021 ($($strings.download.recommended))"
}
Set-ButtonTheme $DlWin10Ltsc

$DlWin11Ltsc = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 198)
    Size = [System.Drawing.Size]::new(240, 34)
    Text = "Windows 11 LTSC 2024"
}
Set-ButtonTheme $DlWin11Ltsc

$DlServer2025 = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 78)
    Size = [System.Drawing.Size]::new(180, 34)
    Text = "Windows Server 2025"
}
Set-ButtonTheme $DlServer2025

$DlServer2022 = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 118)
    Size = [System.Drawing.Size]::new(180, 34)
    Text = "Windows Server 2022"
}
Set-ButtonTheme $DlServer2022

$DlServer2019 = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 158)
    Size = [System.Drawing.Size]::new(180, 34)
    Text = "Windows Server 2019"
}
Set-ButtonTheme $DlServer2019

$DlWin81 = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 78)
    Size = [System.Drawing.Size]::new(220, 34)
    Text = "Windows 8.1"
}
Set-ButtonTheme $DlWin81

$DlOffice2024Installer = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(190, 74)
    Size = [System.Drawing.Size]::new(64, 30)
    Text = "Online"
}
$toolTip.SetToolTip($DlOffice2024Installer, $strings.download.oonline + $strings.download.installer)
Set-ButtonTheme $DlOffice2024Installer

$DlOffice2024ISO = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(262, 74)
    Size = [System.Drawing.Size]::new(56, 30)
    Text = "ISO"
}
$toolTip.SetToolTip($DlOffice2024ISO, $strings.download.oiso)
Set-ButtonTheme $DlOffice2024ISO

$DlOffice2021Installer = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(190, 110)
    Size = [System.Drawing.Size]::new(64, 30)
    Text = "Online"
}
$toolTip.SetToolTip($DlOffice2021Installer, $strings.download.oonline + $strings.download.installer)
Set-ButtonTheme $DlOffice2021Installer

$DlOffice2021ISO = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(262, 110)
    Size = [System.Drawing.Size]::new(56, 30)
    Text = "ISO"
}
$toolTip.SetToolTip($DlOffice2021ISO, $strings.download.oiso)
Set-ButtonTheme $DlOffice2021ISO

$DlOffice2019Installer = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(190, 146)
    Size = [System.Drawing.Size]::new(64, 30)
    Text = "Online"
}
$toolTip.SetToolTip($DlOffice2019Installer, $strings.download.oonline + $strings.download.installer)
Set-ButtonTheme $DlOffice2019Installer

$DlOffice2019ISO = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(262, 146)
    Size = [System.Drawing.Size]::new(56, 30)
    Text = "ISO"
}
$toolTip.SetToolTip($DlOffice2019ISO, $strings.download.oiso)
Set-ButtonTheme $DlOffice2019ISO

$DlOffice2016Installer = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(190, 182)
    Size = [System.Drawing.Size]::new(64, 30)
    Text = "Online"
}
$toolTip.SetToolTip($DlOffice2016Installer, $strings.download.oonline + $strings.download.installer)
Set-ButtonTheme $DlOffice2016Installer

$DlOffice2016ISO = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(262, 182)
    Size = [System.Drawing.Size]::new(56, 30)
    Text = "ISO"
}
$toolTip.SetToolTip($DlOffice2016ISO, $strings.download.oiso)
Set-ButtonTheme $DlOffice2016ISO

$DlOffice2013Installer = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(190, 218)
    Size = [System.Drawing.Size]::new(64, 30)
    Text = "Online"
}
$toolTip.SetToolTip($DlOffice2013Installer, $strings.download.outdated + $strings.download.oonline)
Set-ButtonTheme $DlOffice2013Installer

$DlOffice2013ISO = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(262, 218)
    Size = [System.Drawing.Size]::new(56, 30)
    Text = "ISO"
}
$toolTip.SetToolTip($DlOffice2013ISO, $strings.download.outdated + $strings.download.oiso)
Set-ButtonTheme $DlOffice2013ISO

$l24 = New-Object System.Windows.Forms.Label -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 81)
    ForeColor = $theme.text
    BackColor = $theme.surface
    Text = "Office 2024:"
}

$l21 = New-Object System.Windows.Forms.Label -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 117)
    ForeColor = $theme.text
    BackColor = $theme.surface
    Text = "Office 2021:"
}

$l19 = New-Object System.Windows.Forms.Label -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 153)
    ForeColor = $theme.text
    BackColor = $theme.surface
    Text = "Office 2019:"
}

$l16 = New-Object System.Windows.Forms.Label -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 189)
    ForeColor = $theme.text
    BackColor = $theme.surface
    Text = "Office 2016:"
}

$l13 = New-Object System.Windows.Forms.Label -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 225)
    ForeColor = $theme.text
    BackColor = $theme.surface
    Text = "Office 2013:"
}

$rufus = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 78)
    Size = [System.Drawing.Size]::new(470, 34)
    Text = "Rufus - " + $strings.download.rufus
}
Set-ButtonTheme $rufus

$DlVisio2024 = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 118)
    Size = [System.Drawing.Size]::new(220, 34)
    Text = "Visio 2024 ISO"
}
$toolTip.SetToolTip($DlVisio2024, $strings.download.oiso)
Set-ButtonTheme $DlVisio2024

$DlProject2024 = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(268, 118)
    Size = [System.Drawing.Size]::new(220, 34)
    Text = "Project 2024 ISO"
}
$toolTip.SetToolTip($DlProject2024, $strings.download.oiso)
Set-ButtonTheme $DlProject2024

@($DlWin10, $DlWin11, $DlWin10Ltsc, $DlWin11Ltsc) | ForEach-Object { $dlWindowsPanel.Controls.Add($_) }
@($DlServer2025, $DlServer2022, $DlServer2019) | ForEach-Object { $dlServerPanel.Controls.Add($_) }
@($DlOffice2024Installer, $DlOffice2024ISO, $DlOffice2021Installer, $DlOffice2021ISO, $DlOffice2019Installer, $DlOffice2019ISO, $DlOffice2016Installer, $DlOffice2016ISO, $DlOffice2013Installer, $DlOffice2013ISO, $l24, $l21, $l19, $l16, $l13) | ForEach-Object { $dlOfficePanel.Controls.Add($_) }
@($DlWin81, $rufus, $DlVisio2024, $DlProject2024) | ForEach-Object { $dlExtrasPanel.Controls.Add($_) }
$dlGuideText = New-BodyLabel 18 74 300 26 'Need a direct image? Choose ISO. Installing over an existing setup? Start with Online.' 8.5 $theme.textMuted
$dlGuideText.BackColor = $theme.surface
$dlGuidePanel.Controls.Add($dlGuideText)

@($dlWindowsPanel, $dlServerPanel, $dlOfficePanel, $dlExtrasPanel, $dlGuidePanel) | ForEach-Object { $DlTab.Controls.Add($_) }

function bypass_office_geoblock {
    $result = $mb.Invoke($hidden, $strings.download.bypass, $app, "YesNo", "Information")
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" -Name "CountryCode" -PropertyType String -Value "std::wstring|US" -Force
        $mb.Invoke($strings.download.bypass_done, $app, "OK", "Information")
    }
}
function Show-LinkFailDialog {
    $form = New-Object System.Windows.Forms.Form -Property @{
        Text = $app
        StartPosition = "CenterScreen"
        FormBorderStyle = "FixedDialog"
        MaximizeBox = $false
        MinimizeBox = $false
        TopMost = $true
        AutoSize = $true
        AutoSizeMode = "GrowAndShrink"
        BackColor = $theme.formBack
        Font = [System.Drawing.Font]::new('Segoe UI', 9)
    }

    $tablePanel = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
        AutoSize = $true
        AutoSizeMode = "GrowAndShrink"
        ColumnCount = 1
        RowCount = 2
        Padding = New-Object System.Windows.Forms.Padding(10)
        BackColor = $theme.formBack
    }

    $label = New-Object System.Windows.Forms.Label -Property @{
        Text = $strings.download.linkfail
        AutoSize = $true
        Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
        ForeColor = $theme.text
        BackColor = $theme.formBack
    }
    $tablePanel.Controls.Add($label, 0, 0)

    $btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{
        AutoSize = $true
        AutoSizeMode = "GrowAndShrink"
        WrapContents = $false
        BackColor = $theme.formBack
    }

    $btnRetry = New-Object System.Windows.Forms.Button -Property @{
        Text = $strings.download.btn_retry
        AutoSize = $true
        DialogResult = [System.Windows.Forms.DialogResult]::Retry
    }

    $btnReupload = New-Object System.Windows.Forms.Button -Property @{
        Text = $strings.download.btn_reupload
        AutoSize = $true
    }
    $btnReupload.Add_Click({
        Start-Process "https://malw.link/apps/malwtool"
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Ignore
        $form.Close()
    })

    $btnCancel = New-Object System.Windows.Forms.Button -Property @{
        Text = $strings.download.btn_cancel
        AutoSize = $true
        DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    }

    Set-ButtonTheme $btnRetry
    Set-ButtonTheme $btnReupload
    Set-ButtonTheme $btnCancel

    $btnPanel.Controls.AddRange(@($btnRetry, $btnReupload, $btnCancel))
    $tablePanel.Controls.Add($btnPanel, 0, 1)
    $form.Controls.Add($tablePanel)
    $form.CancelButton = $btnCancel
    return $form.ShowDialog()
}

$DlWin10.Add_Click({
    try {
        $products = Invoke-RestMethod "https://raw.githubusercontent.com/ImMALWARE/$app/main/windl.json"
        $url = 'https://api.gravesoft.dev/msdl/proxy?product_id={0}&sku_id={1}' -f $products.'10'[0], $products.'10'[1 + ($PSUICulture -ne "ru-RU")]
        $response = Invoke-RestMethod $url
        Start-Process $response.ProductDownloadOptions[$response.ProductDownloadOptions.Count - 1].Uri
    } catch {
        $result = Show-LinkFailDialog
        if ($result -eq [System.Windows.Forms.DialogResult]::Retry) {
            $DlWin10.PerformClick()
        }
    }
})

$DlWin11.Add_Click({
    try {
        $products = Invoke-RestMethod "https://raw.githubusercontent.com/ImMALWARE/$app/main/windl.json"
        $url = 'https://api.gravesoft.dev/msdl/proxy?product_id={0}&sku_id={1}' -f $products.'11'[0], $products.'11'[1 + ($PSUICulture -ne "ru-RU")]
        Start-Process (Invoke-RestMethod $url).ProductDownloadOptions[0].Uri
    } catch {
        $result = Show-LinkFailDialog
        if ($result -eq [System.Windows.Forms.DialogResult]::Retry) {
            $DlWin11.PerformClick()
        }
    }
})

$DlWin10Ltsc.Add_Click({
    Start-Process $strings.links.w10ltsc
})

$DlWin11Ltsc.Add_Click({
    Start-Process $strings.links.w11ltsc
})

$DlServer2025.Add_Click({
    Start-Process $strings.links.ws2025
})

$DlServer2022.Add_Click({
    Start-Process $strings.links.ws2022
})

$DlServer2019.Add_Click({
    Start-Process $strings.links.ws2019
})

$DlWin81.Add_Click({
    Start-Process $strings.links.w81
})

$DlOffice2024Installer.Add_Click({
    Start-Process $strings.links.o2024o
    bypass_office_geoblock
})

$DlOffice2024ISO.Add_Click({
    Start-Process $strings.links.o2024i
})

$DlOffice2021Installer.Add_Click({
    Start-Process $strings.links.o2021o
    bypass_office_geoblock
})

$DlOffice2021ISO.Add_Click({
    Start-Process $strings.links.o2021i
})

$DlOffice2019Installer.Add_Click({
    Start-Process $strings.links.o2019o
    bypass_office_geoblock
})

$DlOffice2019ISO.Add_Click({
    Start-Process $strings.links.o2019i
})

$DlOffice2016Installer.Add_Click({
    Start-Process $strings.links.o2016o
    bypass_office_geoblock
})

$DlOffice2016ISO.Add_Click({
    Start-Process $strings.links.o2016i
})

$DlOffice2013Installer.Add_Click({
    Start-Process $strings.links.o2013o
})

$DlOffice2013ISO.Add_Click({
    Start-Process $strings.links.o2013i
})

$rufus.Add_Click({
    New-Item -Path "$env:temp\$app" -ItemType Directory -ErrorAction SilentlyContinue > $null
    $wc.DownloadFile('https://github.com/pbatard/rufus/releases/download/v4.13/rufus-4.13.exe', "$env:temp\$app\rufus.exe")
    Set-Location $env:SystemRoot\System32
    ./cmd.exe /c start "" "$env:temp\$app\rufus.exe"
})

$DlVisio2024.Add_Click({
    Start-Process $strings.links.vis2024
})

$DlProject2024.Add_Click({
    Start-Process $strings.links.proj2024
})

#########

$fnQuickPanel = New-SectionPanel 20 20 430 190 'Quick actions' 'Everyday tools you will likely want first.'
$fnDriversPanel = New-SectionPanel 470 20 430 190 'Drivers' 'Backup and restore helpers for reinstall workflows.'
$fnMaintenancePanel = New-SectionPanel 20 230 880 235 'System and maintenance' 'Cleanup, tweaks, and network workarounds stay unchanged, now grouped more cleanly.'

$winwifipassman = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 78)
    Size = [System.Drawing.Size]::new(185, 34)
    Text = $strings.functions.wifi
}
$tooltip.SetToolTip($winwifipassman, $strings.functions.wifi_desc)
Set-ButtonTheme $winwifipassman

$explorerext = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(219, 78)
    Size = [System.Drawing.Size]::new(185, 34)
    Text = $strings.functions.fileext
}
Set-ButtonTheme $explorerext

$winget = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 122)
    Size = [System.Drawing.Size]::new(185, 34)
    Text = $strings.functions.install + "winget"
}
Set-ButtonTheme $winget

$store = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(219, 122)
    Size = [System.Drawing.Size]::new(185, 34)
    Text = $strings.functions.install + "Microsoft Store"
}
$tooltip.SetToolTip($store, $strings.functions.store_desc)
Set-ButtonTheme $store

$driversbackup = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 90)
    Size = [System.Drawing.Size]::new(185, 34)
    Text = $strings.functions.drivers
}
$tooltip.SetToolTip($driversbackup, $strings.functions.drivers_desc)
Set-ButtonTheme $driversbackup

$driversrestore = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(219, 90)
    Size = [System.Drawing.Size]::new(185, 34)
    Text = $strings.functions.drivers_restore
}
$tooltip.SetToolTip($driversrestore, $strings.functions.drivers_desc)
Set-ButtonTheme $driversrestore

$edgeuninstall = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 84)
    Size = [System.Drawing.Size]::new(270, 34)
    Text = $strings.functions.edge
}
Set-ButtonTheme $edgeuninstall

$delspyfiles = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(304, 84)
    Size = [System.Drawing.Size]::new(270, 34)
    Text = $strings.functions.compattel
}
$tooltip.SetToolTip($delspyfiles, $strings.functions.compattel_desc)
Set-ButtonTheme $delspyfiles

$spicetify = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(590, 84)
    Size = [System.Drawing.Size]::new(270, 34)
    Text = $strings.functions.install + "Spicetify"
}
$tooltip.SetToolTip($spicetify, $strings.functions.spicetify_desc)
Set-ButtonTheme $spicetify

$edit_hosts = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 128)
    Size = [System.Drawing.Size]::new(270, 34)
    Text = $strings.functions.hosts
}
$tooltip.SetToolTip($edit_hosts, $strings.functions.hosts_desc)
Set-ButtonTheme $edit_hosts

@($winwifipassman, $winget, $store) | ForEach-Object { $fnQuickPanel.Controls.Add($_) }
@($driversbackup, $driversrestore) | ForEach-Object { $fnDriversPanel.Controls.Add($_) }
@($edgeuninstall, $spicetify, $edit_hosts) | ForEach-Object { $fnMaintenancePanel.Controls.Add($_) }
@($fnQuickPanel, $fnDriversPanel, $fnMaintenancePanel) | ForEach-Object { $FunctionsTab.Controls.Add($_) }

$val = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideFileExt" -ErrorAction SilentlyContinue
if (!$val -or $val.HideFileExt -ne 0) {
    $fnQuickPanel.Controls.Add($explorerext)
}

if ((Test-Path "$env:SystemRoot\System32\CompatTelRunner.exe") -or (Test-Path "$env:SystemRoot\System32\wsqmcons.exe")) {
    $fnMaintenancePanel.Controls.Add($delspyfiles)
}

$winwifipassman.Add_Click({
    Start-Process powershell -ArgumentList "irm https://raw.githubusercontent.com/ImMALWARE/WinWiFiPassMan/main/WinWiFiPassMan.ps1 | Invoke-Expression" -Verb RunAs
})

$explorerext.Add_Click({
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideFileExt" -Value 0 -Force
    $explorerext.Visible = $false
})

$winget.Add_Click({
    Start-Process powershell -ArgumentList @"
    `$progressPreference = 'silentlyContinue'
    Install-PackageProvider -Name NuGet -Force
    Install-Module -Name Microsoft.WinGet.Client -Force -Repository PSGallery
    Repair-WinGetPackageManager
    Pause
"@ -Verb RunAs
})

$store.Add_Click({
    Start-Process wsreset -ArgumentList "-i" -Verb RunAs -Wait
    $mb.Invoke($strings.functions.store_installing, $app, "OK", "Information")
})

$driversbackup.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
        Description = $strings.functions.drivers_select
        RootFolder = [System.Environment+SpecialFolder]::MyComputer
    }
    if ($dialog.ShowDialog() -eq 'OK') {
        Start-Process powershell -ArgumentList "`$host.UI.RawUI.WindowTitle = '$app - $($strings.functions.drivers)'; pnputil /export-driver * '$($dialog.SelectedPath)'; pause" -Verb RunAs
    }
})

$driversrestore.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
        Description = $strings.functions.drivers_select
        RootFolder = [System.Environment+SpecialFolder]::MyComputer
    }
    if ($dialog.ShowDialog() -eq 'OK') {
        Start-Process powershell -ArgumentList "`$host.UI.RawUI.WindowTitle = '$app - $($strings.functions.drivers_restore)'; pnputil /add-driver '$($dialog.SelectedPath)\*' /subdirs /install; pause" -Verb RunAs
    }
})

$edgeuninstall.Add_Click({
    New-Item -Path "$env:temp\$app" -ItemType Directory -ErrorAction SilentlyContinue
    $wc.DownloadFile('https://raw.githubusercontent.com/AveYo/fox/eab2269a598ad9e8120cf1d598d48384071ff476/Edge_Removal.bat', "$env:temp\$app\Edge_Removal.bat")
    Set-Location $env:SystemRoot\System32
    ./cmd.exe /c start "$env:temp\$app\Edge_Removal.bat"
})

$delspyfiles.Add_Click({
    Start-Process powershell -ArgumentList @"
    `$host.UI.RawUI.WindowTitle = '$app - $($strings.functions.compattel_title)'
    cd '$env:SystemRoot\System32'
    .\takeown.exe /f CompatTelRunner.exe /a
    .\takeown.exe /f wsqmcons.exe /a
    .\icacls.exe CompatTelRunner.exe /grant Administrators:F
    .\icacls.exe wsqmcons.exe /grant Administrators:F
    Stop-Process -Name compattelrunner -Force -ErrorAction SilentlyContinue
    Stop-Process -Name wsqmcons -Force -ErrorAction SilentlyContinue
    Remove-Item CompatTelRunner.exe -Force
    Remove-Item wsqmcons.exe -Force
    pause
"@ -Verb RunAs
})

$spicetify.Add_Click({
    Start-Process powershell -ArgumentList "irm https://raw.githubusercontent.com/spicetify/cli/main/install.ps1 | iex; pause"
})

$edit_hosts.Add_Click({
    Start-Process powershell -ArgumentList @"
        `$host.UI.RawUI.WindowTitle = '$app - $($strings.functions.hosts_title)'
        `$current_hosts = Get-Content -Path 'C:\Windows\System32\drivers\etc\hosts' -Raw
        `$new_hosts = irm https://raw.githubusercontent.com/ImMALWARE/dns.malw.link/master/hosts
        `$start_marker = '### dns.malw.link: hosts file'
        `$end_marker = '### dns.malw.link: end hosts file'

        `$pattern = '(?s)' + [regex]::Escape(`$start_marker) + '.*?' + [regex]::Escape(`$end_marker)

        if (`$current_hosts -match `$pattern) {
            `$updated_hosts = [regex]::Replace(`$current_hosts, `$pattern, `$new_hosts)
        } else {
            `$updated_hosts = `$current_hosts + [Environment]::NewLine + `$new_hosts
        }

        echo `$updated_hosts
        `$updated_hosts | Out-File -Encoding utf8 -FilePath 'C:\Windows\System32\drivers\etc\hosts' -Force
        pause
"@ -Verb RunAs
})

#######

$problemOfficePanel = New-SectionPanel 20 20 430 180 'Office' 'Licenses, uninstall flow, and Office16 folder recovery.'
$problemWindowsPanel = New-SectionPanel 470 20 430 180 'Windows' 'Reset KMS and run the standard system integrity checks.'
$problemSupportPanel = New-SectionPanel 20 220 880 180 'Need help' 'If none of these flows fit, jump straight to the support links and report details.'

$clear_office16 = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 82)
    Size = [System.Drawing.Size]::new(185, 34)
    Text = $strings.problems.cl_o16
}
$tooltip.SetToolTip($clear_office16, $strings.problems.cl_o16_desc)
Set-ButtonTheme $clear_office16

$office_uninstall = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(221, 82)
    Size = [System.Drawing.Size]::new(185, 34)
    Text = $strings.problems.ouninstall
}
Set-ButtonTheme $office_uninstall

$office16_folder = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 126)
    Size = [System.Drawing.Size]::new(388, 34)
    Text = $strings.problems.o16_folder
}
Set-ButtonTheme $office16_folder

$clear_winkms = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 82)
    Size = [System.Drawing.Size]::new(185, 34)
    Text = $strings.problems.cl_win
}
Set-ButtonTheme $clear_winkms

$sfc_scannow = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(221, 82)
    Size = [System.Drawing.Size]::new(185, 34)
    Text = $strings.problems.sfc
}
$tooltip.SetToolTip($sfc_scannow, "sfc /scannow, DISM /Online /Cleanup-Image /RestoreHealth, chkdsk /b /x")
Set-ButtonTheme $sfc_scannow

$otherproblem = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 110)
    Size = [System.Drawing.Size]::new(180, 34)
    Text = $strings.problems.other
}
$tooltip.SetToolTip($otherproblem, $strings.problems.other_desc)
Set-ButtonTheme -Button $otherproblem -Primary

$problemSupportText = New-BodyLabel 18 78 620 24 $strings.problems.other_desc 9 $theme.textMuted
$problemSupportText.BackColor = $theme.surface

@($clear_office16, $office_uninstall, $office16_folder) | ForEach-Object { $problemOfficePanel.Controls.Add($_) }
@($clear_winkms, $sfc_scannow) | ForEach-Object { $problemWindowsPanel.Controls.Add($_) }
@($problemSupportText, $otherproblem) | ForEach-Object { $problemSupportPanel.Controls.Add($_) }
@($problemOfficePanel, $problemWindowsPanel, $problemSupportPanel) | ForEach-Object { $ProblemsTab.Controls.Add($_) }

$clear_office16.Add_Click({
    if (test-path "$env:ProgramFiles\Microsoft Office\Office16\ospp.vbs"){
        $path = "$env:ProgramFiles\Microsoft Office\Office16\"
    }
    elseIf (test-path "${env:ProgramFiles(x86)}\Microsoft Office\Office16\ospp.vbs") {
        $path = "${env:ProgramFiles(x86)}\Microsoft Office\Office16\"
    }
    else {
        $mb.Invoke($strings.problems.cl_o16_err, $app, "OK", "Error")
        exit
    }
    Start-Process powershell -ArgumentList @"
    `$host.ui.RawUI.WindowTitle = '$app - $($strings.problems.cl_o16_title)'
    cd '$path'
    while(`$true){
        `$license = (cscript ospp.vbs /dstatus) | Out-String
        `$match = `$license | Select-String -Pattern 'Last 5 characters of installed product key: (\w{5})'
        if (`$match) {
            `$productKey = `$match.Matches.Groups[1].Value
            cscript ospp.vbs /unpkey:`$productKey
        } else {
            exit
        }
    }
"@ -Verb RunAs -Wait
    $mb.Invoke("OK", $app, "OK", "Information")
})

$office_uninstall.Add_Click({
    if (!(Get-AppxPackage -Name "Microsoft.GetHelp")) {
        if (Get-AppxPackage -Name "Microsoft.WindowsStore") {
            Start-Process "ms-windows-store://pdp/?ProductId=9PKDZBMV1H3T"
            $mb.Invoke($hidden, $strings.problems.ouninstall_nogethelp, $app, "OK", "Warning")
        } else {
            $mb.Invoke($hidden, $strings.problems.nostore, $app, "OK", "Warning")
        }
        return
    }
    Start-Process "ms-contact-support://smc-to-emerald/SARA-UninstallOffice"
})

$office16_folder.Add_Click({
    if (Test-Path "$env:ProgramFiles\Microsoft Office") {
        if (!(Test-Path "$env:ProgramFiles\Microsoft Office\Office16")) {
            New-Item -Path "$env:temp\$app" -ItemType Directory -ErrorAction SilentlyContinue> $null
            $wc.DownloadFile('https://archive.org/download/office16_202502/Office16.zip', "$env:temp\$app\Office16.zip")
            Start-Process powershell -ArgumentList "`$host.UI.RawUI.WindowTitle = '$app - $($strings.problems.o16_title)'; New-Item -Path '$env:ProgramFiles\Microsoft Office\Office16' -ItemType Directory; Set-Location '$env:temp\$app'; Expand-Archive .\Office16.zip -DestinationPath '$env:ProgramFiles\Microsoft Office\Office16'" -Verb RunAs -Wait
            $mb.Invoke("OK", $app, "OK", "Information")
        } else {
            $mb.Invoke($strings.problems.o16_exists, $app, "OK", "Information")
        }
    } else {
        $mb.Invoke($strings.act.notfound.Replace("%p%", "Office"), $app, "OK", "Error")
    }
})

$clear_winkms.Add_Click({
    Start-Process powershell -ArgumentList "`$host.UI.RawUI.WindowTitle = '$app - $($strings.problems.cl_win_title)'; Set-Location $env:SystemRoot\System32; .\slmgr /upk; .\slmgr /cpky" -Verb RunAs
})

$sfc_scannow.Add_Click({
    Start-Process powershell -ArgumentList "`$host.UI.RawUI.WindowTitle = '$app - $($strings.problems.sfc_title)'; Set-Location $env:SystemRoot\System32; .\sfc /scannow; .\Dism /Online /Cleanup-Image /RestoreHealth; Write-Host '$($strings.problems.sfc_note)' .\chkdsk ${(Get-WmiObject Win32_OperatingSystem).SystemDrive} /b /x; .\shutdown /r /t 60 /c '$($strings.problems.sfc_shutdown_desc)'; pause" -Verb RunAs
})

$otherproblem.Add_Click({
    $tabs.SelectedTab = $InfoTab
})

######

$infoOverviewPanel = New-SectionPanel 20 20 400 180 "$app 2.0" 'Core project links collected in one place.'
$infoSupportPanel = New-SectionPanel 440 20 460 180 'Support' 'If something goes sideways, jump straight to chat or open an issue here.'

$malwtool = New-Object System.Windows.Forms.Label -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(18, 82)
    Font = [System.Drawing.Font]::new('Segoe UI', 9)
    ForeColor = $theme.textMuted
    BackColor = $theme.surface
    Text = "$app 2.0"
}

$wiki = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 116)
    Size = [System.Drawing.Size]::new(110, 34)
    Text = $strings.info.wiki
}
Set-ButtonTheme $wiki

$lolzteam = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(144, 116)
    Size = [System.Drawing.Size]::new(110, 34)
    Text = $strings.info.lolz
}
Set-ButtonTheme $lolzteam

$github = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(270, 116)
    Size = [System.Drawing.Size]::new(110, 34)
    Text = "GitHub"
}
Set-ButtonTheme $github

$questions = New-Object System.Windows.Forms.Label -Property @{
    Location = [System.Drawing.Point]::new(18, 82)
    Size = [System.Drawing.Size]::new(400, 24)
    ForeColor = $theme.textMuted
    BackColor = $theme.surface
    Text = $strings.info.questions + $strings.problems.other_desc
}

$telegram = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(18, 116)
    Size = [System.Drawing.Size]::new(130, 34)
    Text = $strings.info.telegram
}
Set-ButtonTheme -Button $telegram -Primary

$lolzteam2 = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(164, 116)
    Size = [System.Drawing.Size]::new(140, 34)
    Text = $strings.info.lolz_write
}
Set-ButtonTheme $lolzteam2

$github2 = New-Object System.Windows.Forms.Button -Property @{
    Location = [System.Drawing.Point]::new(320, 116)
    Size = [System.Drawing.Size]::new(120, 34)
    Text = $strings.info.github
}
Set-ButtonTheme $github2

@($malwtool, $wiki, $lolzteam, $github) | ForEach-Object { $infoOverviewPanel.Controls.Add($_) }
@($questions, $telegram, $lolzteam2, $github2) | ForEach-Object { $infoSupportPanel.Controls.Add($_) }
@($infoOverviewPanel, $infoSupportPanel) | ForEach-Object { $InfoTab.Controls.Add($_) }

$wiki.Add_Click({
    Start-Process "https://wiki.malw.link/apps/malwtool"
})

$lolzteam.Add_Click({
    Start-Process "https://lolz.live/threads/4997821"
})

$github.Add_Click({
    Start-Process "https://github.com/ImMALWARE/$app"
})

$telegram.Add_Click({
    Start-Process "https://t.me/immalware_chat"
})

$lolzteam2.Add_Click({
    $lolzteam.PerformClick()
})

$github2.Add_Click({
    Start-Process "https://github.com/ImMALWARE/$app/issues/new"
})

$form.Controls.Add($headerPanel)
$form.Controls.Add($tabs)
$form.ShowDialog()
