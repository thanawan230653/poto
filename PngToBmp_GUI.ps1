#requires -version 5.1
# PNG -> BMP Converter (GUI) | Save as: PngToBmp_GUI.ps1
# Run: Right-click -> Run with PowerShell (or: powershell -ExecutionPolicy Bypass -File .\PngToBmp_GUI.ps1)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Helpers ---
function Show-Error([string]$msg){
    [System.Windows.Forms.MessageBox]::Show($msg, "Error", "OK", "Error") | Out-Null
}
function Show-Info([string]$msg){
    [System.Windows.Forms.MessageBox]::Show($msg, "Info", "OK", "Information") | Out-Null
}
function Convert-PngToBmp([string[]]$paths, [string]$outDir, [System.Windows.Forms.ListBox]$log){
    if (-not $paths -or $paths.Count -eq 0) { Show-Error "Please select at least 1 PNG file."; return }
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

    $ok = 0; $fail = 0
    foreach($p in $paths){
        try{
            if (-not (Test-Path $p)) { throw "File not found: $p" }
            $ext = [IO.Path]::GetExtension($p)
            if ($ext -notin @(".png",".PNG")) { throw "Not a PNG: $p" }

            $name = [IO.Path]::GetFileNameWithoutExtension($p)
            $dst  = Join-Path $outDir ($name + ".bmp")

            # Load and save
            $img = [System.Drawing.Image]::FromFile($p)
            try{
                $img.Save($dst, [System.Drawing.Imaging.ImageFormat]::Bmp)
            } finally {
                $img.Dispose()
            }

            $ok++
            $log.Items.Add("✅ OK:  " + $dst) | Out-Null
        } catch {
            $fail++
            $log.Items.Add("❌ FAIL: " + $p + " | " + $_.Exception.Message) | Out-Null
        }
    }

    $log.Items.Add("—") | Out-Null
    $log.Items.Add(("Done. Success: {0} | Failed: {1}" -f $ok, $fail)) | Out-Null
}

# --- Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "PNG → BMP Converter (PowerShell GUI)"
$form.Width = 820
$form.Height = 560
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.AllowDrop = $true

$title = New-Object System.Windows.Forms.Label
$title.Text = "🖼️  PNG → BMP Converter"
$title.AutoSize = $true
$title.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(18, 14)
$form.Controls.Add($title)

$hint = New-Object System.Windows.Forms.Label
$hint.Text = "Tip: Drag & Drop PNG files onto this window • or click Browse"
$hint.AutoSize = $true
$hint.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$hint.Location = New-Object System.Drawing.Point(20, 52)
$form.Controls.Add($hint)

# Selected files textbox
$tbFiles = New-Object System.Windows.Forms.TextBox
$tbFiles.Location = New-Object System.Drawing.Point(20, 86)
$tbFiles.Width = 640
$tbFiles.ReadOnly = $true
$tbFiles.Text = "No files selected."
$form.Controls.Add($tbFiles)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse PNG(s)…"
$btnBrowse.Location = New-Object System.Drawing.Point(670, 84)
$btnBrowse.Width = 120
$form.Controls.Add($btnBrowse)

# Output folder
$lbOut = New-Object System.Windows.Forms.Label
$lbOut.Text = "Output folder:"
$lbOut.AutoSize = $true
$lbOut.Location = New-Object System.Drawing.Point(20, 125)
$form.Controls.Add($lbOut)

$tbOut = New-Object System.Windows.Forms.TextBox
$tbOut.Location = New-Object System.Drawing.Point(20, 148)
$tbOut.Width = 640
$tbOut.ReadOnly = $true
$tbOut.Text = (Join-Path $env:USERPROFILE "Desktop\BMP_Output")
$form.Controls.Add($tbOut)

$btnOut = New-Object System.Windows.Forms.Button
$btnOut.Text = "Choose…"
$btnOut.Location = New-Object System.Drawing.Point(670, 146)
$btnOut.Width = 120
$form.Controls.Add($btnOut)

# Convert button
$btnConvert = New-Object System.Windows.Forms.Button
$btnConvert.Text = "Convert"
$btnConvert.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnConvert.Location = New-Object System.Drawing.Point(20, 190)
$btnConvert.Width = 120
$btnConvert.Height = 36
$form.Controls.Add($btnConvert)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear"
$btnClear.Location = New-Object System.Drawing.Point(150, 190)
$btnClear.Width = 90
$btnClear.Height = 36
$form.Controls.Add($btnClear)

$btnOpenOut = New-Object System.Windows.Forms.Button
$btnOpenOut.Text = "Open Output"
$btnOpenOut.Location = New-Object System.Drawing.Point(250, 190)
$btnOpenOut.Width = 120
$btnOpenOut.Height = 36
$form.Controls.Add($btnOpenOut)

# Log box
$log = New-Object System.Windows.Forms.ListBox
$log.Location = New-Object System.Drawing.Point(20, 240)
$log.Width = 770
$log.Height = 260
$form.Controls.Add($log)

# State
$script:selectedPaths = @()

# Dialogs
$ofd = New-Object System.Windows.Forms.OpenFileDialog
$ofd.Filter = "PNG files (*.png)|*.png|All files (*.*)|*.*"
$ofd.Multiselect = $true
$ofd.Title = "Select PNG file(s)"

$fbd = New-Object System.Windows.Forms.FolderBrowserDialog
$fbd.Description = "Choose output folder for BMP files"

# Events
$btnBrowse.Add_Click({
    if ($ofd.ShowDialog() -eq "OK") {
        $script:selectedPaths = $ofd.FileNames
        if ($script:selectedPaths.Count -eq 1) {
            $tbFiles.Text = $script:selectedPaths[0]
        } else {
            $tbFiles.Text = ("{0} files selected" -f $script:selectedPaths.Count)
        }
        $log.Items.Add(("📌 Selected: {0}" -f $tbFiles.Text)) | Out-Null
    }
})

$btnOut.Add_Click({
    $fbd.SelectedPath = $tbOut.Text
    if ($fbd.ShowDialog() -eq "OK") {
        $tbOut.Text = $fbd.SelectedPath
        $log.Items.Add(("📂 Output: {0}" -f $tbOut.Text)) | Out-Null
    }
})

$btnConvert.Add_Click({
    $log.Items.Add("🚀 Converting...") | Out-Null
    Convert-PngToBmp -paths $script:selectedPaths -outDir $tbOut.Text -log $log
})

$btnClear.Add_Click({
    $script:selectedPaths = @()
    $tbFiles.Text = "No files selected."
    $log.Items.Clear()
})

$btnOpenOut.Add_Click({
    if (Test-Path $tbOut.Text) {
        Start-Process explorer.exe $tbOut.Text
    } else {
        Show-Error "Output folder not found."
    }
})

# Drag & Drop
$form.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [Windows.Forms.DragDropEffects]::Copy
    }
})

$form.Add_DragDrop({
    $files = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    $pngs = @($files | Where-Object { [IO.Path]::GetExtension($_) -in @(".png",".PNG") })
    if ($pngs.Count -eq 0) {
        Show-Error "Please drop PNG files only."
        return
    }
    $script:selectedPaths = $pngs
    $tbFiles.Text = if ($pngs.Count -eq 1) { $pngs[0] } else { "{0} files selected (drag-drop)" -f $pngs.Count }
    $log.Items.Add(("📥 Dropped: {0}" -f $tbFiles.Text)) | Out-Null
})

# Footer
$footer = New-Object System.Windows.Forms.Label
$footer.Text = "Built with PowerShell + WinForms | PNG → BMP"
$footer.AutoSize = $true
$footer.ForeColor = [System.Drawing.Color]::Gray
$footer.Location = New-Object System.Drawing.Point(20, 508)
$form.Controls.Add($footer)

# Run
[System.Windows.Forms.Application]::Run($form)
