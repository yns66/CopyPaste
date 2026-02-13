Add-Type -AssemblyName System.Globalization
$ti = (Get-Culture).TextInfo

function Format-Hitap($text) {
    return $ti.ToTitleCase($text.ToLower())
}


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =====================
# AYAR
# =====================
$DefaultPath = "C:\Users\YnS\Desktop\New folder"

# =====================
# FORM
# =====================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Hazır Metinler"
$form.Size = New-Object System.Drawing.Size(600,420)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true

# =====================
# ÜST KLASÖR LABEL
# =====================
$lblFolder = New-Object System.Windows.Forms.Label
$lblFolder.Location = New-Object System.Drawing.Point(10,12)
$lblFolder.Size = New-Object System.Drawing.Size(580,20)
$lblFolder.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$lblFolder.Text = "Klasör seçilmedi"
$form.Controls.Add($lblFolder)

# =====================
# KLASÖR SEÇ BUTONU
# =====================
$btnSelectFolder = New-Object System.Windows.Forms.Button
$btnSelectFolder.Text = "Klasör Seç"
$btnSelectFolder.Location = New-Object System.Drawing.Point(10,35)
$btnSelectFolder.Size = New-Object System.Drawing.Size(100,28)
$form.Controls.Add($btnSelectFolder)

# =====================
# MANUEL DİZİN GİRİŞİ
# =====================
$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Location = New-Object System.Drawing.Point(120,36)
$txtPath.Size = New-Object System.Drawing.Size(460,26)
$txtPath.Font = New-Object System.Drawing.Font("Segoe UI",9)
$form.Controls.Add($txtPath)

# =====================
# HİTAP ALANI
# =====================
$lblHitap = New-Object System.Windows.Forms.Label
$lblHitap.Text = "Hitap (Ahmet Bey / Ayşe Hanım):"
$lblHitap.Location = New-Object System.Drawing.Point(10,70)
$lblHitap.Size = New-Object System.Drawing.Size(220,20)
$form.Controls.Add($lblHitap)

$txtHitap = New-Object System.Windows.Forms.TextBox
$txtHitap.Location = New-Object System.Drawing.Point(230,68)
$txtHitap.Size = New-Object System.Drawing.Size(200,24)
$form.Controls.Add($txtHitap)

$txtHitap.Add_Leave({
    $txtHitap.Text = Format-Hitap $txtHitap.Text
})

$txtHitap.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {

        $txtHitap.Text = Format-Hitap $txtHitap.Text

        if ($listBox.SelectedItem -ne $null) {
            $key = $listBox.SelectedItem
            if ($fileMap.ContainsKey($key)) {

                $content = Get-Content $fileMap[$key] -Raw -Encoding UTF8

                if ($txtHitap.Text.Trim() -ne "") {
                    $prefix = "Merhaba $($txtHitap.Text)`r`n`r`n"
                    $textBox.Text = $prefix + $content
                }
                else {
                    $textBox.Text = $content
                }
            }
        }
    }
})


# =====================
# TXT LİSTESİ
# =====================
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,100)
$listBox.Size = New-Object System.Drawing.Size(250,250)
$form.Controls.Add($listBox)

# =====================
# SAĞ İÇERİK
# =====================
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(270,100)
$textBox.Size = New-Object System.Drawing.Size(310,210)
$textBox.Multiline = $true
$textBox.ScrollBars = "Both"
$textBox.ReadOnly = $true
$textBox.Font = New-Object System.Drawing.Font("Consolas",9)
$form.Controls.Add($textBox)

# =====================
# KOPYALA
# =====================
$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = "Kopyala"
$btnCopy.Location = New-Object System.Drawing.Point(270,320)
$btnCopy.Size = New-Object System.Drawing.Size(120,28)
$form.Controls.Add($btnCopy)

# =====================
# DOSYA MAP
# =====================
$fileMap = @{}

function Load-TxtFiles($basePath) {

    if (-not (Test-Path $basePath)) { return }

    $lblFolder.Text = "Klasör: " + (Split-Path $basePath -Leaf)
    $listBox.Items.Clear()
    $fileMap.Clear()
    $textBox.Clear()

    $files = Get-ChildItem $basePath -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in ".txt",".lnk" }

    foreach ($file in $files) {

        if ($file.Extension -eq ".txt") {
            $listBox.Items.Add($file.Name)
            $fileMap[$file.Name] = $file.FullName
        }
        elseif ($file.Extension -eq ".lnk") {
            $wsh = New-Object -ComObject WScript.Shell
            $sc = $wsh.CreateShortcut($file.FullName)
            if ($sc.TargetPath -and $sc.TargetPath.EndsWith(".txt")) {
                $listBox.Items.Add($file.Name)
                $fileMap[$file.Name] = $sc.TargetPath
            }
        }
    }
}

# =====================
# LİSTE SEÇİMİ (HİTAP EKLİ)
# =====================
$listBox.Add_SelectedIndexChanged({
    $key = $listBox.SelectedItem
    if ($fileMap.ContainsKey($key)) {
        try {
            $content = Get-Content $fileMap[$key] -Raw -Encoding UTF8

            if ($txtHitap.Text.Trim() -ne "") {
                $prefix = "Merhaba $($txtHitap.Text)`r`n`r`n"
                $textBox.Text = $prefix + $content
            }
            else {
                $textBox.Text = $content
            }
        }
        catch {
            $textBox.Text = "Dosya okunamadı!"
        }
    }
})

# =====================
# KOPYALA
# =====================
$btnCopy.Add_Click({
    if ($textBox.Text) {
        [System.Windows.Forms.Clipboard]::SetText($textBox.Text)
        [System.Windows.Forms.MessageBox]::Show("Metin panoya kopyalandı.")
    }
})

# =====================
# ENTER İLE YÜKLE
# =====================
$txtPath.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        if (Test-Path $txtPath.Text) {
            Load-TxtFiles $txtPath.Text
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Geçerli bir dizin giriniz.")
        }
    }
})

# =====================
# BUTON İLE SEÇ
# =====================
$btnSelectFolder.Add_Click({
    $fd = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($fd.ShowDialog() -eq "OK") {
        $txtPath.Text = $fd.SelectedPath
        Load-TxtFiles $fd.SelectedPath
    }
})

# =====================
# FORM AÇILIŞINDA DEFAULT PATH
# =====================
$form.Add_Shown({
    if (Test-Path $DefaultPath) {
        $txtPath.Text = $DefaultPath
        Load-TxtFiles $DefaultPath
    }
})

[void]$form.ShowDialog()
