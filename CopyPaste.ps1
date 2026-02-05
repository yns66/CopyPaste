# SESSİZLİK AYARLARI (en üstte)
$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# EnableVisualStyles sessize alındı
$null = [System.Windows.Forms.Application]::EnableVisualStyles()

# CONFIG DOSYA YOLU (Documents + Gizli klasör)
$docPath = [Environment]::GetFolderPath("MyDocuments")
$configDir = Join-Path $docPath "HazirMetinSecici"
$configFile = Join-Path $configDir "config.txt"

# Klasör yoksa oluştur ve gizle
if (-not (Test-Path $configDir)) {
    $null = New-Item -ItemType Directory -Path $configDir
}
$item = Get-Item $configDir
$item.Attributes = $item.Attributes -bor [IO.FileAttributes]::Hidden

# Config dosyası yoksa oluştur ve gizle
if (-not (Test-Path $configFile)) {
    $null = New-Item -ItemType File -Path $configFile -Force
}
$item2 = Get-Item $configFile
$item2.Attributes = $item2.Attributes -bor [IO.FileAttributes]::Hidden

# FORM OLUŞTURMA
$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.SizeGripStyle = "Hide"
$form.Text = "Hazır Metin Seçici"
$form.Size = New-Object System.Drawing.Size(520,460)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true

$form.Add_FormClosing({
    [System.Windows.Forms.Application]::Exit()
})


$pathBox = New-Object System.Windows.Forms.TextBox
$pathBox.Location = New-Object System.Drawing.Point(10,10)
$pathBox.Size = New-Object System.Drawing.Size(300,25)
$pathBox.ReadOnly = $false

$pathBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        if (Test-Path $pathBox.Text) {
            $pathBox.Text | Set-Content $configFile -Encoding UTF8
            Yukle-TxtDosyalari $pathBox.Text
        }
    }
})


$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Location = New-Object System.Drawing.Point(10,40) # textbox'un altı
$pathLabel.Size = New-Object System.Drawing.Size(300,20)
$pathLabel.Text = "GÖRÜNTÜLEMEK İSTEDİĞİNİZ YOLU SEÇİN"
$pathLabel.ForeColor = [System.Drawing.Color]::Gray



$browseBtn = New-Object System.Windows.Forms.Button
$browseBtn.Location = New-Object System.Drawing.Point(400,8)
$browseBtn.Size = New-Object System.Drawing.Size(80,50)
$browseBtn.Text = "Yol Seçin"

$enterBtn = New-Object System.Windows.Forms.Button
$enterBtn.Location = New-Object System.Drawing.Point(320,8) # Değiştir'in solu
$enterBtn.Size = New-Object System.Drawing.Size(80,50)
$enterBtn.Text = "Enter"

$enterBtn.Add_Click({
    if (Test-Path $pathBox.Text) {
        $pathBox.Text | Set-Content $configFile -Encoding UTF8
        Yukle-TxtDosyalari $pathBox.Text
    }
})


$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,70)
$listBox.Size = New-Object System.Drawing.Size(480,100)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,175)
$textBox.Size = New-Object System.Drawing.Size(480,190)
$textBox.Multiline = $true
$textBox.ReadOnly = $true
$textBox.ScrollBars = "Vertical"

$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(10,375)
$button.Size = New-Object System.Drawing.Size(480,40)
$button.Text = "Panoya Kopyala"

$form.Controls.AddRange(@(
    $pathBox,
    $pathLabel,
    $enterBtn,
    $browseBtn,
    $listBox,
    $textBox,
    $button
))



# TXT DOSYALARINI YÜKLEME FONKSİYONU
$txtFiles = @{}
function Yukle-TxtDosyalari($klasor) {
    $listBox.Items.Clear()
    $txtFiles.Clear()

    if (Test-Path $klasor) {
        Get-ChildItem $klasor -Filter *.txt | ForEach-Object {
            $listBox.Items.Add($_.BaseName)
            $txtFiles[$_.BaseName] = $_.FullName
        }
    }
}

# CONFIG DOSYASINDAN KAYITLI YOLU OKU
$savedPath = Get-Content $configFile -ErrorAction SilentlyContinue
if ($savedPath -and (Test-Path $savedPath)) {
    $pathBox.Text = $savedPath
    Yukle-TxtDosyalari $savedPath
}

# GÖZAT BUTONU
$browseBtn.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Metin klasörünü seçin"

    if ($dialog.ShowDialog() -eq "OK") {
        $pathBox.Text = $dialog.SelectedPath
        $dialog.SelectedPath | Set-Content $configFile -Encoding UTF8
        Yukle-TxtDosyalari $dialog.SelectedPath
    }
})

# LİSTBOX SEÇİMİ DEĞİŞİNCE
$listBox.Add_SelectedIndexChanged({
    if ($listBox.SelectedItem) {
        $textBox.Text = Get-Content $txtFiles[$listBox.SelectedItem] -Raw -Encoding UTF8
    }
})

# PANOYA KOPYALA BUTONU
$button.Add_Click({
    if ($listBox.SelectedItem) {
        Set-Clipboard $textBox.Text
        [System.Windows.Forms.Application]::Exit()
    }
})


# FORMU ÇALIŞTIR (sessize alındı)
$null = [System.Windows.Forms.Application]::Run($form)# SESSİZLİK AYARLARI (en üstte)
$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# EnableVisualStyles sessize alındı
$null = [System.Windows.Forms.Application]::EnableVisualStyles()

# CONFIG DOSYA YOLU (Documents + Gizli klasör)
$docPath = [Environment]::GetFolderPath("MyDocuments")
$configDir = Join-Path $docPath "HazirMetinSecici"
$configFile = Join-Path $configDir "config.txt"

# Klasör yoksa oluştur ve gizle
if (-not (Test-Path $configDir)) {
    $null = New-Item -ItemType Directory -Path $configDir
}
$item = Get-Item $configDir
$item.Attributes = $item.Attributes -bor [IO.FileAttributes]::Hidden

# Config dosyası yoksa oluştur ve gizle
if (-not (Test-Path $configFile)) {
    $null = New-Item -ItemType File -Path $configFile -Force
}
$item2 = Get-Item $configFile
$item2.Attributes = $item2.Attributes -bor [IO.FileAttributes]::Hidden

# FORM OLUŞTURMA
$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.SizeGripStyle = "Hide"
$form.Text = "Hazır Metin Seçici"
$form.Size = New-Object System.Drawing.Size(520,460)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true

$pathBox = New-Object System.Windows.Forms.TextBox
$pathBox.Location = New-Object System.Drawing.Point(10,10)
$pathBox.Size = New-Object System.Drawing.Size(300,25)
$pathBox.ReadOnly = $false

$pathBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        if (Test-Path $pathBox.Text) {
            $pathBox.Text | Set-Content $configFile -Encoding UTF8
            Yukle-TxtDosyalari $pathBox.Text
        }
    }
})


$browseBtn = New-Object System.Windows.Forms.Button
$browseBtn.Location = New-Object System.Drawing.Point(400,8)
$browseBtn.Size = New-Object System.Drawing.Size(80,50)
$browseBtn.Text = "Değiştir"

$enterBtn = New-Object System.Windows.Forms.Button
$enterBtn.Location = New-Object System.Drawing.Point(320,8) # Değiştir'in solu
$enterBtn.Size = New-Object System.Drawing.Size(80,50)
$enterBtn.Text = "Enter"

$enterBtn.Add_Click({
    if (Test-Path $pathBox.Text) {
        $pathBox.Text | Set-Content $configFile -Encoding UTF8
        Yukle-TxtDosyalari $pathBox.Text
    }
})


$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,70)
$listBox.Size = New-Object System.Drawing.Size(480,100)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,175)
$textBox.Size = New-Object System.Drawing.Size(480,190)
$textBox.Multiline = $true
$textBox.ReadOnly = $true
$textBox.ScrollBars = "Vertical"

$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(10,375)
$button.Size = New-Object System.Drawing.Size(480,40)
$button.Text = "Panoya Kopyala"

$form.Controls.AddRange(@(
    $pathBox,
    $pathLabel,
    $enterBtn,
    $browseBtn,
    $listBox,
    $textBox,
    $button
))



# TXT DOSYALARINI YÜKLEME FONKSİYONU
$txtFiles = @{}
function Yukle-TxtDosyalari($klasor) {
    $listBox.Items.Clear()
    $txtFiles.Clear()

    if (Test-Path $klasor) {
        Get-ChildItem $klasor -Filter *.txt | ForEach-Object {
            $listBox.Items.Add($_.BaseName)
            $txtFiles[$_.BaseName] = $_.FullName
        }
    }
}

# CONFIG DOSYASINDAN KAYITLI YOLU OKU
$savedPath = Get-Content $configFile -ErrorAction SilentlyContinue
if ($savedPath -and (Test-Path $savedPath)) {
    $pathBox.Text = $savedPath
    Yukle-TxtDosyalari $savedPath
}

# GÖZAT BUTONU
$browseBtn.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Metin klasörünü seçin"

    if ($dialog.ShowDialog() -eq "OK") {
        $pathBox.Text = $dialog.SelectedPath
        $dialog.SelectedPath | Set-Content $configFile -Encoding UTF8
        Yukle-TxtDosyalari $dialog.SelectedPath
    }
})

# LİSTBOX SEÇİMİ DEĞİŞİNCE
$listBox.Add_SelectedIndexChanged({
    if ($listBox.SelectedItem) {
        $textBox.Text = Get-Content $txtFiles[$listBox.SelectedItem] -Raw -Encoding UTF8
    }
})

# PANOYA KOPYALA BUTONU
$button.Add_Click({
    if ($listBox.SelectedItem) {
        Set-Clipboard $textBox.Text
        [System.Windows.Forms.Application]::Exit()
    }
})

# FORMU ÇALIŞTIR (sessize alındı)
$null = [System.Windows.Forms.Application]::Run($form)
