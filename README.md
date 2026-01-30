Bu scripti, powershell ile çıkartıp EXE ye çevirebilirsiniz.
Gerekli komutlar aşağıdadır.

Invoke-PS2EXE .\TXT Görüntüleyici-Kopyalayıcı.ps1 .\TXT Görüntüleyici-Kopyalayıcı.exe `
  -console `
  -noOutput `
  -requireAdmin:$false `
  -title "Hazır Metin Seçici"


  Ek olarak EXE halide mevcuttur.
