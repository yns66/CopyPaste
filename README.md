Bu scripti, powershell ile EXE ye çevirebilirsiniz.
Gerekli komutlar aşağıdadır.

Invoke-PS2EXE .\CopyPaste.ps1 .\CopyPaste.exe `
>>   -noConsole `
>>   -noOutput `
>>   -requireAdmin:$false `
>>   -icon ".\CopyPaste.ico" `
>>   -title "Hazır Metin Seçici"

EXE halide mevcuttur.

Ek olarak seçilen dizini Belgeler altında "HazirMetinSecici" adında gizli klasörün içindeki config text dosyasından çekmektedir.
