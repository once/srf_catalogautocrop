@powershell -NoProfile -ExecutionPolicy Bypass .\post_ocr.ps1

SRFAutocropDataUploader.exe

@powershell -NoProfile -ExecutionPolicy Bypass .\post_upload.ps1
