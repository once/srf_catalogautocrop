#Folders
$FOLDER_OCR = "06_OCR"

Clear-Host
Write-Host "Подготовка миниатюр и сжатие изображений...."

$files = Get-ChildItem $FOLDER_OCR -recurse | ? {$_.extension -eq ".png"}

for ($i=0; $i -lt $files.Count; $i++) {
    
	$file_full_name = $files[$i].FullName 
	$file_name = $files[$i].Name;
	$file_base_name = $files[$i].BaseName;
	$file_directory = $files[$i].DirectoryName;
	
	Write-Host "Обработка файла: " $file_name;
	
	$new_file_base_name = ([guid]::NewGuid()).guid;
	
	$arguments = "convert -size 180x180 """ + $file_full_name + """ -resize 180x180 -quality 90 +profile ""*"" """+ $file_directory  + "\" + $new_file_base_name +"_mini.jpg"""
	#Write-Output $arguments
	Invoke-Expression "gm.exe $arguments";
	
	$arguments = "convert -size 800x600 """ + $file_full_name + """ -resize 800x600 -quality 90 +profile ""*"" """+ $file_directory  + "\" + $new_file_base_name +".jpg"""
	Invoke-Expression "gm.exe $arguments";
	
	$txt_oldfilename = $file_directory + "\" + $file_base_name +".txt"
	$txt_newfilename = $file_directory + "\" + $new_file_base_name+".txt"

	Invoke-Expression "Rename-Item ""$txt_oldfilename"" ""$txt_newfilename""";
	Invoke-Expression "Remove-Item ""$file_full_name""";
}

