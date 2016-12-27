$dpi_density = "300";
$FOLDER_CATALOGS = "00_CATALOG";
$FOLDER_PAGES = "01_PAGES";

Clear-Host 

Write-Host "Обработка каталогов из папки " $FOLDER_CATALOGS " ... "



$pdf_files = Get-ChildItem $FOLDER_CATALOGS -File -Recurse | ? {$_.extension -in ".pdf"};

for ($i=0; $i -lt $pdf_files.Count; $i++) {
	
	$file_name = $pdf_files[$i].Name;
    $file_name_full = $pdf_files[$i].FullName;
    $file_dir_name = $pdf_files[$i].DirectoryName;
    
    $pos1 = $file_dir_name.IndexOf("\"+$FOLDER_CATALOGS+"\");
	$dir_len = $file_dir_name.Length;
    $file_short_path =  $file_dir_name.substring($pos1,$dir_len-$pos1);

    $city_part = $file_short_path.split("\")[2];
    $distributor_part = $file_short_path.split("\")[3];
	$dates_part = $file_short_path.split("\")[4];
	$city_id = $city_part.split("_")[0];
    $city_name = $city_part.split("_")[1];
    $distributor_id = $distributor_part.split("_")[0];
    $distributor_name = $distributor_part.split("_")[1];
	
    Write-Host "================================="
    Write-Host "Файл каталога: " $file_name
    Write-Host "Город: " $city_name
    Write-Host "Торговая сеть: " $distributor_name
	Write-Host "Сроки акции: " $dates_part
	Write-Host "================================="
    
    Write-Host "Разбиение PDF-файла на отдельные страницы..." 
    
    $target_dir = $file_dir_name.Replace($FOLDER_CATALOGS,$FOLDER_PAGES);

    if ( -Not (Test-Path -Path $target_dir -PathType Container)) {
	    Write-Host "Целевая папка " $target_dir " не найдена и поэтому будет создана.";
        New-Item -ItemType Directory -Force -Path $target_dir
    }

    $output_file_name =  $target_dir + '\image_%02d.png';
    $arguments = 'convert -colorspace rgb -density {0} "{1}" +adjoin "{2}"' -f $dpi_density,$file_name_full,$output_file_name;
    #Write-Output $arguments;
    
    Invoke-Expression "gm.exe $arguments";
}


