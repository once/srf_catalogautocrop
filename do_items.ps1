#Folders
$FOLDER_CATALOGS = "00_CATALOG";
$FOLDER_PAGES = "01_PAGES";
$FOLDER_MAPS = "02_MAPS";
$FOLDER_MAPSETTING = "03_MAPSETTING";
$FOLDER_ITEMS = "04_ITEMS";
$FOLDER_UNPROCESSED= "05_UNPROCESSED";
$FOLDER_OCR = "06_OCR";

$MAPSETTING_CONFIG_FILE = "MAPSETTING.TXT";

Clear-Host

Write-Host "Загрузка трафаретов нарезки...";
Add-Type -Path "microsoft.mshtml.dll"

$maps = @{};
$map_files = Get-ChildItem $FOLDER_MAPS -File -Recurse | ? {$_.extension -in ".aspx",".map",".html",".htm"};

for ($i=0; $i -lt $map_files.Count; $i++) {

	
	
	$map_file_base_name = $map_files[$i].BaseName;
	$file_dir_name = $map_files[$i].DirectoryName;
    
    $pos1 = $file_dir_name.IndexOf("\"+$FOLDER_MAPS+"\");
	$dir_len = $file_dir_name.Length;
    $file_short_path =  $file_dir_name.substring($pos1,$dir_len-$pos1);

    $distributor_part = $file_short_path.split("\")[2];
	$distributor_id = $distributor_part.split("_")[0];
    

	
	if ( -Not (($maps[$distributor_id]) -and ($maps[$distributor_id].GetType().Name -eq "Hashtable")))	{
		$maps[$distributor_id] = @{};
	}
	
	$maps[$distributor_id][$map_file_base_name] = @();
	
	$source = Get-Content -Path $map_files[$i].FullName -Raw;
	#$html = New-Object -ComObject "HTMLFile";
	$html = New-Object mshtml.HTMLDocumentClass
	
	$html.IHTMLDocument2_write($source);

	$areas = $html.getElementsByTagName("area");

	foreach ($area in $areas) {
		
		$coords = $area.getAttribute("coords").split(",");

		$x1 = $coords[0];
		$y1 = $coords[1];
		$x2 = $coords[2];
		$y2 = $coords[3];
		
		$width = $x2-$x1;
		$height = $y2-$y1;
		
		$crop_string = '{0}x{1}+{2}+{3}' -f $width,$height,$x1,$y1;
		
		$maps[$distributor_id][$map_file_base_name] += $crop_string;
	}

}

#Write-Output $maps




$files = Get-ChildItem $FOLDER_PAGES -File -Recurse | ? {$_.extension -in ".png"};

for ($i=0; $i -lt $files.Count; $i++) {
    
		$file_full_name = $files[$i].FullName;
		$file_base_name = $files[$i].BaseName;
		$file_dir_name = $files[$i].DirectoryName;
		$fileindex = $file_base_name.Substring($file_base_name.LastIndexOf("_")+1);
		$pos1 = $file_dir_name.IndexOf("\"+$FOLDER_PAGES+"\");
		$dir_len = $file_dir_name.Length;
		$file_short_path =  $file_dir_name.substring($pos1,$dir_len-$pos1);
		$city_part = $file_short_path.split("\")[2];
		$distributor_part = $file_short_path.split("\")[3];
		$dates_part = $file_short_path.split("\")[4];
		$distributor_id = $distributor_part.split("_")[0];
		<#Write-Host $file_dir_name
		Write-Host $file_short_path
		Write-Host $city_part
		Write-Host $distributor_part
		Write-Host $dates_part#>
		
	

		try {
			#Чтение файла соответствия трафаретов страницам
			$pages_to_maps_file = $file_dir_name.Replace($FOLDER_PAGES,$FOLDER_MAPSETTING) + "\" + $MAPSETTING_CONFIG_FILE;
			$pages_to_maps = @{};
            $pages_to_maps = get-content $pages_to_maps_file -ErrorAction Stop | ConvertFrom-StringData

		}
		catch {
			Write-Warning ("Файл соответствия трафаретов страницам отсутствует ("  + $pages_to_maps_file + ")")
			Break;
		}
		
#		Write-Output $pages_to_maps


		
	Write-Host "Обработка страницы: " $file_full_name;
	

		
	if ($pages_to_maps.Keys.Contains($fileindex)) {
		
		Write-Host "Разрезаем страницу...";

		
		# Определеяем целевую папку
		
		
		$target_dir = $FOLDER_ITEMS + "\" + $city_part + "\" + $distributor_part + "\" + $dates_part;
			
	
			if ( -Not (Test-Path -Path $target_dir -PathType Container)) {
				Write-Host "Целевая папка " $target_dir " не найдена и поэтому будет создана.";
				New-Item -ItemType Directory -Force -Path $target_dir
			}
			
		 $crops_counter=1;
		 $map_index = $pages_to_maps.$fileindex;

           
		 $maps[$distributor_id][$map_index] | ForEach {
				
                $_ | ForEach {
                    
                    $map_settings = $_;
                    $arguments = "convert "+$file_full_name+" -crop "+$map_settings+" "+$target_dir+"\"+$file_base_name+"_"+$crops_counter+".png";
				
				    Invoke-Expression "gm.exe $arguments";
				    #Write-Host "gm.exe $arguments";
				    $crops_counter++;
                }
				
		 }
		
	}
	else {
		
		Write-Warning ("Для страницы " + $fileindex + " не определен трафарет")
			
			$unproc_target_dir = $FOLDER_UNPROCESSED + "\" + $city_part + "\" + $distributor_part + "\" + $dates_part;
			
	
			if ( -Not (Test-Path -Path $unproc_target_dir -PathType Container)) {
				Write-Host "Целевая папка " $unproc_target_dir " не найдена и поэтому будет создана.";
				New-Item -ItemType Directory -Force -Path $unproc_target_dir
			}
			
		Write-Host "Страница скопирована в папку " $unproc_target_dir " для ручной нарезки"
		Copy-Item $file_full_name $unproc_target_dir
		
	}
	
	
}