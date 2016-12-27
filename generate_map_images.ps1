$FOLDER_MAPS = "02_MAPS";
$MAPSIZE_CONFIG_FILE = "_mapsize.txt";
 
Write-Host  -NoNewLine "Генерация схем трафаретов...";

Add-Type -Path "microsoft.mshtml.dll"

$map_files = Get-ChildItem $FOLDER_MAPS -File -Recurse | ? {$_.extension -in ".aspx",".map",".html",".htm"};;

	
for ($i=0; $i -lt $map_files.Count; $i++) {
	
	$map_file_base_name = $map_files[$i].BaseName;
	$file_dir_name = $map_files[$i].DirectoryName;
	
	$mapsize_file_name = $map_files[$i].DirectoryName + "\" + $MAPSIZE_CONFIG_FILE ;
	if ( -Not (Test-Path -Path $mapsize_file_name -PathType Leaf)) {
				Write-Warning ("В папке " + $file_dir_name + " не найден файл " + $MAPSIZE_CONFIG_FILE + " с конфигурацией размеров трафарета. Создайте такой файл, указав в нем размеры трафарета в формате Ширина x Высота");
				Break;
	}
	
	$canvas_dimensions = Get-Content $mapsize_file_name;
	
	$canvas_width = $canvas_dimensions.split("x")[0];
	$canvas_height = $canvas_dimensions.split("x")[1];
    
    $pos1 = $file_dir_name.IndexOf($FOLDER_MAPS+"\");
	$dir_len = $file_dir_name.Length;
    $file_short_path =  $file_dir_name.substring($pos1,$dir_len-$pos1);
    $distributor_part = $file_short_path.split("\")[2];
	
	
	$source = Get-Content -Path $map_files[$i].FullName -Raw;
	$html = New-Object mshtml.HTMLDocumentClass
	$html.IHTMLDocument2_write($source);
	

	$areas = $html.getElementsByTagName("area");

	$draw_objects = "";
	
	$output_file_name = "{0}\{1}.png" -f $file_short_path, $map_file_base_name;
	
	#$arguments = 'convert -size {0}x{1} xc:none -fill none {2}' -f $canvas_width,$canvas_height,$output_file_name;
	#Write-Output $arguments;
	#$Invoke-Expression "gm.exe $arguments";
	
	$draw_objects = "";
	foreach ($area in $areas) {
		
		$coords = $area.getAttribute("coords").split(",");

		$x1 = $coords[0];
		$y1 = $coords[1];
		$x2 = $coords[2];
		$y2 = $coords[3];
			
		$draw_objects += " rectangle {0},{1},{2},{3}" -f $x1,$y1,$x2,$y2;
		#$draw_objects += " line {0},{1},{2},{3}" -f $x1,$y1,$x2,$y2;
		#$draw_objects += " line {0},{1},{2},{3}" -f $x2,$y1,$x1,$y2;
			
		#$arguments = 'mogrify -fill none -stroke blue -strokewidth 2 -draw "{0}" {1}' -f $draw_objects,$output_file_name;
		#Invoke-Expression "gm.exe $arguments";
	}
	
#	$output_file_name = "{0}\{1}.png" -f $file_short_path, $map_file_base_name;
	
	$arguments = 'convert -size {0}x{1} xc:none -fill none -stroke blue -strokewidth 6 -draw "{2}" {3}' -f $canvas_width,$canvas_height,$draw_objects,$output_file_name;
	#Write-Output $arguments;
	Invoke-Expression "gm.exe $arguments";
	Write-Host -NoNewLine "."
}
Write-Host "Готово.";

