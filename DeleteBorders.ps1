########## Удаляем границы всех объектов в файле QlikView:


Write-Host "Enter the absolute path of the QVW-file" -ForegroundColor Yellow
###Проверка является ли введенная строка ссылкой на QV файл
$filepath = Read-host
if (-not (Test-Path $filepath)){
	Write-Host "The file with such path doesn't exists" -ForegroundColor Red
	exit
	}
if (-not ($filepath.Contains(".qvw") -or $filepath.Contains(".QVW"))){
	Write-Host "The file is not qvw-file" -ForegroundColor Red
	exit
	}

# Открываем приложение QlikView	
$objApp = new-object -comobject QlikTech.QlikView
$objApp.WaitForIdle

# Сворачиваем все окна (только для того, чтобы свернуть окно QlikView чтобы перебор объектов происходил без отрисовки визуализации. Простого способа свернуть только окно QV я не нашел)	
$shell = New-Object -ComObject "Shell.Application"
$shell.minimizeall()

# Открываем QV документ
$objDoc = $objApp.OpenDoc($filepath)

Write-Host "There are $($objDoc.NoOfSheets()) sheets in the document" -ForegroundColor Yellow
$Counter = 0

# Перебираем все листы в документе 
for ($i=0;$i -lt ($objDoc.NoOfSheets());$i++){
	Write-Host "Open sheet $($objDoc.Sheets($i).GetProperties().SheetID): $($objDoc.Sheets($i).NoOfSheetObjects()) objects at the sheet" -ForegroundColor Yellow
	$objSheet = $objDoc.ActivateSheet("$($objDoc.Sheets($i).GetProperties().SheetID)")
	
	# Перебираем все объекты в документе и удаляем границы в каждом из них
	foreach ($object in $objSheet.GetSheetObjects()){
		if ($null -ne $object.GetProperties().Layout.Frame.ObjectId){
			$ID =$object.GetProperties().Layout.Frame.ObjectId
		} elseif ($null -ne $object.GetProperties().GraphLayout.Frame.ObjectId){
			$ID =$object.GetProperties().GraphLayout.Frame.ObjectId
		} elseif ($null -ne $object.GetProperties().Frame.ObjectId){
			$ID =$object.GetProperties().Frame.ObjectId
		} else{
			Write-Host "Frame atribute error! Check property path." -ForegroundColor Red
			Exit
		}
		
		$objProperties = $object.GetProperties()
		$NoBorder = 0
		if ($null -ne $objProperties.Layout.Frame){
			if ($objProperties.Layout.Frame.BorderWidth -eq 0) 
				{$NoBorder = 1}
			else 
				{$objProperties.Layout.Frame.BorderWidth = 0}
		} elseif ($null -ne $objProperties.GraphLayout.Frame){
			if ($objProperties.GraphLayout.Frame.BorderWidth -eq 0) 
				{$NoBorder = 1}
			else {$objProperties.GraphLayout.Frame.BorderWidth = 0}
		} elseif ($null -ne $objProperties.Frame){
			if ($objProperties.Frame.BorderWidth -eq 0) 
				{$NoBorder = 1}
			else {$objProperties.Frame.BorderWidth = 0}
		} else{
			Write-Host "Frame atribute error! Check property path." -ForegroundColor Red
		}
		$object.SetProperties($objProperties)
		
		if ($NoBorder -eq 0) 
			{
			Write-Host "$ID-Object Border has been deleted" -ForegroundColor Green
			$Counter += 1
			}
	}
}
# Сохраняем документ и закрываем приложение QV 
$objDoc.Save()
$objDoc.CloseDoc()
$objApp.Quit()
Write-Host "Borders are successfully deleted from $Counter $(if ($Counter -eq 1) {"object"} else {"objects"})" -ForegroundColor Yellow

#Вызывем popup уведомление на случай, если кто-то не развернул обратно окно powershell и не знает когда скрипт отработал
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Script is completed",0,"Done",0x1) | Out-Null

Exit


