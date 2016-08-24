#
# Script.ps1
#
if ($args.Count -gt 0 -and $args[0].LENGTH -ge 0)
{ $targetP = $args[0]
}
else
{ $targetP = Get-Location }

$cont1 = Get-ChildItem $targetP
$Arch = $cont1 | Where Extension -In '.zip','.rar'
#$nonArch = $cont1| Where-Object {$_.Extension -notin '.zip','.rar' }#,'.config'
if ($Arch.LENGTH -eq 0)
{

  foreach ($value in $cont1) {

    if ($value.Extension -eq ".rtf")#docx
    {
      "Wow"
      $value.FullPath
      Print1 ($value)
      break;
    }
    Write-Host $value.FullName
    Write-Host $value
  }
}
else
{

}


$printto = "d:\INSTALL\!office\Bullzip\files\printto.exe"

function Print1 ($file)
{

  # Set environment variables used by the batch file

  $PRINTERNAMe = "Bullzip PDF Printer"
  $PRINTERNAMe
  # PDF Writer - bioPDF

  # Create settings \ runonce.ini
  # $LAPP=$env:LOCALAPPDATA
  $LAPP = $env:APPDATA
  $SF1 = "settings.ini"

  if ($LAPP.LENGTH -eq 0)
  {
    $LAPP = "$env:USERPROFILE\Local Settings\Application Data"
  }
  $settings = "$LAPP\PDF Writer\$PRINTERNAME\$SF1"
  ECHO $settings
  $settFile = $null
  $settingsBackFile = $null
   if (Test-Path "$settings")
  {
    $settFile = (Get-Item $settings)
	 $settingsBackFileName = $SF1 + ".back"
	 $settingsBackFile = Join-Path $settFile.Directory $settingsBackFileName|Get-Item  
	  Remove-Item $settingsBackFile.FullName -Force;
	  Move-Item $settFile.FullName $settingsBackFile.FullName -Force
	  # Get-Item $settingsBackFile
    # rename-item $settFile $settingsBackFileName -Force

    #  $newSett = New-Item $(Join-Path $settFile.Directory ($settFile.name + ".new")) # $newSett = New-Item $(Join-Path $settFile.Directory ($settFile.name + ".new"))
   # $newSett.Replace(($settFile.FullName) ,(Join-Path  $settFile.Directory   $settingsBackFileName ) ,($true) )
   
  }
	else
	{
	 #$settingsBackFileName = $SF1 + ".back"
 
	}
  #(rename "$settings" "$SF1.back")
  $samplesTargetDirName = "Образец"
  $sampleSuffix = "_образец"
	$watermarkText = "образец"
	$samplesTarget = Join-Path $file.Directory $samplesTargetDirName
  $sampleFileName = $file.basename + $sampleSuffix
  # %CD%\out\demo.pdf
  ECHO "Save settings to \" $settings\""
  ECHO "[PDF Printer]" > "$settings"
  ECHO "output=$samplesTarget\$sampleFileName.pdf" >> "$settings"
  ECHO author=PdfSamplesGenerator >> "$settings"
  ECHO showsettings=never >> "$settings"
  ECHO showpdf=no >> "$settings"
  Out-File "$settings" -Append -Encoding "CP1251" -InputObject "watermarktext=$watermarkText"

  ECHO 
@"
  watermarkfontsize=50
  watermarkrotation=c2c
  watermarkcolor=
  watermarkfontname=arial.ttf
  watermarkoutlinewidth=2
  watermarklayer=top
  watermarkverticalposition=center
  watermarkhorizontalposition=center
	PrinterFirstPage=1
	PrinterLastPage=2
"@
	>> "$settings"

  # ECHO "watermarktext=$watermarkText" >> "$settings"
  ECHO confirmoverwrite=no >> "$settings"
  "$file.FullName"
  "{$file.FullName}"
  & $printto ("""" + $file.FullName + """") ("""" + $PRINTERNAME + """")


  # $printto.exe "in\example.rtf" "$PRINTERNAME"
  ECHO ERRORLEVEL=$lastexitcode
  if ($settingsBackFile -ne $null -and $settingsBackFile.Exists) #(Test-Path "$settings.back")
  { 
	 # Remove-Item -Force $settings
	 # move-Item -Force $settingsBackFile.FullName $SF1 
   }
  elseif (Test-Path $settingsBackFileName)
  {
    Rename-Item -Force $settingsBackFileName $settfile.name
  }
}


#Foreach-Object {
#    $content = Get-Content $_.FullName

#    #filter and save content to the original file
#    $content | Where-Object {$_ -match 'step[49]'} | Set-Content $_.FullName

#    #filter and save content to a new file 
#    $content | Where-Object {$_ -match 'step[49]'} | Set-Content ($_.BaseName + '_out.log')
#}
