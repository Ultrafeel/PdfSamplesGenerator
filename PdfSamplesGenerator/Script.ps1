#
# Script.ps1
#
if ($args.Count -gt 0 -and $args[0].LENGTH -ge 0)
{  $targetP =  $args[0]
}
else
{ $targetP = Get-Location }

$cont1 = Get-ChildItem $targetP 
$Arch = $cont1| Where Extension -in '.zip','.rar'
#$nonArch = $cont1| Where-Object {$_.Extension -notin '.zip','.rar','.config' }
 if ($Arch.Length -eq 0 )
 {

  foreach ($value in $cont1){
 
   if ($value.Extension -eq ".docx")
   {
   "Wow"
	   $value.FullPath
	   Print1 ($value)
    break; 
    }
	Write-Host   $value.FullName
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

$PRINTERNAMe="Bullzip PDF Printer"
$PRINTERNAMe
# PDF Writer - bioPDF

# Create settings \ runonce.ini
# $LAPP=$env:LOCALAPPDATA
$LAPP=$env:APPDATA
$SF1="settings.ini"

IF ($LAPP.length -eq 0) 
{	$LAPP="$env:USERPROFILE\Local Settings\Application Data" }
$settings="$LAPP\PDF Writer\$PRINTERNAME\$SF1"
ECHO $settings
	$settFile=$null
	$settingsBackFile=$null
	$settingsBackFileName  =  $settFile.name + ".back"
IF (Test-Path "$settings" )
	{
		$settFile = (Get-Item $settings)
		 rename-item $settFile $settingsBackFileName -Force
	$settingsBackFile = Join-Path $settFile.Directory $settingsBackFileName|Get-Item 
	}
	#(rename "$settings" "$SF1.back")
$samplesTargetDirName = "Образец"
	$sampleSuffix = "_образец"
	$samplesTarget = Join-Path $file.Directory $samplesTargetDirName
	$sampleFileName = $file.basename + $sampleSuffix
	# %CD%\out\demo.pdf
ECHO "Save settings to \"$settings\""
ECHO "[PDF Printer]" >  "$settings"
ECHO "output=$samplesTarget\$sampleFileName.pdf" >> "$settings"
ECHO author=Demo Script >> "$settings"
ECHO showsettings=never >> "$settings"
ECHO showpdf=no >> "$settings"
ECHO "watermarktext=Batch Demo" >>  "$settings"
ECHO confirmoverwrite=no >>  "$settings"
 "$file.FullName"
	 "{$file.FullName}"
& $printto (""""+$file.FullName+"""") (""""+$PRINTERNAME+"""")


# $printto.exe "in\example.rtf" "$PRINTERNAME"
ECHO ERRORLEVEL=$lastexitcode
IF ($settingsBackFile -ne $null -and $settingsBackFile.Exists) #(Test-Path "$settings.back")
	{Rename-Item -Force $settingsBackFile.FullName $settfile.Name}
	elseif (Test-Path $settingsBackFileName)
	{
		Rename-Item -Force $settingsBackFileName $settfile.Name
	}
}


#Foreach-Object {
#    $content = Get-Content $_.FullName

#    #filter and save content to the original file
#    $content | Where-Object {$_ -match 'step[49]'} | Set-Content $_.FullName

#    #filter and save content to a new file 
#    $content | Where-Object {$_ -match 'step[49]'} | Set-Content ($_.BaseName + '_out.log')
#}