#
# Script.ps1
#


function printto
{
	param( [string]$file, [string]$printer)

	if ($printer -ne $null)
	{
		Start-Process  –FilePath $file -ArgumentList $printer -Verb "printto"  -Wait
	}
	else
	{
		Start-Process  –FilePath $file  -Verb "print"   -Wait

	}

}
$printto =  #= "d:\INSTALL\!office\Bullzip\files\printto.exe"
$pdftk  =  Get-Command "pdftk" -ErrorAction SilentlyContinue
if ($pdftk -eq $null)
{ $pdftk  = "C:\Program Files (x86)\PDFtk\bin\pdftk.exe" }

$u7z  =  Get-Command "7z" -ErrorAction SilentlyContinue
if ($u7z -eq $null)
{ $u7z  = "C:\Program Files (x86)\Universal Extractor\bin\7z.exe" }

function WaitForFile ($file)
{
	[int]$i=10000
	for (; $i -gt 0 -and !(Test-Path $file); $i--) 
		{ Start-Sleep 10 }
	return ($i -gt 0);
}

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
function msgBoxRetryCancel($x)
{
		# “Продолжить” или “Отменить”

	$OUTPUT=  [System.Windows.Forms.MessageBox]::Show($x, 
		'Генератор PDF образцов:PowerShell', 
	[Windows.Forms.MessageBoxButtons]::RetryCancel , 
	[Windows.Forms.MessageBoxIcon]::Exclamation, #Information
		 [Windows.Forms.MessageBoxDefaultButton]::Button1)
	# [Windows.Forms.MessageBoxOptions]::ServiceNotification
	return $OUTPUT;
}
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
  $samplesTargetDirName = "Образцы"
  $sampleSuffix = "_образец"
	$watermarkText = "OBRAZEC" # "образец"
	$samplesTarget = Join-Path -Path $($file.Directory) -ChildPath $samplesTargetDirName
  $sampleFileName = $file.basename + $sampleSuffix


	while (!(Test-Path $samplesTarget))
	{
		$msg = "Нет папки `“$samplesTargetDirName`” по пути `“$($file.Directory)`”"
		$OUTPUT = msgBoxRetryCancel($msg)
		if ($OUTPUT -eq [System.Windows.Forms.DialogResult]::Cancel ) 
		{ 
			throw $msg;
		} 
	}

	$outFileS =	"$samplesTarget\$sampleFileName" 
	$outFile =	("$outFileS"+ ".pdf")
	
	 # %CD%\out\demo.pdf
  ECHO "Save settings to \" $settings\""
  #ECHO "[PDF Printer]" > "$settings"
  #ECHO "output=$outFile" 			 >> "$settings"
  #ECHO author=PdfSamplesGenerator 	>> "$settings"
  #ECHO showsettings=never 			>> "$settings"
  #ECHO showpdf=no >> "$settings"

 # Out-File "$settings" -Append -Encoding "unicode" -InputObject 	"$watermarkText"
 
	# Нельзя добавлять в конце пробелы!!
  Out-File "$settings" -Encoding "unicode" -InputObject @"
[PDF Printer]
  output=$outFile
  author=PdfSamplesGenerator
  showsettings=never
  showpdf=no
  watermarktext=$watermarkText
  watermarkfontsize=70
  watermarkrotation=c2c
  watermarkcolor=
  watermarkfontname=arial.ttf
  watermarkoutlinewidth=2
  watermarklayer=top
  watermarkverticalposition=center
  watermarkhorizontalposition=center
  confirmoverwrite=no
"@ 
	#	>> "$settings"
	#PrintToPrinter=Foxit Reader PDF Printer
	#PrinterFirstPage=1
	#PrinterLastPage=2

  # ECHO "watermarktext=$watermarkText" >> "$settings"
 # ECHO >> "$settings"
  "$file.FullName"

  printto ("""" + $file.FullName + """") ("""" + $PRINTERNAME + """") -ErrorAction Continue -ErrorVariable ProcessError
  #  Silently
	If ($ProcessError) {

	  echo  $file.FullName + "не имеет печатающей программы" 

		return;
	}


  # $printto.exe "in\example.rtf" "$PRINTERNAME"
  $ptERRORLEVEL=$lastexitcode
	# if ($ptERRORLEVEL -eq 0) $res11 = & $pdftk "$outFile" dump_data | Select-string -Pattern "PageMediaNumber: ([0-9]*)"
	  if ( WaitForFile($outFile) )
	{
		$res11 = & $pdftk "$outFile" dump_data | Select-string -Pattern "PageMediaNumber: ([0-9]+)"

		$numOfPages = ($res11.Matches[0].Groups[1].value)
		if ($numOfPages -gt 8)
		{
		$outFileCut = ("$samplesTarget\$sampleFileName"+"C.pdf") #($outFile
		& $pdftk "$outFile" cat 1-8 output $outFileCut  verbose
		if (Test-path  $outFileCut)
		{
			Remove-Item $outFile -Force;
			Move-Item $outFileCut $outFile -Force
		}
		}
	}
	
  if ($settingsBackFile -ne $null -and $settingsBackFile.Exists) #(Test-Path "$settings.back")
  { 
	 Remove-Item -Force $settings
	 move-Item -Force $settingsBackFile.FullName $SF1 
   }
  elseif (Test-Path $settingsBackFileName)
  {
    Rename-Item -Force $settingsBackFileName $settfile.name
  }
}

if ($args.Count -gt 0 -and $args[0].LENGTH -ge 0)
{ $targetP = $args[0]
}
else
{ $targetP = Get-Location }

$cont1 = Get-ChildItem $targetP


$archs = 

((	"7z	"	)	,		"7z"	) 				,
((	"xz	"	)	,		"XZ"	)				,
((	"zip	"	)	,		"ZIP"	)			,
((	"gz","gzip","tgz"	)	,	"GZIP"	)		,
((	"bz2","bzip2","tbz2","tbz"),	"BZIP2"	)	,
((	"tar"	)	,		"TAR"	)				,
((	"wim","swm"	)	,		"WIM"	),
((	"lzma"	)	,		"LZMA"	)	 ,
((	"rar"	)	,		"RAR"	)	 ,
((	"cab"	)	,		"CAB"	)	 ,
((	"arj"	)	,		"ARJ"	)	 ,
((	"z","taz")	,		"Z"	)		 ,
((	"cpio"	)	,		"CPIO"	)
	


$ArchsExts = ($archs | ForEach-Object { $_[0]})

 $cont1 | Where Extension -In '.zip','.rar'
#$nonArch = $cont1| Where-Object {$_.Extension -notin '.zip','.rar' }#,'.config'


 # foreach ($value in $cont1) 
 For($i = 1; $i -le $cont1.count; $i++)
{ 
 
$value = $cont1[$i]
Write-Progress -Activity “Generating samples” -status “ file $value” `
-percentComplete ($i / $cont1.count*100)
# $cont1 | Select name

if ($value.Attributes -band  [System.IO.FileAttributes]::Directory)
	  {
		continue;  
	  }
   Write-Host $value.FullName
    Write-Host $value
	$fExt = $value.Extension
	if ($fExt -ne $null)
	{ $fExt.TrimStart('.') }
    if ($fExt -in $ArchsExts)#docx
    {

	}
	else  
	{
		if ($value.Extension -in ".rtf", ".cdr", ".jpg", ".tif", ".tiff", ".doc", ".docx", ".indd" )
		{

			Print1 ($value)
		}
		elseif ($value.Extension -eq ".pdf")
		{
			Print1 ($value)
		}
		
      $value.FullPath
      #break;
    }
   }




#Foreach-Object {
#    $content = Get-Content $_.FullName

#    #filter and save content to the original file
#    $content | Where-Object {$_ -match 'step[49]'} | Set-Content $_.FullName

#    #filter and save content to a new file 
#    $content | Where-Object {$_ -match 'step[49]'} | Set-Content ($_.BaseName + '_out.log')
#}






#( " 7z " ) , ( " 7z " )
#( " xz " ) , ( " XZ " )
#( " zip " ) , ( " ZIP " )
#( " gz gzip tgz " ) , ( " GZIP " )
#( " bz2 bzip2 tbz2 tbz " ) , ( " BZIP2 " )
#( " tar " ) , ( " TAR " )
#( " wim swm " ) , ( " WIM " )
#( " lzma " ) , ( " LZMA " )
#( " rar " ) , ( " RAR " )
#( " cab " ) , ( " CAB " )
#( " arj " ) , ( " ARJ " )
#( " z taz " ) , ( " Z " )
#( " cpio " ) , ( " CPIO " )
