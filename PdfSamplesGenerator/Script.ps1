#
# Script.ps1
#

function EchoA
{
  for ($i = 0; $i -lt $args.length; $i++)
  {
    "Arg $i is <$($args[$i])>"
  }
}
function printto
{
  param([string]$file,[string]$printer)
  #$err1;
  if ($printer -ne $null)
  {
    Start-Process –FilePath $file -ArgumentList $printer -Verb "printto" -Wait -ErrorVariable err1
  }
  else
  {
    Start-Process –FilePath $file -Verb "print" -Wait -ErrorVariable err1

  }
  if ($err1 -ne $null)
  {
    return $false
  }
  return $true;

}
$printto = #= "d:\INSTALL\!office\Bullzip\files\printto.exe"
$pdftk = Get-Command "pdftk" -ErrorAction SilentlyContinue
if ($pdftk -eq $null)
{ $pdftk = "C:\Program Files (x86)\PDFtk\bin\pdftk.exe" }

$u7z = Get-Command "7z" -ErrorAction SilentlyContinue
if ($u7z -eq $null)
{ $u7z = "C:\Program Files (x86)\Universal Extractor\bin\7z.exe" }

function WaitForFile ($file)
{
  [int]$i = 10000
  for (; $i -gt 0 -and !(Test-Path $file); $i --)
  { Start-Sleep 10 }
  return ($i -gt 0);
}

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
function msgBoxRetryCancel ($x)
{
  # “Продолжить” или “Отменить”

  $OUTPUT = [System.Windows.Forms.MessageBox]::Show($x,
    'Генератор PDF образцов:PowerShell',
    [Windows.Forms.MessageBoxButtons]::RetryCancel,
    [Windows.Forms.MessageBoxIcon]::Exclamation,#Information
    [Windows.Forms.MessageBoxDefaultButton]::Button1)
  # [Windows.Forms.MessageBoxOptions]::ServiceNotification
  return $OUTPUT;
}

#directory to process
function Print1 ($file, [string]$obrazcyParentDir)
{
	if ($obrazcyParentDir -eq $null)
	{
		$obrazcyParentDir = $file.Directory
	}
  # Set environment variables used by the batch file

  $PRINTERNAMe = "Bullzip PDF Printer"
  $PRINTERNAMe
  # PDF Writer - bioPDF

  # Create settings \ runonce.ini
  # $LAPP=$env:LOCALAPPDATA
  $LAPP = $env:APPDATA
  $SF1 = "settings.ini"

  if ($LAPP.length -eq 0)
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
    $settingsBackFile = Join-Path $settFile.Directory $settingsBackFileName | Get-Item
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

  $samplesTarget = Join-Path -Path $obrazcyParentDir -ChildPath $samplesTargetDirName
  $sampleFileName = $file.basename + $sampleSuffix


  while (!(Test-Path $samplesTarget))
  {
    $msg = "Нет папки `“$samplesTargetDirName`” по пути `“$obrazcyParentDir`”"
    $OUTPUT = msgBoxRetryCancel ($msg)
    if ($OUTPUT -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
      throw $msg;
    }
  }

  $outFileS = "$samplesTarget\$sampleFileName"
  $outFile = ("$outFileS" + ".pdf")

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

  $ptSuccess = printto $file.FullName $PRINTERNAME
  #  Silently-ErrorVariable ProcessError -ErrorAction Continue 
  if (!$ptSuccess) {

    ECHO $($file.FullName) + "не имеет печатающей программы"

    return;
  }


  # $printto.exe "in\example.rtf" "$PRINTERNAME"
  #  $ptERRORLEVEL = $lastexitcode
  # if ($ptERRORLEVEL -eq 0) $res11 = & $pdftk "$outFile" dump_data | Select-string -Pattern "PageMediaNumber: ([0-9]*)"
  if (WaitForFile ($outFile))
  {
    $res11 = & $pdftk "$outFile" dump_data | Select-String -Pattern "PageMediaNumber: ([0-9]+)"

    $numOfPages = ($res11.Matches[0].Groups[1].value)
    if ($numOfPages -gt 8)
    {
      $outFileCut = ("$samplesTarget\$sampleFileName" + "C.pdf") #($outFile
      & $pdftk "$outFile" cat 1-8 output $outFileCut verbose
      if (Test-Path $outFileCut)
      {
        Remove-Item $outFile -Force;
        Move-Item $outFileCut $outFile -Force
      }
    }
  }

  if ($settingsBackFile -ne $null -and $settingsBackFile.Exists) #(Test-Path "$settings.back")
  {
    Remove-Item -Force $settings
    Move-Item -Force $settingsBackFile.FullName $SF1
  }
  elseif (Test-Path $settingsBackFileName)
  {
    Rename-Item -Force $settingsBackFileName $settfile.name
  }
}
$docExtensions1 = ".rtf",".cdr",".jpg",".tif",".tiff",".doc",".docx",".indd"
$docExtensions = $docExtensions1 + ".pdf"
function AlgA_Iter
{
	params($value, $obrazcyParentDir)
	if ($value.Extension -in $docExtensions1)
    {

      Print1 ($value, $obrazcyParentDir)
    }
    elseif ($value.Extension -eq ".pdf")
    {
      Print1 ($value, $obrazcyParentDir)
    }

    $value.FullPath
}

if ($args.Count -gt 0 -and $args[0].length -ge 0)
{ 
	$targetP = $args[0]
}
else
{ $targetP = Get-Location }

function Algs([string]$targetP1, [Boolean]$algAForB, $obrazcyParentDir)
{
	
	[Boolean]$algAOnly = $algAForB

$cont1 = Get-ChildItem $targetP1


$archs =

(("7z"),"7z"),
(("xz"),"XZ"),
(("zip"),"ZIP"),
(("gz","gzip","tgz"),"GZIP"),
(("bz2","bzip2","tbz2","tbz"),"BZIP2"),
(("tar"),"TAR"),
(("wim","swm"),"WIM"),
(("lzma"),"LZMA"),
(("rar"),"RAR"),
(("cab"),"CAB"),
(("arj"),"ARJ"),
(("z","taz"),"Z"),
(("cpio"),"CPIO")



$ArchsExts = ($archs | ForEach-Object { $_[0] })

#$nonArch = $cont1| Where-Object {$_.Extension -notin '.zip','.rar' }#,'.config'


# foreach ($value in $cont1) 
for ($iF = 1; $iF -le $cont1.Count; $iF++)
{

  $value = $cont1[$iF]
  Write-Progress -Activity “Generating samples” -Status “ file $value” `
     -PercentComplete ($iF / $cont1.Count * 100)
  # $cont1 | Select name

  if ($value.Attributes -band [System.IO.FileAttributes]::Directory)
  {
    continue;
  }
  Write-Host $value.FullName
  Write-Host $value
  $fExt = $value.Extension
  if ($fExt -ne $null)
  { $fExt = $fExt.TrimStart('.') }
  if ($fExt -in $ArchsExts)
  {

	  #algB
	   if ($algAOnly)
	  {
		  continue;
	}
    $archContT = & $u7z "l" "-slt" $value.FullName

    $aafiles = {@()}.Invoke(); #System.Collections.ObjectModel.Collection`1[System.Management.Automation.PSObject]

    $fileL = $null;
    $pathFoundMode = $false;
    $dirFoundMode = $false;

    for ($i = 0; $i -lt $archContT.Count;++ $i)
    {
      [string]$acString = $archContT[$i]
      if ($acString.length -eq 0)
      {
        if ($fileL -ne $null -and $dirFoundMode)
        {
          $aafiles.Add($fileL)
        }
          $pathFoundMode = $false;
          $dirFoundMode = $false;

		$fileL = $null
      }
      else {

        if (!$pathFoundMode)
        {
          $archParse = $acString | Select-String -Pattern "Path = (.*)"

          if ($archParse -ne $null -and $archParse.Matches[0].Success)
          {
            $pathFoundMode = $true;

            if ($fileL -eq $null)
            {
              $fileL = New-Object PSObject -Property @{ Path = $archParse.Matches[0].Groups[1].value; isdir = $false }
            }
			else
			  {
				   $fileL.Path = $archParse.Matches[0].Groups[1].value #never!
			}
          }
			$archParse
        }
        elseif (!$dirFoundMode)
        {
          $archParse = $acString | Select-String -Pattern "Folder = (.*)"

          if ($archParse -ne $null -and $archParse.Matches[0].Success)
          {
            $dirFoundMode = $true;
            if ($fileL -ne $null)
            {
              $fileL.isdir = ($archParse.Matches[0].Groups[1].value -eq "+")

            }
            else
            { $fileL } # never!
          }

        }

      }
     # $aafiles
    }


	# если папок нет
	  if (($aafiles| Where-Object { $_.isdir }).Length -eq 0)
	  {
		  
		#[System.Reflection.Assembly]::LoadWithPartialName("System.IO.Path")
		$TMPfiltFile = "ExtList.extList"	#Get-Item ($value.Directory.ToString()+  [System.IO.Path]::GetTempFileName()
		 $TMPfullP =   $value.FullName + "ext"
		  $oldWD = Get-Location
		  cd $value.Directory
		out-file	$TMPfiltFile -Encoding "utf8" -InputObject 	( ($docExtensions|% {"*" + $_}) -join "`n")  
	 & $u7z "e" $value.Name "-o$TMPfullP" "-i@$TMPfiltFile" "-y"
		cd $oldWD
		  Algs (Get-Item $TMPfullP)  $true  $value.Directory 
		  Remove-Item $TMPfullP -Force
		 Remove-Item  $TMPfiltFile -Force
	  
	  }
  <#  $archCont = $archContT | Select-String -Pattern "Path = (.*)"

    $archCont
    $archCont2 = $archCont | Where-Object { $_.Matches[0].Success } | Select-Object -Skip 1 | select @{ name = "FName"; Expression = { $_.Matches[0].Groups[1].value } }

    $archCont2
    $TMPfullP = $value.FullName + "ext"
    foreach ($archF in $archCont2)
    {
      $arPath = $archF.FName
      $apSpl = Split-Path $arPath
      $apSplit = $arPath.Split('\')
      $apSplit

      & $u7z "e" $value.FullName "-o$TMPfullP" "-i!$arPath" "-y"
    }
	  #>
  }
  else
  {
	 
    continue; #TODO
    AlgA_Iter $value  $obrazcyParentDir 
    #break;
  }
}


}

Algs $targetP   $false  $null 

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
