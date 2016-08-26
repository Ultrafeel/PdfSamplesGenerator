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

   return  $err1
  #if ($err1 -ne $null)
  #{
  #  return $false
  #}
  #return $true;

}
$printto =  "d:\INSTALL\!office\Bullzip\files\printto.exe"
$pdftk = Get-Command "pdftk" -ErrorAction SilentlyContinue
if ($pdftk -eq $null)
{ $pdftk = "C:\Program Files (x86)\PDFtk\bin\pdftk.exe" }

$u7z = Get-Command "7z" -ErrorAction SilentlyContinue
if ($u7z -eq $null)
{ $u7z = "C:\Program Files (x86)\Universal Extractor\bin\7z.exe" }

function WaitForFile ($file)
{
  [int]$i = 100
  for (; $i -gt 0 -and !(Test-Path $file); $i --)
  { Start-Sleep -Milliseconds 10 }
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
#  ECHO $settings
  $settFile = $null
  $settingsBackFile = $null
  if (Test-Path "$settings")
  {
    $settFile = (Get-Item $settings)
    $settingsBackFileName = Join-Path $settFile.Directory ($SF1 + ".back" )
    $settingsBackFile =  $settingsBackFileName | Get-Item -ErrorAction SilentlyContinue
    Remove-Item $settingsBackFile -Force -ErrorAction SilentlyContinue;
    Move-Item $settFile.FullName $settingsBackFile.FullName -Force
    # Get-Item $settingsBackFile
    # rename-item $settFile $settingsBackFileName -Force

    #  $newSett = New-Item $(Join-Path $settFile.Directory ($settFile.name + ".new")) # $newSett = New-Item $(Join-Path $settFile.Directory ($settFile.name + ".new"))
    # $newSett.Replace(($settFile.FullName) ,(Join-Path  $settFile.Directory   $settingsBackFileName ) ,($true) )

  }
  else
  {
    #$settingsBackFileName = Join-Path $settFile.Directory ($SF1 + ".back" )

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
  #ECHO "Save settings to \" $settings\""
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
  watermarkfontsize=40
  watermarkrotation=c2c
  watermarkcolor=
  watermarkfontname=arial.ttf
  watermarkoutlinewidth=2
  watermarklayer=top
  watermarkverticalposition=center
  watermarkhorizontalposition=center
  confirmoverwrite=no
  showprogressfinished=yes
"@

# TODO: showprogress=yes
  #
	# confirmnewfolder=yes
<#  suppresserrors=no
    rememberlastfoldername=yes
  openfolder=no
  showsaveas=nofile

  device=pdfwrite
  textalphabits=4
  graphicsalphabits=4
  author=PdfSamplesGenerator
  title=
  subject=
  keywords=	  #>

  #	>> "$settings"
  #PrintToPrinter=Foxit Reader PDF Printer
  #PrinterFirstPage=1
  #PrinterLastPage=2

  # ECHO "watermarktext=$watermarkText" >> "$settings"
  # ECHO >> "$settings"

  $ptErr = printto "`"$($file.FullName)`"" "`"$PRINTERNAME`""
  #$ptSuccess  Silently-ErrorVariable ProcessError -ErrorAction Continue 
 
  if (   i$settingsBackFile -ne $null -and $settingsBackFile.Exists) #(Test-Path "$settings.back")
  {
    Remove-Item -Force $settings
    Move-Item -Force $settingsBackFile.FullName $SF1
  }
  elseif (Test-Path $settingsBackFileName)
  {
    Rename-Item -Force $settingsBackFileName $settfile.name
  }
	

 if ($ptErr -ne $null) {

    ECHO ("`"$($file.FullName)`"" + " не имеет печатающей программы :"  + $ptErr.ToString())

    return;
  }


  # $printto.exe "in\example.rtf" "$PRINTERNAME"
  #  $ptERRORLEVEL = $lastexitcode
  # if ($ptERRORLEVEL -eq 0) $res11 = & $pdftk "$outFile" dump_data | Select-string -Pattern "PageMediaNumber: ([0-9]*)"
  do {
	if (WaitForFile ($outFile))
  {
    $res11 = & $pdftk "$outFile" dump_data | Select-String -Pattern "PageMediaNumber: ([0-9]+)"

    $numOfPages = ($res11.Matches[0].Groups[1].value)
    if ($numOfPages -gt 8)
    {

		#TODO: cut pdf first
      $outFileCut = ("$samplesTarget\$sampleFileName" + ".pdf8cut") #($outFile
      & $pdftk "$outFile" cat 1-8 output $outFileCut verbose
      if (Test-Path $outFileCut)
      {
        Remove-Item $outFile -Force;
        Move-Item $outFileCut $outFile -Force
      }
    }
	  break;
  }
   } while(msgBoxRetryCancel("Не конвертируется файл `"$($file.FullName)`" Нажмите `"отмена`" чтобы пропустить его  и Retry что бы подождать") -eq [System.Windows.Forms.DialogResult]::Retry)


}
$docExtensions1 = ".rtf",".cdr",".jpg",".tif",".tiff",".doc",".docx",".indd"
$docExtensions = $docExtensions1 + ".pdf"
function AlgA_Iter
{
	param($value, $obrazcyParentDir)
	if ($value.Extension -in $docExtensions1)
    {

      Print1 $value $obrazcyParentDir
    }
    elseif ($value.Extension -eq ".pdf")
    {
      Print1 $value $obrazcyParentDir
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
for ($iF = 0; $iF -lt $cont1.Count; $iF++)
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
				$path1 = $archParse.Matches[0].Groups[1].value	
				$pathAr1 = $path1.Split('\')
				$fileL = New-Object PSObject -Property @{ Path = $path1; 
					isdir = $false ;
					pathAr =$pathAr1 ; 
					depth = $pathAr1.Count }
            }
			else
			  {
				   $fileL.Path = $archParse.Matches[0].Groups[1].value #never!
			}
          }
			#todo
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
    }

     # $aafiles
		  $aafiles[0].PathAr
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
		 Remove-Item  $TMPfiltFile -Force
		cd $oldWD
		  Algs (Get-Item $TMPfullP)  $true  $value.Directory 
		  Remove-Item $TMPfullP -Force	-Recurse
	  
	  }
	  else
	  {	  # $aafiles |  Measure-Object -Property depth  -
		  $aFfiles	=	$aafiles | Where-Object { !$_.isdir -and ($_.depth -gt 1) }
		   # 	| Sort -Property depth -Descending

		  #первая, если вторая будет глубже - не подходит
		   [int]$depth = 1
		  $deepest_firstIndex  = 0
		  for ($iA = 0; $iA -lt $aFfiles.Count; ++$iA)
		  {	  
			   $afile  = $aFfiles[$iA] ;
			  if ($afile.depth -gt $depth)
			  { 
				  $depth	= $afile.depth
				  $deepest_firstIndex =  $iA
			  }
			  elseif ($afile.depth -lt $depth)
			  { break; }
		  }
		  $arTargdirSplit =  ($aFfiles[$deepest_firstIndex].pathAr | select -SkipLast 1 )
		  $arTargdir = $arTargdirSplit	-join "\";
		  $aFfilesTargFolder = ($aFfiles  | Where-Object { (Compare-Object -ReferenceObject ($_.pathAr |select -First ($depth-1) ) -DifferenceObject $arTargdirSplit -SyncWindow 0) -eq $null } )	# $_.path	-like  "$arTargdir\*"

		 $pretendent1 = @();

		do
		  {
		   foreach ( $mask in ("1_.pdf", "*.pdf" ))
		  {

			$pretendent1 = $aFfilesTargFolder | Where-Object { $_.pathAr[$_.pathAr.Count - 1] -like $mask}
		
		   if ( $pretendent1.Count -gt 0)  
			{ break; }

		}
			if ( $pretendent1.Count -gt 0)  
			{ break; }
	 	  foreach ( $mask in ( ("*.jpg","*.jpeg" ), @("*.pdf"), @("*.tif"), @("*.cdr") ,
			    @("telo..*.doc" , "telo..*.docx"), @("*.doc" , "*.docx")))
		  {

			$pretendent1 = $aFfilesTargFolder | Where-Object { 
				$_.pathAr[$_.pathAr.Count - 1] -like $mask[0] -and 
				( ($mask.Count -le 1 ) -or ($_.pathAr[$_.pathAr.Count - 1] -like $mask[1]) ) 
			}

		 }
		} while($false)
		  $logFile =  $MyInvocation.ScriptName + ".log"

	   if ( $pretendent1.Count -gt 0)  
		{ 
		$pretendent1	 
	

		
		}
		else
		  {


		 }

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
	 if (!$algAForB)
   { continue; }#TODO
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
