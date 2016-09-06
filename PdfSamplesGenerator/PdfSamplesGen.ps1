#
# Script.ps1
#


#Parameters : this.ps1 [TargetLocation] 
# otherwise   - uses Current script file location

#Add-Type -assemblyname  Bullzip.PdfWriter
#$env:CommonProgramFiles 
#$bres= Add-Type -Path "$env:CommonProgramFiles\Bullzip\PDF Printer\API\Microsoft.NET\Framework\v4.0\Bullzip.PDFWriter.dll" -PassThru
##[Bullzip.PdfWriter.PdfInternal]
#$gh1 = New-Object  Bullzip.PdfWriter.ComPdfInternal
#$gh1.pdftk()
##$gh1 = New-Object  PdfWriter.PdfInternal.Ghostscript
#$ErrorActionPreference =  Inquire #"SilentlyContinue" 

#TODO
Set-StrictMode -Version 2.0

$waterMPDF = "d:\!Work\Pdf_c\_Образец_ВодЗнак.pdf"
function EchoA
{
  for ($i = 0; $i -lt $args.length; $i++)
  {
    "Arg $i is <$($args[$i])>"
  }
}

$logFile = $null

function Wait-KeyPress2 ($keysToSkip)
{
  #Write-Host $prompt , $skipMessage	$prompt='Press "S" key to skip this',

  #$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
  $doSleep = $false;
  if ($Host.UI.RawUI.KeyAvailable)
  {
    $key1 = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    if (@($keysToSkip| % {[char]$_}) -inotcontains $key1.Character) # -icontains $key1.Character
    {
		if ($key1.Character -ne 0)
      {
		  $Host.UI.RawUI.FlushInputBuffer()
	  }
      $doSleep = $true;
    }
	  else
	  {
		  Write-Debug "KeyPressed to stop" 
	  }
	  
		Write-Debug "KeyPressed withresult doSleep $doSleep :"
		Write-Debug $key1
	
  }
  else
  { $doSleep = $true; }

  if ($doSleep)
  {
    Start-Sleep -Milliseconds 100
    return $false;
  }
  else
  {
    $Host.UI.RawUI.FlushInputBuffer()
    return $true;
  }
  #do {

  #} until 
  #		[console]::KeyAvailable

  #[console]::ReadKey("NoEcho,IncludeKeyDown")
  # $Host.UI.RawUI.KeyAvailable

}

function printto
{
  param([string]$file,[string]$printer)
  #$err1;

  [System.Diagnostics.ProcessStartInfo]$startInfo = New-Object System.Diagnostics.ProcessStartInfo #-Args $file #System.Diagnostics.ProcessStartInfo


  $startInfo.FileName = $file
  $startInfo.Arguments = @( $printer)

  #$startInfo.RedirectStandardOutput = $true
  $startInfo.CreateNoWindow = $false
  $startInfo.UseShellExecute = $true
  #$startInfo.Username = "DOMAIN\Username"
  #$startInfo.Password = $password
  if ($printer -ne $null)
  {
    #  if ($procForFile.Verbs -notcontains "printto")
    #{    return  "type notcontains 'printto'"		 }

    $startInfo.Verb = "printto"
    #Start-Process –FilePath $file -ArgumentList $printer -Verb "printto" -Wait -ErrorVariable err1
  }
  else
  {
    echo "Convert $file error "
    return $null
    #  " $startInfo.Verb =  "print" 
    #if ($procForFile.Verbs -notcontains "print")
    #{ return  "type notcontains 'print'"		   }

    #Start-Process –FilePath $file -Verb "print" -Wait -ErrorVariable err1

  }

  $process = New-Object System.Diagnostics.Process
  $process.StartInfo = $startInfo

  $errP = $null
  $process.Start() | Write-Debug
  #$standardOut = $process.StandardOutput.ReadToEnd()
  #$process.WaitForExit()

  if (!$?)
  {
    $errP = $Error[0]
  }


  $process



  return $errP;

  # return $err1


  #if ($err1 -ne $null)
  #{
  #  return $false
  #}
  #return $true;

}
#$printto = "d:\INSTALL\!office\Bullzip\files\printto.exe"

$pdftk = Get-Command "pdftk" -ErrorAction SilentlyContinue
if ($pdftk -eq $null)
{ $pdftk = Get-Command "C:\Program Files (x86)\PDFtk\bin\pdftk.exe" }

if (!(Test-Path $waterMPDF))
{

  $waterMPDF = Split-Path -Parent $pdftk.Path
  $waterMPDF = Split-Path -Parent $waterMPDF
  $waterMPDF = Join-Path $waterMPDF "Образец_ВодЗнак.pdf"
}
function Get-PdfNumOfPages ([string]$outFile)
{
  $dumpData1 = & $pdftk "$outFile" dump_data
  if (!$?)
  { return 0; }
  $res11 = $dumpData1 | Select-String -Pattern "NumberOfPages: ([0-9]+)" #"PageMediaNumber: ([0-9]+)"

  $numOfPages = ($res11.Matches[0].Groups[1].value)
  return [int]$numOfPages
}

function Cut_PdfTo8 ($inFile,$outFileCut)
{
  $numOfPages = Get-PdfNumOfPages $inFile
  [int]$numOfPagesN = [int]$numOfPages
  if ($numOfPagesN -gt 8)
  {

    # Write-Debug  $DebugPreference = "Continue" 
     $outTK =  & $pdftk "$inFile" cat 1-8 output $outFileCut verbose dont_ask 
    if ($?)
    { return (8 - $numOfPagesN) }
	  else
	  { Write-Warning $outTK }
  }
  elseif ($numOfPagesN -ge 1)
  {
    return $numOfPagesN;
  }
  Write-Warning "Cut Pdf To 8 pages something wrong"
  return $numOfPagesN
}


#TODO reorder
$u7z = Get-Command "7z" -ErrorAction SilentlyContinue
#if ($u7z -eq $null)
#{ }
if ($u7z -eq $null)
{
  $reg1 = Get-Item -Path Registry::HKEY_CURRENT_USER\SOFTWARE\7-Zip
  if ($reg1 -ne $null)
  { $u7z = Get-Command ($reg1.Getvalue("Path") + "\7z.exe") }

}
if ($u7z -eq $null)
{ $u7z = Get-Command "C:\Program Files (x86)\Universal Extractor\bin\7z.exe" }
if ($u7z -eq $null)
{ Write-Warning "No 7z!!!" }

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


function PrintByRegCommand ([string]$file,[string]$printer)
{
  $err1 = $true

  $ftype1 = $null
  if ($file -like "*.jpg" -or $file -like "*.jpeg" -or $file -like "`"*.jpg`"" -or $file -like "`"*.jpeg`"")
  {
    $ftype1 = "jpegfile"
  }
  elseif ($file -like "*.tif" -or $file -like "*.tiff" -or $file -like "`"*.tif`"" -or $file -like "`"*.tiff`"")
  {
    $ftype1 = "TIFImage.Document"
  }
  if ($ftype1 -ne $null)
  {
    $reg1 = Get-Item -Path Registry::HKEY_CLASSES_ROOT\$ftype1\shell\printto\Command
    if ($reg1 -ne $null)
    { $printt2 = $reg1.Getvalue("").replace("`"%1`"",$file).replace("`"%2`"",$printer).replace("%3","").replace("%4","")
		$errF = $null
      #	 Start-Process  $printt2	-ErrorVariable $errF
      Invoke-Expression "& $printt2" -ErrorVariable $err1
		#cmd /c $printt2
	  
	  #$ptERRORLEVEL = $null
   #   $ptERRORLEVEL = $lastexitcode
   #   if ($ptERRORLEVEL -ne 0)
   #   { $err1 = $ptERRORLEVEL }
   #   else
   #   {
   #     $err1 = $null
   #   }
    }
  }
  return $err1;
}
#directory to process
function Print1 ($file,[string]$obrazcyParentDir)
{

  $checkExistance = $false;
  if ($obrazcyParentDir -eq $null -or ($obrazcyParentDir.length -eq 0))
  {
    $obrazcyParentDir = $file.Directory
  }
  else
  {
    $checkExistance = $true
  }
  #TODO
  #echo "obrazcyParentDir = $obrazcyParentDir"
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
  if ($checkExistance)
  {
    $iOF = 1;
    while (Test-Path $outFile)
    {
      $outFile = ("$outFileS(" + $iOF.ToString() + ").pdf")
      ++ $iOF;
    }
  }

  # Out-File "$settings" -Append -Encoding "unicode" -InputObject 	"$watermarkText"
  if (Test-Path $outFile)
  {
    Remove-Item -Force $outFile
  }

  $outFileCut = $null
  if ($file.Extension -eq ".pdf")
  {

    $outFileCut = ("$outFileS" + ".8cut.pdf") #($outFile


    $cutRes = Cut_PdfTo8 $file.FullName $outFileCut
    if (($cutRes -lt 0) -and (Test-Path $outFileCut))
    {
      Remove-Item $outFile -Force -ErrorAction SilentlyContinue;

      # not work & $pdftk $outFileCut multistamp "`"$waterMPDF`"" output $outFile
      #  Move-Item $outFileCut $outFile -Force


    }
    elseif ($cutRes -eq 0)
    {

      $em3 = "  Cannot convert $($file.FullName).  Cannot cut tmp file $outFile"

      Write-Warning $em3
      "[$(get-date)] $em3" >> $logFile
      Remove-Item $outFile -Force -ErrorAction SilentlyContinue;

      return
    }
    else
    {
      $outFileCut = $null;
    }
  }
  #------------------------------

  # Set environment variables used by the batch file

  $PRINTERNAMe = "Bullzip PDF Printer"
  # $PRINTERNAMe
  # PDF Writer - bioPDF

  # Create settings \ runonce.ini
  # $LAPP=$env:LOCALAPPDATA
  $LAPP = $env:APPDATA
  $SF1 = "settings.ini"

  if ($LAPP.length -eq 0)
  {
    $LAPP = "$env:USERPROFILE\Local Settings\Application Data"
  }
  $settingsDir = "$LAPP\PDF Writer\$PRINTERNAME"
  $settings = "$settingsDir\$SF1"
  #  ECHO $settings
  $settFile = $null
  $settingsBackFile = $null
  if (Test-Path "$settings")
  {
    $settFile = (Get-Item $settings)
    $settingsBackFileName = Join-Path ($settFile.Directory) ($SF1 + ".back")
    $settingsBackFile = $settingsBackFileName | Get-Item -ErrorAction SilentlyContinue
    if ($settingsBackFile -ne $null) # (Test-Path $settingsBackFileName)
    {
      Remove-Item $settingsBackFile -Force -ErrorAction SilentlyContinue
    }
    Move-Item $settFile.FullName $settingsBackFileName -Force
    # Get-Item $settingsBackFile
    # rename-item $settFile $settingsBackFileName -Force

    #  $newSett = New-Item $(Join-Path $settFile.Directory ($settFile.name + ".new")) # $newSett = New-Item $(Join-Path $settFile.Directory ($settFile.name + ".new"))
    # $newSett.Replace(($settFile.FullName) ,(Join-Path  $settFile.Directory   $settingsBackFileName ) ,($true) )

  }
  else
  {
    $settingsBackFileName = Join-Path $settingsDir ($SF1 + ".back")

  }
  #(rename "$settings" "$SF1.back")

  # %CD%\out\demo.pdf
  #ECHO "Save settings to \" $settings\""
  #ECHO "[PDF Printer]" > "$settings"
  #ECHO "output=$outFile" 			 >> "$settings"
  #ECHO author=PdfSamplesGenerator 	>> "$settings"
  #ECHO showsettings=never 			>> "$settings"
  #ECHO showpdf=no >> "$settings"

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
  autorotatepages=none
  showprogressfinished=yes
"@

  if (Test-Path $waterMPDF)
  {

    Out-File "$settings" -Append -Encoding "unicode" -InputObject @"
  superimpose=$waterMPDF
  superimposeresolution=
  superimposelayer=bottom
"@
  }
  # TODO: showprogress=yes
  #
  # confirmnewfolder=yes
  <# 
	suppresserrors=no
    rememberlastfoldername=yes
	rememberlastfilename=no
  openfolder=no
  showsaveas=nofile
  autorotatepages=none
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

  $fileToPrint = $null
  if ($outFileCut -eq $null)
  {
    $fileToPrint = $($file.FullName)
  }
  else
  {
    $fileToPrint = $($outFileCut)
  }
  echo "Основной этап конвертации `"$($file.FullName)`" "
  [System.Diagnostics.Process]$ptProc,$errP = printto "`"$fileToPrint`"" "`"$PRINTERNAME`""

  #$ptSuccess  Silently-ErrorVariable ProcessError -ErrorAction Continue 
  #$ptProcErrOut = $ptProc.StandardError.ReadToEnd()
  $endStat = $ptProc.HasExited

  #if ($ptErr -ne $null) {

  #  $errPrint = ("`"$($file.FullName)`"" + " не имеет печатающей программы :" + $ptErr.ToString())
  #  Write-Warning $errPrint
  #  "[$(get-date)] $errPrint" >> $logFile
  # }
  # else

  # $printto.exe "in\example.rtf" "$PRINTERNAME"
  #  $ptERRORLEVEL = $lastexitcode
  # if ($ptERRORLEVEL -eq 0) $res11 = & $pdftk "$outFile" dump_data | Select-string -Pattern "PageMediaNumber: ([0-9]*)"

  for ([int]$iW = 0; $true;++ $iW)
  {
    if (Test-Path ($outFile))
    {
      if ($outFileCut -eq $null)
      {
        $outFileCut = ("$outFileS" + ".pdf8cut") #($outFile


        $cutRes = Cut_PdfTo8 $outFile $outFileCut
        if (($cutRes -lt 0) -and (Test-Path $outFileCut))
        {
          Remove-Item $outFile -Force;
          Move-Item $outFileCut $outFile -Force
        }
        elseif ($cutRes -eq 0)
        {
          $em3 = "  Cannot convert $($file.FullName).  Cannot cut tmp file $outFile"

          Write-Warning $em3
          "[$(get-date)] $em3" >> $logFile
          Remove-Item $outFile -Force -ErrorAction SilentlyContinue;

        }
        $outFileCut = $null
      }

      if ($iW -gt 0)
      {
        Write-Host "Конвертация файла `"$($file.FullName)`" завершилась"
      }
      break;
    }
    elseif ($ptProc.HasExited -and $ptProc.ExitCode -eq 1)
    {
      #не очень понятно почему, - возможно. для прог, которые остаются закрытыми.
      continue
    }
    elseif ($ptProc.HasExited -and $ptProc.ExitCode -ge 2)
    {
      $errPrint = ("`"$($file.FullName)`"" + " не конвертируется ")
      Write-Warning $errPrint
      "[$(get-date)] $errPrint" >> $logFile
      break;

    }
    elseif ($errP -ne $null)
    {

      $errP2 = $null;
      if ($errP.Exception.InnerException.NativeErrorCode -eq 1155)
      {

        $errP2 = PrintByRegCommand "`"$($file.FullName)`"" "`"$PRINTERNAME`""
        if ($errP2 -eq $null)
        { continue }
      }

      $errPrint = ("`"$($file.FullName)`"" + " не конвертируется, возможно не имеет печатающей программы :" + $errP.ToString())
      Write-Warning $errPrint
      "[$(get-date)] $errPrint" >> $logFile
      break;

    }

    else
    {
      $keysToSkip = 's'
      if ($iW -eq 0)
      {
        Start-Sleep -Milliseconds 1000
        continue;
      }
      elseif ($iW -eq 1)
      {
        Write-Host "Конвертация файла `"$($file.FullName)`" затянулась. Нажмите `"S`" чтобы пропустить его"
        Start-Sleep -Milliseconds 100
        continue

      }
	  $StopPressed=	Wait-KeyPress2 $keysToSkip
      if ($StopPressed)
      {
        Write-Host "Конвертация файла `"$($file.FullName)`" пропущена"
        break;
      }
    }

  } # for 
  #else 
  if ($outFileCut -ne $null)
  {
    Remove-Item $outFileCut -Force;
  }

  if ($settingsBackFile -ne $null -and $settingsBackFile.Exists) #(Test-Path "$settings.back")
  {
    Remove-Item -Force $settings
    Move-Item -Force $settingsBackFile.FullName $SF1 -ErrorAction SilentlyContinue
  }
  elseif (Test-Path $settingsBackFileName)
  {
    Remove-Item -Force $settings
    Move-Item -Force $settingsBackFileName $SF1 -ErrorAction SilentlyContinue
  }


} #Print1

$docExtensions1 = ".rtf",".cdr",".jpg",".tif",".tiff",".doc",".docx",".indd"
$docExtensions = $docExtensions1 + ".pdf"


$archs =

(("zip"),"ZIP"),
(("rar"),"RAR"),
(("7z"),"7z"),
(("arj"),"ARJ"),
(("gz","gzip","tgz"),"GZIP"),
(("bz2","bzip2","tbz2","tbz"),"BZIP2")
#,
#(("xz"),"XZ"),
#(("tar"),"TAR"),
#(("wim","swm"),"WIM"),
#(("lzma"),"LZMA"),
#(("cab"),"CAB"),
#(("z","taz"),"Z"),
#(("cpio"),"CPIO")

function AlgA_Iter
{
  param($value,$obrazcyParentDir)
  if ($docExtensions1 -contains $value.Extension)
  {

    Print1 $value $obrazcyParentDir
  }
  elseif ($value.Extension -eq ".pdf")
  {
    Print1 $value $obrazcyParentDir
  }

  echo $value.FullName | Write-Debug
}

#not in PS2 !! echo $MyInvocation.PSCommandPath

if ($args.Count -gt 0 -and $args[0].length -ge 0)
{
  $targetP = $args[0]
}
else
{
  #$MyInvocation.
  #TODO
  $par2 = Split-Path -Parent $MyInvocation.MyCommand.Path
  $targetP = $par2
  if ($targetP -eq $null)
  {
    $targetP = (Get-Item $PSCommandPath).Directory

  }

  echo "Скрипт запущен для $targetP"
  while ($targetP -eq $null -or !(Test-Path $targetP))
  {
    echo "Не получилось найти $targetP . Введите вручную"
    $targetP = Read-Host -Prompt "Не получилось найти $targetP . Введите вручную"
    if ($targetP -eq $null)
    {
      echo "Не получилось найти $targetP . Выходим"
      return
    }
  }
  # Get-Location 
}


function ExtractSpecified
{
  param($value,$wildCardFArray)

  $TMPfullP = $value.FullName + "ext"
  if ($wildCardFArray.Count -gt 1)
  {
    $TMPfiltFile = "ExtList.extList" #Get-Item ($value.Directory.ToString()+  [System.IO.Path]::GetTempFileName()
    $oldWD = Get-Location
    cd $value.Directory
    Out-File $TMPfiltFile -Force -Encoding "utf8" -InputObject (($wildCardFArray | % { "*" + $_ }) -join "`n")

    # Write-Debug  - set 	$DebugPreference = "Continue" 
    & $u7z "e" $value.name "-o$TMPfullP" "-i@$TMPfiltFile" "-y" | Write-Debug
    Remove-Item $TMPfiltFile -Force
    cd $oldWD
  }
  else
  {
    $oldWD = Get-Location
    cd $value.Directory
    $wd = $wildCardFArray[0]
    & $u7z "e" $value.name "-o$TMPfullP" "-i!$wd" "-y" | Write-Debug
    cd $oldWD
  }

  return $TMPfullP;

}

function Algs ([string]$targetP1,[boolean]$algAForB,$obrazcyParentDir)
{
  $logFile = ((Get-Item $MyInvocation.ScriptName).Directory).FullName + ".log"


  [boolean]$algAOnly = $algAForB

  $cont1 = Get-ChildItem $targetP1




  $ArchsExts = ($archs | ForEach-Object { $_[0] })

  #$nonArch = $cont1| Where-Object { '.zip','.rar' -notcontains  $_.Extension}#,'.config'


  # foreach ($value in $cont1) 
  for ($iF = 0; $iF -lt $cont1.Count; $iF++)
  {

    $value = $cont1[$iF]

    if (!$algAForB)
    {
      Write-Progress -Activity “Generating samples” -Status “ file $value” `
         -PercentComplete ($iF / $cont1.Count * 100)
    }
    # $cont1 | Select name

    if ($value.Attributes -band [System.IO.FileAttributes]::Directory)
    {
      continue;
    }
    #  Write-Host $value.FullName
    # Write-Host $value
    $fExt = $value.Extension
    if ($fExt -ne $null)
    { $fExt = $fExt.TrimStart('.') }
    if ($ArchsExts -contains $fExt)
    {

      #algB
      if ($algAOnly)
      {
        continue;
      }
      $archContT = & $u7z "l" "-slt" $value.FullName

      $aafiles = { @() }.Invoke(); #System.Collections.ObjectModel.Collection`1[System.Management.Automation.PSObject]

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
                  isdir = $false;
                  pathAr = $pathAr1;
                  depth = $pathAr1.Count }
              }
              else
              {
                $fileL.Path = $archParse.Matches[0].Groups[1].value #never!
              }
            }
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
      $aafiles[0].pathAr | Out-Host
      # если папок нет
      if (@( $aafiles | Where-Object { $_.isdir }).length -eq 0)
      {

        #[System.Reflection.Assembly]::LoadWithPartialName("System.IO.Path")
        $wildCardFArray = $docExtensions
        $TMPfullP = ExtractSpecified $value $wildCardFArray
        Algs (Get-Item $TMPfullP) $true $value.Directory
        Remove-Item $TMPfullP -Force -Recurse

      }
      else
      { # $aafiles |  Measure-Object -Property depth  -
        $aFfiles = $aafiles | Where-Object { !$_.isdir -and ($_.depth -gt 1) }
        # 	| Sort -Property depth -Descending

        #первая, если вторая будет глубже - не подходит
        [int]$depth = 1
        $deepest_firstIndex = 0
        for ($iA = 0; $iA -lt $aFfiles.Count;++ $iA)
        {
          $afile = $aFfiles[$iA];
          if ($afile.depth -gt $depth)
          {
            $depth = $afile.depth
            $deepest_firstIndex = $iA
          }
          elseif ($afile.depth -lt $depth)
          { break; }
        }
        $arTargdirSplit = ($aFfiles[$deepest_firstIndex].pathAr | select -SkipLast 1)
        $arTargdir = $arTargdirSplit -join "\";
        $aFfilesTargFolder = ($aFfiles | Where-Object { (Compare-Object -ReferenceObject ($_.pathAr | select -First ($depth + (-1))) -DifferenceObject $arTargdirSplit -SyncWindow 0) -eq $null }) # $_.path	-like  "$arTargdir\*"
        # @($aFfiles[$deepest_firstIndex])

        #currently  One file
        $pretendent1 = @();

        do
        {
          foreach ($mask in ("1_.pdf","*.pdf"))
          {

            $pretendent1 = $aFfilesTargFolder | Where-Object { $_.pathAr[$_.pathAr.Count + (-1)] -like $mask }

            if ($pretendent1.Count -gt 0)
            { break; }

          }
          if ($pretendent1.Count -gt 0)
          { break; }
          foreach ($mask in (("*.jpg","*.jpeg"),@( "*.pdf"),@( "*.tif"),@( "*.cdr"),
              @( "telo..*.doc","telo..*.docx"),@( "*.doc","*.docx")))
          {

            $pretendent1 = $aFfilesTargFolder | Where-Object {
              $_.pathAr[$_.pathAr.Count +(-1)] -like $mask[0] -or
              (($mask.Count -le 1) -or
                ($_.pathAr[$_.pathAr.Count + (- 1)] -like $mask[1]))
            }

            if ($pretendent1.Count -gt 0)
            { break; }
          }
        } while ($false)


        if ($pretendent1.Count -gt 0)
        {
          $wildCardFArray2 = @( $pretendent1[0].Path) # $pretendent1 |% { $_.Path } 	 

          $TMPfullP = ExtractSpecified $value $wildCardFArray2
          Algs (Get-Item $TMPfullP) $true $value.Directory
          Remove-Item $TMPfullP -Force -Recurse

        }
        else
        {
          "[$(get-date)] архив `"$($value.Fullname)`" не содержит искомых файлов" >> $logFile

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
      #if (!$algAForB)
      #{ continue; } 
      AlgA_Iter $value $obrazcyParentDir
      #break;
    }
  }


}

# $IMag = New-Object -ComObject "ImageMagickObject.MagickImage.1"
# $msgs = $IMag.Convert "logo:" -format "%m,%h,%w" info: 
# $msgs = $IMag.Convert("logo:","-format","%m,%h,%w","info:")	   $targetP\$($pd1)C.pdf
#			$pd1 =  "cc.pdf" # "ТЕОРИЯ АВТОМАТИЧЕСКОГО УПРАВЛЕНИЯ ДЛЯ «ЧАЙНИКОВ» tau_dummy.pdf"
#$IMag.Convert( "$targetP\$pd1[0-7]" , "-delete 8--1")


Algs $targetP $false $null

echo "Обработка $targetP завершена. Скрипт завершён."
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
