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

$logFile = $null

function Wait-KeyPress2 ($keysToSkip)
{
  #Write-Host $prompt , $skipMessage	$prompt='Press "S" key to skip this',

  #$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
  $doSleep = $false;
  if ($Host.UI.RawUI.KeyAvailable)
  {

    if ($host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") -notin $keysToSkip)
    { $Host.UI.RawUI.FlushInputBuffer()
      $doSleep = $true;
    }
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

  [System.Diagnostics.ProcessStartInfo]$procForFile = New-Object System.Diagnostics.ProcessStartInfo -Args $file #System.Diagnostics.ProcessStartInfo

  if ($printer -ne $null)
  {
    #  if ($procForFile.Verbs -notcontains "printto")
    #{    return  "type notcontains 'printto'"		 }

    Start-Process –FilePath $file -ArgumentList $printer -Verb "printto" -Wait -ErrorVariable err1
  }
  else
  {
    #if ($procForFile.Verbs -notcontains "print")
    #{ return  "type notcontains 'print'"		   }

    Start-Process –FilePath $file -Verb "print" -Wait -ErrorVariable err1

  }
  if ($err1 -ne $null)
  {
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

        #	 Start-Process  $printt2	-ErrorVariable err1
        cmd /c $printt2

        $ptERRORLEVEL = $lastexitcode
        if ($ptERRORLEVEL -ne 0)
        { $err1 = $ptERRORLEVEL }
        else
        {
          $err1 = $null
        }
      }
    }
  }


  return $err1


  #if ($err1 -ne $null)
  #{
  #  return $false
  #}
  #return $true;

}
$printto = "d:\INSTALL\!office\Bullzip\files\printto.exe"

$pdftk = Get-Command "pdftk" -ErrorAction SilentlyContinue
if ($pdftk -eq $null)
{ $pdftk = "C:\Program Files (x86)\PDFtk\bin\pdftk.exe" }

function Get-PdfNumOfPages([string]$outFile)
{
	  $dumpData1 = 	& $pdftk "$outFile" dump_data 
      $res11 = $dumpData1 | Select-String -Pattern "NumberOfPages: ([0-9]+)" #"PageMediaNumber: ([0-9]+)"

      $numOfPages = ($res11.Matches[0].Groups[1].value)
	return  [int]$numOfPages  
 }

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
function Print1 ($file,[string]$obrazcyParentDir)
{
  if ($obrazcyParentDir -eq $null)
  {
    $obrazcyParentDir = $file.Directory
  }
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
  $settings = "$LAPP\PDF Writer\$PRINTERNAME\$SF1"
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
    $settingsBackFileName = Join-Path $settFile.Directory ($SF1 + ".back" )

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
  if (Test-Path $outFile)
  {
    Remove-Item -Force $outFile
  }
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

  if ($settingsBackFile -ne $null -and $settingsBackFile.Exists) #(Test-Path "$settings.back")
  {
    Remove-Item -Force $settings
    Move-Item -Force $settingsBackFile.FullName $SF1
  }
  elseif (Test-Path $settingsBackFileName)
  {
    Remove-Item -Force $settfile
    Move-Item -Force $settingsBackFileName $settfile.name
  }


  if ($ptErr -ne $null) {

    $errPrint = ("`"$($file.FullName)`"" + " не имеет печатающей программы :" + $ptErr.ToString())
    Write-Warning $errPrint
    "[$(get-date)] $errPrint" >> $logFile
    return;
  }


  # $printto.exe "in\example.rtf" "$PRINTERNAME"
  #  $ptERRORLEVEL = $lastexitcode
  # if ($ptERRORLEVEL -eq 0) $res11 = & $pdftk "$outFile" dump_data | Select-string -Pattern "PageMediaNumber: ([0-9]*)"

  for ([int]$iW = 0; $true; ++$iW)
  {
    if (Test-Path ($outFile))
    {
		$numOfPages = Get-PdfNumOfPages $outFile
		   if ([int]$numOfPages -gt 8)
      {

        #TODO: cut pdf first
        $outFileCut = ("$samplesTarget\$sampleFileName" + ".pdf8cut") #($outFile
					# Write-Debug  $DebugPreference = "Continue" 
        & $pdftk "$outFile" cat 1-8 output $outFileCut verbose dont_ask | Write-Debug
        if (Test-Path $outFileCut)
        {
          Remove-Item $outFile -Force;
          Move-Item $outFileCut $outFile -Force
        }
      }
      if ($iW -gt 0)
      {
        Write-Host "Конвертация файла `"$($file.FullName)`" завершилась"
      }
      break;
    }
    else
    {
      $keysToSkip = 's'
      if ($iW -eq 0)
      {
        Write-Host "Конвертация файла `"$($file.FullName)`" затянулась. Нажмите `"S`" чтобы пропустить его"
        Start-Sleep -Milliseconds 100
        continue

      }
      if (Wait-KeyPress2 $keysToSkip)
      {
        Write-Host "Конвертация файла `"$($file.FullName)`" пропущена"
        break;
      }
    }

  } # for 


}
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
  if ($value.Extension -in $docExtensions1)
  {

    Print1 $value $obrazcyParentDir
  }
  elseif ($value.Extension -eq ".pdf")
  {
    Print1 $value $obrazcyParentDir
  }

  $value.FullName | Write-Debug
}

if ($args.Count -gt 0 -and $args[0].length -ge 0)
{
  $targetP = $args[0]
}
else
{ $targetP = Get-Location }


function ExtractSpecified
{
  param($value,$wildCardFArray)

  $TMPfullP = $value.FullName + "ext"
  if ($wildCardFArray.Count -gt 1)
  {
    $TMPfiltFile = "ExtList.extList" #Get-Item ($value.Directory.ToString()+  [System.IO.Path]::GetTempFileName()
    $oldWD = Get-Location
    cd $value.Directory
    Out-File $TMPfiltFile -Encoding "utf8" -InputObject (($wildCardFArray | % { "*" + $_ }) -join "`n")

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

  #$nonArch = $cont1| Where-Object {$_.Extension -notin '.zip','.rar' }#,'.config'


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
    if ($fExt -in $ArchsExts)
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
      $aafiles[0].pathAr
      # если папок нет
      if (($aafiles | Where-Object { $_.isdir }).length -eq 0)
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
        $aFfilesTargFolder = ($aFfiles | Where-Object { (Compare-Object -ReferenceObject ($_.pathAr | select -First ($depth - 1)) -DifferenceObject $arTargdirSplit -SyncWindow 0) -eq $null }) # $_.path	-like  "$arTargdir\*"
        # @($aFfiles[$deepest_firstIndex])

        #currently  One file
        $pretendent1 = @();

        do
        {
          foreach ($mask in ("1_.pdf","*.pdf"))
          {

            $pretendent1 = $aFfilesTargFolder | Where-Object { $_.pathAr[$_.pathAr.Count - 1] -like $mask }

            if ($pretendent1.Count -gt 0)
            { break; }

          }
          if ($pretendent1.Count -gt 0)
          { break; }
          foreach ($mask in (("*.jpg","*.jpeg"),@( "*.pdf"),@( "*.tif"),@( "*.cdr"),
              @( "telo..*.doc","telo..*.docx"),@( "*.doc","*.docx")))
          {

            $pretendent1 = $aFfilesTargFolder | Where-Object {
              $_.pathAr[$_.pathAr.Count - 1] -like $mask[0] -or
              (($mask.Count -le 1) -or
                ($_.pathAr[$_.pathAr.Count - 1] -like $mask[1]))
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
      if (!$algAForB)
      { continue; } #TODO
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
