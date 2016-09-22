#
# Script1Develop.ps1
#
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

$files = $targetP | Get-ChildItem -Filter "*.cdr"
$file = $files[0]
$cdraw = New-Object -Com  CorelDRAW.Application

function PublishCorelDraw($cdr_doc, [string]$outFile)
{

 

$cdDocToPrint = $cdr_doc  #$cdraw.ActiveDocument
#PDFSettings 
$pdfSett = $cdDocToPrint.PDFSettings

$pdfSett.PublishRange =  3  # 3 == VGCore.pdfPageRange  "VGCore.pdfExportRange.pdfPageRange"
#$Error[0].Exception.HResult
$pdfSett.PageRange = "1-9"
#$pdfSett.ShowDialog() 
$cdDocToPrint.PublishToPDF($outFile)

#$prs  = $cdDocToPrint.PrintSettings
#$prs|gm
#$prs.Copies = 3
#$prs.PrintRange = 3 # 3 == PrnPrintRange VGCore.prnPageRange
#$prs.PageRange = "1-9"
#$prs.Options.PrintJobInfo = True
##With .PostScript
##.DownloadType1 = True
##.Level = prnPSLevel3

# $cdDocToPrint.O
#$cdDocToPrint.PrintOut
$cdDocToPrint.Close
}


$cdr_doc =  $cdraw.OpenDocument($file.FullName)  #$cdraw.OpenDocument($file.FullName) AsCopy AsCopy
echo $cdr_doc.Pages.Count
for ($i = 0; $i -lt 9 -and ($i -lt $cdr_doc.Pages.Count); ++$i )
{

  $page = $cdr_doc.Pages[$i]
  echo $page.Layers.Count
   $page.Layers[1]|gm|Out-Host

  $laFrs, $laLast =  $page.Layers[1] , $page.Layers[$page.Layers.Count] 
  echo $laFrs.Name , $laLast.Name
  $watLayer =  $page.CreateLayer("Watermark")
  $watermForCorel  = 'd:\!Work\Pdf_c\Waterm.pdf'
	$enc = [system.Text.Encoding]::GetEncoding(1252)
$consumerkey ="xvz1evFS4wEEPTGEFPHBog"
$encconsumerkey= $enc.GetBytes($watermForCorel)
  #$watermForCorel.   # 0x80131501
  $watLayer.Import($encconsumerkey) 
}
# Layers
# $cdraw.
PublishCorelDraw $cdr_doc ($file.FullName + ".pdf")
$cdraw.Quit()
#$cdrawScript = New-Object -Com  CorelDRAW.CorelScript
#$cdraw.Optimization
##$cdrawScript.FileOpen($file.FullName)
#$cdrawScript |gm
$cdr_doc