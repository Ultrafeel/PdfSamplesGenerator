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
#$cdDocToPrint.PrintOut()
$cdDocToPrint.Close
}
function Invoke-Method0 {
  param(
    [__ComObject] $object,
    [String] $methodName,
    $methodParameters
  )
  $output = $object.GetType().InvokeMember($methodName,"InvokeMethod",$NULL,$object,$methodParameters)
  if ( $output ) { $output }
}
function Invoke-Method2 {
  param(
    [__ComObject] $object,
    [String] $methodName,
    $methodParameters
  )
  $output = $object.GetType().InvokeMember($methodName,[System.Reflection.BindingFlags]::InvokeMethod,
	  $NULL,$object,$methodParameters)
  if ( $output ) { $output }
}



$cdr_doc =  $cdraw.OpenDocument($file.FullName)  #$cdraw.OpenDocument($file.FullName) AsCopy AsCopy
$cdr_doc.SetDocVisible($false)
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
	#Read-Host -Prompt "wowo"
#	$watLayer.GetType().InvokeMember("Import", [System.Reflection.BindingFlags]::InvokeMethod,
#    $null,  ## Binder
#    $watLayer,  ## Target
#    ([Object[]]@($watermForCorel)),  ## Args
#    $null,  ## Modifiers
#    $null,  ## Culture
#    ([String[]]$NamedParameters)  ## NamedParameters
#)
	 [System.Reflection.Assembly]::LoadWithPartialName("System.Runtime.InteropServices")
	[System.Runtime.InteropServices.BStrWrapper]$ww2= ($watermForCorel)
	[System.String]$ww4 =  $watermForCorel
	Invoke-Method2 $watLayer "Import" $ww4
  $watLayer.Import($ww4 )
}


<#I have managed to get this working using the InvokeMember method of System.__ComObject. In order to pass multiple parameters to the method, simply enclose them in parentheses.

An example of the line of code is shown here:

PS C:> $usercontacts=[System.__ComObject].InvokeMember("GetSharedDefaultFolder" [System.Reflection.BindingFlags]::InvokeMethod,$null,$mapi,($user,10))

$user is the recipient object previously set up. $mapi is the MAPI namespace object (also set up previously).

#>
# Layers
# $cdraw.
PublishCorelDraw $cdr_doc ($file.FullName + ".pdf")
$cdraw.Quit()
#$cdrawScript = New-Object -Com  CorelDRAW.CorelScript
#$cdraw.Optimization
##$cdrawScript.FileOpen($file.FullName)
#$cdrawScript |gm
$cdr_doc