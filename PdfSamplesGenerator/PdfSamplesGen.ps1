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
#$ErrorActionPreference =  "Inquire" #
$DebugPreference = "Continue"
Set-StrictMode -Version 2.0
$scriptStartDate = Get-Date

function EchoA
{
  for ($i = 0; $i -lt $args.length; $i++)
  {
    "Arg $i is <$($args[$i])>"
  }
}

#trap
#{
#	"Hey trap Error: $_"
#	continue	
#}
$logFile = $null
 function WarnAndLog ($message)
{
    if ($logFile -eq $null)
    {
        Write-Debug "logFile null"   
    }
      Write-Warning $message
     "[$(get-date)] $message" >> $logFile
}
#TODO:
$waitPeriodMs = 200
function Wait-KeyPress2 ($keysToSkip)
{
  #Write-Host $prompt , $skipMessage	$prompt='Press "S" key to skip this',

  #$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
  $doSleep = $false;
  Write-Host -NoNewline "_"
  if ($Host.UI.RawUI.KeyAvailable)
  {
    Write-Host -NoNewline ","
    #Write-Debug "KeyPressed to stop"
    $key1 =  $host.UI.RawUI.ReadKey("AllowCtrlC,IncludeKeyDown,IncludeKeyUp")
    #$host.UI.RawUI.ReadKey("NoEcho") 
    #[console]::ReadKey("NoEcho,AllowCtrlC") 
    if ($key1 -eq $null -or (!$key1.KeyDown) -or @( $keysToSkip | % { [char]$_ }) -inotcontains $key1.Character) # -icontains $key1.Character
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
    Write-Debug " =$([int]$key1.Character)= $key1 =kd: $($key1.KeyDown)."

  }
  else
  { $doSleep = $true; }

  if ($doSleep)
  {
    Start-Sleep -Milliseconds $waitPeriodMs
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

function release-comobject ($ref)
{
  if ($ref -eq $null)
  {
    return
  }
  while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($ref) -gt 0) {}
  # [System.GC]::Collect()
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
    # echo "Convert $file error "
    # return $null
    Write-Output "Verb = print for : $file "
    $startInfo.Verb = "print"
    $startInfo.Arguments = $null
    #if ($procForFile.Verbs -notcontains "print")
    #{ return  "type notcontains 'print'"		   }

    #Start-Process –FilePath $file -Verb "print" -Wait -ErrorVariable err1

  }

  $process = New-Object System.Diagnostics.Process
  $process.StartInfo = $startInfo

  $errP = $null

  #Return Values:
  #true if a process resource is started; false if no new process resource is started (for example, if an existing process is reused).



  $newStarted = $null
  try
  {
    $newStarted = $process.Start()

    if (!$?)
    {
      $errP = $Error[0]
    }
  }
  catch
  {
    Write-Debug " process.Start  catched:$_ "
    $errP = $_
  }
  #$standardOut = $process.StandardOutput.ReadToEnd()
  #$process.WaitForExit()

  $process

  return $errP;

  # return $err1


  #if ($err1 -ne $null)
  #{
  #  return $false
  #}
  #return $true;

}

function print_to_usingDefault
{
  param([string]$file,[string]$printer)
  $null = (Get-WmiObject -ComputerName . -Class Win32_Printer -Filter "Name='$printer'").SetDefaultPrinter()
  printto $file
}

#$printto = "d:\INSTALL\!office\Bullzip\files\printto.exe"

$pdftk = Get-Command "pdftk" -ErrorAction SilentlyContinue
if ($pdftk -eq $null)
{ $pdftk = Get-Command "C:\Program Files (x86)\PDFtk\bin\pdftk.exe" }
$waterMPDF = "d:\!Work\Pdf_c\_Образец_ВодЗнак.pdf"

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
    $outTK = & $pdftk "$inFile" cat 1-8 output $outFileCut verbose dont_ask
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
      #-ErrorVariable $err1
      $pOutp = Invoke-Expression "& $printt2"
      #cmd /c $printt2
      if ($?)
      {
        $err1 = $null
      }
      else
      {
        $err1 = $Error[0]
        if ($err1 -eq $null)
        {
          $err1 = $pOutp
        }

      }
      Write-Debug "Print $file as $ftype1 $($err1 -eq $null)"
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


function PrintCorelDrawInternal ([string]$fileToPrint,[string]$printer,$cdraw)
{
  trap
  {
    Write-Debug "PrintCorelDrawInternal trap: $_"
    #TODO:
    if ([intptr]::Size * 8 -ne 64)
    {
      return "PrintCorelDrawInternal trap:32 $_ !!"
    }
    else
    {
      return "64!! PrintCorelDrawInternal trap: $_"
    }
    #return $true	
  }
  if ($cdraw -eq $null)
  {
    $cdraw = New-Object -Com CorelDRAW.Application
    $cdraw.Visible = $false
  }

  $cdDocToPrint = $cdraw.OpenDocument($fileToPrint) #$cdraw.OpenDocument($file.FullName) AsCopy AsCopy
  #$cdDocToPrint.SetDocVisible($false)

  $prs = $cdDocToPrint.PrintSettings
  #$prs|gm
  # $prs.Copies = 3
  $prs.PrintRange = 3 # 3 == PrnPrintRange VGCore.prnPageRange
  $prs.PageRange = "1-8"
  #$prs.Options.PrintJobInfo = $True

  if ($prs.Printer.Name -ne $printer)
  {
    for ($iPr = 0; $iPr -lt $cdraw.Printers.Count; $iPr++)
    {
      $pr2 = $null
      $pr2 = $cdraw.Printers($iPr)
      if ($pr2 -ne $null -and ($pr2.Name -eq $printer))
      {
        $prs.Printer = $pr2
        break;
      }

    }
  }

  if ($prs.Printer -eq $null)
  { return $true }
  if ($prs.Printer.Name -ne $printer)
  {
    return $true
  }

  if (!$prs.Printer.Ready())
  {
    return $true
  }


  #With .PostScript
  #.DownloadType1 = True
  #.Level = prnPSLevel3
  $cdDocToPrint.PrintOut()
  $cdDocToPrint.Close()

}
$cdraw = $null
    function SharpPrintCorelDraw ($using1, $cast)
    {

               $code = @"
				using System;
 $using1//using Corel.Interop.VGCore;

namespace CorelInterop1
{
    public class Class1
    {
        static public bool Print(string fileToPrint, string printer)
        {
             Type pia_type = Type.GetTypeFromProgID("CorelDRAW.Application");
            object cdraw1 = Activator.CreateInstance(pia_type);
            IVGApplication cdraw = cdraw1 as IVGApplication;//Application
            //var fileToPrint = @"d:\!Work\Pdf_c\Тестовый каталог\(011-1-1-48929)(А4).cdr";
            IVGDocument cdDocToPrint = cdraw.OpenDocument(fileToPrint, 0); //cdraw.OpenDocument(file.FullName) AsCopy AsCopy
            //cdDocToPrint.SetDocVisible(false)

            IPrnVBAPrintSettings prs = cdDocToPrint.PrintSettings;//
             prs.Copies = 1;
            prs.PrintRange = PrnPrintRange.prnPageRange;//Corel.Interop. 3 == PrnPrintRange VGCore.prnPageRange
            prs.PageRange = "1-8";
            if (prs.Printer.Name != printer)
            {
                for (int iPr = 0; iPr < cdraw.Printers.Count; iPr++)
                {
                    IPrnVBAPrinter pr2 = null;//Printer IPrnVBAPrinter
                    try
                    {
                        pr2 = cdraw.Printers[iPr];
                    }
                    catch (System.ArgumentException )
                    {

                        continue;
                    }
                    if (pr2 != null && (pr2.Name == printer))
                    { 
                        //prs.Printer = Convert.ChangeType(pr2, prs.Printer
                            prs.Printer = $cast pr2;//(Printer)
                       /* PropertyInfo propertyInfo = prs.GetType().GetProperty("Printer");
                        if (propertyInfo == null)
                        {
                            //using  ;
                            //VGCore.IPrnVBAPrintSettings.
                            //Type t = prs.Printer.GetType();
                        }
                        else
                            propertyInfo.SetValue(prs, pr2, null);
                        */
//Convert.ChangeType(pr2, propertyInfo.PropertyType)
                       // prs.Printer = Convert.ChangeType(pr2, prs.Printer.GetType());
                        //pr2;
                        break;
                    }

                }
            }

            if (prs.Printer == null)
            { return true; }
            if (prs.Printer.Name != printer)
            {
                return true;
            }

            if (!prs.Printer.Ready)
            {
                return true;
            }


            // #With .PostScript
            //#.DownloadType1 = True
            //#.Level = prnPSLevel3;
            cdDocToPrint.PrintOut();//Corel.Interop.VGCore.
            ((IVGDocument)cdDocToPrint).Close();
            return false;
        }
	}
}
"@
       return $code;
    }
function PrintCorelDraw ([string]$fileToPrint,[string]$printer)
{
  trap
  {
    Write-Debug "PrintCorelDraw trap: $_"
    return $true
  }
  if ($cdraw -eq $null)
  {
    $cdraw = New-Object -Com CorelDRAW.Application
    $cdraw.Visible = $false
  }
  $err = PrintCorelDrawInternal $fileToPrint $printer $cdraw
  if ($err -ne $null)
  {
    if ([intptr]::Size * 8 -ne 64)
    {
      #C:\Windows\SysNative\WindowsPowerShell\v1.0\PowerShell.exe -File $MyInvocation.MyCommand.Path -YourParam1 $YourParam1 -YourParam2 $YourParam2

    }
    else
    {
      $func = (Get-Command PrintCorelDrawInternal);
      $jobCD = Start-Job -ScriptBlock $func.ScriptBlock -ArgumentList @( $fileToPrint,$printer) -RunAs32
      Wait-Job $jobCD | Out-Null
      $_out1 = $null
      try {
        $_out1 = $jobCD.ChildJobs[0].Output[0]
      }
      catch
      {
        Write-Debug " p$jobCD.ChildJobs[0].Output[0]  catched:$_ ||| $_out1"
        $_out1 = $_
      }

      if ($_out1 -eq $null -and $jobCD.State -eq "Completed")
      {
        return $_out1
      }
      else
      {
        
 
        $vgcore = [System.Reflection.Assembly]::LoadWithPartialName("Corel.Interop.VGCore")
        # $comcore = [System.Reflection.Assembly]::LoadWithPartialName("VGCoreAuto.tlb")

        $usingVGCore = "using Corel.Interop.VGCore;"#
         $usingVGCoreCom = "VGCore"
          $printerCast =  "(Printer)";
          $code2 = SharpPrintCorelDraw  $usingVGCore $printerCast  
         $addErr = $null
        try
        {
          Add-Type -ErrorVariable addErr  -ReferencedAssemblies @( $vgcore) -TypeDefinition $code2 -Language CSharp | Out-Null
        }
        catch
        {
          if ($_.FullyQualifiedErrorId -eq "TYPE_ALREADY_EXISTS,Microsoft.PowerShell.Commands.AddTypeCommand")
            {
              Write-Debug " Add-Type  already:$_ "
                $addErr = $null;}
         else { Write-Debug " Add-Type  catched:$_ " }
        }
        if ($addErr -ne $null)
         {      
            $addErr = $null
            $code3  = SharpPrintCorelDraw $usingVGCore ""
            try
            {
              Add-Type -ErrorVariable addErr -ReferencedAssemblies @( $vgcore) -TypeDefinition $code3 -Language CSharp | Out-Null
            }
            catch
            {
              if ($_.FullyQualifiedErrorId -eq "TYPE_ALREADY_EXISTS,Microsoft.PowerShell.Commands.AddTypeCommand")
            {    Write-Debug " Add-Type  already:$_ "
               $addErr = $null;}
              else { Write-Debug " Add-Type  catched:$_ " }
             }
         }

        $err = [CorelInterop1.Class1]::Print($fileToPrint,$printer)

        if ($err -eq $false)
        {
          $err = $null

        }


      }
    }
  }
  return $err
}
function TestFileWritable ($file1)
{
  if (!(Test-Path $file1))
  {
    return $false
  }

  try {
    [IO.File]::OpenWrite($file1).Close();
    $true
  }
  catch {
    $false }
}
function PrinInDesignInternal ([string]$fileToPrint,[string]$printer,$InDesign1)
{

    trap
    {
           if ($InDesign1 -ne $null)
           {   $InDesign1.Quit(0x6e6f2020)  }
        return @("InDesign.Quit PrinInDesignInternal error trap :  $_" , $_)

    }
    Write-Warning  "PrinInDesignInternal"

  if ($InDesign1 -eq $null)
  {
    $InDesign1 = New-Object -Com InDesign.Application

  }
  if ($InDesign1 -eq $null)
  {
 
    return $false
  }
  <#$Awindows =  $InDesign.ActiveWindow 
	#$Awindows.Minimize()
	$windows1 =  $InDesign.Windows
	$window1 = $windows1.FirstItem()
	if ($window1 -ne $null)
	{
		$window1.Minimize()
	}#>
  #idDefault = 0x44666c74,
  #idOpenOriginal = 0x4f704f72,
  #idOpenCopy = 0x4f704370
  $indd_doc = $InDesign1.Open($fileToPrint,$false,0x4f704370)

    if ($indd_doc -eq $null)
    {
    #       typedef [helpstring("Options for saving a document before closing or quitting.")]
    #enum {
    #    idNo = 0x6e6f2020,
    #    idAsk = 0x61736b20,
    #    idYes = 0x79657320
    #} idSaveOptions;

       $InDesign.Quit(0x6e6f2020)
        return "indq"
    }

  # $links1 = $indd_doc.Links
  #$indd_doc.SetDocVisible($false)"InDesign.idOpenOptions.idOpenCopy"
  #$PRINTERNAMe = "Bullzip PDF Printer"
  #  $printPresetName = $printer.Replace(" ","_") + "8";
  #  $printPreset = $InDesign.PrinterPresets.Item($printPresetName)
  #  if ($printPreset -eq $null)
  #  {
  #	  $printPreset = $InDesign.PrinterPresets.Add($printPresetName)
  #$printPreset.Printer = $printer
  #$printPreset.Sequence = "1-8"
  #	  }

  $printPref = $indd_doc.PrintPreferences
  #$printPref.PageRange = [System.Runtime.InteropServices.BStrWrapper]"1-9"
  $printPref.Printer = $printer
  $indd_doc.PrintOut($false)
   # BackgroundTask

 #   $taskCount = $InDesign1.BackgroundTasks.Count
	#if ($taskCount -eq 0) {
 #   	Write-Host("There is no BackgroundTask task.")
 #   }
 #   Else
 #   {
	#	$str = ""
	#	For ($i = 1 ; $i -lt $taskCount; ++$i)
	#	{	
 #           $myIdleTask = $InDesign1.BackgroundTasks.Item($i)
	#		$str = $str +  "BackgroundTask   $($myIdleTask.id):  $($myIdleTask.name) \r\n"
 #         }
		
	#	Write-Host($str)
	#}

  #Return value: A list of task states for task that finished. Type: Array of idTaskState enumerators
  $taskStates = $InDesign1.WaitForAllTasks()
  #Constant idAsk = 1634954016 (&H61736b20)
  #Constant idNo = 1852776480 (&H6e6f2020)
  #    Default member of InDesign.idSaveOptions
  #    Does not save changes.
  $indd_doc.Close(1852776480)


}
  $InDesign = $null

 function PrinInDesign ([string]$fileToPrint,[string]$printer)
{
  if ($InDesign -eq $null)
  {
  #  $InDesign = New-Object -Com InDesign.Application

  }
  if ($InDesign -eq $null)
  {
 
   # return $false
  }

     $func = (Get-Command  PrinInDesignInternal);

    for ($tryI = 2; $tryI -gt 0; --$tryI)
    {
      $jobCD = Start-Job -ScriptBlock $func.ScriptBlock -ArgumentList @( $fileToPrint,$printer, $InDesign)
    #TODO timeout  
        $tmout =  60
      Wait-Job $jobCD  -Timeout $tmout | Out-Host
    if ($jobCD.State -ne "Completed")
    {
        $nottryAgain = $jobCD.State -eq "Running"
         if ($nottryAgain)
        {
            Write-Warning "maybe InDesign document wrong links"
             $tryI = 0 
        }
         Write-Warning "timeout -> Stop InDesign"
        Stop-Job $jobCD  | Out-Host
        try {
            Get-Process  -Name "InDesign" | foreach {
               Write-Warning "fail stop : $_ "
                Stop-Process $_ -Force   

            }
       # Write-Warning $killRes
            }
        catch
        {
           Write-Debug " stop ind : $_"     
      }
            continue
        #return $true;          
     }
      $_out1 = $null
      try {
        $_out1 = $jobCD.ChildJobs[0].Output[0]
      }
      catch
      {
        Write-Debug "p  $jobCD.ChildJobs[0].Output[0]  catched:$_ ||| $_out1"
        $_out1 = $_
      }

      if ($_out1 -eq $null -and $jobCD.State -eq "Completed")
      {
          #success
        return $_out1
      }
    else
    { 

        continue    

     }
  }
    return $true
}
  $samplesTargetDirName = "Образцы"

#directory to process
function Print1 ($file,[string]$obrazcyParentDir, [string]$targetName)
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

  #echo "obrazcyParentDir = $obrazcyParentDir"
  $sampleSuffix = "_образец"
  $watermarkText = "OBRAZEC" # "образец"

  $samplesTarget = Join-Path -Path $obrazcyParentDir -ChildPath $samplesTargetDirName
  $sampleFileName = $null
    
   if (  $targetName -ne $null)
  {
    $sampleFileName = $targetName + $sampleSuffix

  }
  else {

  $sampleFileName = $file.basename + $sampleSuffix
  }

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
  if ($file.basename -like "*!NO")
  {
      #TODO
      Write-Debug "Suspicioue $file"
      #msgBoxRetryCancel ("Wy $($file.FullName)?")
  }

  if ($checkExistance)
  {
    $iOF = 1;
    while (Test-Path $outFile)
    {
      $outFile = ("$($outFileS)_0" + $iOF.ToString() + ".pdf")
      ++ $iOF;
    }
  }

  # Out-File "$settings" -Append -Encoding "unicode" -InputObject 	"$watermarkText"

  #$outFileItem = Get-Item $outFile	-ErrorAction SilentlyContinue
  #if ($outFileItem -ne $null -and ( (New-TimeSpan -start $outFileItem.LastWriteTime -End $scriptStartDate) -gt 0)) # 4 version Test-Path -Path $outFile -OlderThan $scriptStartDate)
  if (Test-Path $outFile)
  {
    Remove-Item $outFile -Force
  }

  $outFileCut = $null
  if ($file.Extension -eq ".pdf")
  {

    $outFileCut = Join-Path $obrazcyParentDir ("$sampleFileName" + ".8cut.pdf") #($outFile


    $cutRes = Cut_PdfTo8 $file.FullName $outFileCut
    if (($cutRes -lt 0) -and (Test-Path $outFileCut))
    {
      Remove-Item $outFile -Force -ErrorAction SilentlyContinue;

      # not work & $pdftk $outFileCut multistamp "`"$waterMPDF`"" output $outFile
      #  Move-Item $outFileCut $outFile -Force


    }
    elseif ($cutRes -eq 0)
    {

      WarnAndLog "  Cannot convert $($file.FullName).  Cannot cut tmp file $outFile"

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
  watermarkfontsize=180
  watermarkrotation=c2c
  watermarkcolor=#DFDFDF
  watermarkfontname=arial.ttf
  watermarkoutlinewidth=2
  watermarklayer=top
  watermarkverticalposition=center
  watermarkhorizontalposition=center
  confirmoverwrite=no
  autorotatepages=none
  showprogressfinished=yes
"@



  if ((".jpg",".jpeg") -inotcontains $file.Extension) #расширения без фона пропускаем
  {
    if (Test-Path $waterMPDF)
    {
      $watermarkSuperimposelayer = ""

      $watermarkSuperimposelayer = "bottom"


      #pdf со слоеми не прозрачный, походу из за огранечений как не векторный. 
      #$watermarkSuperimposelayer = "top"

      #  professional version -  superimposeresolution=vector
      Out-File "$settings" -Append -Encoding "unicode" -InputObject @"
  superimpose=$waterMPDF
  superimposeresolution=
  superimposelayer=$watermarkSuperimposelayer
"@
    }
  }



  <#
	
	-dPDFSETTINGS=/screen   (screen-view-only quality, 72 dpi images)
-dPDFSETTINGS=/ebook    (low quality, 150 dpi images)
-dPDFSETTINGS=/printer  (high quality, 300 dpi images)
-dPDFSETTINGS=/prepress (high quality, color preserving, 300 dpi imgs)
-dPDFSETTINGS=/default  (almost identical to /screen)
	
	 target=ebook - pdf qulity.

	#>
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

  #trap
  #{
  #  Out-Host "Something wrong $_"
  #  continue
  #}

  echo "Основной этап конвертации `"$($file.FullName)`" "

  $errP = $false
  [System.Diagnostics.Process]$ptProc = $null
  if ($file.Extension -eq ".cdr")
  {
    $errP = PrintCorelDraw $fileToPrint $PRINTERNAME
  }
  elseif ($file.Extension -eq ".indd")
  {
    $errP = PrinInDesign $fileToPrint $PRINTERNAME

    if ($errP -ne $null)
    {
      WarnAndLog " InDesign проблемы"
      Write-Debug "PrinInDesign fail : $errP "
    }
  }
  if ($errP -ne $null)
  {
    [System.Diagnostics.Process]$ptProc0,$errP0 = printto "`"$fileToPrint`"" "`"$PRINTERNAME`""
    if ($errP -ne $false -and $errP0 -eq $null -and $ptProc0 -ne $null -and $ptProc0.HasExited -eq $null)
    {
      Write-Debug "Failed second try"
    }
    else
    {
      $ptProc = $ptProc0
      $errP = $errP0
    }
  }

  #$ptSuccess  Silently-ErrorVariable ProcessError -ErrorAction Continue 
  #$ptProcErrOut = $ptProc.StandardError.ReadToEnd()
  $endStat = $null
  if ($ptProc -ne $null)
  {
    $endStat = $ptProc.HasExited
  }
  #if ($ptErr -ne $null) {

  #  $errPrint = ("`"$($file.FullName)`"" + " не имеет печатающей программы :" + $ptErr.ToString())
  # $printto.exe "in\example.rtf" "$PRINTERNAME"
  #  $ptERRORLEVEL = $lastexitcode
  # if ($ptERRORLEVEL -eq 0) $res11 = & $pdftk "$outFile" dump_data | Select-string -Pattern "PageMediaNumber: ([0-9]*)"

  for ([int]$iW = 0; $true;++ $iW)
  {
    if (TestFileWritable ($outFile)) #Test-Path
    {

      if ($outFileCut -eq $null)
      {
        $outFileCut = ("$outFileS" + ".8cut.pdf") 

        $cutRes = Cut_PdfTo8 $outFile $outFileCut
        if (($cutRes -lt 0) -and (Test-Path $outFileCut))
        {
          Remove-Item $outFile -Force;
          Move-Item $outFileCut $outFile -Force
        }
        elseif ($cutRes -eq 0)
        {
          WarnAndLog "  Cannot convert $($file.FullName).  Cannot cut tmp file $outFile"

          Remove-Item $outFile -Force -ErrorAction SilentlyContinue;

        }
        $outFileCut = $null
      } #outFileCut

      if ($iW -gt 0)
      {
        Write-Host "Конвертация файла `"$($file.FullName)`" завершилась"
      }
      break;
    }
    elseif ($ptProc -ne $null)
    {
      if ($ptProc.HasExited -and $ptProc.ExitCode -eq 1)
      {
        #1 -, но печатает, такое поведение не очень понятно почему, например у AcrobReader
        #- возможно. для прог, которые остаются открытыми.
        #continue
      }
      elseif ($ptProc.HasExited -and $ptProc.ExitCode -eq 8985 -and $file.Extension -eq ".rtf")
      {
        #непонятно
        #continue
      }
      elseif ($ptProc.HasExited -and $ptProc.ExitCode -ge 2 -and ($iW -gt (100000 / $waitPeriodMs)))
      {
        WarnAndLog ("`"$($file.FullName)`"" + " не конвертируется. Errcode = " + $ptProc.ExitCode)
        break;

      }
      if (($iW -lt 10 -or (($iW % 10) -eq 0)))
      {
        if (($ptProc -ne $null) -and $ptProc.HasExited -eq $null)
        {
          Write-Debug "($ptProc).HasExited -eq null N $iW"
        }
        else
        {
          Write-Debug "($ptProc) N $iW"

        }
      }
      #if ($iW -gt (100000 / $waitPeriodMs) -and  $ptProc.HasExited -eq $null) #TODO
      #{
      #	$errP = "не удалось напечатать"
      #}
    } # ptProc
    if (($iW -lt 10 -or (($iW % 10) -eq 0)))
    {
      Write-Host -NoNewline "."
      # Write-Debug -no "($ptProc).HasExited -eq null N $iW"

    }
    if ($errP -ne $null)
    {

      $errP2 = $null;
      try
      {
        if ($errP.Exception.InnerException.NativeErrorCode -eq 1155)
        {

          $errP2 = PrintByRegCommand "`"$($file.FullName)`"" "`"$PRINTERNAME`""
          if ($errP2 -eq $null)
          {
            $errP = $null
            continue
          }
        }
      }
      catch
      {
        Write-Debug "non exception errP: $errP	, $_"
      }

      WarnAndLog ("`"$($file.FullName)`"" + " не конвертируется, возможно не имеет печатающей программы :" + $errP.ToString())

      break;

    }
    else
    {

      $keysToSkip = 's'
      if ($iW -eq 0)
      {
        Start-Sleep -Milliseconds $(10 * $waitPeriodMs)
        continue;
      }
      elseif ($iW -eq 1)
      {
        Write-Host "Конвертация файла `"$($file.FullName)`" затянулась. Нажмите `"S`" чтобы пропустить его"
        Start-Sleep -Milliseconds $waitPeriodMs
        continue

      }
      elseif ($iW -gt 90 -and ($file.Extension -eq ".indd")) #TODO
      {
        WarnAndLog "Конвертация файла `"$($file.FullName)`" прервана, т.к. затянулась"
        Remove-Item $outFile -Force -ErrorAction SilentlyContinue;
        break;

      } 
      elseif ($iW -gt 9900) #TODO
      {
        WarnAndLog "Конвертация файла `"$($file.FullName)`" прервана, т.к. затянулась"
        Remove-Item $outFile -Force -ErrorAction SilentlyContinue;
        break;

      }
      $StopPressed = Wait-KeyPress2 $keysToSkip
      if ($StopPressed)
      { #$ptProc.
        WarnAndLog "Конвертация файла `"$($file.FullName)`" пропущена по запросу пользователя"
        Remove-Item $outFile -Force -ErrorAction SilentlyContinue;
        break;
      }
    }

  } # for 
  Write-Host "После  файла `"$($file.FullName)`" "
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


#not in PS2 !! echo $MyInvocation.PSCommandPath

if ($args.Count -gt 0 -and $args[0].length -ge 0)
{
  $targetP = $args[0]
}
else
{
  #$MyInvocation.
  $par2 = Split-Path -Parent $MyInvocation.MyCommand.Path
  $targetP = $par2
  if ($targetP -eq $null)
  {
    $targetP = (Get-Item $PSCommandPath).Directory

  }
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
  echo "Скрипт запущен для $targetP"
function SetLog
 {   
     $thisScr =  Get-Item $MyInvocation.ScriptName
     #(Join-Path $targetP $samplesTargetDirName)
       $Script:logFile = Join-Path $targetP ($thisScr.Name + ".log")
   Write-Output "Log file (not created if no error): `"$logFile`"  "
  }
  SetLog
function ExtractSpecified
{
  param($value,$wildCardFArray)

  $TMPfullP = $value.FullName + "ext"
 try {
  Remove-Item $TMPfullP -Force -Recurse | Out-Null
     }
    catch{
    Write-Debug "remove old arch $_"      
    }
     
     $out7z =  $null
  if ($wildCardFArray.Count -gt 1)
  {
    $TMPfiltFile = "ExtList.extList" #Get-Item ($value.Directory.ToString()+  [System.IO.Path]::GetTempFileName()
    $oldWD = Get-Location
    cd $value.Directory
    Out-File $TMPfiltFile -Force -Encoding "utf8" -InputObject (($wildCardFArray | % { "*" + $_ }) -join "`n")

    # Write-Debug  - set 	$DebugPreference = "Continue" 
    trap
    {
      Write-Debug  " u7z 1 trapped:$_ "
      continue
    }
    $out7z = & $u7z "e" "$($value.Name)" "-o$TMPfullP" "-i@$TMPfiltFile" "-y"
     if (!$?)
       {     
             Remove-Item $TMPfullP -Force -Recurse
          $TMPfullP = $null
        }
     Remove-Item $TMPfiltFile -Force
    cd $oldWD
  }
  else
  {
    $oldWD = Get-Location
    cd $value.Directory
    $wd = $wildCardFArray[0]
    trap
    {
      Write-Debug " u7z 2 trapped:$_ "
      continue
    }
    $out7z =  & $u7z "e" $value.Name "-o$TMPfullP" "-i!$wd" "-y"
          if (!$?)
       {     
             Remove-Item $TMPfullP -Force -Recurse
          $TMPfullP = $null
        }
    cd $oldWD
  }
  Write-Debug " u7z = $out7z"

  return $TMPfullP;

}

function Algs ([string]$targetP1,[boolean]$algAForB,$obrazcyParentDir, [string]$targetName)
{
  #$logFile = ((Get-Item $MyInvocation.ScriptName).Directory).FullName + ".log"
   # WarnAndLog "hey"
     
  [boolean]$algAOnly = $algAForB

  $cont1 = @( Get-ChildItem $targetP1 )




  $ArchsExts = ($archs | ForEach-Object { $_[0] })
  #$supportedDocs = "rtf", "cdr", "jpg", "tif", "doc", "docx, "indd", "pdf"
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
      Write-Output " Обработка архива $($value.FullName)"
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
      Write-debug "aafiles[0].pathAr = $($aafiles[0].pathAr)" 
 
      if ($TRUE)
      { # $aafiles |  Measure-Object -Property depth  -
        $aFfiles = @($aafiles | 
            Where-Object { 
                !$_.isdir -and ($_.depth -gt 0)
            } )
        # 	| Sort -Property depth -Descending

        #первая, если вторая будет глубже - не подходит
        [int]$depth = 0
        $deepest_firstIndex = 0
        for ($iA = 0; $iA -lt $aafiles.Count;++ $iA)
        {
          $afile = $aafiles[$iA];
            $fd = $afile.depth
            if ($afile.isdir)
            { $fd++ } 
          if ($fd -gt $depth)
          {
            $depth = $fd
            $deepest_firstIndex = $iA
          }
          elseif ($fd -lt $depth)
          { break; }
        }
        $aTrgetFPath = @( $aafiles[$deepest_firstIndex].pathAr)
          if ($aafiles[$deepest_firstIndex].isdir)
          {  $aTrgetFPath =  $aTrgetFPath + @(".") }
        if ($aTrgetFPath.Count -lt ($depth ))
        {
          Write-Debug "pathAr strange"
          $aafiles[$deepest_firstIndex].pathAr = $aafiles[$deepest_firstIndex].Path.Split('\')
          $aTrgetFPath = $aafiles[$deepest_firstIndex].pathAr

        }
        $arTargdirSplit = ($aTrgetFPath | select -First ($aTrgetFPath.Count + (-1)))
        $arTargdir = $null
        if ($arTargdirSplit -eq $null)
          {  $arTargdir = "\" }
        else
          { $arTargdir = $arTargdirSplit -join "\"; }
        $aFfilesTargFolder = ($aFfiles | Where-Object {
           ( ($arTargdirSplit -eq $null) -and ($_.depth -eq 1) ) -or
             (Compare-Object -ReferenceObject ($_.pathAr | select -First ($depth + (-1))) -DifferenceObject $arTargdirSplit -SyncWindow 0) -eq $null 
        }) # $_.path	-like  "$arTargdir\*"
        # @($aFfiles[$deepest_firstIndex])


        #currently  One file
        $pretendent1 = @();

        do
        {
          foreach ($mask in ("1_.pdf","*.pdf"))
          {

            $pretendent1 = @( $aFfilesTargFolder    |
                 Where-Object {
                      $_.pathAr[$_.pathAr.Count + (-1)] -like $mask 
                 })

            if ($pretendent1.Count -gt 0)
            { break; }

          }
          if ($pretendent1.Count -gt 0)
          { break; }
          foreach ($mask in (("*.jpg","*.jpeg"),@( "*.pdf"),@( "*.tif"),@( "*.cdr"),
              @( "telo..*.doc","telo..*.docx"),@( "*.doc","*.docx")))
          {

            $pretendent1 = @( $aFfilesTargFolder | Where-Object {
                $_.pathAr[$_.pathAr.Count + (-1)] -like $mask[0] -or
                (($mask.Count -gt 1) -and
                  ($_.pathAr[$_.pathAr.Count + (- 1)] -like $mask[1]))
              })

            if ($pretendent1.Count -gt 0)
            { break; }
          }
        } while ($false)


        if ($pretendent1.Count -gt 0)
        {
          $wildCardFArray2 = @( $pretendent1[0].Path) # $pretendent1 |% { $_.Path } 	 

          $TMPfullP = ExtractSpecified $value $wildCardFArray2
                   Write-debug   "TMPfullP = $TMPfullP "
         $extracteFolder = Get-Item $TMPfullP
          if ($extracteFolder -ne $null)
          {

            Algs $extracteFolder $true $value.Directory $value.BaseName
          }
          Remove-Item $TMPfullP -Force -Recurse | Out-Null

        }
        else
        {
          WarnAndLog " архив `"$($value.Fullname)`" не содержит искомых файлов"

        }

      } #TRUE

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
    elseif ($docExtensions -contains ".$fExt")
    {
      #if (!$algAForB)
      #{ continue; } 
      Print1 $value $obrazcyParentDir $targetName
      #break;
    }
    else 
      { Write-Debug "Skipped = $value" }
  }


}


# $IMag = New-Object -ComObject "ImageMagickObject.MagickImage.1"
# $msgs = $IMag.Convert "logo:" -format "%m,%h,%w" info: 
# $msgs = $IMag.Convert("logo:","-format","%m,%h,%w","info:")	   $targetP\$($pd1)C.pdf
#			$pd1 =  "cc.pdf" # "ТЕОРИЯ АВТОМАТИЧЕСКОГО УПРАВЛЕНИЯ ДЛЯ «ЧАЙНИКОВ» tau_dummy.pdf"
#$IMag.Convert( "$targetP\$pd1[0-7]" , "-delete 8--1")

# $Printers = Get-WmiObject -Class Win32_Printer $Printers|where { $_.Default }
$initiallyDefaultPrinter = Get-WmiObject -ComputerName . -Class Win32_Printer -Filter "Default=True"

Algs $targetP $false $null $null


<#
$defprinter = (Get-WmiObject -ComputerName . -Class Win32_Printer -Filter "Default=True").Name
$null = (Get-WmiObject -ComputerName . -Class Win32_Printer -Filter "Name='My Desired Printer'").SetDefaultPrinter()
get-childitem "\\nas\directory" | % { Start-Process -FilePath $_.VersionInfo.FileName –Verb Print -PassThru }
$null = (Get-WmiObject -ComputerName . -Class Win32_Printer -Filter "Name='$defprinter'").SetDefaultPrinter()
#>
if ($initiallyDefaultPrinter -ne $null)
{
  $initiallyDefaultPrinter.SetDefaultPrinter() | Out-Null
}
echo "Обработка `"$targetP`" завершена. Скрипт завершён."
release-comobject $cdraw
release-comobject $InDesign
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
