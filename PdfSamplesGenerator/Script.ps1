#
# Script.ps1
#
if ($args.Count -gt 0 -and $args[1].LENGTH -andgt 0)
{  $targetP =  $args[1]
}
else
{ $targetP = Get-Location }

$cont1 = Get-ChildItem $targetP 
$Arch = $cont1| Where Extension -in '.zip','.rar'
#$nonArch = $cont1| Where-Object {$_.Extension -notin '.zip','.rar','.config' }
 if ($Arch.Length -eq 0 )
 {

  foreach ($value in $nonArch){
 
   if ($value.Extension -eq "docx")
   {
   "Wow"
    break; 
    }
  Write-Host $value
	  }
}
else
{

}

#Foreach-Object {
#    $content = Get-Content $_.FullName

#    #filter and save content to the original file
#    $content | Where-Object {$_ -match 'step[49]'} | Set-Content $_.FullName

#    #filter and save content to a new file 
#    $content | Where-Object {$_ -match 'step[49]'} | Set-Content ($_.BaseName + '_out.log')
#}