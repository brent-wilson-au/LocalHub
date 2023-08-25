$SearchPath = 'C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Outlook'
$SearchFilter = 'Microsoft.Office.Interop.Outlook.dll'
$PathToAssembly = Get-ChildItem -LiteralPath $SearchPath -Filter $SearchFilter -Recurse |
    Select-Object -ExpandProperty FullName -Last 1

if ($PathToAssembly) {
    Add-Type -LiteralPath $PathToAssembly
}
else {
    throw "Could not find '$SearchFilter'"
}
    #Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
Add-type -AssemblyName "System.Runtime.Interopservices"

$Now = Get-Date
Write-Host $now.ToString("u") "Starting ..."

try
{
$outlook = [Runtime.Interopservices.Marshal]::GetActiveObject('Outlook.Application')
    $outlookWasAlreadyRunning = $true
}
catch
{
    try
    {
        $Outlook = New-Object -comobject Outlook.Application
        $outlookWasAlreadyRunning = $false
    }
    catch
    {
        write-host "You must exit Outlook first."
        exit
    }
}
$namespace = $Outlook.GetNameSpace("MAPI")
$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)
$emails = $inbox.Items 

#$tstemails = $inbox.Items | Select-Object SenderEmailAddress -Unique
#write-host "tst " + $tstemails.Length

#The $emails variable will contain a list of just those emails that were sent from SSRS.  Now, for each of those, we can apply some processing.  For instance:
# Process the reports by recipient

$totalEmails = $emails.Count
write-host "Total emails in Inbox:" $totalEmails
$uniqueEmails = @("a")
#$uniqueEmails.Add("a")
$recCount = 0

$startTime = Get-Date

ForEach ($bemail in $emails)
{
#    write-host $uniqueEmails.Length
  
    if ($uniqueEmails.Contains($bemail.SenderEmailAddress))  
    { 
 #       Write-Host "Skipped"          
    }
    else {
     
        $uniqueEmails +=$bemail.SenderEmailAddress
  #      write-host "Added"
    
    }
    $recCount ++
   if (($recCount % 1000) -eq 0)
    {
#        Write-Host -NoNewLine "`r$a% complete"
        $Now = Get-Date
        Write-Host -NoNewline "`r"$now.ToString("u") " Emails processed:" $recCount  " Array size:" $uniqueEmails.Length
    }
#     write-host $recCount
#    write-host $bemail.SenderEmailAddress
#    write-host $bemail.ReceivedTime
    
    
}
# Close outlook if it wasn't opened before running this script
if ($outlookWasAlreadyRunning -eq $false)
{
    Get-Process "*outlook*" | Stop-Process -Force
}


$Now = Get-Date
write-host " "
write-host "Started Reading " $startTime
write-host "Ended   Reading " $Now
Write-Host $now.ToString("u") "Sorting"
$uniqueEmails = $uniqueEmails| Sort-Object
$Now = Get-Date
Write-Host $now.ToString("u") "Finished Sorting"

$Now = Get-Date
Write-Host $now.ToString("u") "Writing Output"
$fName = (Get-Date -Format "yyyy-MM-dd-HH-mm-ss").tostring() + ".txt"
Foreach ($uniqueOne in $uniqueEmails)
{
    if ($uniqueOne -ne "a")
    {
        $splitName = $uniqueOne.split("@")
        Add-Content -Path $fName -Value ($uniqueOne + "," + $splitName[0] + "," + $splitName[1])
    }

}
# $uniqueEmails | Out-File -Append $fName
$Now = Get-Date
Write-Host $now.ToString("u") "Finished Writing"

#write-host "***** Starting "
#foreach ($sender in $uniqueEmails)
#{
 #   Write-Host $sender
#}