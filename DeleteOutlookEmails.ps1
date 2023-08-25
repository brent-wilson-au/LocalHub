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

#Add-Type -assembly "Microsoft.Office.Interop.Outlook"
#add-type -assembly "System.Runtime.Interopservices"
Add-Type -AssemblyName System.Windows.Forms

write-host "Starting"
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

# Get the filename of addresses to remove
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('MyDocuments') }
$FileBrowser.Title = "Please select file"
#$FileBrowser.Filter = “Text files (*.txt)|*.txt”
$null = $FileBrowser.ShowDialog()
write-host "Selected File - " $FileBrowser.filename    
$addressList = Get-content $FileBrowser.filename

write-host "Read in "  $addressList.count " email addresses to remove emails"

$namespace = $Outlook.GetNameSpace("MAPI")
$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)
$deletedItems = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderDeletedItems)
$emails = $inbox.Items 

#$tstemails = $inbox.Items | Select-Object SenderEmailAddress -Unique
#write-host "tst " + $tstemails.Length

#The $emails variable will contain a list of just those emails that were sent from SSRS.  Now, for each of those, we can apply some processing.  For instance:
# Process the reports by recipient


write-host "There are" $emails.count "emails in the Inbox"
#$uniqueEmails = @("a")
#$uniqueEmails.Add("a")
$recCount = 0

$inboxCount = $inbox.Items.Count - 1


for ($recNum = $inboxCount; $recNum -ge 1; $recNum--) 
{
# Walk backwards from the bottom of the email list
# If we find a match, then delete it
# NOTE: We are starting at the bottom of the list as a delete will shuffle all the items below up one spot and we could miss one
#   write-host "Found " $inbox.Items.Item($recNum).SenderEmailAddress " " $recNum

    $oneSender = $inbox.Items.Item($recNum).SenderEmailAddress

    if ($addresslist.Contains($oneSender)) 
    {
    
        Write-host "found record for " $inbox.Items.Item($recNum).SenderEmailAddress "at item " $recnum
        $recCount ++

    }

#    Write-Host "Elapsed Time After IF Item: $($elapsed.Elapsed.ToString())"
    if (($recNum % 100) -eq 0)
    {
#        Write-Host -NoNewLine "`r$a% complete"
        $Now = Get-Date
#        Write-Host -NoNewline "`r"$now.ToString("u") " Emails processed:" $recNum  " Emails Found:" $recCount
        Write-Host $now.ToString("u") " Emails processed:" $recNum  " Emails Found:" $recCount
        #       Write-Host $now.ToString("u") " Emails processed:" $recCount  " Array size:" $uniqueEmails.Length    
    }
}


#ForEach ($Fromaddress in $addressList)
#{
# Get all emails for that sender
#    Write-Host "looking for " $Fromaddress
#    $emailsFromSender = $inbox.Items | where-object { $_.sender -eq $Fromaddress }
#    write-host "Processing " $emailsFromSender.Count " emails for " $Fromaddress

#for ($c = $mailboxFolders.Items.count; $c -ge 1;$c--) {
#    [void]$inbox.Folders['test'].Items[$c].Move($moveTarget)    
#write-host $uniqueEmails.Length
#    Write-Host "Moving emails to DeletedItems for email address " $Fromaddress
    
    
#    if ($uniqueEmails.Contains($bemail.SenderEmailAddress))  
#    { 
 #       Write-Host "Skipped"          
#    }
#    else {
     
#        $uniqueEmails +=$bemail.SenderEmailAddress
  #      write-host "Added"
    
 #   }
 #  $recCount ++
 #  if (($recCount % 5000) -eq 0)
 #   {
 #       $Now = Get-Date
 #       Write-Host $now.ToString("u") $recCount
 #       Write-Host "Array size " $uniqueEmails.Length
 #   }
#     write-host $recCount
#    write-host $bemail.SenderEmailAddress
#    write-host $bemail.ReceivedTime
    
    
#}
#}
# Close outlook if it wasn't opened before running this script
if ($outlookWasAlreadyRunning -eq $false)
{
    Get-Process "*outlook*" | Stop-Process -Force
}

$Now = Get-Date
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
$uniqueEmails | Out-File -Append $fName
$Now = Get-Date
Write-Host $now.ToString("u") "Finished Writing"

#write-host "***** Starting "
#foreach ($sender in $uniqueEmails)
#{
 #   Write-Host $sender
#}