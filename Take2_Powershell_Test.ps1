
###########################################################################
###########   Take2 PowerShell Test Script v1.0 by Abdul-Rehman Karim #####
###########################################################################

# !!!!!!!! PLEASE ENSIRE YOU RUN THE WHOLE SCRIPT AT ONCE AND THE USERS.CSV FILE IS PLACED IN C:\TEMP

cls
write-host "`nHi!`n`nPlease esnure Users.csv is placed in the C:\temp folder in order for this script to work.`n" -ForegroundColor Cyan
read-host -Prompt "Press enter to continue"

try {

$csv = Import-Csv C:\Temp\Users.csv -ErrorAction Stop

}

    Catch { Write-Host "`nSorry, we had trouble importing the CSV, please ensure it is named Users.csv and stored in C:\temp.`n`nExiting script...`n" -ForegroundColor white -BackgroundColor red; sleep 2; exit }


# Question 1
$csv | select -first 5 | ft -wrap
write-host "Please see the answer for question 1 above. We have imported the CSV into a variable as requested and displayed only the first 5 rows for aesthetic purposes to the screen.`n" -ForegroundColor green

# Question 2
read-host -Prompt "Press enter to continue"
Write-Host "`nFor question 2, there are a total of"$csv.count"users in the imported list.`n" -ForegroundColor green

# Question 3
read-host -Prompt "Press enter to continue"
[int]$allmailboxgb = 0
$csv.mailboxsizegb | % { $allmailboxgb = $allmailboxgb + $_ }
Write-Host "`nFor question 3, the total size of mailboxes accross all sites is"$allmailboxgb"GB`n" -ForegroundColor green

# Question 4
read-host -Prompt "Press enter to continue"
$userswithdiffnames = @()
foreach ($user in $csv) {

if ($user.EmailAddress -ceq $user.UserPrincipalName) {
#Do nothing..
}

else {
#Adding to array for users with different names.
$userswithdiffnames += $user
}

}
        
Write-Host "`nFor question 4,"$userswithdiffnames.count"users exist with different emailaddresses to userprincipalnames, with case sensitivity in mind. Please see list below.`n" -ForegroundColor green
$userswithdiffnames | ft

# Question 5
read-host -Prompt "Press enter to continue"
[int]$NYCmailboxgb = 0
$nycusers = $csv | ? {$_.site -eq 'NYC'}
$nycusers.mailboxsizegb | % { $nycmailboxgb = $nycmailboxgb + $_ }
Write-Host "`nFor question 5, the total size of mailboxes at the NYC site is"$NYCmailboxgb"GB`n" -ForegroundColor green


# Question 6
read-host -Prompt "Press enter to continue"
$employeeusers = $csv | ? {$_.accounttype -eq 'Employee'}
$employeeslargerthan10GB = @()
$mbgb = $null
$employeeusers | % { [int]$mbgb = $_.mailboxsizegb; if ($mbgb -gt '10') { $employeeslargerthan10GB += $_} else {} }
Write-Host "`nFor question 6, there are a total of"$employeeslargerthan10gb.count"users with acount type 'Employee' with maiilboxes larger than 10GB.`n" -ForegroundColor green

# Question 7
read-host -Prompt "Press enter to continue"
$domain2users = $nycusers | where {$_.emailaddress -like "*domain2*"} | Sort-Object mailboxsizegb -Descending | Select-Object -ExpandProperty EmailAddress
$wordarray = $domain2users.Replace("@domain2.com", "")
[string]$wordstring = $wordarray
Write-Host "`nFor question 7, I really liked the hidden message and it was definitely the funest question!`n`nPlease see the answer below.`n`n"$wordstring"`n" -ForegroundColor green

# Question 8
read-host -Prompt "Press enter to continue"



$groupdata = $csv | Group-Object Site
$summary = @()

Foreach ($Site in $GroupData) {

$TotalEmployees = $null
$TotalContractor = $null
[int]$TotalmailobxsizeGBInt = $null
[string]$TotalmailobxsizeGB = $null
[string]$AverageMailboxsizeGB = $null
[int]$AverageMailboxsizeGBInt = $null

$TotalEmployees = $csv | ? {$_.Site -eq $site.name -and $_.AccountType -eq 'Employee' }
$TotalContractor = $csv | ? {$_.Site -eq $site.name -and $_.AccountType -eq 'Contractor' }
$csv | ? {$_.Site -eq $site.name } | select -ExpandProperty MailboxsizeGB | % { $TotalmailobxsizeGBint += $_ }
[string]$TotalmailobxsizeGB = "$TotalmailobxsizeGBint" + "GB"
$AverageMailboxsizeGBInt = $TotalmailobxsizeGBInt / $site.count
$AverageMailboxsizeGB = "$AverageMailboxsizeGBint" + "GB"


$data = [PSCustomObject]@{ 
         
          Site =   Write-Output $site.name
          TotalUserCount = write-output $site.count 
          EmployeeCount = Write-Output $TotalEmployees.Count
          ContractorCount = Write-Output $TotalContractor.Count
          TotalMailboxSizeGB = Write-Output $TotalmailobxsizeGB
          AverageMailboxSizeGB = Write-Output $AverageMailboxsizeGB

          }
          $summary += $data

          }
          $repotfile = New-Item -force -Path "C:\temp\Report_$((Get-Date).ToString('dd-MM-yyyy')).csv"
          $summary | Export-Csv -Force -Path $repotfile -NoTypeInformation
          $summary | ft

Write-Host "`nFor question 8, please see the data summary above, please note, I have also exported this to a newly generated report file in C:\temp.`n`nThank you!!" -ForegroundColor green
