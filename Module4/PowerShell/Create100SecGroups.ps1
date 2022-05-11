# Script to create large amount of security groups.
# The purpose of this script is to be able to ilustrate the difference between scripts
# "Call Graph from PowerShell - Demo 1.ps1" and "Call Graph from PowerShell - Demo 2.ps1" when it comes to pagination.

Connect-AzureAD

# Create Groups
[int]$numberGroups = 150
for ([int]$i = 1; $i -le $numberGroups; $i++) {
    New-AzureADGroup -DisplayName "Dummy Group $i" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
}

# Remove Groups
#$groupsToDelete = Get-AzureADGroup -SearchString "Dummy Group" -All $true
#foreach ($groupToDelete in $groupsToDelete){
#    Remove-AzureADGroup -ObjectId $groupToDelete.ObjectId
#}

