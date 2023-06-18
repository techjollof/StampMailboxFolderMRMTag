# This function will automatically change current folder to the current extracted folder location
function prompt {
    $p = Split-Path -leaf -path (Get-Location)
    "$p> "
}
prompt

# Retrive the data from the TXT file
$RetirveValue = Get-Content -Path ".\EWSRequiredParameters.txt" | ForEach-Object {
    ConvertFrom-StringData $($_ -join [Environment]::NewLine)
}

#Remapping value from the retrived data set
$RetirveParValue = @{
    TargetFolderName 			        = 	$RetirveValue.TargetFolderName
    ArchiveOrRetentionTagRawRetentionId =  	$RetirveValue.ArchiveOrRetentionTagRawRetentionId
    RetentionFlagsValue			        =	$RetirveValue.RetentionFlagsValue
    ArchiveOrRetentionPeriodInDays     	=   $RetirveValue.ArchiveOrRetentionPeriodInDays
    TenantInitialDomain			        =	$RetirveValue.TenantInitialDomain
    AzureEWSApplicationClientId		    = 	$RetirveValue.AzureEWSApplicationClientId
    TargetUserAccountsCsv			    =	$RetirveValue.TargetUserAccountsCsv
    TargetFolderLocation			    = 	$RetirveValue.TargetFolderLocation
    ArchiveOrRetainAction               =   $RetirveValue.ArchiveOrRetainAction
}

#Cast the values to the main program file
.\EWSMRMPolicyTagAssignment.ps1 @RetirveParValue