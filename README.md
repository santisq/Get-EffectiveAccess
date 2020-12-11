# Get-EffectiveAccess

.DESCRIPTION
Gets the ACL of an object and translate it to readable data.
The input parameter can either be the object’s Name, DistinguishedName or CanonicalName:
    Get-EffectiveAccess ExampleOU
    Get-EffectiveAccess domainName.com/ExampleOU
    Get-EffectiveAccess OU=ExampleOU,DC=domainName,DC=com

.OUTPUT
Excel file or GridView (OutGrid switch).

.EXAMPLE
Pipeline examples:
    Get-ADOrganizationalUnit -Filter {Name -eq 'ExampleOU'} | Get-EffectiveAccess
    Get-ADComputer -Filter 'Name -like "ExampleCOMP*"' -SearchBase 'OU=ExampleOU,DC=domainName,DC=com' | Get-EffectiveAccess 

There are 3 available switches: IncludeIAM, IncludeOprhan & OutGrid:
    IncludeIAM: To include all groups managed by IAM team.
    IncludeOrphan: Any IdentityReference in the ACL that begins with "S-1-*" – If you want to display them use this switch.
    OutGrid: The function will use Out-GridView instead of exporting to Excel.

#Requires -Modules ActiveDirectory, ImportExcel

.AUTHOR
    Santiago Squarzon
 
