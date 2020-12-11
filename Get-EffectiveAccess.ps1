Function Get-EffectiveAccess {
<#
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
    Get-ADOrganizationalUnit -Filter {Name -eq 'Web'} | Get-EffectiveAccess
    Get-ADComputer -Filter 'Name -like "VRTVA*"' -SearchBase 'OU=Other,OU=Azure,OU=CIO Segment Servers,DC=dir,DC=svc,DC=accenture,DC=com' | Get-EffectiveAccess 
    Get-ADOrganizationalUnit -SearchBase 'OU=IAM,DC=dir,DC=svc,DC=accenture,DC=com' -SearchScope Subtree -Filter * | Get-EffectiveAccess 

There are 3 available switches: IncludeIAM, IncludeOprhan & OutGrid:
    IncludeIAM: To include all groups managed by IAM team.
    IncludeOrphan: Any IdentityReference in the ACL that begins with "S-1-*" – If you want to display them use these switch.
    OutGrid: The function will use Out-GridView instead of exporting to Excel.

#Requires -Modules ActiveDirectory, ImportExcel

.AUTHOR
    Santiago Squarzon
#>

[CmdletBinding()]
Param(
      [Parameter(
        Mandatory,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName
        )]
      [Alias('DistinguishedName')]
      $parameter,
      
      [Parameter(Mandatory=$false)]
      [switch]$IncludeIAM,
      [switch]$IncludeOrphan,
      [switch]$OutGrid
      )

Begin {

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Validate that the required Powershell modules are installed for the Current User.
Try {Import-Module ActiveDirectory -ErrorAction Stop}
Catch {Write-Warning "Need ActiveDirectory Powershell Module to run this script. Can be downloaded from:
https://www.microsoft.com/en-us/download/details.aspx?id=45520";break}

Try {Import-Module ImportExcel -ErrorAction Stop}
Catch {Write-Warning "Need ImportExcel Powershell Module to export the Access Control List.
You can use 'Install-Module ImportExcel -Scope CurrentUser' to install it.";break}
############################################################################
$elapsed = [System.Diagnostics.Stopwatch]::StartNew()
$GUID = @{}
$exportPath = "$env:USERPROFILE\Documents\ACLs\"
$domain = Get-ADRootDSE
$switches = @{
    generic = $(cat "$psscriptroot\genericExclude.txt")
    orphan = "S-1-*"
    iamman = $(cat "$psscriptroot\iamGroups.txt")
    }
}#End of Begin {} block.

Process {
#Begin of Process {} block and parameter validation.
$ACLGrid = @()
ForEach ($param in $parameter) {
If ($param -match "CN=[a-zA-Z0-9]+|OU=[a-zA-Z0-9]+|DC=[a-zA-Z0-9]+") {
    $sort = Get-ADobject $param -Properties CanonicalName
    }
ElseIf ($param -match "\w+/\w+.*") {
    $lastindex = $param.Substring($param.LastIndexOf('/')+1)
    $sort = Get-ADObject -Filter "Name -eq '$($lastindex)'" -Properties CanonicalName | ?{$_.CanonicalName -eq $param}
    }
Else {$sort = Get-ADobject -Filter "Name -eq '$($param)'" -Properties CanonicalName}
If (!$sort) {Write-Warning "Cannot find an Object or Organizational Unit with name '$($param.ToUpper())' under: $($domain.DefaultNamingContext).";return}
#End of input paramater validation.
###################################################
#.NET integration when there is more than one OU with the provided input parameter.
If ($sort.Count -gt 1) {

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Multiple OUs found'
$form.Size = New-Object System.Drawing.Size(650,600)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'Fixed3D'
$form.TopMost = $true
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(430,510)
$OKButton.Size = New-Object System.Drawing.Size(100,40)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$OKButton.Font = New-Object System.Drawing.Font("Tahoma",9,[System.Drawing.FontStyle]::Regular)
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(530,510)
$CancelButton.Size = New-Object System.Drawing.Size(100,40)
$CancelButton.Font = New-Object System.Drawing.Font("Tahoma",9,[System.Drawing.FontStyle]::Regular)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(580,40)
$label.Font = New-Object System.Drawing.Font("Tahoma",8,[System.Drawing.FontStyle]::Regular)
$label.Text = "The following OUs with the name '$($param.ToUpper())' were found under: $($domain.defaultNamingContext)`nSelect the one to get the Access Control List."
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,70)
$listBox.Size = New-Object System.Drawing.Size(620,430)
$listBox.Font = New-Object System.Drawing.Font("Tahoma",13,[System.Drawing.FontStyle]::Regular)
$listBox.HorizontalScrollbar = $false

$sort.CanonicalName | Sort-Object | %{
    [void]$listBox.Items.Add($_)
    }
$form.Controls.Add($listBox)

$width = ($listBox.Items | Measure-Object -Maximum -Property Length).Maximum * 9
If ($width -gt $listBox.Size.Width) {
    $listBox.HorizontalExtent = $width
    $listBox.HorizontalScrollbar = $true
    }

$form.Topmost = $true
$result = $form.ShowDialog()

If ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $sort = $sort | ?{$_.CanonicalName -match $listBox.Selecteditem}
    } Else {return}
}#End of .NET integration.
##################################################################################
#Mapping the Domain's Schema & Extended Rights GUIDs.
If ($GUID.Count -eq 0) {
If ((Test-Path "$exportPath\GUID.map\GUIDmap.csv") -and (Get-Item "$exportPath\GUID.map\GUIDmap.csv").Length -ge 350000) {
$csv = Import-Csv "$exportPath\GUID.map\GUIDmap.csv"
$i=0;$csv | %{
Write-Progress -Activity "Importing System.GUID Map..." -Status $_.Value -PercentComplete ($i++/($csv.Count)*100)
$GUID[[system.GUID]$_.Key]=$_.Value
    }
} Else {
$ErrorActionPreference = 'SilentlyContinue'
Write-Progress "Importing SchemaIDGUID..." -Status $domain.schemaNamingContext -PercentComplete 50
$schemaIDs = Get-ADObject -SearchBase $domain.schemaNamingContext -LDAPFilter '(schemaIDGUID=*)' -Properties name, schemaIDGUID
Write-Progress "Importing ExtendedRights..." -Status "CN=Extended-Rights,$($domain.configurationNamingContext)" -PercentComplete 100
$extendedRigths = Get-ADObject -SearchBase "CN=Extended-Rights,$($domain.configurationNamingContext)" -LDAPFilter '(objectClass=controlAccessRight)' -Properties name, rightsGUID

$i=0;$schemaIDs | %{
Write-Progress -Activity "Generating System.GUID Map..." -Status $_.Name -PercentComplete ($i++/($schemaIDs.Count+$extendedRigths.Count)*100)
$GUID.add([System.GUID]$_.schemaIDGUID,$_.name)
}
$extendedRigths | %{
Write-Progress -Activity "Generating System.GUID Map..." -Status $_.Name -PercentComplete ($i++/($schemaIDs.Count+$extendedRigths.Count)*100)
$GUID.add([System.GUID]$_.rightsGUID,$_.name)}
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path "$exportPath\GUID.map" | Out-Null
$GUID.GetEnumerator() | Select Key,Value | Export-Csv "$exportPath\GUID.map\GUIDmap.csv" -NoTypeInformation
    }
}
###################################################################################
#Getting the complete ACL for the provided object or container and sorting it using the GUID Map.
$ACL = (Get-ACL "AD:$sort").Access
$ACL | %{
    $ACLGrid += [PSCustomObject]@{
        'Identity Reference' = $_.IdentityReference
        'Access Control Type' = $_.AccessControlType
        'Active Directory Rights' = $_.ActiveDirectoryRights
        'Object Type' = $(
             If ($_.ObjectType -eq '00000000-0000-0000-0000-000000000000'){
                'All Objects (Full Control)'}Else{$GUID[$_.ObjectType]
            })
        'Inherited Object Type' = $(
            If ($_.InheritedObjectType -eq '00000000-0000-0000-0000-000000000000'){
                'Applied to Any Inherited Object'}Else{$GUID[$_.InheritedObjectType]
            })
        'Inheritance Type' = $_.InheritanceType
        'Is Inherited' = $_.IsInherited
            }
        }
#Checking if the switches are present to sort the output object.
If ($IncludeIAM.IsPresent -and $IncludeOrphan.IsPresent) {
    $aclExport = $ACLGrid | ?{$_.'identity reference' -notmatch $switches.generic} |
        Sort-Object 'identity reference'#,'Access Control Type','object type'
}ElseIf ($IncludeIAM.IsPresent) {
    $aclExport = $ACLGrid | ?{$_.'identity reference' -notmatch "$($switches.generic)|$($switches.orphan)"} |
        Sort-Object 'identity reference'#,'Access Control Type','object type'
}ElseIf ($IncludeOrphan.IsPresent) {
    $aclExport = $ACLGrid | ?{$_.'identity reference' -notmatch "$($switches.generic)|$($switches.iamman)"} |
        Sort-Object 'identity reference'#,'Access Control Type','object type'
}Else {
    $aclExport = $ACLGrid | ?{$_.'identity reference' -notmatch "$($switches.generic)|$($switches.iamman)|$($switches.orphan)"} |
        Sort-Object 'identity reference'#,'Access Control Type','object type'
}
#End of switch validation.
##################################
If ($OutGrid.IsPresent) {$aclExport | Out-GridView -Title "ACL for $($sort.CanonicalName)"}
Else {    
#Begin of Excel output formatting.
If (!(Test-Path $exportPath)){New-Item -ItemType Directory -Force -Path $exportPath | Out-Null}
$folder = "$(Get-Date -format ddMMyyhh)\"
If (!(Test-Path $folder)){New-Item -ItemType Directory -Force -Path $exportPath$folder | Out-Null}
$filename = "ACL - $($sort.CanonicalName -replace '/','_').xlsx"

$xls = $aclExport |
    Export-Excel ($exportPath+$folder+$filename) -StartRow 2 -TableName acl -WorksheetName $sort.Name -AutoFilter -FreezePane 3 -PassThru

$sheet = $xls.Workbook.Worksheets[$sort.Name]

$sheet.Row(1).Height = 50;$sheet.Column(1).Width = 33;$sheet.Column(2).Width = 20;$sheet.Column(3).Width = 32
$sheet.Column(4).Width = 23;$sheet.Column(5).Width = 26;$sheet.Column(6).Width = 17;$sheet.Column(7).Width = 13

Set-Format -Range A1:G$($acl.count) -Worksheet $sheet -FontSize 9 -FontName "Tahoma" -VerticalAlignment Center -WrapText
Set-Format -Range A1 -value "Access Control List for $($sort.CanonicalName)" -Worksheet $sheet -Bold -FontSize 18
Set-Format -Range A1:G1 -Merge -Worksheet $sheet -VerticalAlignment Top
Set-Format -Range A2:G2 -Worksheet $sheet -Bold
Set-Format -Range B3:B$($acl.Count) -Worksheet $sheet -HorizontalAlignment Center
Set-Format -Range A3:G$($acl.Count) -Worksheet $sheet -WrapText

Close-ExcelPackage $xls
Write-Host "Output file:`n$exportPath$folder$filename"
ii $exportPath$folder
           }#End of Excel formatting.
        }#End of ForEach loop.
  }#End of Process {} block.
End {
$elapsed.Stop()
Write-Host "`nElapsed Time: $($elapsed.Elapsed.Seconds) seconds."
    }
}