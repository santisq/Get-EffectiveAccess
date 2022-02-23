# requires -Modules ActiveDirectory
$ErrorActionPreference = 'Stop'

function Get-EffectiveAccess {
[CmdletBinding()]
param(
[Parameter(
    Mandatory,
    ValueFromPipelineByPropertyName
)]
[ValidatePattern(
    '(?:(CN=([^,]*)),)?(?:((?:(?:CN|OU)=[^,]+,?)+),)?((?:DC=[^,]+,?)+)$'
)][string]$DistinguishedName,
[switch]$IncludeOrphan
)

    begin {
        $GUIDMap = @{}
        $domain = Get-ADRootDSE
        $z = '00000000-0000-0000-0000-000000000000'
        $hash = @{
            SearchBase  = $domain.schemaNamingContext
            LDAPFilter  = '(schemaIDGUID=*)'
            Properties  = 'name', 'schemaIDGUID'
            ErrorAction = 'SilentlyContinue'
        }
        $schemaIDs = Get-ADObject @hash 

        $hash = @{
            SearchBase  = "CN=Extended-Rights,$($domain.configurationNamingContext)"
            LDAPFilter  = '(objectClass=controlAccessRight)'
            Properties  = 'name', 'rightsGUID'
            ErrorAction = 'SilentlyContinue'
        }
        $extendedRigths = Get-ADObject @hash

        foreach($i in $schemaIDs) {
            if(-not $GUIDMap.ContainsKey([System.GUID]$i.schemaIDGUID)) {
                $GUIDMap.add([System.GUID]$i.schemaIDGUID, $i.name)
            }
        }
        foreach($i in $extendedRigths) {
            if(-not $GUIDMap.ContainsKey([System.GUID]$i.rightsGUID)) {
                $GUIDMap.add([System.GUID]$i.rightsGUID, $i.name)
            }
        }
    }

    process {
        $object = Get-ADObject $DistinguishedName
        $acls = (Get-ACL "AD:$object").Access
        
        $result = foreach($acl in $acls) {
            $objectType = (
                $GUIDMap[$acl.ObjectType],
                'All Objects (Full Control)'
            )[$acl.ObjectType -eq $z]
            
            $inheritedObjType = (
                $GUIDMap[$acl.InheritedObjectType],
                'Applied to Any Inherited Object'
            )[$acl.InheritedObjectType -eq $z]

            [PSCustomObject]@{
                Name = $object.Name
                IdentityReference = $acl.IdentityReference
                AccessControlType = $acl.AccessControlType
                ActiveDirectoryRights = $acl.ActiveDirectoryRights
                ObjectType = $objType
                InheritedObjectType = $inheritedObjType
                InheritanceType = $acl.InheritanceType
                IsInherited = $acl.IsInherited
            }
        }
        
        if($IncludeOrphan.IsPresent) {
            return $result | Sort-Object IdentityReference
        }
        
        $result | Where-Object { -not $_.IdentityReference.StartsWith('S-1-5-21') } |
        Sort-Object IdentityReference        
    }
}
