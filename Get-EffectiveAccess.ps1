function Get-EffectiveAccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidatePattern('(?:(CN=([^,]*)),)?(?:((?:(?:CN|OU)=[^,]+,?)+),)?((?:DC=[^,]+,?)+)$')]
        [alias('ObjectGUID','ObjectSID','DistinguishedName','sAMAccountName')]
        [string] $Identity
    )

    begin {
        $GUIDMap = @{}
        $domain  = Get-ADRootDSE
        $guid    = [guid]::Empty
        $hash    = @{
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
            if(-not $GUIDMap.ContainsKey([guid] $i.schemaIDGUID)) {
                $GUIDMap.Add([guid] $i.schemaIDGUID, $i.name)
            }
        }
        foreach($i in $extendedRigths) {
            if(-not $GUIDMap.ContainsKey([guid] $i.rightsGUID)) {
                $GUIDMap.Add([guid] $i.rightsGUID, $i.name)
            }
        }
    }

    process {
        try {
            $object = Get-ADObject $DistinguishedName
            $acls   = (Get-ACL "AD:$object").Access

            foreach($acl in $acls) {
                if($guid.Equals($acl.ObjectType)) {
                    $objectType = 'All Objects (Full Control)'
                }
                elseif($GUIDMap.ContainsKey($acl.ObjectType)) {
                    $objectType = $GUIDMap[$acl.ObjectType]
                }
                else {
                    $objectType = $acl.ObjectType
                }

                if($guid.Equals($acl.InheritedObjectType)) {
                    $inheritedObjType = 'Applied to Any Inherited Object'
                }
                elseif($GUIDMap.ContainsKey($acl.InheritedObjectType)) {
                    $inheritedObjType = $GUIDMap[$acl.InheritedObjectType]
                }
                else {
                    $inheritedObjType = $acl.InheritedObjectType
                }

                [PSCustomObject]@{
                    Name                  = $object.Name
                    IdentityReference     = $acl.IdentityReference
                    AccessControlType     = $acl.AccessControlType
                    ActiveDirectoryRights = $acl.ActiveDirectoryRights
                    ObjectType            = $objectType
                    InheritedObjectType   = $inheritedObjType
                    InheritanceType       = $acl.InheritanceType
                    IsInherited           = $acl.IsInherited
                }
            }
        }
        catch {
            $PSCmdlet.WriteError($_)
        }
    }
}