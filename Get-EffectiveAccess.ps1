function Get-EffectiveAccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidatePattern('(?:(CN=([^,]*)),)?(?:((?:(?:CN|OU)=[^,]+,?)+),)?((?:DC=[^,]+,?)+)$')]
        [alias('DistinguishedName')]
        [string] $Identity,

        [parameter()]
        [alias('Domain')]
        [string] $Server
    )

    begin {
        $guid    = [guid]::Empty
        $GUIDMap = @{}

        if($PSBoundParameters.ContainsKey('Server')) {
            $domain = Get-ADRootDSE -Server $Server
        }
        else {
            $domain = Get-ADRootDSE
        }

        $params = @{
            SearchBase  = $domain.schemaNamingContext
            LDAPFilter  = '(schemaIDGUID=*)'
            Properties  = 'name', 'schemaIDGUID'
            ErrorAction = 'SilentlyContinue'
        }
        $adObjParams = @{
            Properties = 'nTSecurityDescriptor'
        }

        if($PSBoundParameters.ContainsKey('Server')) {
            $params['Server']  = $Server
            $adObjParams['Server'] = $Server
        }
        $schemaIDs = Get-ADObject @params

        $params['SearchBase'] = "CN=Extended-Rights,$($domain.configurationNamingContext)"
        $params['LDAPFilter'] = '(objectClass=controlAccessRight)'
        $params['Properties'] = 'name', 'rightsGUID'
        $extendedRigths = Get-ADObject @params

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
            $adObjParams['Identity'] = $Identity
            $object = Get-ADObject @adObjParams

            foreach($acl in $object.nTSecurityDescriptor.Access) {
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