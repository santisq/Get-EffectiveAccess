# Get-EffectiveAccess

### DESCRIPTION
Gets the ACL of an object and translate it to readable data.

### USAGE
```
PS \> Get-ADOrganizationalUnit -Filter {Name -eq 'ExampleOU'} | Get-EffectiveAccess
PS \> Get-EffectiveAccess -DistinguishedName 'OU=ExampleOU,DC=domainName,DC=com'
PS \> Get-ADGroup exampleGroup | Get-EffectiveAccess -IncludeOrphan
```
### SWITCH
- **`IncludeOrphan`** By default, the function will filter all orphaned ACLs. Use this switch to include all `IdentityReference` that begins with `S-1-*`

### Requirements
- PowerShell v5.1
- ActiveDirectory PS Module
