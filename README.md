# Get-EffectiveAccess

### DESCRIPTION
Gets the ACL of an object and translate it to readable data.

### USAGE
```
PS \> Get-ADOrganizationalUnit -Filter "Name -eq 'ExampleOU'" | Get-EffectiveAccess | Out-GridView
PS \> Get-EffectiveAccess -DistinguishedName 'OU=ExampleOU,DC=domainName,DC=com' | Out-GridView
PS \> $effectiveAccess = Get-ADGroup exampleGroup | Get-EffectiveAccess -IncludeOrphan
PS \> Get-ADOrganizationalUnit -Filter * | Select -First 10 | Get-EffectiveAccess | Out-GridView
```
### SWITCH
- **`IncludeOrphan`** By default, the function will filter all orphaned ACLs. Use this switch to include all `IdentityReference` that begin with `S-1-*`

### Requirements
- PowerShell v5.1
- ActiveDirectory PS Module

### EXAMPLE OUTPUT WITH Out-GridView

![exampleoutput](/effectiveAccess.png?raw=true)
