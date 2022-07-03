# Get-EffectiveAccess

## Description

PowerShell function that tries to give a friendly translatation of [`Get-Acl`](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/get-acl?view=powershell-7.2) into human readable data. The function is designed exclusively for Active Directory, and requires the [__ActiveDirectory Module__](https://docs.microsoft.com/en-us/powershell/module/activedirectory/?view=windowsserver2022-ps).

## Examples


- Get the _Effective Access_ of the Organizational Unit named `ExampleOU`:

```powershell
Get-ADOrganizationalUnit -Filter "Name -eq 'ExampleOU'" |
    Get-EffectiveAccess | Out-GridView
```

- Same as above but using the OU's `DistinguishedName` attribute:

```powershell
Get-EffectiveAccess -Identity 'OU=ExampleOU,DC=domainName,DC=com' | Out-GridView
```

- Store the _Effective Access_ of the group named `exampleGroup` in a variable:

```powershell
$effectiveAccess = Get-ADGroup exampleGroup | Get-EffectiveAccess
```

- Get the _Effective Access_ of the first 10 OUs found in the Domain:

```powershell
Get-ADOrganizationalUnit -Filter * | Select -First 10 |
    Get-EffectiveAccess | Out-GridView
```

## Sample output with `Out-GridView`

![exampleoutput](/Screenshot/effectiveAccess.png?raw=true)

## Requirements

- PowerShell v5.1+
- ActiveDirectory PS Module