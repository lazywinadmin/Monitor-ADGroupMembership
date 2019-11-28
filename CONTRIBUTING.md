# Contributing to ADSIPS module

Contributions are welcome via pull requests and issues. Before submitting a pull request, please make sure all tests pass.

## Guidelines

* Don't use Write-Host
* Use Verb-Noun format, Check Get-Verb for approved verb
* Always use explicit parameter names, don't assume position
* Don't use Aliases
* Should support Credential input
* If you want to show informational information use Write-Verbose
* If you use Verbose, show the name of the function, you can do this:

```powershell
# Define this variable a the beginning
$ScriptName = (Get-Variable -name MyInvocation -Scope 0 -ValueOnly).Mycommand

# Show your verbose message this way
Write-Verbose -Message "[$ScriptName] Querying system X"
```

* You need to have Error Handling (TRY/CATCH)
* Return terminating error using ```$PSCmdlet.ThrowTerminatingError($_)```
* Think about the next guy, document your function, help them understand what you are achieving, give at least one example
* Implement appropriate WhatIf/Confirm support if you function is changing something

## TODO (not in a specific order)

* [ ] Pick one of the issues
* [ ] Add more Pester test
* [x] Integrate with CI/CD (Appveyor) for auto publish to gallery
  - **NOTE:** have since migrated to Azure Pipelines
* [x] Migrate to Azure Pipelines
* [ ] Update documentation, possibly add support for [ReadTheDocs](http://docs.readthedocs.io/en/latest/index.html) to help Adoption
* [ ] Set-ADSIComputer
* [ ] Set-ADSIGroup (also with -ManagedBy, and -ManagerCanUpdateMembership [example](https://blogs.technet.microsoft.com/blur-lines_-powershell_-author_shirleym/2013/10/07/manager-can-update-membership-list-part-1/))
* [ ] Set-ADSIObject
* [ ] New-User -CopyFrom <account/DN/GUID>
* [ ] Set-ADSIOrganizationalUnit
* [ ] Restore-ADSIAccount
* [ ] Unlock-ADSIAccount
* [x] Search-ADSIAccount (retrieve disabled account, expired, expiring,...)
* [ ] ACL functions
* [ ] GPO functions
* [ ] Set-ADSIDomainMode
* [ ] Set-ADSIForestMode
* [ ] Get-ADSIAccountResultantPasswordReplicationPolicy
* [ ] Get-ADSIDomainControllerPasswordReplicationPolicy
* [ ] Add-ADSIDomainControllerPasswordReplicationPolicy
* [ ] Remove-ADSIDomainControllerPasswordReplicationPolicy
* [x] Get-ADSIDefaultDomainPasswordPolicy
* [ ] Set-ADSIDefaultDomainPasswordPolicy
* [ ] Get-ADSIDomainControllerPasswordReplicationPolicyUsage
* [ ] Get-ADSIDomainControllerPasswordReplicationPolicyUsage
* [x] Get-ADSIFineGrainedPasswordPolicy
* [ ] Get-ADSIAccountResultantPasswordReplicationPolicy
* [ ] Set-ADSIAccountPassword
* [ ] Clear-ADSIAccountExpiration
* [ ] Find expensive queries
* [ ] Improve Get TokenSize