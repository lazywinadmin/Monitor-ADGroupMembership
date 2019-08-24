<#PSScriptInfo

.VERSION 2.0.6

.GUID 28826cd6-0760-4e00-aae6-1330e60118ee

.AUTHOR Francois-Xavier Cat

.COMPANYNAME LazyWinAdmin.Com

.COPYRIGHT (c) 2016 Francois-Xavier Cat. All rights reserved. Licensed under The MIT License (MIT)

.TAGS ActiveDirectory Group GroupMembership Monitor Report ADSI Quest

.LICENSEURI https://github.com/lazywinadmin/Monitor-ADGroupMembership/blob/master/LICENSE

.PROJECTURI https://github.com/lazywinadmin/Monitor-ADGroupMembership

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES


#>

<#
.DESCRIPTION
	This script is monitoring group(s) in Active Directory and send an email when someone is changing the membership.
	It will also report the Change History made for this/those group(s).

.SYNOPSIS
	This script is monitoring group(s) in Active Directory and send an email when someone is changing the membership.

.PARAMETER Group
	Specify the group(s) to query in Active Directory.
	You can also specify the 'DN','GUID','SID' or the 'Name' of your group(s).
	Using 'Domain\Name' will also work.

.PARAMETER SearchRoot
	Specify the DN, GUID or canonical name of the domain or container to search. By default, the script searches the entire sub-tree of which SearchRoot is the topmost object (sub-tree search). This default behavior can be altered by using the SearchScope parameter.

.PARAMETER SearchScope
	Specify one of these parameter values
		'Base' Limits the search to the base (SearchRoot) object.
			The result contains a maximum of one object.
		'OneLevel' Searches the immediate child objects of the base (SearchRoot)
			object, excluding the base object.
		'Subtree'   Searches the whole sub-tree, including the base (SearchRoot)
			object and all its child objects.

.PARAMETER GroupScope
	Specify the group scope of groups you want to find. Acceptable values are: 
		'Global'; 
		'Universal'; 
		'DomainLocal'.

.PARAMETER GroupType
	Specify the group type of groups you want to find. Acceptable values are: 
		'Security';
		'Distribution'.

.PARAMETER File
	Specify the File where the Group are listed. DN, SID, GUID, or Domain\Name of the group are accepted.

.PARAMETER EmailServer
    Specify the Email Server IPAddress/FQDN.
    
.PARAMETER EmailPort
    Specify the port for the Email Server

.PARAMETER EmailTo
	Specify the Email Address(es) of the Destination. Example: fxcat@fx.lab

.PARAMETER EmailFrom
	Specify the Email Address of the Sender. Example: Reporting@fx.lab

.PARAMETER EmailEncoding
	Specify the Body and Subject Encoding to use in the Email.
	Default is ASCII.

.PARAMETER Server
	Specify the Domain Controller to use.
	Aliases: DomainController, Service

.PARAMETER HTMLLog
	Specify if you want to save a local copy of the Report.
	It will be saved under the directory "HTML".

.PARAMETER IncludeMembers
	Specify if you want to include all members in the Report.
	
.PARAMETER AlwaysReport
	Specify if you want to generate a Report each time.

.PARAMETER OneReport
	Specify if you want to only send one email with all group report as attachment

.PARAMETER ExtendedProperty
	Specify if you want to add Enabled and PasswordExpired Attribute on members in the report
	
.EXAMPLE
	.\AD-GROUP-Monitor_MemberShip.ps1 -Group "FXGroup" -EmailFrom "From@Company.com" -EmailTo "To@Company.com" -EmailServer "mail.company.com"

	This will run the script against the group FXGROUP and send an email to To@Company.com using the address From@Company.com and the server mail.company.com.

.EXAMPLE
	.\AD-GROUP-Monitor_MemberShip.ps1 -Group "FXGroup","FXGroup2","FXGroup3" -EmailFrom "From@Company.com" -Emailto "To@Company.com" -EmailServer "mail.company.com"

	This will run the script against the groups FXGROUP,FXGROUP2 and FXGROUP3  and send an email to To@Company.com using the address From@Company.com and the Server mail.company.com.

.EXAMPLE
	.\AD-GROUP-Monitor_MemberShip.ps1 -Group "FXGroup" -EmailFrom "From@Company.com" -Emailto "To@Company.com" -EmailServer "mail.company.com" -Verbose

	This will run the script against the group FXGROUP and send an email to To@Company.com using the address From@Company.com and the server mail.company.com. Additionally the switch Verbose is activated to show the activities of the script.

.EXAMPLE
	.\AD-GROUP-Monitor_MemberShip.ps1 -Group "FXGroup" -EmailFrom "From@Company.com" -Emailto "Auditor@Company.com","Auditor2@Company.com" -EmailServer "mail.company.com" -Verbose

	This will run the script against the group FXGROUP and send an email to Auditor@Company.com and Auditor2@Company.com using the address From@Company.com and the server mail.company.com. Additionally the switch Verbose is activated to show the activities of the script.

.EXAMPLE
	.\AD-GROUP-Monitor_MemberShip.ps1 -SearchRoot 'FX.LAB/TEST/Groups' -Emailfrom Reporting@fx.lab -Emailto "Catfx@fx.lab" -EmailServer 192.168.1.10 -Verbose

	This will run the script against all the groups present in the CanonicalName 'FX.LAB/TEST/Groups' and send an email to catfx@fx.lab using the address Reporting@fx.lab and the server 192.168.1.10. Additionally the switch Verbose is activated to show the activities of the script.

.EXAMPLE
	.\AD-GROUP-Monitor_MemberShip.ps1 -file .\groupslist.txt -Emailfrom Reporting@fx.lab -Emailto "Catfx@fx.lab" -EmailServer 192.168.1.10 -Verbose

	This will run the script against all the groups present in the file groupslists.txt and send an email to catfx@fx.lab using the address Reporting@fx.lab and the server 192.168.1.10. Additionally the switch Verbose is activated to show the activities of the script.

.EXAMPLE
	.\AD-GROUP-Monitor_MemberShip.ps1 -server DC01.fx.lab -file .\groupslist.txt -Emailfrom Reporting@fx.lab -Emailto "Catfx@fx.lab" -EmailServer 192.168.1.10 -Verbose

	This will run the script against the Domain Controller "DC01.fx.lab" on all the groups present in the file groupslists.txt and send an email to catfx@fx.lab using the address Reporting@fx.lab and the server 192.168.1.10. Additionally the switch Verbose is activated to show the activities of the script.

.EXAMPLE
	.\AD-GROUP-Monitor_MemberShip.ps1 -server DC01.fx.lab:389 -file .\groupslist.txt -Emailfrom Reporting@fx.lab -Emailto "Catfx@fx.lab" -EmailServer 192.168.1.10 -Verbose

	This will run the script against the Domain Controller "DC01.fx.lab" (on port 389) on all the groups present in the file groupslists.txt and send an email to catfx@fx.lab using the address Reporting@fx.lab and the server 192.168.1.10. Additionally the switch Verbose is activated to show the activities of the script.

.EXAMPLE
	.\AD-GROUP-Monitor_MemberShip.ps1 -group "Domain Admins" -Emailfrom Reporting@fx.lab -Emailto "Catfx@fx.lab" -EmailServer 192.168.1.10 -EmailEncoding 'ASCII' -HTMLlog

	This will run the script against the group "Domain Admins" and send an email (using the encoding ASCII) to catfx@fx.lab using the address Reporting@fx.lab and the SMTP server 192.168.1.10. It will also save a local HTML report under the HTML Directory.

.INPUTS
	System.String

.OUTPUTS
	Email Report

.NOTES
	NAME:	AD-GROUP-Monitor_MemberShip.ps1
	AUTHOR:	Francois-Xavier Cat 
	EMAIL:	info@lazywinadmin.com
	WWW:	www.lazywinadmin
	Twitter:@lazywinadm

	REQUIREMENTS:
		-Read Permission in Active Directory on the monitored groups
		-Quest Active Directory PowerShell Snapin
		-A Scheduled Task (in order to check every X seconds/minutes/hours)

	VERSION HISTORY:
	1.0 	2012.02.01
		Initial Version

	1.1 	2012.03.13
		CHANGE to monitor both Domain Admins and Enterprise Admins

	1.2 	2013.09.23
		FIX issue when specifying group with domain 'DOMAIN\Group'
		CHANGE Script Format (BEGIN, PROCESS, END)
		ADD Minimal Error handling. (TRY CATCH)

	1.3 	2013.10.05
		CHANGE in the PROCESS BLOCK, the TRY CATCH blocks and placed
		 them inside the FOREACH instead of inside the TRY block
		ADD support for Verbose
		CHANGE the output file name "DOMAIN_GROUPNAME-membership.csv"
		ADD a Change History File for each group(s)
		 example: "GROUPNAME-ChangesHistory-yyyyMMdd-hhmmss.csv"
		ADD more Error Handling
		ADD a HTML Report instead of plain text
		ADD HTML header
		ADD HTML header for change history

	1.4 	2013.10.11
		CHANGE the 'Change History' filename to
		 "DOMAIN_GROUPNAME-ChangesHistory-yyyyMMdd-hhmmss.csv"
		UPDATE Comments Based Help
		ADD Some Variable Parameters
		
	1.5 	2013.10.13
		ADD the full Parameter Names for each Cmdlets used in this script
		ADD Alias to the Group ParameterName
		
	1.6 	2013.11.21
		ADD Support for Organizational Unit (SearchRoot parameter)
		ADD Support for file input (File Parameter)
		ADD ParamaterSetNames and parameters GroupType/GroupScope/SearchScope
		REMOVE [mailaddress] type on $Emailfrom and $EmailTo to make the script available to PowerShell 2.0
		ADD Regular expression validation on $Emailfrom and $EmailTo
	
	1.7 	2013.11.23
		ADD ValidateScript on File Parameter
		ADD Additional information about the Group in the Report
		CHANGE the format of the $changes output, it will now include the DateTime Property
		UPDATE Help
		ADD DisplayName Property in the report
		
	1.8 	2013.11.27
		Minor syntax changes
		UPDATE Help
	
	1.8.1 	2013.12.29
		Rename to AD-GROUP-Monitor_MemberShip

	1.8.2 	2014.02.17
		Update Notes

	2.0	2014.05.04
		ADD Support for ActiveDirectory module (AD module is use by default)
		ADD failover to Quest AD Cmdlet if AD module is available
		RENAME GetQADGroupParams variable to ADGroupParams

	2.0.1 	2015.01.05
		REMOVE the DisplayName property from the email
		ADD more clear details/Comments
		RENAME a couple of Verbose and Warning Messages
		FIX the DN of the group in the Summary
		FIX SearchBase/SearchRoot Parameter which was not working with AD Module
		FIX Some other minor issues
		ADD Check to validate data added to $Group is valid
		ADD Server Parameter to be able to specify a domain controller

	2.0.2	2015.01.14
		FIX an small issue with the $StateFile which did not contains the domain
		ADD the property Name into the final output.
		ADD Support to export the report to a HTML file (-HTMLLog) It will save
			the report under the folder HTML
		ADD Support for alternative Email Encoding: Body and Subject. Default is ASCII.
	
	2.0.3   2017.06.30
		ADD 'IncludeMembers' Switch to list all members in the report.
		ADD 'AlwaysReport' switch to send report each run.
		ADD 'OneReport' Switch to have it only send one email, with each group 
			report in the mail as attachment.
		ADD 'ExtendedProperty' switch to add Enabled and PasswordExpired to the member list.

	2.0.4   2017.06.30
		FIX Minor typos
		Update CATCH Blow to throw the error
		Update verbose and messages to include the script name
	
	2.0.5 2017.07.04
		FIX member liste showing in report were the history, not current.

	2.0.6 2019.08.13 @revoice1
		Fix issue where $Members only contain one item
		https://github.com/lazywinadmin/Monitor-ADGroupMembership/pull/42/files

	TODO:
		-Add Switch to make the Group summary Optional (info: Description,DN,CanonicalName,SID, Scope, Type)
			-Current Member Count, Added Member count, Removed Member Count
		-Switch to Show all the Current Members (Name, Department, Role, EMail)
		-Possibility to Ignore some groups
		-Email Credential
		-Recursive Membership search
		-Switch to save a local copy of the HTML report (maybe put this by default)
#>

#requires -version 3.0

[CmdletBinding(DefaultParameterSetName = "Group")]
PARAM(
    [Parameter(ParameterSetName = "Group", Mandatory = $true, HelpMessage = "You must specify at least one Active Directory group")]
    [ValidateNotNull()]
    [Alias('DN', 'DistinguishedName', 'GUID', 'SID', 'Name')]
    [string[]]$Group,

    [Parameter(ParameterSetName = "OU", Mandatory = $true)]
    [Alias('SearchBase')]
    [string[]]$SearchRoot,
    
    [Parameter(ParameterSetName = "OU")]
    [ValidateSet("Base", "OneLevel", "Subtree")]
    [string]$SearchScope,

    [Parameter(ParameterSetName = "OU")]
    [ValidateSet("Global", "Universal", "DomainLocal")]
    [String]$GroupScope,

    [Parameter(ParameterSetName = "OU")]
    [ValidateSet("Security", "Distribution")]
    [string]$GroupType,

    [Parameter(ParameterSetName = "File", Mandatory = $true)]
    [ValidateScript({ Test-Path -Path $_ })]
    [string[]]$File,

    [Parameter()]
    [Alias('DomainController', 'Service')]
    $Server,

    [Parameter(Mandatory = $true, HelpMessage = "You must specify the Sender Email Address")]
    [ValidatePattern("[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?")]
    [string]$EmailFrom,

    [Parameter(Mandatory = $true, HelpMessage = "You must specify the Destination Email Address")]
    [ValidatePattern("[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?")]
    [string[]]$EmailTo,

    [Parameter(Mandatory = $true, HelpMessage="You must specify the Email Server to use (IPAddress or FQDN)")]
    [string]$EmailServer,

    [Parameter(Mandatory = $true, HelpMessage="You must specify the Email Server Port")]
    [string]$EmailPort,

    [Parameter()]
    [ValidateSet("ASCII", "UTF8", "UTF7", "UTF32", "Unicode", "BigEndianUnicode", "Default")]
    [string]$EmailEncoding = "ASCII",

    [Parameter()]
    [Switch]$HTMLLog,

    [Parameter()]
    [Switch]$IncludeMembers,

    [Parameter()]
    [Switch]$AlwaysReport,

    [Parameter()]
    [Switch]$OneReport,

    [Parameter()]
    [Switch]$ExtendedProperty
)
Begin
{
    try
    {
        $EmailCred = (Get-Credential);
        # Retrieve the Script name
        $ScriptName = $MyInvocation.MyCommand;

        # Set the Paths Variables and create the folders if not present
        $ScriptPath = (Split-Path -Path ((Get-Variable -Name MyInvocation).Value).MyCommand.Path);
        $ScriptPathOutput = $ScriptPath + "\Output";

        if (-not(Test-Path -Path $ScriptPathOutput))
        {
            Write-Verbose -Message "[$ScriptName][Begin] Creating the Output Folder : $ScriptPathOutput";
            New-Item -Path $ScriptPathOutput -ItemType Directory | Out-Null;
        }

        $ScriptPathChangeHistory = $ScriptPath + "\ChangeHistory";

        if (-not(Test-Path -Path $ScriptPathChangeHistory))
        {
            Write-Verbose -Message "[$ScriptName][Begin] Creating the ChangeHistory Folder : $ScriptPathChangeHistory";
            New-Item -Path $ScriptPathChangeHistory -ItemType Directory | Out-Null;
        }

        # Set the Date and Time variables format
        $DateFormat = Get-Date -Format "ddMMMyyyy-HH:mm:ss";
        #$ReportDateFormat = Get-Date - Format "yyy\MM\dd HH:mm:ss"

        # Active Directory Module
        if (Get-Module -Name ActiveDirectory -ListAvailable) # verify AD Module is installed
        {
            Write-Verbose -Message "[$ScriptName][Begin] Active Directory Module";

            # Verify AD Module is loaded
            if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue -ErrorVariable ErrorBeginGetADModule))
            {
                Write-Verbose -Message "[$ScriptName][Begin] Active Directory Module - Loading";
                Import-Module -Name ActiveDirectory -ErrorAction Stop -ErrorVariable ErrorBeginAddADModule;
                Write-Verbose -Message "[$ScriptName][Begin] Active Directory Module - Loaded"
                $script:ADModule = $true;
            }
            else
            {
                Write-Verbose -Message "[$ScriptName][Begin] Active Directory module seems loaded";
                $script:ADModule = $true
            }
        }
        else # else try to load Quest AD Cmdlets
        {
            Write-Verbose -Message "[$ScriptName][Begin] Quest AD Snapin";
            if (-not (Get-PSSnapin -Name Quest.ActiveRoles.ADManagement -ErrorAction Stop -ErrorVariable ErrorBeginGetQuestAD))
            {
                Write-Verbose -Message "[$ScriptName][Begin] Quest Active Directory - Loading";
                Add-PSSnapin -Name Quest.ActiveRoles.ADManagement -ErrorAction Stop -ErrorVariable ErrorBeginAddQuestAD;
                Write-Verbose -Message "[$ScriptName][Begin] Quest Active Directory - Loaded";
                $script:QuestADSnapin = $true;
            }
            else
            {
                Write-Verbose "[$ScriptName][Begin] Quest AD Snapin seems loaded";
                $script:QuestADSnapin = $true;
            }
        }

        Write-Verbose -Message "[$ScriptName][Begin] Setting HTML Variables";

        # HTML Report Settings
        $Report = "<p style=`"background-color: white; font-family: consolas; font-size: 9pt;`">" +
        "<strong>Report Time:</strong> $DateFormat <br>" +
        "<strong>Account:</strong> $env:userdomain\$($env:username) on $($env:computername)" +
        "</p>";

        $Head = "<style>" +
        "body { background-color: white; font-family: consolas; font-size: 11pt; }" +
        "table { border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse; }" +
        "th { border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color:`"#00297a`"; font-color: white; }" +
        "td { border-width: 1px; padding-right: 2px; padding-left: 2px; padding-top: 0px; padding-bottom: 0px; border-style: solid; border-color: black; background-color: white; }" +
        "</style>";

        $Head2 = "<style>" +
        "body { background-color: white; font-family: consolas; font-size: 9pt; }" +
        "table { border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse; }" +
        "th { border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: `"#c0c0c0`"; }" +
        "td { border-width: 1px; padding-right: 2px; padding-left: 2px; padding-top: 0px; padding-bottom: 0px; border-style: solid; border-color: black; background-color: white; }" +
        "</style>";
    } #try
    catch
    {
        Write-Warning -Message "[$ScriptName][Begin] Something went wrong";

        # Quest AD Cmdlet Errors
        if ($ErrorBeginGetQuestAD) { Write-Warning -Message "[$ScriptName][Begin] Can't find the Quest Active Directory Snapin"; }
        if ($ErrorBeginAddQuestAD) { Write-Warning -Message "[$ScriptName][Begin] Can't load the Quest Active Directory Snapin"; }

        # AD Module Errors
        if ($ErrorBeginGetADModule) { Write-Warning -Message "[$ScriptName][Begin] Can't find the Active Directory Module"; }
        if ($ErrorBeginAddADModule) { Write-Warning -Message "[$ScriptName][Begin] Can't load the Active Directory Module"; }
    } #catch
} #Begin

Process
{
    try
    {
        # SearchRoot parameter specified
        if ($PSBoundParameters['SearchRoot'])
        {
            Write-Verbose -Message "[$ScriptName][Process] SearchRoot specified";

            foreach ($item in $SearchRoot)
            {
                # ADGroup Splatting
                $ADGroupParams = @{ };

                # Active Directory Module
                if ($ADModule)
                {
                    $ADGroupParams.SearchBase = $item;

                    # Server Specified
                    if ($PSBoundParameters['Server']) { $ADGroupParams.Server = $Server; }
                }

                if ($QuestADSnapin)
                {
                    $ADGroupParams.SearchRoot = $item;

                    # Server Specified
                    if ($PSBoundParameters['Server']) { $ADGroupParams.Service = $Server; }
                }

                # SearchScope parameter specified
                if ($PSBoundParameters['SearchScope'])
                {
                    Write-Verbose -Message "[$ScriptName][Process] SearchScope specified";

                    $ADGroupParams.SearchScope = $SearchScope;
                }

                # GroupScope parameter specified
                if ($PSBoundParameters['GroupScope'])
                {
                    Write-Verbose -Message "[$ScriptName][Process] GroupScope specified";

                    # Active Directory Module Parameter
                    if ($ADModule) { $ADGroupParams.Filter = "GroupScope -eq `'$GroupScope`'"; }
                    # Quest Active Directory Snapin Parameter
                    else { $ADGroupParams.GroupScope = $GroupScope; }
                }

                # GroupType parameter specified
                if ($PSBoundParameters['GroupType'])
                {
                    Write-Verbose -Message "[$ScriptName][Process] GroupType specified";

                    if ($ADModule)
                    {
                        if ($ADGroupParams.Filter)
                        {
                            $ADGroupParams.Filter = "$($ADGroupParams.Filter) -and GroupCategory -eq `'$GroupType`'";
                        }
                        else
                        {
                            $ADGroupParams.Filter = "GroupCategory -eq '$GroupType'";
                        }
                    }
                    # Quest Active Directory Snapin
                    else
                    {
                        $ADGroupParams.GroupType = $GroupType;
                    }
                }

                if ($ADModule)
                {
                    if (-not($ADGroupParams.Filter)) { $ADGroupParams.Filter = "*"; }

                    Write-Verbose -Message "[$ScriptName][Process] AD Module - Querying...";

                    # Add the groups to the variable $Group
                    $GroupSearch = Get-ADGroup @ADGroupParams;

                    if ($GroupSearch)
                    {
                        $Group += $GroupSearch.DistinguishedName;
                        Write-Verbose -Message "[$ScriptName][Process] OU: $item";
                    }
                }

                if ($QuestADSnapin)
                {
                    Write-Verbose -Message "[$ScriptName][Process] Quest AD Snapin - Querying...";

                    # Add the gruops to the variable $Group
                    $GroupSearchQuest = Get-QADGroup @ADGroupparams;

                    if ($GroupSearchQuest)
                    {
                        $Group += $GroupSearchQuest.DN;
                        Write-Verbose -Message "[$ScriptName][Process] OU: $item";
                    }
                }
            } # foreach ($item in $OU)
        } # if ($PSBoundParameters['SearchRoot'])

        # File parameter specified
        if ($PSBoundParameters['File'])
        {
            Write-Verbose -Message "[$ScriptName][Process] File";

            foreach ($item in $File)
            {
                Write-Verbose -Message "[$ScriptName][Process] Loading File: $item";

                $FileContent = Get-Content -Path $File;

                if ($FileContent)
                {
                    $Group += Get-Content -Path $File;
                }
            } # foreach ($item in $File)
        } # if ($PSBoundParameters['File'])

        # Group or SearchRoot or File parameters specified
        foreach ($item in $Group)
        {
            try
            {
                Write-Verbose -Message "[$ScriptName][Process] Group: $item...";

                # Splatting for the AD Group Request
                $GroupSplatting = @{ };
                $GroupSplatting.Identity = $item;

                # Group Information
                if ($ADModule)
                {
                    Write-Verbose -Message "[$ScriptName][Process] Active Directory Module";

                    if ($PSBoundParameters['Server']) { $GroupSplatting.Server = $Server; }

                    # Look for Group
                    $GroupName = Get-ADGroup @GroupSplatting -Properties * -ErrorAction Continue -ErrorVariable ErrorProcessGetADGroup;
                    Write-Verbose -Message "[$ScriptName][Process] Extracting Domain Name from $($GroupName.CanonicalName)";
                    $DomainName = ($GroupName.CanonicalName -split '/')[0];
                    $RealGroupName = $GroupName.Name;
                }
                if ($QuestADSnapin)
                {
                    Write-Verbose -Message "[$ScriptName][Process] Quest Active Directory Snapin";

                    # Add the Server if Specified
                    if ($PSBoundParameters['Server']) { $GroupSplatting.Service = $Server; }

                    # Look for Group
                    $GroupName = Get-QADGroup $GroupSplatting -ErrorAction Continue -ErrorVariable ErrorProcessGetQADGroup;
                    $DomainName = $($GroupName.domain.name);
                    $RealGroupName = $GroupName.name;
                }

                if ($GroupName)
                {
                    $GroupMemberSplatting = @{ };
                    $GroupMemberSplatting.Identity = $GroupName;

                    # Get GroupName Membership
                    if ($ADModule)
                    {
                        Write-Verbose -Message "[$ScriptName][Process] Group: $item - Querying Membership (AD Module)";

                        # Add the Server if specified
                        if ($PSBoundParameters['Server']) { $GroupMemberSplatting.Server = $Server; }

                        $MemberObjs = Get-ADGroupMember @GroupMemberSplatting -Recursive -ErrorAction Stop -ErrorVariable ErrorProcessGetADGroupMember;
                        [Array]$Members = $MemberObjs | Where-Object {$_.objectClass -eq "user" } | Get-ADUser -Properties PasswordExpired | Select-Object -Property *,@{ Name = 'DN'; Expression = { $_.DistinguishedName } };
                        $Members += $MemberObjs | Where-Object {$_.objectClass -eq "computer" } | Get-ADComputer -Properties PasswordExpired | Select-Object -Property *,@{ Name = 'DN'; Expression = { $_.DistinguishedName } };
                    }

                    if ($QuestADSnapin)
                    {
                        Write-Verbose -Message "[$ScriptName][Process] Group: $item - Querying Membership (Quest AD Snapin)";

                        # Add the SErver if specified
                        if ($PSBoundParameters['Server']) { $GroupMemberSplatting.Service = $Server; }

                        $Members = Get-QADGroupMember @GroupMemberSplatting -Indirect -ErrorAction Stop -ErrorVariable ErrorProcessGetQADGroupMember #| Select-Object -Property *,@{ Name = 'DistinguishedName'; Expression = { $_.dn } };
                    }

                    # no members, add some info in $members to avoid the $null
                    # if the value is $null the compare-object won't work
                    if (-not ($Members))
                    {
                        Write-Verbose -Message "[$ScriptName][Process] Gruop: $item is empty";

                        $Members = New-Object -TypeName PSObject -Property @{
                            Name = "No User or Group"
                            SamAccountName = "No User or Group"
                        };
                    }

                    # GroupName Membership File
                    # if the file doesn't exist, assume we don't have a record to refer to
                    $StateFile = "$($DomainName)_$($RealGroupName)-membership.csv";

                    if (-not (Test-Path -Path (Join-Path -Path $ScriptPathOutput -ChildPath $StateFile)))
                    {
                        Write-Verbose -Message "[$ScriptName][Process] $item - The following file did not exist: $StateFile";
                        Write-Verbose -Message "[$ScriptName][Process] $item - Exporting the current membership information into the file: $StateFile";

                        $Members | Export-Csv -Path (Join-Path -Path $ScriptPathOutput -ChildPath $StateFile) -NoTypeInformation -Encoding Unicode;
                    }
                    else
                    {
                        Write-Verbose -Message "[$ScriptName][Process] $item - The following file exists: $StateFile";
                    }

                    # GroupName Membership File is compared with the current GruopName Membership
                    Write-Verbose -Message "[$ScriptName][Process] $item - Comparing Current and Before";

                    $ImportCSV = Import-Csv -Path (Join-Path -Path $ScriptPathOutput -ChildPath $StateFile) -ErrorAction Stop -ErrorVariable ErrorProcessImportCSV;

                    $Changes = Compare-Object -DifferenceObject $ImportCSV -ReferenceObject $Members -ErrorAction Stop -ErrorVariable ErrorProcessCompareObject -Property Name, SamAccountName, DN |
                    Select-Object @{ Name = "DateTime"; Expression = { Get-Date -Format "yyyyMMdd-hh:mm:ss" } }, @{
                        n = 'State'; e = {
                            if ($_.SideIndicator -eq "=>") { "Removed" }
                            else { "Added" }
                        }
                    }, DisplayName, Name, SamAccountName, DN | Where-Object { $_.name -notlike "*no user or group*" }

                    Write-Verbose -Message "[$ScriptName][Process] $item - Compare Block Done!";

                    <# Troubleshooting
                    Write-Verbose -Message "[$ScriptName]IMPORTCSV var";
                    $ImportCSV | fl -Property Name, SamAccountName, DN;

                    Write-Verbose -Message "[$ScriptName]MEMBER";
                    $Members | fl -Property Name, SamAccountName, DN

                    Write-Verbose -Message "[$ScriptName]CHANGE";
                    $Changes;
                    #>

                    # Changes Found
                    if ($Changes -or $AlwaysReport)
                    {
                        Write-Verbose -Message "[$ScriptName][Process] $item - Some changes found";
                        $Changes | Select-Object -Property DateTime, State, Name, SamAccountName, DN;

                        # Change History
                        # Get the Past Changes History
                        Write-Verbose -Message "[$ScriptName][Process] $item - Get the change history for this group";
                        $ChangesHistoryFiles = Get-ChildItem -Path $ScriptPathChangeHistory\$($DomainName)_$($RealGroupName)-ChangeHistory.csv -ErrorAction SilentlyContinue;
                        Write-Verbose -Message "[$ScriptName][Process] $item - Change history files: $(($ChangesHistoryFiles|Measure-Object).Count)";

                        # Process each history changes
                        if ($ChangesHistoryFiles)
                        {
                            $InfoChangeHistory = @();

                            foreach ($file in $ChangesHistoryFiles.FullName)
                            {
                                Write-Verbose -Message "[$ScriptName][Process] $item - Change history files - Loading $file";

                                # Import the file and show the $file creation time and its content
                                $ImportedFile = Import-Csv -Path $file -ErrorAction Stop -ErrorVariable ErrorProcessImportCSVChangeHistory;

                                foreach ($obj in $ImportedFile)
                                {
                                    $Output = "" | Select-Object -Property DateTime, State, DisplayName, Name, SamAccountName, DN;
                                    $Output.DateTime = $obj.DateTime;
                                    $Output.State = $obj.State;
                                    $Output.DisplayName = $obj.DisplayName;
                                    $Output.Name = $obj.Name;
                                    $Output.SamAccountName = $obj.SamAccountName;
                                    $Output.DN = $obj.DN;
                                    $InfoChangeHistory = $InfoChangeHistory + $Output;
                                } # foreach ($obj in $ImportedFile)
                            } # foreach ($file in $ChangesHistoryFiles.FullName)

                            Write-Verbose -Message "[$ScriptName][Process] $item - Change History Process Completed";
                        } # if ($ChangeHistoryFiles)

                        if ($IncludeMembers)
                        {
                            $InfoMembers = @();
                            Write-Verbose -Message "[$ScriptName][Process] $item - Full Member List - Loading";

                            foreach ($obj in $Members)
                            {
                                $Output = "" | Select-Object -Property Name, SamAccountName, DN, Enabled, PasswordExpired;
                                $Output.Name = $obj.Name;
                                $Output.SamAccountName = $obj.SamAccountName;
                                $Output.DN = $obj.DistinguishedName;
                                
                                if ($ExtendedProperty)
                                {
                                    $Output.Enabled = $obj.Enabled;
                                    $Output.PasswordExpired = $obj.PasswordExpired;
                                }

                                $InfoMembers = $InfoMembers + $Output;
                            } #foreach ($obj in $Members)

                            Write-Verbose -Message "[$ScriptName][Process] $item - Full Member List Process Completed";
                        } # if ($IncludeMembers)

                        if ($Changes)
                        {
                            Write-Verbose -Message "[$ScriptName][Process] $item - Save Changes to a ChangesHistory File";

                            if (-not (Test-Path -Path (Join-Path -Path $ScriptPathChangeHistory -ChildPath "$($DomainName)_$($RealGroupName)-ChangeHistory.csv")))
                            {
                                $Changes | Export-Csv -Path (Join-Path -Path $ScriptPathChangeHistory -ChildPath "$($DomainName)_$($RealGroupName)-ChangeHistory.csv") -NoTypeInformation -Encoding Unicode;
                            }
                            else
                            {
                                $Changes | Export-Csv -Path (Join-Path -Path $ScriptPathChangeHistory -ChildPath "$($DomainName)_$($RealGroupName)-ChangeHistory.csv") -NoTypeInformation -Append -Encoding Unicode;
                            }
                        } # if ($Changes)

                        # Email
                        Write-Verbose -Message "[$ScriptName][Process] $item - Preparing the notification email...";

                        $EmailSubject = "PS MONITORING - $($GroupName.SamAccountName) Membership Change";

                        # Preparing the body of the email
                        $Body = "<h2>Group: $($GroupName.SamAccountName)</h2>";
                        $Body += "<p style=`"background-color: white; font-family: consolas; font-size: 8pt;`">";
                        $Body += "<u>Group Description:</u> $($GroupName.Description)<br>";
                        $Body += "<u>Group DistinguishedName:</u> $($GroupName.DistinguishedName)<br>";
                        $Body += "<u>Group CanonicalName:</u> $($GroupName.CanonicalName)<br>";
                        $Body += "<u>Group SID:</u> $($GroupName.Sid.Value)<br>";
                        $Body += "<u>Group Scope/Type:</u> $($GroupName.GroupScope) / $($GroupName.GroupType)<br>";
                        $Body += "</p>";

                        if ($Changes)
                        {
                            $Body += "<h3>Membership Change</h3>";
                            $Body += "<i>The membership of this group changed. See the following Added or Removed members.</i>";

                            # Removing the old DisplayName Property
                            $Changes = $Changes | Select-Object -Property DateTime, State, Name, SamAccountName, DN;

                            $Body += $changes | ConvertTo-Html -head $Head | Out-String;
                            $Body += "<br><br><br>";
                        } # if ($Changes)
                        else
                        {
                            $Body += "<h3>Membership Change</h3>";
                            $Body += "<i>No changes.</i>";
                        }

                        if ($InfoChangeHistory)
                        {
                            # Removing the old DisplayName Property
                            $InfoChangeHistory = $InfoChangeHistory | Select-Object -Property DateTime, State, Name, SamAccountName, DN;

                            $Body += "<h3>Change History</h3>";
                            $Body += "<i>List of previous changes on this group observed by the script</i>";
                            $Body += $InfoChangeHistory | Sort-Object -Property DateTime -Descending | ConvertTo-Html -Fragment -PreContent $Head2 | Out-String;
                        } # if ($InfoChangeHistory)

                        if ($InfoMembers)
                        {
                            $Body += "<h3>Members</h3>";
                            $Body += "<i>List of all members</i>";
                            $Body += $InfoMembers | Sort-Object -Property SamAccountName -Descending | ConvertTo-Html -Fragment -PreContent $Head2 | Out-String;
                        } # if ($InfoMembers)

                        $Body = $Body -replace "Added", "<font color=`"blue`"><b>Added</b></font>";
                        $Body = $Body -replace "Removed", "<font color=`"red`"><b>Removed</b></font>";
                        $Body += $Report;

                        if (-not ($OneReport))
                        {
                            $mailParam = @{
                                To = $EmailTo
                                From = $EmailFrom
                                Subject = $EmailSubject
                                Body = $Body
                                SmtpServer = $EmailServer
                                Port = $EmailPort
                                Credential = $EmailCred
                            }

                            Send-MailMessage @mailParam -UseSsl -BodyAsHtml;

                            Write-Verbose -Message "[$ScriptName][Process] $item - Email Sent.";
                        } # if (-not ($OneReport))

                        # GroupName Membership export to CSV
                        Write-Verbose -Message "[$ScriptName][Process] $item - Exporting the current membership to $StateFile";
                        $Members | Export-Csv -Path (Join-Path -Path $ScriptPathOutput -ChildPath $StateFile) -NoTypeInformation -Encoding Unicode;

                        if ($PSBoundParameters['HTMLLog'])
                        {
                            $ScriptPathHTML = $ScriptPath + "\HTML";

                            if (-not (Test-Path -Path $ScriptPathHTML))
                            {
                                Write-Verbose -Message "[$ScriptName][Process] Creating the HTML Folder : $ScriptPathHTML";
                                New-Item -Path $ScriptPathHTML -ItemType Directory | Out-Null;
                            }

                            # Define HTML File Name
                            $HTMLFileName = "$($DomainName)_$($RealGroupName)-$(Get-Date -Format 'yyyyMMdd_HHmmss').html";
                            $Body | Out-File -FilePath (Join-Path -Path $ScriptPathHTML -ChildPath $HTMLFileName);
                        } # if ($PSBoundParameters['HTMLLog'])

                        if ($OneReport)
                        {
                            # Create HTML Directory if it does not exist
                            $ScriptPathOneReport = $ScriptPath + "\OneReport";
                            if (-not (Test-Path -Path $ScriptPathOneReport))
                            {
                                Write-Verbose -Message "[$ScriptName][Process] Creating the OneReport Folder : $ScriptPathOneReport";
                                New-Item -Path $ScriptPathOneReport -ItemType Directory | Out-Null;
                            }

                            # Define HTML File Name
                            $HTMLFileName = "$($DomainName)_$($RealGroupName)-$(Get-Date -Format 'yyyyMMdd_HHmmss').html";

                            # Save HTML File
                            $Body | Out-File -FilePath (Join-Path -Path $ScriptPathOneReport -ChildPath $HTMLFileName);
                        } # if ($OneReport)
                    } # if ($Changes -or $AlwaysReport)
                    else
                    {
                        Write-Verbose -Message "[$ScriptName][Process] $item - No Change";
                    }
                } # if ($GroupName)
                else
                {
                    Write-Verbose -Message "[$ScriptName][Process] $item -Group can't be found";
                }
            } # try
            catch
            {
                Write-Warning -Message "[$ScriptName][Process] Something went wrong";

                # Quest Snapin Errors
                if ($ErrorProcessGetQADGroup) { Write-Warning -Message "[$ScriptName][Process] Quest AD - Error querying the group $item in Active Directory"; }
                if ($ErrorProcessGetQADGroupMember) { Write-Warning -Message "[$ScriptName][Process] Quest AD - Error querying the group $item members in Active Directory"; }

                # Active Directory Module Errors
                if ($ErrorProcessGetADGroup) { Write-Warning -Message "[$ScriptName][Process] AD Module - Error querying the gruop $item in Active Directory"; }
                if ($ErrorProcessGetADGroupMember) { Write-Warning -Message "[$ScriptName][Process] AD Module - Error querying the group $item members in Active Directory"; }

                # Import CSV Errors
                if ($ErrorProcessImportCSV) { Write-Warning -Message "[$ScriptName][Process] Error importing $StateFile"; }
                if ($ErrorProcessCompareObject) { Write-Warning -Message "[$ScriptName][Process] Error when comparing"; }
                if ($ErrorProcessImportCSVChangeHistory) { Write-Warning -Message "[$ScriptName][Process] Error importing $file"; }

                Write-Warning -Message $_.Exception.Message;
            } # catch
        } # foreach ($item in $Group)

        if ($OneReport)
        {
            [string[]]$attachments = @();

            foreach ($a in (Get-ChildItem $ScriptPathOneReport))
            {
                $attachments.Add($a.fullname);
            }
            
            $mailParam = @{
                To = $EmailTo
                From = $EmailFrom
                Subject = "PS Monitoring - Membership Report"
                Body = "<h2>See Report in Attachment</h2>"
                SmtpServer = $EmailServer
                Port = $EmailPort
                Credential = $EmailCred
                Attachments = $attachments
            }

            Send-MailMessage @mailParam -UseSsl -BodyAsHtml;
            
            foreach ($a in $attachments)
            {
                $a.Dispose();
            }

            Get-ChildItem $ScriptPathOneReport | Remove-Item -Force -Confirm:$false
            Write-Verbose -Message "[$ScriptName][Process] OneReport - Email Sent.";
        } # if ($OneReport)
    } #try
    catch
    {
        Write-Warning -Message "[$ScriptName][Process] Something went wrong";
        throw $_;
    } #catch
} #Process
End
{
    Write-Verbose -Message "[$ScriptName][End] Script Completed";
}