<#
.NOTES
    Author: Robert D. Biddle
    Date: 01/27/2017
.Synopsis
    Custom Active Directory Sync process for Office 365 Federated Tenants
.DESCRIPTION
    Finds all Domains using Federated authentication in Office365
     and then searches Active Directory for users with UPNs matching a federated domain
     and then attempts to sync those users 
.EXAMPLE
#>
function Global:Sync-FederatedTenant {
    [CmdletBinding(DefaultParametersetName="Set 1")]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # Active Directory Domain Controller FQDN
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$false,ParameterSetName = "Set 1")]
        [parameter(ParameterSetName = "Set 2")]
        [Parameter(HelpMessage="FQDN of Active Directory Domain Controller")]
        [String]
        $DomainControllerFQDN,

        # Credential for Active Directory
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$false,ParameterSetName = "Set 1")]
        [parameter(ParameterSetName = "Set 2")]
        [Parameter(HelpMessage="PSCredential object for Active Directory")]
        [PSCredential]
        $CredentialForActiveDirectory = (Get-Credential),

        # Credential for Office365
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$false,ParameterSetName = "Set 1")]
        [parameter(ParameterSetName = "Set 2")]
        [Parameter(HelpMessage="PSCredential object for Office 365")]
        [PSCredential]
        $CredentialForOffice365,

        # Search All Tenants for Federated Domains
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$false,ParameterSetName = "Set 1")]
        [Parameter(HelpMessage="Search ALL Partner Tenants for Federated Domains and Sync Related Users")]
        [Switch]
        $AllTenants,

        # Sync Users associated with specified Federated Domain only
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$false,ParameterSetName = "Set 2")]
        [Parameter(HelpMessage="Sync only users with a UserPrincipalName matching this Federated Domain")]
        [String]
        $FederatedDomain,

        # Search Only Specified Tenant for Federated Domains
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$false,ParameterSetName = "Set 2")]
        #[ValidatePattern("^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$")]
        [Parameter(HelpMessage="Sync only Tenant specified by TenantID.  TenantID must be a valid GUID string")]
        [String]
        $TenantID
    )
    function Local:Get-UsersToSync ($TenantID, $FederatedDomain) {
            $Users365 = Get-MsolUser -TenantId $TenantID
            $filter = "userPrincipalName -like `"*$($FederatedDomain)`""
            $UsersAD =  Get-ADUser -Properties * -Filter $filter
            $UsersToSync = $UsersAD
            # Add immutableID property to objects and populate values
            $UsersToSync | ForEach-Object {
                $guid = $_.ObjectGUID
                $immutableID = [System.Convert]::ToBase64String($guid.tobytearray())
                $_ | Add-Member -MemberType NoteProperty -Name ImmutableId -Value $immutableID -force -ErrorAction SilentlyContinue
            }

            # Add ExistsIn365 property to objects and set to $false
            $UsersToSync | ForEach-Object {
                $_ | Add-Member -MemberType NoteProperty -Name ExistsIn365 -Value $false -force -ErrorAction SilentlyContinue
            }
            # Change SyncedTo365 value to true for objects that exist in both AD & 365
            $UsersToSync | Where-Object UserPrincipalName -In $Users365.UserPrincipalName | ForEach-Object {
                    $_.ExistsIn365 = $true
                }
            $UsersToSync | Where-Object ImmutableId -In $Users365.ImmutableId | ForEach-Object {
                    $_.ExistsIn365 = $true
                }

            # Add SyncComplete property to objects and set to $false
            $UsersToSync | ForEach-Object {
                $_ | Add-Member -MemberType NoteProperty -Name SyncComplete -Value $false -force -ErrorAction SilentlyContinue
            }
            # Change SyncComplete value to true after verifying synced attributes
            $UsersToSync | Where-Object UserPrincipalName -In $Users365.UserPrincipalName | ForEach-Object {
                    $currentUser = $_
                    $matching365User = $Users365 | Where-Object UserPrincipalName -eq $currentUser.UserPrincipalName
                    if ( ($currentUser.GivenName -like $matching365User.FirstName) `
                        -and ($currentUser.Surname -like $matching365User.LastName) `
                        -and ($currentUser.DisplayName -like $matching365User.DisplayName) `
                        -and ($currentUser.ImmutableId -like $matching365User.ImmutableId)                
                        ){
                            $currentUser.SyncComplete = $true
                        }
                }
            # Return Objects
            $UsersToSync
    }

    # Add in secure credential handling for automated use here...
    # $CredentialForOffice365 = 
    # $CredentialForActiveDirectory = 
    # Connect-MsolService -Credential $CredentialForOffice365
    #
    Add-Type -AssemblyName System.Web # Provides support for generating random passwords

    # Test for existing MsolService connection
    if ((Get-MsolCompanyInformation -ErrorAction SilentlyContinue) -ne $true) {
        # Connect to Office365
        if($CredentialForOffice365){
            Connect-MsolService -Credential $CredentialForOffice365
        }Else{Connect-MsolService}
    }
    # Check for ActiveDirectory Module
    if (!(Get-Module ActiveDirectory)) {
        # Try to import ActiveDirectory
        Import-Module ActiveDirectory
        if (!(Import-Module ActiveDirectory -ErrorAction SilentlyContinue)) {
            # Connect to Active Directory to import ActiveDirectory module if needed - This would require DomainAdmin while Success Above does not
            $DomainControllerPSSession = New-PSSession -Name $DomainControllerFQDN -ComputerName $DomainControllerFQDN -Credential $CredentialForActiveDirectory
            Import-Module ActiveDirectory -PSSession $DomainControllerPSSession -Global
        }
    }
    # Get lists of Office365 Tenants
    switch ($PSCmdlet.ParameterSetName) {
        "Set 1" {$tenants = Get-MsolPartnerContract -All}
        "Set 2" {$tenants = (Get-MsolPartnerContract -All) | Where-Object TenantId -eq $TenantID}
    }
    # Add Federated Domains to objects
    $tenants | Add-Member -MemberType NoteProperty -Name FederatedDomain  -Value $null
    $tenants | ForEach-Object {
        $currentTenant = $_
        $currentTenantId = $currentTenant.TenantId.GUID
        $DomainList = Get-MsolDomain -TenantId $currentTenantId | Where-Object Authentication -eq Federated
        $currentFederatedDomains = @()
        $DomainList | ForEach-Object {
            $currentDomain = $_
            $obj = New-Object -TypeName psobject
            $obj | Add-Member -NotePropertyName Name -NotePropertyValue $currentDomain.Name
            $obj | Add-Member -NotePropertyName Status -NotePropertyValue $currentDomain.Status
            $obj | Add-Member -NotePropertyName Authentication -NotePropertyValue $currentDomain.Authentication
            $currentFederatedDomains += $obj
        }
        $currentTenant.FederatedDomain = $currentFederatedDomains
    }
    # Attempt to Confirm non-verified Domains
    $tenants | Where-Object {$_.FederatedDomain.Authentication -like 'Federated' -and $_.FederatedDomain.Status -NotLike 'Verified'} | ForEach-Object {
        $currentTenant = $_
        $currentTenantId = $currentTenant.TenantId.GUID
        $currentDomainList = $currentTenant.FederatedDomain.Name
        $currentDomainList | ForEach-Object {
            $currentDomain = $_
            If(Get-MsolDomainVerificationDns -DomainName $currentDomain -TenantId $currentTenantId){
                $TXTrecordToSet = (Get-MsolDomainVerificationDns -DomainName $currentDomain -TenantId $currentTenantId -Mode DnsTxtRecord).Text
                Write-Output -Message "$($currentDomain) TXT Record of $TXTrecordToSet has not been verified, attempting verification now..."
                Confirm-MsolDomain -TenantId $currentTenantId -DomainName $currentDomain -ErrorAction SilentlyContinue
                <#  if ((Get-MsolDomain -DomainName $currentDomain -TenantId $currentTenantId).Status -notlike 'Verified' -and (Get-MsolDomain -DomainName $currentDomain -TenantId $currentTenantId).Authentication -like 'Federated') {
                    Set-MsolDomainAuthentication -DomainName $currentDomain -TenantId $currentTenantId -Authentication Managed
                    Confirm-MsolDomain -TenantId $currentTenantId -DomainName $currentDomain
                    Set-MsolDomainAuthentication -DomainName $currentDomain -TenantId $currentTenantId -Authentication Federated
                }#>
            }
        }
    }
    # Get list of Tenants with Federated Domains
    If($FederatedDomain) {
        $federatedTenants = $tenants | Where-Object {$_.FederatedDomain.Name -Like $FederatedDomain}
    }
    if ($AllTenants) {
        $federatedTenants = $tenants | Where-Object {$_.FederatedDomain.Status -like 'Verified'}
    }Else{
        Write-Error -Message "No -FederatedDomain Specified"
    }
    # Sync Stuff
    $federatedTenants | ForEach-Object {
        $currentTenant = $_
        $currentTenantId = $currentTenant.TenantId.GUID
        $currentDomainList = $currentTenant.FederatedDomain.Name
        Write-Output "Starting work on Tenant: $currentTenantId"
        Write-Output "Tenant has the following Federated Domains: $currentDomainList"
        $currentDomainList| ForEach-Object {
            $currentTenantFederatedDomain = $_
            $UsersToSync = Get-UsersToSync -TenantID $currentTenantId -FederatedDomain $currentTenantFederatedDomain
            Write-Output "Starting Sync of $currentTenantFederatedDomain"
            $UsersToSync | Sort-Object ExistsIn365, SyncComplete, DisplayName | Format-Table UserPrincipalName, DisplayName, ExistsIn365, SyncComplete
            # Create New Office365 Users for each user that has not been synced
            $UsersToSync | Where-Object ExistsIn365 -eq $false | ForEach-Object {
                $password = [System.Web.Security.Membership]::GeneratePassword(16,0)
                New-MsolUser -TenantId $currentTenant.TenantId.GUID `
                    -UserPrincipalName "$($_.UserPrincipalName)" `
                    -DisplayName "$($_.DisplayName)" `
                    -FirstName "$($_.GivenName)" `
                    -LastName "$($_.Surname)" `
                    -ImmutableId $_.immutableID `
                    -Password $password
            }
            # Sync attributes for Users that are not fully synced
            $UsersToSync | Where-Object ExistsIn365 -eq $true | Where-Object SyncComplete -eq $false | ForEach-Object {
                $currentUser365 = Get-MsolUser -TenantId $currentTenantId -UserPrincipalName $_.UserPrincipalName
                if(!($currentUser365)){
                    $currentUser365 = Get-MsolUser -TenantId $currentTenantId | Where-Object ImmutableId -Match $_.ImmutableId
                }
                if($_.DisplayName -notlike $currentUser365.DisplayName){
                    Set-MsolUser -TenantId $currentTenantId -UserPrincipalName $_.UserPrincipalName -DisplayName $_.DisplayName
                }
                if($_.GivenName -notlike $currentUser365.FirstName){
                    Set-MsolUser -TenantId $currentTenantId -UserPrincipalName $_.UserPrincipalName -FirstName $_.GivenName
                }
                if($_.Surname -notlike $currentUser365.LastName ){
                    Set-MsolUser -TenantId $currentTenantId -UserPrincipalName $_.UserPrincipalName -LastName $_.Surname
                }
                if($_.immutableID -notlike $currentUser365.immutableID){
                    Set-MsolUser -TenantId $currentTenantId -UserPrincipalName $_.UserPrincipalName -ImmutableId $_.immutableID
                }
            }
            $UsersToSync = Get-UsersToSync -TenantID $currentTenantId -FederatedDomain $currentTenantFederatedDomain
            Write-Output "Sync of $currentTenantFederatedDomain complete"
            $UsersToSync | Sort-Object ExistsIn365, SyncComplete, DisplayName | Format-Table UserPrincipalName, DisplayName, ExistsIn365, SyncComplete
        }
    }
}
