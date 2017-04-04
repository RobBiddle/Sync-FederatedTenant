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
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$false,ParameterSetName = "Set 1")]
        [Parameter(ParameterSetName = "Set 2")]
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
                $_ | Add-Member -MemberType NoteProperty -Name ImmutableId -Value $immutableID
            }

            # Add ExistsIn365 property to objects and set to $false
            $UsersToSync | ForEach-Object {
                $_ | Add-Member -MemberType NoteProperty -Name ExistsIn365 -Value $false
            }
            # Change SyncedTo365 value to true for objects that exist in both AD & 365
            $UsersToSync | Where-Object UserPrincipalName -In $Users365.UserPrincipalName | ForEach-Object {
                    $_.ExistsIn365 = $true
                }

            # Add SyncComplete property to objects and set to $false
            $UsersToSync | ForEach-Object {
                $_ | Add-Member -MemberType NoteProperty -Name SyncComplete -Value $false
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

    # Connect to Active Directory
    $DomainControllerPSSession = New-PSSession -Name $DomainControllerFQDN -ComputerName $DomainControllerFQDN -Credential $CredentialForActiveDirectory
    Import-Module ActiveDirectory -PSSession $DomainControllerPSSession -Global

    # Get lists of Office365 Tenants
    switch ($PSCmdlet.ParameterSetName) {
        "Set 1" {$tenants = Get-MsolPartnerContract -All}
        "Set 2" {$tenants = (Get-MsolPartnerContract -All) | Where-Object TenantId -eq $TenantID}
    }
    # Add Federated Domains to objects
    $tenants | ForEach-Object {
        $_ | Add-Member -MemberType NoteProperty -Name FederatedDomain `
            -Value "$((Get-MsolDomain -TenantId $_.TenantId.GUID | Where-Object Authentication -eq Federated).Name)"
        $_ | Add-Member -MemberType NoteProperty -Name DomainStatus `
            -Value "$((Get-MsolDomain -TenantId $_.TenantId.GUID | Where-Object Authentication -eq Federated).Status)"
    }

    # Get list of Tenants with Federated Domains
    If($FederatedDomain) {
        $federatedTenants = $tenants | Where-Object FederatedDomain -Like $FederatedDomain 
    }
    if ($AllTenants) {
        $federatedTenants = $tenants | Where-Object FederatedDomain | Where-Object DomainStatus -like 'Verified'
    }Else{
        Write-Error -Message "No -FederatedDomain Specified"
    }
    # Sync Stuff
    $federatedTenants | ForEach-Object {
        $currentTenant = $_
        $currentTenantId = $currentTenant.TenantId.GUID
        $FederatedDomain = $currentTenant.FederatedDomain
        $FederatedDomain | ForEach-Object {
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
                if($_.DisplayName -notlike $currentUser365.DisplayName){
                    Set-MsolUser -TenantId $currentTenantId -UserPrincipalName $_.UserPrincipalName -DisplayName $_.DisplayName
                }
                if($_.GivenName -notlike $currentUser365.FirstName){
                    Set-MsolUser -TenantId $currentTenantId -UserPrincipalName $_.UserPrincipalName -FirstName $_.GivenName
                }
                if($_.Surname -notlike $currentUser365.LastName ){
                    Set-MsolUser -TenantId $currentTenantId -UserPrincipalName $_.UserPrincipalName -DisplayName $_.Surname
                }
                if($_.immutableID -notlike $currentUser365.immutableID){
                    Set-MsolUser -TenantId $currentTenantId -UserPrincipalName $_.UserPrincipalName -DisplayName $_.immutableID
                }
            }
            $UsersToSync = Get-UsersToSync -TenantID $currentTenantId -FederatedDomain $currentTenantFederatedDomain
            Write-Output "Sync of $currentTenantFederatedDomain complete"
            $UsersToSync | Sort-Object ExistsIn365, SyncComplete, DisplayName | Format-Table UserPrincipalName, DisplayName, ExistsIn365, SyncComplete
        }
    }
}
