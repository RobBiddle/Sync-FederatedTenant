# Sync-FederatedTenant
Syncs Active Directory with one or all Office 365 Partner Federated Domains

For use in scenarios matching the following conditions:
1. User email is NOT in Office 365 mailbox

This is for syncing user account data from AD to O365 for use with Federated Authentication.  
A Random password is set on the Office 365 user account since it is not used (authentication is redirected to federation server)
This does not cause problems unless using Office 365 services which do not support federated authentication
