<#
.DESCRIPTION
       This is a script that import Receive Connectors from a CSV file on a server.
       You must use the CSV file created with the Export-ReceiveConnector.ps1 script

.EXAMPLE
       .\Import-ReceiveConnector.ps1 -Server E2016-01 -InputFile .\Sample_Export.csv
       Creates the Receive Connector using the data in the Sample_Export.csv file.

#>
[CmdletBinding()]
Param(
       [Parameter(Mandatory)][string]$Server,
       [Parameter(Mandatory)][string]$InputFile
)

Try{
       $ReceiveConnectors = Import-CSV $InputFile -ErrorAction Stop
} 
Catch {
       Write-Host "Error trying to open file $InputFile"
       Exit
}



Foreach ($MyConnector in $ReceiveConnectors){
       Write-Host "Creating Connector $($MyConnector.Name) on server $Server" -BackgroundColor Yellow -ForegroundColor Blue
       $Properties = @{    
       AuthTarpitInterval = $MyConnector.AuthTarpitInterval
       TarpitInterval = $MyConnector.TarpitInterval
       MaxAcknowledgementDelay = $MyConnector.MaxAcknowledgementDelay
       ConnectionInactivityTimeout = $MyConnector.ConnectionInactivityTimeout
       ConnectionTimeout = $MyConnector.ConnectionTimeout
       MaxInboundConnectionPercentagePerSource = $MyConnector.MaxInboundConnectionPercentagePerSource
       MaxLogonFailures = $MyConnector.MaxLogonFailures
       MaxProtocolErrors = $MyConnector.MaxProtocolErrors
       MaxLocalHopCount = $MyConnector.MaxLocalHopCount
       MaxInboundConnectionPerSource = $MyConnector.MaxInboundConnectionPerSource
       MaxHopCount = $MyConnector.MaxHopCount
       MaxRecipientsPerMessage = $MyConnector.MaxRecipientsPerMessage
       MaxInboundConnection = $MyConnector.MaxInboundConnection
       Name = $MyConnector.Name
 
       #DistinguishedName = $MyConnector.DistinguishedName #Parameter set automatically (combination of name, server, domain, ...)
       Fqdn = $MyConnector.Fqdn
       SizeEnabled = $MyConnector.SizeEnabled
       TransportRole = $MyConnector.TransportRole
       MessageRateSource = $MyConnector.MessageRateSource
       ExtendedProtectionPolicy = $MyConnector.ExtendedProtectionPolicy
       ProtocolLoggingLevel = $MyConnector.ProtocolLoggingLevel
       AuthMechanism = $MyConnector.AuthMechanism
       MessageRateLimit = $MyConnector.MessageRateLimit
       #AcceptConsumerMail =  [System.Convert]::ToBoolean($($MyConnector.AcceptConsumerMail)) #Parameter reserved for internal Microsoft use
       AdvertiseClientSettings =  [System.Convert]::ToBoolean($($MyConnector.AdvertiseClientSettings))
       #BareLinefeedRejectionEnabled =  [System.Convert]::ToBoolean($($MyConnector.BareLinefeedRejectionEnabled)) #Parameter only available with Set-ReceiveConnector
       DomainSecureEnabled =  [System.Convert]::ToBoolean($($MyConnector.DomainSecureEnabled))
       EnableAuthGSSAPI =  [System.Convert]::ToBoolean($($MyConnector.EnableAuthGSSAPI))
       #LiveCredentialEnabled =  [System.Convert]::ToBoolean($($MyConnector.LiveCredentialEnabled)) #Parameter reserved for internal Microsoft use
       LongAddressesEnabled =  [System.Convert]::ToBoolean($($MyConnector.LongAddressesEnabled))
       OrarEnabled =  [System.Convert]::ToBoolean($($MyConnector.OrarEnabled))
       #ProxyEnabled =  [System.Convert]::ToBoolean($($MyConnector.ProxyEnabled)) #Parameter reserved for internal Microsoft use
       RejectReservedSecondLevelRecipientDomains =  [System.Convert]::ToBoolean($($MyConnector.RejectReservedSecondLevelRecipientDomains))
       RejectReservedTopLevelRecipientDomains =  [System.Convert]::ToBoolean($($MyConnector.RejectReservedTopLevelRecipientDomains))
       RejectSingleLabelRecipientDomains =  [System.Convert]::ToBoolean($($MyConnector.RejectSingleLabelRecipientDomains))
       RequireEHLODomain =  [System.Convert]::ToBoolean($($MyConnector.RequireEHLODomain))
       RequireTLS =  [System.Convert]::ToBoolean($($MyConnector.RequireTLS))
       SuppressXAnonymousTls =  [System.Convert]::ToBoolean($($MyConnector.SuppressXAnonymousTls))
       BinaryMimeEnabled =  [System.Convert]::ToBoolean($($MyConnector.BinaryMimeEnabled))
       ChunkingEnabled =  [System.Convert]::ToBoolean($($MyConnector.ChunkingEnabled))
       DeliveryStatusNotificationEnabled =  [System.Convert]::ToBoolean($($MyConnector.DeliveryStatusNotificationEnabled))
       EightBitMimeEnabled =  [System.Convert]::ToBoolean($($MyConnector.EightBitMimeEnabled))
       Enabled =  [System.Convert]::ToBoolean($($MyConnector.Enabled))
       EnhancedStatusCodesEnabled =  [System.Convert]::ToBoolean($($MyConnector.EnhancedStatusCodesEnabled))
       PipeliningEnabled =  [System.Convert]::ToBoolean($($MyConnector.PipeliningEnabled))
       #SmtpUtf8Enabled =  [System.Convert]::ToBoolean($($MyConnector.SmtpUtf8Enabled)) #Parameter reserved for internal Microsoft use
       MaxHeaderSize = $($MyConnector.MaxHeaderSize) + "KB"
       MaxMessageSize = $($MyConnector.MaxMessageSize) + "MB"
       
}

IF(-not [string]::IsNullOrEmpty($($MyConnector.ServiceDiscoveryFqdn))){$properties.Add("ServiceDiscoveryFqdn",$($MyConnector.ServiceDiscoveryFqdn))}
IF(-not [string]::IsNullOrEmpty($($MyConnector.TlsDomainCapabilities))){$properties.Add("TlsDomainCapabilities" , $($MyConnector.TlsDomainCapabilities -join ";"))}
IF(-not [string]::IsNullOrEmpty($($MyConnector.Bindings))){$properties.Add("Bindings",$($MyConnector.Bindings -join ";"))}
IF(-not [string]::IsNullOrEmpty($($MyConnector.RemoteIPRanges))){$properties.Add("RemoteIPRanges",$($MyConnector.RemoteIPRanges -join ";"))}
IF(-not [string]::IsNullOrEmpty($($MyConnector.Banner))){$properties.Add("Banner",$($MyConnector.Banner))}
IF(-not [string]::IsNullOrEmpty($($MyConnector.Comment))){$properties.Add("Comment",$($MyConnector.Comment))}
IF(-not [string]::IsNullOrEmpty($($MyConnector.DefaultDomain))){$properties.Add("DefaultDomain",$($MyConnector.DefaultDomain))}
IF(-not [string]::IsNullOrEmpty($($MyConnector.TlsCertificateName))){$properties.Add("TlsCertificateName",$($MyConnector.TlsCertificateName))}

#Bindings = ($MyConnector.Bindings -join ";")
#RemoteIPRanges = ($MyConnector.RemoteIPRanges -join ";")
#TlsDomainCapabilities = ($MyConnector.TlsDomainCapabilities -join ";")
#Banner = $MyConnector.Banner
# Comment = $MyConnector.Comment
# DefaultDomain = $MyConnector.DefaultDomain
#ServiceDiscoveryFqdn = $MyConnector.ServiceDiscoveryFqdn
# TlsCertificateName = $MyConnector.TlsCertificateName

If ($($MyConnector.PermissionGroups) -match 'Custom'){
       Write-Host "Warning: This connector have AD Extended permissions like ms-Exch-SMTP-Accept-Any-Recipient or ms-Exch-SMTP-Accept-Any-Sender - please use Add-ADPermissions with the -ExtendedRights parameter" -ForegroundColor Yellow
} Else {
       IF(-not [string]::IsNullOrEmpty($($MyConnector.PermissionGroups))){$properties.Add("PermissionGroups",$($MyConnector.PermissionGroups))}
}

#PermissionGroups = $MyConnector.PermissionGroups


#$Properties

Try{
       New-ReceiveConnector -Server $Server @Properties -ErrorAction Stop
} catch {
       Write-Host "Issue creating this connector - connector NOT created - See below error:" -ForegroundColor red
       $ErrorMessage = $_.Exception.Message
       Write-Host "$ErrorMessage" -ForegroundColor DarkMagenta
}

}

Write-Host "Finished..." -ForegroundColor Green