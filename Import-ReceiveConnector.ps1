<#
.DESCRIPTION
       This is a script that import Receive Connectors from a CSV file on a server.
       
#>
[CmdletBinding()]
Param(
       [Parameter(Mandatory)][string]$Server,
       [Parameter(Mandatory)]$InputFile
)

Try{
       $ReceiveConnectors = Import-CSV $InputFile -ErrorAction Stop
} 
Catch {
       Write-Host "Error trying to open file $InputFile"
       Exit
}



Foreach ($MyConnector in $ReceiveConnectors){

    $Properties = @{    
                        Server = $Server              
                        AcceptConsumerMail = $MyConnector.AcceptConsumerMail
                        AdvertiseClientSettings = $MyConnector.AdvertiseClientSettings
                        AuthMechanism = $MyConnector.AuthMechanism
                        AuthTarpitInterval = $MyConnector.AuthTarpitInterval
                        Banner = $MyConnector.Banner
                        BareLinefeedRejectionEnabled = $MyConnector.BareLinefeedRejectionEnabled
                        BinaryMimeEnabled = $MyConnector.BinaryMimeEnabled
                        Bindings = ($MyConnector.Bindings -split ";")
                        ChunkingEnabled = $MyConnector.ChunkingEnabled
                        Comment = $MyConnector.Comment
                        ConnectionInactivityTimeout = $MyConnector.ConnectionInactivityTimeout
                        ConnectionTimeout = $MyConnector.ConnectionTimeout
                        DefaultDomain = $MyConnector.DefaultDomain
                        DeliveryStatusNotificationEnabled = $MyConnector.DeliveryStatusNotificationEnabled
                        DistinguishedName = $MyConnector.DistinguishedName
                        DomainSecureEnabled = $MyConnector.DomainSecureEnabled
                        EightBitMimeEnabled = $MyConnector.EightBitMimeEnabled
                        EnableAuthGSSAPI = $MyConnector.EnableAuthGSSAPI
                        Enabled = $MyConnector.Enabled
                        EnhancedStatusCodesEnabled = $MyConnector.EnhancedStatusCodesEnabled
                        ExtendedProtectionPolicy = $MyConnector.ExtendedProtectionPolicy
                        Fqdn = $MyConnector.Fqdn
                        LiveCredentialEnabled = $MyConnector.LiveCredentialEnabled
                        LongAddressesEnabled = $MyConnector.LongAddressesEnabled
                        MaxAcknowledgementDelay = $MyConnector.MaxAcknowledgementDelay
                        MaxHeaderSize = $($MyConnector.MaxHeaderSize/1KB)
                        MaxHopCount = $MyConnector.MaxHopCount
                        MaxInboundConnection = $MyConnector.MaxInboundConnection
                        MaxInboundConnectionPercentagePerSource = $MyConnector.MaxInboundConnectionPercentagePerSource
                        MaxInboundConnectionPerSource = $MyConnector.MaxInboundConnectionPerSource
                        MaxLocalHopCount = $MyConnector.MaxLocalHopCount
                        MaxLogonFailures = $MyConnector.MaxLogonFailures
                        MaxMessageSize = $($MyConnector.MaxMessageSize/1MB)
                        MaxProtocolErrors = $MyConnector.MaxProtocolErrors
                        MaxRecipientsPerMessage = $MyConnector.MaxRecipientsPerMessage
                        MessageRateLimit = $MyConnector.MessageRateLimit
                        MessageRateSource = $MyConnector.MessageRateSource
                        Name = $MyConnector.Name
                        OrarEnabled = $MyConnector.OrarEnabled
                        PermissionGroups = $MyConnector.PermissionGroups
                        PipeliningEnabled = $MyConnector.PipeliningEnabled
                        ProtocolLoggingLevel = $MyConnector.ProtocolLoggingLevel
                        ProxyEnabled = $MyConnector.ProxyEnabled
                        RejectReservedSecondLevelRecipientDomains = $MyConnector.RejectReservedSecondLevelRecipientDomains
                        RejectReservedTopLevelRecipientDomains = $MyConnector.RejectReservedTopLevelRecipientDomains
                        RejectSingleLabelRecipientDomains = $MyConnector.RejectSingleLabelRecipientDomains
                        RemoteIPRanges = ($MyConnector.RemoteIPRanges -join ";")
                        RequireEHLODomain = $MyConnector.RequireEHLODomain
                        RequireTLS = $MyConnector.RequireTLS
                        ServiceDiscoveryFqdn = $MyConnector.ServiceDiscoveryFqdn
                        SizeEnabled = $MyConnector.SizeEnabled
                        SmtpUtf8Enabled = $MyConnector.SmtpUtf8Enabled
                        SuppressXAnonymousTls = $MyConnector.SuppressXAnonymousTls
                        TarpitInterval = $MyConnector.TarpitInterval
                        TlsCertificateName = $MyConnector.TlsCertificateName
                        TlsDomainCapabilities = ($MyConnector.TlsDomainCapabilities -join ";")
                        TransportRole = $MyConnector.TransportRole
    }

New-ReceiveConnector @Properties
}

Write-Host "Finished..." -ForegroundColor Green