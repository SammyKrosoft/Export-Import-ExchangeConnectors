<#PSScriptInfo

.VERSION 1.0

.GUID 31c04a6c-bbb1-4dd6-9204-5e6b89c2d426

.AUTHOR sammy

.COMPANYNAME 

.COPYRIGHT 

.TAGS 

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES

#>

<#
.DESCRIPTION
       This is a script that exports Receive Connectors from a given server into
       a CSV file. Multivalued properties are joined with a semicolon ";".
       Import script will split the joined properties using ";" and add to the 
       destination Receive Connector object.

       Note: the default export file will be saved in the current user's "My Documents" folder.
       Use the -OutputFile parameter to change the default output file path and name.

.EXAMPLE
       .\Export-ReceiveConnector.ps1 -Server E2016-01 -OutputFile .\Sample_Export.csv
       Exports all Receive Connectors of server E2016-01 into a file named Sample_Export.csv file on the current directory.
       This CSV file can be used with the Import-ReceiveConnectors.ps1 script to create the same connector on different servers.

#>
[CmdletBinding()]
Param(
       [Parameter(Mandatory = $true)][string]$Server,
       [Parameter(Mandatory = $false)][string]$OutputFile = "$($env:USERPROFILE)\Documents\ReceiveConnectorExport_$Server_$(Get-Date -Format yyyy-MM-dd-hh-mm-ss).csv"
)

$CustomObjectColl = @()

$ReceiveConnectors = Get-ReceiveConnector

Foreach ($MyConnector in $ReceiveConnectors){

    $Properties = @{    AcceptConsumerMail = $MyConnector.AcceptConsumerMail
                        AdvertiseClientSettings = $MyConnector.AdvertiseClientSettings
                        AuthMechanism = $MyConnector.AuthMechanism
                        AuthTarpitInterval = $MyConnector.AuthTarpitInterval
                        Banner = $MyConnector.Banner
                        BareLinefeedRejectionEnabled = $MyConnector.BareLinefeedRejectionEnabled
                        BinaryMimeEnabled = $MyConnector.BinaryMimeEnabled
                        Bindings = ($MyConnector.Bindings -join ";")
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
                        MaxHeaderSize = $MyConnector.MaxHeaderSize.ToKB()
                        MaxHopCount = $MyConnector.MaxHopCount
                        MaxInboundConnection = $MyConnector.MaxInboundConnection
                        MaxInboundConnectionPercentagePerSource = $MyConnector.MaxInboundConnectionPercentagePerSource
                        MaxInboundConnectionPerSource = $MyConnector.MaxInboundConnectionPerSource
                        MaxLocalHopCount = $MyConnector.MaxLocalHopCount
                        MaxLogonFailures = $MyConnector.MaxLogonFailures
                        MaxMessageSize = $MyConnector.MaxMessageSize.ToMB()
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

$CustomObject = New-Object -TypeName PSObject -Prop $Properties
$CustomObjectColl += $CustomObject
}

$CustomObjectColl | Export-csv -NoTypeInformation -Path $OutputFile

Notepad $OutputFile