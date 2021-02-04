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
        $Properties = @{}
        If (-not [String]::IsNullOrEmpty($MyConnector.AuthTarpitInterval)){$Properties.Add("AuthTarpitInterval",$MyConnector.AuthTarpitInterval)}
        If (-not [String]::IsNullOrEmpty($MyConnector.TarpitInterval)){$Properties.Add("TarpitInterval",$MyConnector.TarpitInterval)}
        If (-not [String]::IsNullOrEmpty($MyConnector.MaxAcknowledgementDelay)){$Properties.Add("MaxAcknowledgementDelay",$MyConnector.MaxAcknowledgementDelay)}
        If (-not [String]::IsNullOrEmpty($MyConnector.ConnectionInactivityTimeout)){$Properties.Add("ConnectionInactivityTimeout",$MyConnector.ConnectionInactivityTimeout)}
        If (-not [String]::IsNullOrEmpty($MyConnector.ConnectionTimeout)){$Properties.Add("ConnectionTimeout",$MyConnector.ConnectionTimeout)}
        If (-not [String]::IsNullOrEmpty($MyConnector.MaxInboundConnectionPercentagePerSource)){$Properties.Add("MaxInboundConnectionPercentagePerSource",$MyConnector.MaxInboundConnectionPercentagePerSource)}
        If (-not [String]::IsNullOrEmpty($MyConnector.MaxLogonFailures)){$Properties.Add("MaxLogonFailures",$MyConnector.MaxLogonFailures)}
        If (-not [String]::IsNullOrEmpty($MyConnector.MaxProtocolErrors)){$Properties.Add("MaxProtocolErrors",$MyConnector.MaxProtocolErrors)}
        If (-not [String]::IsNullOrEmpty($MyConnector.MaxLocalHopCount)){$Properties.Add("MaxLocalHopCount",$MyConnector.MaxLocalHopCount)}
        If (-not [String]::IsNullOrEmpty($MyConnector.MaxInboundConnectionPerSource)){$Properties.Add("MaxInboundConnectionPerSource",$MyConnector.MaxInboundConnectionPerSource)}
        If (-not [String]::IsNullOrEmpty($MyConnector.MaxHopCount)){$Properties.Add("MaxHopCount",$MyConnector.MaxHopCount)}
        If (-not [String]::IsNullOrEmpty($MyConnector.MaxRecipientsPerMessage)){$Properties.Add("MaxRecipientsPerMessage",$MyConnector.MaxRecipientsPerMessage)}
        If (-not [String]::IsNullOrEmpty($MyConnector.MaxInboundConnection)){$Properties.Add("MaxInboundConnection",$MyConnector.MaxInboundConnection)}
        If (-not [String]::IsNullOrEmpty($MyConnector.Name)){$Properties.Add("Name",$MyConnector.Name)}
        #If (-not [String]::IsNullOrEmpty($MyConnector.Fqdn)){$Properties.Add("Fqdn",$MyConnector.Fqdn)}
        If (-not [String]::IsNullOrEmpty($MyConnector.SizeEnabled)){$Properties.Add("SizeEnabled",$MyConnector.SizeEnabled)}
        If (-not [String]::IsNullOrEmpty($MyConnector.TransportRole)){$Properties.Add("TransportRole",$MyConnector.TransportRole)}
        If (-not [String]::IsNullOrEmpty($MyConnector.MessageRateSource)){$Properties.Add("MessageRateSource",$MyConnector.MessageRateSource)}
        If (-not [String]::IsNullOrEmpty($MyConnector.ExtendedProtectionPolicy)){$Properties.Add("ExtendedProtectionPolicy",$MyConnector.ExtendedProtectionPolicy)}
        If (-not [String]::IsNullOrEmpty($MyConnector.ProtocolLoggingLevel)){$Properties.Add("ProtocolLoggingLevel",$MyConnector.ProtocolLoggingLevel)}
        If (-not [String]::IsNullOrEmpty($MyConnector.AuthMechanism)){
        
            If ($MyConnector.AuthMechanism -match "ExchangeServer"){$Properties.Add("Fqdn",$Server)}
            $Properties.Add("AuthMechanism",$MyConnector.AuthMechanism)}        
        
        
        If (-not [String]::IsNullOrEmpty($MyConnector.MessageRateLimit)){$Properties.Add("MessageRateLimit",$MyConnector.MessageRateLimit)}
        If (-not [String]::IsNullOrEmpty($($MyConnector.AdvertiseClientSettings))){$Properties.Add("AdvertiseClientSettings",[System.Convert]::ToBoolean($($MyConnector.AdvertiseClientSettings)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.DomainSecureEnabled))){$Properties.Add("DomainSecureEnabled",[System.Convert]::ToBoolean($($MyConnector.DomainSecureEnabled)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.EnableAuthGSSAPI))){$Properties.Add("EnableAuthGSSAPI",[System.Convert]::ToBoolean($($MyConnector.EnableAuthGSSAPI)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.LongAddressesEnabled))){$Properties.Add("LongAddressesEnabled",[System.Convert]::ToBoolean($($MyConnector.LongAddressesEnabled)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.OrarEnabled))){$Properties.Add("OrarEnabled",[System.Convert]::ToBoolean($($MyConnector.OrarEnabled)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.RejectReservedSecondLevelRecipientDomains))){$Properties.Add("RejectReservedSecondLevelRecipientDomains",[System.Convert]::ToBoolean($($MyConnector.RejectReservedSecondLevelRecipientDomains)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.RejectReservedTopLevelRecipientDomains))){$Properties.Add("RejectReservedTopLevelRecipientDomains",[System.Convert]::ToBoolean($($MyConnector.RejectReservedTopLevelRecipientDomains)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.RejectSingleLabelRecipientDomains))){$Properties.Add("RejectSingleLabelRecipientDomains",[System.Convert]::ToBoolean($($MyConnector.RejectSingleLabelRecipientDomains)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.RequireEHLODomain))){$Properties.Add("RequireEHLODomain",[System.Convert]::ToBoolean($($MyConnector.RequireEHLODomain)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.RequireTLS))){$Properties.Add("RequireTLS",[System.Convert]::ToBoolean($($MyConnector.RequireTLS)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.SuppressXAnonymousTls))){$Properties.Add("SuppressXAnonymousTls",[System.Convert]::ToBoolean($($MyConnector.SuppressXAnonymousTls)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.BinaryMimeEnabled))){$Properties.Add("BinaryMimeEnabled",[System.Convert]::ToBoolean($($MyConnector.BinaryMimeEnabled)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.ChunkingEnabled))){$Properties.Add("ChunkingEnabled",[System.Convert]::ToBoolean($($MyConnector.ChunkingEnabled)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.DeliveryStatusNotificationEnabled))){$Properties.Add("DeliveryStatusNotificationEnabled",[System.Convert]::ToBoolean($($MyConnector.DeliveryStatusNotificationEnabled)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.EightBitMimeEnabled))){$Properties.Add("EightBitMimeEnabled",[System.Convert]::ToBoolean($($MyConnector.EightBitMimeEnabled)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.Enabled))){$Properties.Add("Enabled",[System.Convert]::ToBoolean($($MyConnector.Enabled)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.EnhancedStatusCodesEnabled))){$Properties.Add("EnhancedStatusCodesEnabled",[System.Convert]::ToBoolean($($MyConnector.EnhancedStatusCodesEnabled)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.PipeliningEnabled))){$Properties.Add("PipeliningEnabled",[System.Convert]::ToBoolean($($MyConnector.PipeliningEnabled)))}
        If (-not [String]::IsNullOrEmpty($($MyConnector.MaxHeaderSize) + "KB")){$Properties.Add("MaxHeaderSize",$($MyConnector.MaxHeaderSize) + "KB")}
        If (-not [String]::IsNullOrEmpty($($MyConnector.MaxMessageSize) + "MB")){$Properties.Add("MaxMessageSize",$($MyConnector.MaxMessageSize) + "MB")}
        IF(-not [string]::IsNullOrEmpty($($MyConnector.ServiceDiscoveryFqdn))){$properties.Add("ServiceDiscoveryFqdn",$($MyConnector.ServiceDiscoveryFqdn))}
        IF(-not [string]::IsNullOrEmpty($($MyConnector.TlsDomainCapabilities))){$properties.Add("TlsDomainCapabilities" , $($MyConnector.TlsDomainCapabilities -split ";"))}
        IF(-not [string]::IsNullOrEmpty($($MyConnector.Bindings))){$properties.Add("Bindings",$($MyConnector.Bindings -join ";"))}
        IF(-not [string]::IsNullOrEmpty($($MyConnector.RemoteIPRanges))){$properties.Add("RemoteIPRanges",$($MyConnector.RemoteIPRanges -split ";"))}
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