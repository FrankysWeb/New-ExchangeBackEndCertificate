<#
.SYNOPSIS
	Renew Exchange BackEnd certificate
.DESCRIPTION
	Creates a new self signed certificated for Exchange Server BackEnd
.EXAMPLE
	.\New-ExchangeBackEndCertificate.ps1
.NOTES
	Author:  Frank Zoechling
	Website: https://www.frankysweb.de
	Twitter: @FrankysWeb
#>

write-host "Create new Exchange BackEnd certificate..." -foregroundcolor yellow
$NewCert = Get-ExchangeCertificate | where {$_.FriendlyName -eq "Microsoft Exchange" -and $_.IsSelfSigned -eq $true} | New-ExchangeCertificate -Force -PrivateKeyExportable $true
$NewCert
$Thumbprint = $NewCert.Thumbprint

write-host "Copy new certificate to Trusted Root Certification Authorities..." -foregroundcolor yellow
$SourceStore = New-Object  -TypeName System.Security.Cryptography.X509Certificates.X509Store  -ArgumentList My, LocalMachine
$SourceStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)

$cert = $SourceStore.Certificates | Where-Object { $_.Thumbprint -eq $Thumbprint}
$DestStore = New-Object  -TypeName System.Security.Cryptography.X509Certificates.X509Store  -ArgumentList root, LocalMachine
$DestStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
$DestStore.Add($cert)
 
$SourceStore.Close()
$DestStore.Close()

write-host "Change Exchange Server BackEnd certificate to new certificate..." -foregroundcolor yellow
Import-Module WebAdministration
$site = Get-ChildItem -Path "IIS:\Sites" | where {( $_.Name -eq "Exchange Back End" )}
$binding = $site.Bindings.Collection | where {$_.protocol -eq 'https' -and $_.bindingInformation -eq '*:444:'}
$binding.AddSslCertificate($Thumbprint, "my")

write-host "Done!" -foregroundcolor yellow