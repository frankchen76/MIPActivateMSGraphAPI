# Create self-signed certificate - Demo.ps1 - Get all guest users and they last logon time

Param(
    # The path where the exported certificates (PFX and CER) will be saved; defaults to your user's Downloads folder
    [Parameter()]
    [String]$SaveCertificateTo = '~\Downloads',

    # The name of the certificate; how you want it to be identified in a friendly manner; defaults to "Graph certificate"
    [Parameter()]
    [String]$CertificateSubject = 'Locked down certificate',

    # The certificate store where your certificate will be created; defaults to your user account's "Personal" folder
    [Parameter()]
    [ValidateScript({
            if ($_ -ilike 'Cert:\*' -and (Test-Path -Path $_)) { $true }
            else { throw 'Invalid certificate store path' }
        })]
    [String]$CertificateStore = 'Cert:\CurrentUser\My',

    # When the certificate should expire; defaults to one year after the script runs
    [Parameter()]
    [DateTime]$ExpiresOn = (Get-Date).AddYears(10),

    # The certificate's password for securing the private key
    [Parameter(Mandatory)]
    [SecureString]$CertificatePassword
)

# Create a certificate
$graphCertificate = New-SelfSignedCertificate -Subject $CertificateSubject -CertStoreLocation $CertificateStore -NotAfter $ExpiresOn

# Set up the base path and file name, without extension
$certificatePath = Join-Path -Path $SaveCertificateTo -ChildPath $CertificateSubject

# Export a certificate without private key to disk
$graphCertificate | Export-Certificate -Type CERT -FilePath "$certificatePath.cer" | Out-Null

# Export a certificate with private key and password to disk
$graphCertificate | Export-PfxCertificate -Password $CertificatePassword -FilePath "$certificatePath.pfx" | Out-Null