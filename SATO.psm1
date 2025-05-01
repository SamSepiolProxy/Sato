
$manifest = Import-PowerShellDataFile "$PSScriptRoot\SATO.psd1"
$version = $manifest.ModuleVersion
$host.ui.RawUI.WindowTitle = "SATO v$version"


$banner = @"
  
   ____/\\\\\\\\\\\____        _____/\\\\\\\\\____        __/\\\\\\\\\\\\\\\___        _______/\\\\\______        
    __/\\\/////////\\\_        ___/\\\\\\\\\\\\\\\__        _\///////\\\/////__        _____/\\\///\\\____        
     _\//\\\______\///__        __/\\\/////////\\\_        _______\/\\\_______        ___/\\\/__\///\\\__         
      __\////\\\_________        _\/\\\_______\/\\\_        _______\/\\\_______        __/\\\______\//\\\_       
       _____\////\\\______        _\/\\\\\\\\\\\\\\\_        _______\/\\\_______        _\/\\\_______\/\\\_      
        _______\////\\\___        _\/\\\/////////\\\_        _______\/\\\_______        _\//\\\______/\\\__     
         __/\\\______\//\\\__      _\/\\\_______\/\\\_        _______\/\\\_______        __\///\\\__/\\\____    
          _\///\\\\\\\\\\\/___      _\/\\\_______\/\\\_        _______\/\\\_______        ____\///\\\\\/_____    
           ___\///////////____       _\///________\///__        _______\///________        _______\/////_______
                                                                                                             
                 _________  Secure Azure Token Operations  ___________                                       
                            Version $version                                                  
                          by Edrian Miranda aka ObiKuro                                     

"@


Write-Host $banner -ForegroundColor DarkGreen

Write-Host "--------------------------------------------------------------------------------------------------------------------------" -ForegroundColor DarkGreen








$modulesPath = Join-Path -Path $PSScriptRoot -ChildPath 'modules'
$scripts = Get-ChildItem -Path "$modulesPath\*.ps1" -ErrorAction SilentlyContinue

foreach ($script in $scripts) {
    try {
        . $script.FullName
    } catch {
        Write-Error "Failed to import $($script.FullName): $_"
    }
}


$PredefinedScopes = @{
    MsGraph = "https://graph.microsoft.com/.default offline_access openid"
    MSTeams = "https://api.spaces.skype.com/.default offline_access openid"
    Office = "https://manage.office.com/.default offline_access openid"
    Outlook = "https://outlook.office365.com/.default offline_access openid"
    WinGraph = "https://graph.windows.net/.default offline_access openid"
    CoreARM = "https://management.core.windows.net/.default offline_access openid"
    MaARM = "https://management.azure.com/.default offline_access openid"
    IntuneMam = "https://intunemam.microsoftonline.com/.default offline_access openid"
    SharePoint = "https://$SharePointTenantName$AdminSuffix.sharepoint.com/Sites.FullControl.All offline_access openid"
    OneDrive = "https://officeapps.live.com/.default offline_access openid"
    KeyVault = "https://vault.azure.net/.default offline_access openid"
}


$PredefinedGrantTypes = @(
    "client_credentials",
    "password",
    "refresh_token",
    "device_code",
    "jwt_assertion",
    "jwt_assertion_sign"
)


function Invoke-Sato {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("client_credentials", "password", "refresh_token", "device_code", "jwt_assertion", "jwt_assertion_sign")]
        [string]$GrantType,

        [Parameter(Mandatory = $true)]
        [string]$TenantID,

        [Parameter()]
        [string]$ClientID = "d3590ed6-52b3-4102-aeff-aad2292ab01c", # Default to Microsoft Office client ID

        [Parameter()]
        [string]$ClientSecret,

        [Parameter()]
        [string]$Username,

        [Parameter()]
        [string]$Password,

        [Parameter()]
        [string]$Scope = "https://graph.windows.net/.default offline_access openid",

        [Parameter()]
        [string]$RefreshToken,

        [Parameter()]
        [switch]$Decode,

        [Parameter()]
        [switch]$UseCAE,

        [Parameter()]
        [string]$AppID,

        [Parameter()]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,

        [Parameter()]
        [string]$CertificatePath,

        [Parameter()]
        [string]$KeyVaultName,

        [Parameter()]
        [string]$CertName,

        [Parameter()]
        [string]$KeyToken,

        [Parameter(Mandatory = $false)]
        [ValidateSet("MsGraph", "MSTeams", "Office", "Outlook", "WinGraph", "CoreARM", "MaARM", "IntuneMam", "SharePoint", "OneDrive", "KeyVault")]
        [string]$PredefinedScope,

        [Parameter()]
        [string]$UserAgent = 'azsdk-net-Identity/1.11.4 (.NET Framework 4.8.9290.0; Microsoft Windows 10.0.19045 )'
    )

    # Set up the HTTP headers for the authentication requests.
    $headers = @{
        "Host"                          = "login.microsoftonline.com"
        "X-Client-Sku"                  = "MSAL.Desktop"
        "X-Client-Ver"                  = "4.61.3.0"
        "X-Client-Os"                   = "Windows 10 Pro"
        "Client-Request-Id"             = "b86490af-0d7c-4510-ac4f-eb0b6a9c2ff0"
        "Return-Client-Request-Id"      = "true"
        "X-App-Name"                    = "UnknownClient"
        "X-App-Ver"                     = "0.0.0.0"
        "Content-Type"                  = "application/x-www-form-urlencoded"
        "X-Ms-Client-Request-Id"        = "6121e2fb-8d02-4559-954a-7f7b24ddb757"
        "X-Ms-Return-Client-Request-Id" = "true"
        "User-Agent"                    = $UserAgent
    }

    if ($PredefinedScope) {
        $Scope = $PredefinedScopes[$PredefinedScope]
    }

    
    if ($GrantType -eq "jwt_assertion" -and !$Certificate) {
        if ($CertificatePath) {
            try {
                Write-Host "Loading certificate from file: $CertificatePath" -ForegroundColor Cyan
                $securePwd = Read-Host "Enter the certificate password" -AsSecureString
                $Certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
                $Certificate.Import($CertificatePath, $securePwd, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
            } catch {
                Write-Error "Failed to load certificate from path: $CertificatePath. Error: $_"
                return
            }
        } else {
            Write-Error "Either a Certificate variable or CertificatePath must be provided for JWT assertion."
            return
        }
    }

    switch ($GrantType) {
        "password" {
            $response = Get-PasswordToken -TenantID $TenantID -ClientID $ClientID -Username $Username -Password $Password -Scope $Scope -Headers $Headers
        }

        "client_credentials" {
            $response = Get-ClientCredentialsToken -TenantID $TenantID -ClientID $ClientID -ClientSecret $ClientSecret -Scope $Scope -Headers $Headers
        }

        "refresh_token" {
            $response = Get-RefreshToken -TenantID $TenantID -ClientID $ClientID -RefreshToken $RefreshToken -Scope $Scope -Headers $Headers
        }

        "device_code" {
            $response = Get-DeviceCodeToken -TenantID $TenantID -ClientID $ClientID -Scope $Scope -UseCAE:$UseCAE -Headers $Headers
        }

        "jwt_assertion" {
            if ($Certificate) {
                Write-Host "Using local certificate for JWT assertion" -ForegroundColor Cyan
                $response = Get-CertificateToken -ClientCertificate $Certificate -TenantID $TenantID -AppID $AppID -Scope $Scope -Headers $Headers
            } else {
                Write-Error "A certificate must be provided for JWT assertion."
                return
            }
        }

        "jwt_assertion_sign" {
            if ($KeyVaultName -and $CertName -and $KeyToken) {
                Write-Host "Using Azure Key Vault for JWT assertion signing" -ForegroundColor DarkGreen
                $response = Get-KeyVaultSignedJwt -TenantID $TenantID -AppID $AppID -KeyVaultName $KeyVaultName -CertName $CertName -KeyToken $KeyToken -Scope $Scope
                
            } else {
                Write-Error "Key Vault details must be provided for JWT assertion signing."
                return
            }
        }

        default {
            Write-Error "Unsupported grant type: $GrantType"
            return
        }
    }

    if ($response) {
        Write-Host "Access Token:" -ForegroundColor DarkGreen
        Write-Output $response.access_token

        if ($response.refresh_token) {
            Write-Host "Refresh Token:" -ForegroundColor DarkGreen
            Write-Output $response.refresh_token
        }

        if ($Decode) {
            Decode-Jwt -Token $response.access_token
        }
    }
}


Export-ModuleMember -Function Invoke-Sato