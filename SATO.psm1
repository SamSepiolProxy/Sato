
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

# map aliases → GUIDs
$predefinedClientIDs = @{
    o365mgmt      = '00b41c95-dab0-4487-9791-b9d2c32c80f2'
    azcli         = '04b07795-8ddb-461a-bbee-02f9e1bf7b46'
    azps          = '1950a258-227b-4e31-a9cf-717495945fc2'
    teams         = '1fec8e78-bce4-4aaf-ab1b-5451cc387264'
    msteams       = '1fec8e78-bce4-4aaf-ab1b-5451cc387264'
    msoffice      = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
    aadps         = '1b730954-1685-4b74-9bfd-dac224a7b894'
    msedge        = 'ecd6b820-32c2-49b6-98a6-444530e5a77a'
    edge          = 'ecd6b820-32c2-49b6-98a6-444530e5a77a'
    msbroker      = '29d9ed98-a469-4536-ade2-f981bc1d605e'
    broker        = '29d9ed98-a469-4536-ade2-f981bc1d605e'
    companyportal = '9ba1a5c7-f17a-4de9-a1f1-6178c8d51223'
}

$PredefinedGrantTypes = @(
    "client_credentials",
    "password",
    "refresh_token",
    "device_code",
    "jwt_assertion",
    "jwt_assertion_sign", 
    "estsauthcookie"
)

$uaMap = @{
    edge            = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0"
    edge_windows    = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0"
    edge_android    = "Mozilla/5.0 (Linux; Android 10; SM-G973F) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.6045.66 Mobile Safari/537.36 EdgA/118.0.2088.66"
    chrome          = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
    chrome_windows  = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
    chrome_android  = "Mozilla/5.0 (Linux; Android 10; K) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Mobile Safari/537.3"
    chrome_macos    = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
    chrome_linux    = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.3"
    chrome_ios      = "Mozilla/5.0 (iPhone; CPU iPhone OS 15_8 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) CriOS/119.0.6045.109 Mobile/15E148 Safari/604."
    firefox         = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/119.0"
    firefox_windows = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/119.0"
    firefox_macos   = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:109.0) Gecko/20100101 Firefox/119.0"
    firefox_ubuntu  = "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/119.0"
    safari          = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.1 Safari/605.1.1"
    safari_macos    = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.1 Safari/605.1.1"
    safari_ios      = "Mozilla/5.0 (iPhone; CPU iPhone OS 16_6_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.6 Mobile/15E148 Safari/604."
}


function Invoke-Sato {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("client_credentials", "password", "refresh_token", "device_code", "jwt_assertion", "jwt_assertion_sign", "estsauthcookie")]
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
        [string]$Cookie,

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
        [string]$SaveToVar = "access_token",

        [Parameter(Mandatory = $false)]
        [ValidateSet(
            'o365mgmt',
            'azcli',
            'azps',
            'teams',
            'msteams',
            'msoffice',
            'aadps',
            'msedge',
            'edge',
            'msbroker',
            'broker',
            'companyportal'
        )]
        [string]$PredefinedClientID,

        [Parameter()]
        [string]$UserAgent = 'azsdk-net-Identity/1.11.4 (.NET Framework 4.8.9290.0; Microsoft Windows 10.0.19045 )',

        [Parameter(Mandatory = $false)]
        [ValidateSet(
            'edge','edge_windows','edge_android',
            'chrome','chrome_windows','chrome_android','chrome_macos','chrome_linux','chrome_ios',
            'firefox','firefox_windows','firefox_macos','firefox_ubuntu',
            'safari','safari_macos','safari_ios'
        )]
        [string]$PredefinedUserAgent
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

        # if they picked a predefined alias, swap in its GUID
    if ($PredefinedClientID) {
        $ClientID = $predefinedClientIDs[$PredefinedClientID]
    }

    if ($PredefinedUserAgent) {
        $UserAgent = $uaMap[$PredefinedUserAgent]
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

        "estsauthcookie" {
            $response = Get-EstsAuthCookieToken -ClientID $ClientID -Scope $Scope -ESTSAuthCookie:$Cookie
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
        Set-Variable -Name $SaveToVar -Value $response -Scope Global
        Write-Host "Tokens saved to variable: `$${SaveToVar}" -ForegroundColor Green

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