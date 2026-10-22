<#
.SYNOPSIS
    Shows the charge stations and the active transaction from the 50five GraphQL API.

.DESCRIPTION
    PowerShell equivalent of get_charge_stations and get_active_transaction in
    custom_components/evcnet/api.py. Asks for e-mail and password, runs only read-only queries.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File scripts\show_status.ps1
    powershell -ExecutionPolicy Bypass -File scripts\show_status.ps1 -Raw
#>
param(
    # Also print the full JSON responses
    [switch]$Raw
)

$ErrorActionPreference = "Stop"

$GraphqlUrl = "https://lms.servicelayer.platform-01.plugz.dev/graphql"
$Headers = @{ "plugz-application-id" = "com.plugz.app50five"; "Accept" = "application/json" }

$StationFields = @"
    id name commId reference state softwareVersion physicalNumber gpsAddress gpsCity lastCommunicationDate
    manufacturerType { vendor model }
    accessOptions { accessType authorizationMode chargeCard publishedOnMap }
    channels { id channelNo evseId globalStatus statuses lastStatusChangeDate power amp phases connectorType capabilities }
"@

$TransactionFields = @"
    transactionId transactionStatus energyDelivered durationCharging durationIdle startDate endDate updateDate
    channelId evseId loyaltyCardIdentifier loyaltyCardType totalAmount currency
"@

$LoginMutations = @(
    'mutation Login($email: String!, $password: String!) { loginFiftyFive(email: $email, password: $password) { access_token expires_in } }',
    # Deprecated, but kept as fallback like in api.py
    'mutation Login($email: String!, $password: String!) { login(email: $email, password: $password) { access_token expires_in } }'
)

function Invoke-Gql([string]$Query, [hashtable]$Variables = @{}) {
    $body = @{ query = $Query; variables = $Variables } | ConvertTo-Json -Depth 10 -Compress
    Invoke-RestMethod -Uri $GraphqlUrl -Method Post -Headers $Headers -ContentType "application/json" -Body $body
}

function Show-Errors($Response) {
    foreach ($e in @($Response.errors)) {
        if ($e) { Write-Warning ("{0}: {1}" -f (@($e.path) -join "."), $e.message) }
    }
}

function Join-Value($Value) {
    if ($null -eq $Value) { return "" }
    (@($Value) | ForEach-Object { "$_" }) -join ", "
}

function Format-Kwh($Wh) {
    if ($null -eq $Wh) { return "" }
    "{0:N3} kWh" -f ($Wh / 1000)
}

function Format-Kw($W) {
    if (-not $W) { return "" }
    "{0:N1} kW" -f ($W / 1000)
}

# --- Credentials ---
$email = Read-Host "E-mail"
$securePassword = Read-Host "Password" -AsSecureString
$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
try {
    $password = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
} finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
}

# --- Login ---
$token = $null
foreach ($mutation in $LoginMutations) {
    $response = Invoke-Gql $mutation @{ email = $email; password = $password }
    $result = $response.data.loginFiftyFive
    if (-not $result) { $result = $response.data.login }
    if ($result -and $result.access_token) {
        $token = $result.access_token
        break
    }
    Show-Errors $response
}
$password = $null
if (-not $token) {
    Write-Error "Login failed."
}
$Headers["Authorization"] = "Bearer $token"
Write-Host "Logged in." -ForegroundColor Green

# --- Charge stations ---
$response = Invoke-Gql "query Stations { getCustomerChargeStations { $StationFields } }"
Show-Errors $response
$stations = @($response.data.getCustomerChargeStations) | Where-Object { $_ }

Write-Host "`n=== Charge stations ($(@($stations).Count))" -ForegroundColor Cyan
foreach ($s in $stations) {
    [pscustomobject]@{
        Name              = $s.name
        Id                = $s.id
        CommId            = $s.commId
        Reference         = $s.reference
        State             = $s.state
        Vendor            = $s.manufacturerType.vendor
        Model             = $s.manufacturerType.model
        Software          = $s.softwareVersion
        PhysicalNumber    = $s.physicalNumber
        Address           = (@($s.gpsAddress, $s.gpsCity) | Where-Object { $_ }) -join ", "
        LastCommunication = $s.lastCommunicationDate
        AccessType        = Join-Value ($s.accessOptions | ForEach-Object { $_.accessType })
        AuthorizationMode = Join-Value ($s.accessOptions | ForEach-Object { $_.authorizationMode })
        ChargeCard        = Join-Value ($s.accessOptions | ForEach-Object { $_.chargeCard })
        PublishedOnMap    = Join-Value ($s.accessOptions | ForEach-Object { $_.publishedOnMap })
    } | Format-List

    Write-Host "Channels of $($s.name)" -ForegroundColor Yellow
    @($s.channels) | Where-Object { $_ } | ForEach-Object {
        [pscustomobject]@{
            No           = $_.channelNo
            Id           = $_.id
            EvseId       = $_.evseId
            Status       = $_.globalStatus
            Statuses     = Join-Value $_.statuses
            StatusSince  = $_.lastStatusChangeDate
            MaxPower     = Format-Kw $_.power
            Amp          = $_.amp
            Phases       = $_.phases
            Connector    = $_.connectorType
            Capabilities = Join-Value $_.capabilities
        }
    } | Format-List
}

# --- Active transaction ---
$response = Invoke-Gql "query Active { lmsActiveTransaction { $TransactionFields } }"
Show-Errors $response
$active = $response.data.lmsActiveTransaction

Write-Host "`n=== Active transaction" -ForegroundColor Cyan
if ($active) {
    [pscustomobject]@{
        TransactionId    = $active.transactionId
        Status           = $active.transactionStatus
        Energy           = Format-Kwh $active.energyDelivered
        DurationCharging = $active.durationCharging
        DurationIdle     = $active.durationIdle
        Start            = $active.startDate
        End              = $active.endDate
        LastUpdate       = $active.updateDate
        ChannelId        = $active.channelId
        EvseId           = $active.evseId
        Card             = $active.loyaltyCardIdentifier
        CardType         = $active.loyaltyCardType
        Amount           = (@($active.totalAmount, $active.currency) | Where-Object { $null -ne $_ }) -join " "
    } | Format-List
} else {
    Write-Host "No active transaction."
}

if ($Raw) {
    Write-Host "`n=== Raw: getCustomerChargeStations" -ForegroundColor Cyan
    $stations | ConvertTo-Json -Depth 10
    Write-Host "`n=== Raw: lmsActiveTransaction" -ForegroundColor Cyan
    $active | ConvertTo-Json -Depth 10
}
