# Fetch token with managed identity
function Get-AzToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $ResourceUri,
        [Switch]$AsHeader
    ) 
    $Context = [Microsoft.Azure.Commands.Common.Authentication.Abstractions.AzureRmProfileProvider]::Instance.Profile.DefaultContext
    $Token = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate($context.Account, $context.Environment, $context.Tenant.Id.ToString(), $null, [Microsoft.Azure.Commands.Common.Authentication.ShowDialog]::Never, $null, $ResourceUri).AccessToken
    if ($AsHeader) {
        return @{Headers = @{Authorization = "Bearer $Token" } }
    }
    return $Token    
}

$token = Get-AzToken -ResourceUri 'https://graph.microsoft.com/' -AsHeader

#### Step 2: Mapping all the parameters and calling Cloud PC endpoint
# Set headers
$headers = $token.Headers
# BaseURI for graph calls
$baseUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint"
# Format date 60 days ago
$date = (Get-Date).AddDays(-60) # Change according to days you want to look back
$dateFormat = 'yyyy-MM-dd'
$actualDate = Get-Date -Date $date -Format $dateFormat

$formattedDate = $actualDate + "T00:00:00.000Z"

$params = @"
{
    "top": 50,
    "skip": 0,
    "search": "",
    "filter": "((TotalUsageInHour eq 0)) and (NeverSignedIn eq true) and (CreatedDate le $formattedDate)",
    "select": ["CloudPcId", "IntuneDeviceId", "ManagedDeviceName", "UserPrincipalName", "TotalUsageInHour", "LastActiveTime", "PcType", "CreatedDate", "DeviceType"],
    "orderBy": ["TotalUsageInHour"]
}
"@

# Retrieve all devices and make a variable with all grace period devices
try {
    $allCloudPCs = Invoke-RestMethod -Uri "$baseUri/cloudPCs" -Headers $headers -Method GET -ErrorAction Stop
    $allGracePeriodDevices = $allCloudPCs.value | where-object {$_.status -eq "inGracePeriod"}
} catch {
    Write-Error "Failed to retrieve Cloud PCs: $_"
    exit
}

# Retrieve utilization Cloud PCs list
try {
    $cloudPCs = Invoke-RestMethod -Uri "$baseUri/reports/getTotalAggregatedRemoteConnectionReports" -Headers $headers -Method POST -ContentType "application/json" -Body $params -ErrorAction Stop
} catch {
    Write-Error "Failed to retrieve Cloud PC utilization reports: $_"
    exit
}

#### Step 3: Creating the generic functions
# Function to send email via Microsoft Graph API
function Send-GraphEmail {
    param (
        [Parameter(Mandatory = $true)]
        [string]$from,
        [Parameter(Mandatory = $true)]
        [string]$to,
        [Parameter(Mandatory = $true)]
        [string]$subject,
        [Parameter(Mandatory = $true)]
        [string]$body
    )

    $emailMessage = @{
        Message = @{
            Subject = $subject
            Body = @{
                ContentType = "Text"
                Content = $body
            }
            ToRecipients = @(@{EmailAddress = @{Address = $to}})
            From = @{
                EmailAddress = @{
                    Address = $from
                }
            }
        }
        SaveToSentItems = "true"
    }

    $emailMessageJson = $emailMessage | ConvertTo-Json -Depth 10
    $sendMailUri = "https://graph.microsoft.com/v1.0/me/sendMail"
    try {
        Invoke-RestMethod -Uri $sendMailUri -Headers $headers -Method POST -ContentType "application/json" -Body $emailMessageJson -ErrorAction Stop
    } catch {
        Write-Error "Failed to send email: $_"
    }
}

# Function to add a specific license to a user
function Add-UserLicense {
    param (
        [Parameter(Mandatory = $true)]
        [string]$userId,
        [Parameter(Mandatory = $true)]
        [string]$skuId
    )

    $addLicenseBody = @{
        "addLicenses" = @($skuId)
        "removeLicenses" = @()
    }

    $addLicenseJson = $addLicenseBody | ConvertTo-Json
    $addLicenseUri = "https://graph.microsoft.com/v1.0/users/$userId/assignLicense"
    try {
        Invoke-RestMethod -Uri $addLicenseUri -Headers $headers -Method POST -ContentType "application/json" -Body $addLicenseJson -ErrorAction Stop
    } catch {
        Write-Error "Failed to add license: $_"
    }
}

# Function to remove a specific license from a user
function Remove-UserLicense {
    param (
        [Parameter(Mandatory = $true)]
        [string]$userId,
        [Parameter(Mandatory = $true)]
        [string]$skuId
    )

    $removeLicenseBody = @{
        "addLicenses" = @()
        "removeLicenses" = @($skuId)
    }

    $removeLicenseJson = $removeLicenseBody | ConvertTo-Json
    $removeLicenseUri = "https://graph.microsoft.com/v1.0/users/$userId/assignLicense"
    try {
        Invoke-RestMethod -Uri $removeLicenseUri -Headers $headers -Method POST -ContentType "application/json" -Body $removeLicenseJson -ErrorAction Stop
    } catch {
        Write-Error "Failed to remove license: $_"
    }
}

# Function to fetch SKU ID dynamically based on Cloud PC type
function Get-SkuId {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$customObj
    )

    # Fetch SKU details
    try {
        $licenseDetails = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -Headers $headers -Method GET -ErrorAction Stop
    } catch {
        Write-Error "Failed to retrieve SKU details: $_"
        return $null
    }

    # Split SKU name from customObj object
    $splitObj = $customObj.PcType.Split(" ")
    $splitObj = $splitObj[3].Split("/")
    $splitObj = $splitObj.Split("v")

    $cpu = $splitObj[0]
    $cpu = $cpu + "C"
    $memory = $splitObj[2]
    $storage = $splitObj[3]

    # Searching for license and finding SkuID
    $searchString = "CPC_E_" + $cpu + "_" + $memory + "_" + $storage
    $skuId = ($licenseDetails.value | Where-Object { $_.skuPartNumber -eq $searchString }).skuId

    return $skuId
}

#### Step 4: Check grace period devices for utilization
foreach ($device in $allGracePeriodDevices) {
    # Define customObj
    $customObj = [PSCustomObject]@{
        CloudPcId = $device.CloudPcId
        IntuneDeviceId = $device.IntuneDeviceId
        ManagedDeviceName = $device.ManagedDeviceName
        UserPrincipalName = $device.UserPrincipalName
        TotalUsageInHour = $device.TotalUsageInHour
        LastActiveTime = $device.LastActiveTime
        PcType = $device.PcType
        CreatedDate = $device.CreatedDate
        DeviceType = $device.DeviceType
    }

    # Verify if user has started using device
    $isPresent = $cloudPCs.values | Where-Object {$_ -like "*$($device.id)*"}
    if ($isPresent) {
        $skuId = Get-SkuId -customObj $customObj
        if ($skuId) {
            $userId = $customObj.UserPrincipalName
            Add-UserLicense -userId $userId -skuId $skuId
        }
    } else {
        Write-Host "Cloud PC: $($device.id) is not in use.."
    }
}

#### Step 5: Notify user for underutilization and remove license
# Logic to identify and optimize license usage
foreach ($cpc in $cloudPCs.values) {
    $customObj = [PSCustomObject]@{
        CloudPcId = $cpc[0]
        IntuneDeviceId = $cpc[1]
        ManagedDeviceName = $cpc[2]
        UserPrincipalName = $cpc[3]
        TotalUsageInHour = $cpc[4]
        LastActiveTime = $cpc[5]
        PcType = $cpc[6]
        CreatedDate = $cpc[7]
        DeviceType = $cpc[8]
    }
    $isPresent = $allCloudPCs.value | Where-Object {$_.id -eq $customObj.CloudPcId}
    if ($isPresent -and $isPresent.Status -eq "provisioned") {
        # Sending email to the user
        $from = "no-reply@domain.com"
        $to = $customObj.UserPrincipalName
        $subject = "Are you still using your Cloud PC?"
        $body = "Dear user, your Cloud PC with ID $($customObj.CloudPcId) is underutilized and will therefore be reallocated. Should you still need your Cloud PC, please sign into it, and you will automatically keep it."
        Send-GraphEmail -from $from -to $to -subject $subject -body $body

        # Remove license from user (Direct assigned license)    
        $userId = $customObj.UserPrincipalName
        $skuId = Get-SkuId -customObj $customObj
        if ($skuId) {
            Remove-UserLicense -userId $userId -skuId $skuId
            Write-Host "License has been successfully removed from user"
        }
    }
}