#Requires -Modules Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Reports
# Microsoft 365 Attendance Report v3

$TenantId = $env:TENANT_ID
$ClientId = $env:CLIENT_ID
$ClientSecret = $env:CLIENT_SECRET

$OfficeIPPrefixes=@("114.9.96.","101.255.148.")
#$ExportFolder="D:\M365_Attendance"
#$ExportFile=Join-Path $ExportFolder "Attendance_Logs.csv"

$Modules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Reports"
)	

foreach ($Module in $Modules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        Install-Module $Module -Scope CurrentUser -Force -AllowClobber
    }
}

if ($env:RUNNER_TEMP) {
    $ExportFolder = Join-Path $env:RUNNER_TEMP "Attendance"
}
else {
    $ExportFolder = "D:\M365_Attendance"
}

if (!(Test-Path $ExportFolder)) {
    New-Item -ItemType Directory -Path $ExportFolder -Force | Out-Null
}



# $TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$TimeStamp = [System.TimeZoneInfo]::ConvertTime(
    (Get-Date),
    $JakartaTZ
).ToString("yyyyMMdd_HHmmss")

$FileName = "Attendance_Logs_$TimeStamp.csv"
$ExportFile = Join-Path $ExportFolder $FileName
$ErrorLog=Join-Path $ExportFolder "Attendance_Errors.log"

# ==========================================================
# REPORT DATE (Yesterday - Jakarta WIB)
# ==========================================================

$JakartaTZ = [System.TimeZoneInfo]::FindSystemTimeZoneById("SE Asia Standard Time")

$Yesterday = (Get-Date).Date.AddDays(-1)

$StartLocal = [datetime]::SpecifyKind(
    $Yesterday,
    [System.DateTimeKind]::Unspecified
)

$EndLocal = [datetime]::SpecifyKind(
    $Yesterday.AddDays(1).AddSeconds(-1),
    [System.DateTimeKind]::Unspecified
)

$StartDate = [System.TimeZoneInfo]::ConvertTimeToUtc(
    $StartLocal,
    $JakartaTZ
).ToString("yyyy-MM-ddTHH:mm:ssZ")

$EndDate = [System.TimeZoneInfo]::ConvertTimeToUtc(
    $EndLocal,
    $JakartaTZ
).ToString("yyyy-MM-ddTHH:mm:ssZ")

$TargetDate = $Yesterday.ToString("yyyy-MM-dd")

#if(Test-Path $ExportFile){
# $Archive="Attendance_Logs_"+(Get-Date).ToString("yyyyMMdd_HHmmss")+".csv"
# Rename-Item $ExportFile $Archive -Force
#}

$Sec=ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$Cred=New-Object System.Management.Automation.PSCredential($ClientId,$Sec)
Connect-MgGraph -TenantId $TenantId -Credential $Cred -NoWelcome

$Users=Get-MgUser -All -Property DisplayName,UserPrincipalName|?{$_.UserPrincipalName -notlike "*.onmicrosoft.com*"}|sort DisplayName
$Report=@()
""|Out-File $ErrorLog

$i=0
foreach($User in $Users){
$i++
Write-Progress -Activity "Attendance" -Status "$i / $($Users.Count)" -PercentComplete (($i/$Users.Count)*100)
try{
$Logs=Get-MgAuditLogSignIn -Filter "userPrincipalName eq '$($User.UserPrincipalName)' and createdDateTime ge $StartDate and createdDateTime le $EndDate and isInteractive eq true" -All -ErrorAction Stop|
?{
$OK=$false
foreach($P in $OfficeIPPrefixes){if($_.IPAddress -like "$P*"){$OK=$true;break}}
$_.Status.ErrorCode -eq 0 -and $OK -and $_.DeviceDetail
}|sort CreatedDateTime

if($Logs.Count){
$f=$Logs|select -First 1
$l=$Logs|select -Last 1
$Report+=[pscustomobject]@{
Date=$TargetDate
DisplayName=$User.DisplayName
UserEmail=$User.UserPrincipalName
Morning_First_Access=([System.TimeZoneInfo]::ConvertTime([datetimeoffset]$f.CreatedDateTime,$JakartaTZ)).ToString("yyyy-MM-dd HH:mm:ss")
Night_Last_Access=([System.TimeZoneInfo]::ConvertTime([datetimeoffset]$l.CreatedDateTime,$JakartaTZ)).ToString("yyyy-MM-dd HH:mm:ss")
LoginCount=$Logs.Count
DeviceOS=$f.DeviceDetail.OperatingSystem
Browser=$f.DeviceDetail.Browser
IPAddress=$f.IPAddress
Status="Present"
}
}else{
$Report+=[pscustomobject]@{Date=$TargetDate;DisplayName=$User.DisplayName;UserEmail=$User.UserPrincipalName;Morning_First_Access="";Night_Last_Access="";LoginCount=0;DeviceOS="";Browser="";IPAddress="";Status="Absent"}
}
}catch{
Add-Content $ErrorLog "$($User.UserPrincipalName): $($_.Exception.Message)"
$Report+=[pscustomobject]@{Date=$TargetDate;DisplayName=$User.DisplayName;UserEmail=$User.UserPrincipalName;Morning_First_Access="";Night_Last_Access="";LoginCount=0;DeviceOS="";Browser="";IPAddress="";Status="Error"}
}
}
# ==========================================================
# Export Report to Local Drive
# ==========================================================

$Report |
    Sort-Object DisplayName |
    Export-Csv $ExportFile -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Local report saved:"
Write-Host $ExportFile

# ==========================================================
# Upload Report to SharePoint
# ==========================================================

$DriveId = $env:DRIVE_ID

try{

    Write-Host ""
    Write-Host "Uploading Attendance_Logs.csv to SharePoint..."

    Invoke-MgGraphRequest `
        -Method PUT `
	-Uri ("https://graph.microsoft.com/v1.0/drives/$DriveId/root:/Attendance/{0}:/content" -f $FileName) `
        -Body ([System.IO.File]::ReadAllBytes($ExportFile)) `
        -ContentType "text/csv"


    Write-Host "SharePoint upload successful."

}
catch{

    Write-Warning "SharePoint upload failed."

    Write-Warning $_.Exception.Message

}

Disconnect-MgGraph | Out-Null

Write-Host ""
Write-Host "Attendance process completed."
