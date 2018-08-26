Param(
    [Parameter(Mandatory=$TRUE)][ValidateNotNullOrEmpty()][string]$METHOD,
    [Parameter(Mandatory=$FALSE)][string]$SN_INPUT_FILE,
    [Parameter(Mandatory=$FALSE)][string]$SN_QUERY,
    [Parameter(Mandatory=$FALSE)][string]$SN_FILTERS,
    [Parameter(Mandatory=$FALSE)][string]$SN_OUTPUT_FILE,
    [Parameter(Mandatory=$FALSE)][string]$SN_JSON_RESULT,
    [Parameter(Mandatory=$FALSE)][string]$SHORT_DESCRIPTION,
    [Parameter(Mandatory=$FALSE)][string]$DESCRIPTION,
    [Parameter(Mandatory=$FALSE)][string]$CALLER_ID,
    [Parameter(Mandatory=$FALSE)][string]$ASSIGNMENT_GROUP,
    [Parameter(Mandatory=$FALSE)][string]$U_CI_CLASS_NAME,
    [Parameter(Mandatory=$FALSE)][string]$U_SERIAL_NUMBER,
    [Parameter(Mandatory=$FALSE)][string]$WORK_NOTES,
    [Parameter(Mandatory=$FALSE)][string]$COMMENTS,
    [Parameter(Mandatory=$FALSE)][string]$ASSIGNED_TO
)

$METHOD = $METHOD.ToUpper()

If (!$SN_INPUT_FILE) {
        if ($METHOD -eq "PATCH" -or $METHOD -eq "POST") {
            Write-Output ""
            Write-Output "Invalid input, try this -> -SN_INPUT_FILE 'c:/temp/JsonContent.json'"
            Write-Output ""
            Exit
        } else {
            if ($($SN_QUERY.Length -lt 3) -or $($SN_FILTERS.Length -lt 3) -and $METHOD -ne "XLSX") {
                Write-Output ""
                Write-Output "For '$METHOD' method please provide '-SN_QUERY and -SN_FILTERS' parameters, try this -> -SN_QUERY 'number=INC123456' -SN_FILTERS 'number,description'"
                Write-Output ""
                Exit
            } else {
                $SYSPARM_FIELDS = "&sysparm_fields=" + $SN_FILTERS
            }
        }
    } else {
        if (Test-Path $SN_INPUT_FILE) {
            if ($($SN_QUERY.Length -lt 4) -or $($SN_FILTERS.Length -lt 4) -and $METHOD -ne "POST") {
                if ($METHOD -eq "PATCH" -and $SN_QUERY.Length -gt 4) { $BODY_R = (Get-Content -Raw $SN_INPUT_FILE | ConvertFrom-Json); $SYSPARM_FIELDS = "&sysparm_fields=" + 'number'
                } else {
                    Write-Output ""
                    Write-Output "For '$METHOD' method please provide '-SN_QUERY and -SN_FILTERS' parameters, try this -> -SN_QUERY 'number=INC123456' -SN_FILTERS 'number,description'"
                    Write-Output ""
                    Exit                
                }
            } else {
                $BODY_R = (Get-Content -Raw $SN_INPUT_FILE | ConvertFrom-Json)
                $SYSPARM_FIELDS = "&sysparm_fields=" + $SN_FILTERS
            }
        } else {
            Write-Output ""
            Write-Output "This file is empty, try this -> -SN_FILE 'c:/temp/JsonContent.json'"
            Write-Output ""
            Exit
        }
}

$SN_USER = "yourSnUserName"
$SN_PASS = "yourSnPassword"
$CURRENT_USER = $env:UserName
$BASE64AUTHINFO = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $SN_USER, $SN_PASS)))
$HEADERS = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$HEADERS.Add('Authorization',('Basic {0}' -f $BASE64AUTHINFO))
If ($HEADERS -notcontains "application/json") { $HEADERS.Add('Accept','application/json') }
$SN_INSTANCE = "https://yourSnInstance.service-now.com"
$SN_INSTANCE_API = "$SN_INSTANCE/api/now/table/"
$PROXY = "http://yourProxy.com:8080"
$APP_JSON = "application/json"
$SN_TABLE_API = "incident"
if (!$SN_JSON_RESULT) { $SN_JSON_RESULT = "JsonFileResult.json" }
if (!$SN_OUTPUT_FILE) { $SN_OUTPUT_FILE = "incident" }
$TIMEOUT = "120"

switch($METHOD){
    "XLSX" {
        $SN_OUTPUT_FILE = "$SN_OUTPUT_FILE.xlsx"
        if (Test-Path $SN_OUTPUT_FILE) { Remove-Item -Path $SN_OUTPUT_FILE -Force }
        $SN_QUERY_API = "sysparm_query=" + $SN_QUERY + $SYSPARM_FIELDS
        $SN_URI = $SN_INSTANCE + '/incident_list.do' + '?' + $METHOD +'&' + $SN_QUERY_API + $SYSPARM_FIELDS
        Invoke-WebRequest -Headers $HEADERS -Uri $SN_URI -Proxy $PROXY -OutFile $SN_OUTPUT_FILE -TimeoutSec "$TIMEOUT"
    }
    "CSV" {
        $SN_OUTPUT_FILE = "$SN_OUTPUT_FILE.csv"
        if (Test-Path $SN_OUTPUT_FILE) { Remove-Item -Path $SN_OUTPUT_FILE -Force }
        $SN_QUERY_API = "sysparm_query=" + $SN_QUERY + $SYSPARM_FIELDS
        $SN_URI = $SN_INSTANCE + '/incident_list.do' + '?' + $METHOD +'&' + $SN_QUERY_API + $SYSPARM_FIELDS
        Invoke-WebRequest -Headers $HEADERS -Uri $SN_URI -Proxy $PROXY -OutFile $SN_OUTPUT_FILE -TimeoutSec "$TIMEOUT"
    }
    "PDF" {
        $SN_OUTPUT_FILE = "$SN_OUTPUT_FILE.pdf"
        if (Test-Path $SN_OUTPUT_FILE) { Remove-Item -Path $SN_OUTPUT_FILE -Force }
        $SN_QUERY_API = "sysparm_query=" + $SN_QUERY + $SYSPARM_FIELDS
        $SN_URI = $SN_INSTANCE + '/incident_list.do' + '?' + $METHOD +'&' + $SN_QUERY_API + $SYSPARM_FIELDS
        Invoke-WebRequest -Headers $HEADERS -Uri $SN_URI -Proxy $PROXY -OutFile $SN_OUTPUT_FILE -TimeoutSec "$TIMEOUT"
    }
    "UNL" {
        $SN_OUTPUT_FILE = "$SN_OUTPUT_FILE.unl"
        if (Test-Path $SN_OUTPUT_FILE) { Remove-Item -Path $SN_OUTPUT_FILE -Force }
        $SN_QUERY_API = "sysparm_query=" + $SN_QUERY + $SYSPARM_FIELDS
        $SN_URI = $SN_INSTANCE + '/incident_list.do' + '?' + $METHOD +'&' + $SN_QUERY_API + $SYSPARM_FIELDS
        Invoke-WebRequest -Headers $HEADERS -Uri $SN_URI -Proxy $PROXY -OutFile $SN_OUTPUT_FILE -TimeoutSec "$TIMEOUT"
    }
    "GET" {
        $SN_QUERY_API = "sysparm_query=" + $SN_QUERY + $SYSPARM_FIELDS
        $SN_URI = $SN_INSTANCE_API + $SN_TABLE_API + '?' + $SN_QUERY_API + $SYSPARM_FIELDS
        $R_API = Invoke-WebRequest -Headers $HEADERS -Method $METHOD -Uri $SN_URI -Proxy $PROXY -ContentType "$APP_JSON" -TimeoutSec "$TIMEOUT"
        $R_JSON = $R_API.Content | ConvertFrom-Json | ConvertTo-Json
        $R_JSON
        Write-Output "$R_JSON" > $SN_JSON_RESULT
        Write-Output "STATUS: $($R_API.StatusDescription),STATUS CODE: $($R_API.StatusCode),LENGTH: $($R_API.RawContentLength)"
    }
    "POST" {
        $SN_URI = $SN_INSTANCE_API + $SN_TABLE_API
        $BODY_JSON = $BODY_R | ConvertTo-Json
        $R_API = Invoke-WebRequest -Headers $HEADERS -Method $METHOD -Uri $SN_URI -Body $BODY_JSON -Proxy $PROXY -ContentType "$APP_JSON" -TimeoutSec "$TIMEOUT"
        $R_JSON = $R_API.Content | ConvertFrom-Json | ConvertTo-Json
        $R_JSON
        Write-Output "$R_JSON" > $SN_JSON_RESULT
        Write-Output "STATUS: $($R_API.StatusDescription),STATUS CODE: $($R_API.StatusCode),LENGTH: $($R_API.RawContentLength)"
    }
    "PATCH" {
        $SYSPARM_FIELDS = "&sysparm_fields=sys_id"
        $SN_QUERY_API = "sysparm_query=" + $SN_QUERY + $SYSPARM_FIELDS
        $SN_URI = $SN_INSTANCE_API + $SN_TABLE_API + '?' + $SN_QUERY_API
        $R_API = Invoke-WebRequest -Headers $HEADERS -Method 'GET' -Uri $SN_URI -Body $BODY_JSON -Proxy $PROXY -ContentType "$APP_JSON" -TimeoutSec "$TIMEOUT"
        $R_ID = ($R_API.Content | ConvertFrom-Json).Result
        $SN_SYSID = $($R_ID.sys_id)
        $SN_URI = $SN_INSTANCE_API + $SN_TABLE_API + "/$SN_SYSID"
        $BODY_JSON = $BODY_R | ConvertTo-Json
        $BODY_JSON
        $R_API = Invoke-WebRequest -Headers $HEADERS -Method $METHOD -Uri $SN_URI -Body $BODY_JSON -Proxy $PROXY -ContentType "$APP_JSON" -TimeoutSec "$TIMEOUT"
        $R_JSON = $R_API.Content | ConvertFrom-Json | ConvertTo-Json
        $R_JSON
        Write-Output "$R_JSON" > $SN_JSON_RESULT
        Write-Output "STATUS: $($R_API.StatusDescription),STATUS CODE: $($R_API.StatusCode),LENGTH: $($R_API.RawContentLength)"
    }
    default {
        Write-Output ""
        Write-Output " --> Invalid input, for option METHOD choose 'GET' or 'POST' or 'PATCH', try this -> -METHOD 'GET'"
        Write-Output ""
        Exit
    }
}