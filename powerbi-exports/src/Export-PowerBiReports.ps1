param(
    $Destination = (Resolve-Path .),
    [Parameter(Mandatory=$true)]
    $BearerToken,
    [Switch]$Grupper,
    [Switch]$Distrikter,
    $GrupperFile = "Grupper.csv",
    $GrupperFileIdField = "GruppeID",
    $GrupperFileNameField = "Gruppe",
    $DistrikterFile = "Distrikter.csv",
    $DistrikterFileNameField = "Distrikt",
    $ChunkSize = 10,
    $Skip = 0
)

$ErrorActionPreference = "STOP"

$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36 Edg/113.0.1774.35"


function Export-ReportForDistrikt {
    param (
        $DistriktName
    )

    Write-Host "Eksporterer for ${DistriktName}"

    $activityId = [System.Guid]::NewGuid().ToString("d")

    $result = Invoke-WebRequest -UseBasicParsing -Uri "https://wabi-europe-north-b-redirect.analysis.windows.net/export/reports/d4644347-1a7e-493a-ba7c-01ca397f8b70/asyncexports" `
        -Method "POST" `
        -WebSession $session `
        -Headers @{
        "authority"                  = "wabi-europe-north-b-redirect.analysis.windows.net"
        "accept"                     = "application/json, text/plain, */*"
        "accept-encoding"            = "gzip, deflate, br"
        "accept-language"            = "en-US,en;q=0.9,da;q=0.8"
        "authorization"              = $BearerToken
        "activityid"                 = $activityId
        "powerbi-reportsectioncount" = "2"
        "referer"                    = "https://app.powerbi.com/"
        "x-powerbi-hostenv"          = "Power BI Web"
    } `
        -ContentType "application/json;charset=UTF-8" `
        -Body "{`"format`":`"pdf`",`"powerBIReportConfiguration`":{`"settings`":{`"locale`":`"en-US`",`"timeZoneId`":`"Romance Standard Time`",`"excludeHiddenPages`":false},`"payload`":`"{\`"objectId\`":\`"e73ffc96-08f2-788b-184b-93cee73bbfee\`",\`"type\`":99,\`"explorationState\`":\`"{\\\`"version\\\`":\\\`"1.3\\\`",\\\`"filters\\\`":{\\\`"byExpr\\\`":[{\\\`"name\\\`":\\\`"Filtereeb4b4ea64b689d75a87\\\`",\\\`"type\\\`":\\\`"RelativeDate\\\`",\\\`"filter\\\`":{\\\`"Version\\\`":2,\\\`"From\\\`":[{\\\`"Name\\\`":\\\`"d\\\`",\\\`"Entity\\\`":\\\`"dimtid\\\`",\\\`"Type\\\`":0}],\\\`"Where\\\`":[{\\\`"Condition\\\`":{\\\`"Between\\\`":{\\\`"Expression\\\`":{\\\`"Column\\\`":{\\\`"Expression\\\`":{\\\`"SourceRef\\\`":{\\\`"Source\\\`":\\\`"d\\\`"}},\\\`"Property\\\`":\\\`"Dato\\\`"}},\\\`"LowerBound\\\`":{\\\`"DateSpan\\\`":{\\\`"Expression\\\`":{\\\`"DateAdd\\\`":{\\\`"Expression\\\`":{\\\`"Now\\\`":{}},\\\`"Amount\\\`":-4,\\\`"TimeUnit\\\`":3}},\\\`"TimeUnit\\\`":3}},\\\`"UpperBound\\\`":{\\\`"DateSpan\\\`":{\\\`"Expression\\\`":{\\\`"DateAdd\\\`":{\\\`"Expression\\\`":{\\\`"Now\\\`":{}},\\\`"Amount\\\`":-1,\\\`"TimeUnit\\\`":3}},\\\`"TimeUnit\\\`":3}}}}}]},\\\`"expression\\\`":{\\\`"Column\\\`":{\\\`"Expression\\\`":{\\\`"SourceRef\\\`":{\\\`"Entity\\\`":\\\`"dimtid\\\`"}},\\\`"Property\\\`":\\\`"Dato\\\`"}},\\\`"howCreated\\\`":1},{\\\`"name\\\`":\\\`"Filter5d5f717fe247ec01302b\\\`",\\\`"type\\\`":\\\`"Categorical\\\`",\\\`"filter\\\`":{\\\`"Version\\\`":2,\\\`"From\\\`":[{\\\`"Name\\\`":\\\`"g\\\`",\\\`"Entity\\\`":\\\`"Grupper\\\`",\\\`"Type\\\`":0}],\\\`"Where\\\`":[{\\\`"Condition\\\`":{\\\`"In\\\`":{\\\`"Expressions\\\`":[{\\\`"Column\\\`":{\\\`"Expression\\\`":{\\\`"SourceRef\\\`":{\\\`"Source\\\`":\\\`"g\\\`"}},\\\`"Property\\\`":\\\`"Distrikt\\\`"}}],\\\`"Values\\\`":[[{\\\`"Literal\\\`":{\\\`"Value\\\`":\\\`"'${DistriktName}'\\\`"}}]]}}}]},\\\`"expression\\\`":{\\\`"Column\\\`":{\\\`"Expression\\\`":{\\\`"SourceRef\\\`":{\\\`"Entity\\\`":\\\`"Grupper\\\`"}},\\\`"Property\\\`":\\\`"Distrikt\\\`"}},\\\`"howCreated\\\`":1}]},\\\`"sections\\\`":{\\\`"ReportSection\\\`":{\\\`"filters\\\`":{\\\`"byExpr\\\`":[{\\\`"name\\\`":\\\`"Filter486862a21e8e09d730c0\\\`",\\\`"type\\\`":\\\`"RelativeDate\\\`",\\\`"filter\\\`":{\\\`"Version\\\`":2,\\\`"From\\\`":[{\\\`"Name\\\`":\\\`"d\\\`",\\\`"Entity\\\`":\\\`"dimtid\\\`",\\\`"Type\\\`":0}],\\\`"Where\\\`":[{\\\`"Condition\\\`":{\\\`"Between\\\`":{\\\`"Expression\\\`":{\\\`"Column\\\`":{\\\`"Expression\\\`":{\\\`"SourceRef\\\`":{\\\`"Source\\\`":\\\`"d\\\`"}},\\\`"Property\\\`":\\\`"Dato\\\`"}},\\\`"LowerBound\\\`":{\\\`"DateSpan\\\`":{\\\`"Expression\\\`":{\\\`"DateAdd\\\`":{\\\`"Expression\\\`":{\\\`"Now\\\`":{}},\\\`"Amount\\\`":-4,\\\`"TimeUnit\\\`":3}},\\\`"TimeUnit\\\`":3}},\\\`"UpperBound\\\`":{\\\`"DateSpan\\\`":{\\\`"Expression\\\`":{\\\`"DateAdd\\\`":{\\\`"Expression\\\`":{\\\`"Now\\\`":{}},\\\`"Amount\\\`":-1,\\\`"TimeUnit\\\`":3}},\\\`"TimeUnit\\\`":3}}}}}]},\\\`"expression\\\`":{\\\`"Column\\\`":{\\\`"Expression\\\`":{\\\`"SourceRef\\\`":{\\\`"Entity\\\`":\\\`"dimtid\\\`"}},\\\`"Property\\\`":\\\`"Dato\\\`"}},\\\`"howCreated\\\`":1}]},\\\`"visualContainers\\\`":{}}},\\\`"objects\\\`":{\\\`"merge\\\`":{\\\`"outspacePane\\\`":[{\\\`"properties\\\`":{\\\`"expanded\\\`":{\\\`"expr\\\`":{\\\`"Literal\\\`":{\\\`"Value\\\`":\\\`"true\\\`"}}}}}]}}}\`"}`"}}"

    $reportRequestId = ($result | ConvertFrom-Json).id
    $reportRequestId
}

function Export-ReportForGruppe {
    param(
        $GruppeId,
        $GruppeNavn
    )

    Write-Host "Eksporterer for ${GruppeNavn}"

    $activityId = [System.Guid]::NewGuid().ToString("d")

    $result = Invoke-WebRequest -UseBasicParsing -Uri "https://wabi-europe-north-b-redirect.analysis.windows.net/export/reports/d4644347-1a7e-493a-ba7c-01ca397f8b70/asyncexports" `
        -Method "POST" `
        -WebSession $session `
        -Headers @{
        "authority"                  = "wabi-europe-north-b-redirect.analysis.windows.net"
        "accept-encoding"            = "gzip, deflate, br"
        "accept-language"            = "en-US,en;q=0.9,da;q=0.8"
        "authorization"              = $BearerToken
        "activityid"                 = $activityId
        "powerbi-reportsectioncount" = "2"
        "x-powerbi-hostenv"          = "Power BI Web"
    } `
        -ContentType "application/json;charset=UTF-8" `
        -Body "{`"format`":`"pdf`",`"powerBIReportConfiguration`":{`"settings`":{`"locale`":`"en-US`",`"timeZoneId`":`"Romance Standard Time`",`"excludeHiddenPages`":false},`"payload`":`"{\`"objectId\`":\`"28cd0166-4908-846f-61a3-0597e0e082f4\`",\`"type\`":99,\`"explorationState\`":\`"{\\\`"version\\\`":\\\`"1.0\\\`",\\\`"filters\\\`":{\\\`"byExpr\\\`":[{\\\`"name\\\`":\\\`"Filterd08fa18140bd39d72c69\\\`",\\\`"type\\\`":\\\`"Categorical\\\`",\\\`"filter\\\`":{\\\`"Version\\\`":2,\\\`"From\\\`":[{\\\`"Name\\\`":\\\`"g\\\`",\\\`"Entity\\\`":\\\`"Grupper\\\`",\\\`"Type\\\`":0}],\\\`"Where\\\`":[{\\\`"Condition\\\`":{\\\`"In\\\`":{\\\`"Expressions\\\`":[{\\\`"Column\\\`":{\\\`"Expression\\\`":{\\\`"SourceRef\\\`":{\\\`"Source\\\`":\\\`"g\\\`"}},\\\`"Property\\\`":\\\`"GruppeID\\\`"}}],\\\`"Values\\\`":[[{\\\`"Literal\\\`":{\\\`"Value\\\`":\\\`"${GruppeId}L\\\`"}}]]}}}]},\\\`"expression\\\`":{\\\`"Column\\\`":{\\\`"Expression\\\`":{\\\`"SourceRef\\\`":{\\\`"Entity\\\`":\\\`"Grupper\\\`"}},\\\`"Property\\\`":\\\`"GruppeID\\\`"}},\\\`"howCreated\\\`":1}]},\\\`"sections\\\`":{}}\`"}`"}}"    
    $reportRequestId = ($result | ConvertFrom-Json).id
    $reportRequestId
}

function Get-RequestedReport {
    param(
        $reportRequestId,
        $itemName
    )

    Write-Host "Afventer resultat for $itemName " -NoNewLine

    do {
        Start-Sleep -Seconds 2
        $statusResult = Invoke-WebRequest -UseBasicParsing -Uri "https://wabi-europe-north-b-redirect.analysis.windows.net/export/reports/d4644347-1a7e-493a-ba7c-01ca397f8b70/asyncexports/${reportRequestId}/status" `
            -WebSession $session `
            -Headers @{
            "authority"                  = "wabi-europe-north-b-redirect.analysis.windows.net"
            "accept-encoding"            = "gzip, deflate, br"
            "accept-language"            = "en-US,en;q=0.9,da;q=0.8"
            "authorization"              = $BearerToken
            "activityid"                 = $activityId
            "powerbi-reportsectioncount" = "2"
            "x-powerbi-hostenv"          = "Power BI Web"
        }
        $statusDocument = ($statusResult.Content | ConvertFrom-Json)
        $status = $statusDocument.status
        Write-Host "." -NoNewLine
    } while ($status -lt 3)

    Write-Host " " -NoNewline
    $fileResponse = Invoke-WebRequest -UseBasicParsing -Uri "https://wabi-europe-north-b-redirect.analysis.windows.net/export/reports/d4644347-1a7e-493a-ba7c-01ca397f8b70/asyncexports/${reportRequestId}/file" `
        -WebSession $session `
        -Proxy http://127.0.0.1:8888 `
        -Headers @{
        "authority"                  = "wabi-europe-north-b-redirect.analysis.windows.net"
        "accept"                     = "application/json, text/plain, */*"
        "accept-encoding"            = "gzip, deflate, br"
        "accept-language"            = "en-US,en;q=0.9,da;q=0.8"
        "activityid"                 = $activityId
        "authorization"              = $BearerToken
        "powerbi-reportsectioncount" = "2"
        "referer"                    = "https://app.powerbi.com/"
        "x-powerbi-hostenv"          = "Power BI Web"
    }
    $fileName = "{0} {1}{2}" -f $statusDocument.reportName, $itemName, $statusDocument.resourceFileExtension
    $fullPath = Join-Path $Destination $fileName.Replace("/", "_").Replace("\", "-")
    [System.IO.File]::WriteAllBytes($fullPath, $fileResponse.Content)
    Write-Host "Gemt $fileName"
}

function Select-Chunk {
    <#
    .SYNOPSIS
    Chunks pipeline input.
    
    .DESCRIPTION
    Chunks (partitions) pipeline input into arrays of a given size.
    
    By design, each such array is output as a *single* object to the pipeline,
    so that the next command in the pipeline can process it as a whole.
    
    That is, for the next command in the pipeline $_ contains an *array* of
    (at most) as many elements as specified via -ReadCount.
    
    .PARAMETER InputObject
    The pipeline input objects binds to this parameter one by one.
    Do not use it directly.
    
    .PARAMETER ReadCount
    The desired size of the chunks, i.e., how many input objects to collect
    in an array before sending that array to the pipeline.
    
    0 effectively means: collect *all* inputs and output a single array overall.
    
    .EXAMPLE
    1..7 | Select-Chunk 3 | ForEach-Object { "$_" }
    
    1 2 3
    4 5 6
    7
    
    The above shows that the ForEach-Object script block receive the following
    three arrays: (1, 2, 3), (4, 5, 6), and (, 7)
    #>
    
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([object[]])]
    param (
        [Parameter(ValueFromPipeline)] 
        $InputObject
        ,
        [Parameter(Mandatory, Position = 0)]
        [ValidateRange(0, [int]::MaxValue)]
        [int] $ReadCount
    )
        
    begin {
        $q = [System.Collections.Generic.Queue[object]]::new($ReadCount)
    }
        
    process {
        $q.Enqueue($InputObject)
        if ($q.Count -eq $ReadCount) {
            , $q.ToArray()
            $q.Clear()
        }
    }
        
    end {
        if ($q.Count) {
            , $q.ToArray()
        }
    }
  
}


If ($Distrikter -or $Grupper) {
    Write-Host "Eksporterer til ${Destination}"
}

function ExecuteFile {
    param(
        $FilePath,
        [scriptblock]$ItemScript
    )

    Write-Host "Importerer $FilePath"
    $data = Get-Content $FilePath -Encoding UTF8 | ConvertFrom-Csv -Delimiter ";"
    $chunks = $data | Select-Object -Skip $Skip | Select-Chunk $ChunkSize
    $chunks | ForEach-Object  {
        $chunk = $_
        $requested = $chunk | ForEach-Object {
            Invoke-Command $ItemScript
        }
        $requested | ForEach-Object {
            $itm = $_
            Get-RequestedReport -reportRequestId $itm.ReportId -itemName $itm.ItemName
        }
    }
    
    Write-Host ""

}

If ($Distrikter) {

    ExecuteFile -FilePath $DistrikterFile -ItemScript {
        $name = $_.$DistrikterFileNameField
        $reportId = Export-ReportForDistrikt -DistriktName $name
        [PSCustomObject]@{
            Name     = $name
            ReportId = $reportId
            ItemName = $name
        }
    }
}

If ($Grupper) {
    ExecuteFile -FilePath $GrupperFile -ItemScript {
        $name = $_.$GrupperFileNameField
        $id = $_.$GrupperFileIdField
        $reportId = Export-ReportForGruppe -GruppeId $id -GruppeNavn $name
        [PSCustomObject]@{
            Name     = $name
            Id       = $id
            ReportId = $reportId
            ItemName = "${id} ${name}"
        }
    }
}

Write-Host "OK" -ForegroundColor Green
Write-Host ""
