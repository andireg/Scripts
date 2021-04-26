param(
    [string] $organization,
    [string] $project,
    [string] $dashboardId,
    [string] $detailWidgetId,
    [string] $detailWikiPage,
    [string] $teamsWebhook,
    [string] $personalAccessToken
)

$uriPrefix = "https://dev.azure.com/$($organization)/$($project)/_apis"
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("PersonalAccessToken:$personalAccessToken"))
$header = @{Authorization = ("Basic {0}" -f $base64AuthInfo) }

function Get-BuildPipelines() {
    $response = Invoke-RestMethod `
        -Uri "$($uriPrefix)/build/definitions?api-version=6.0" `
        -Method Get `
        -ContentType "application/json" `
        -Headers $header    
    return $response.value | 
    ForEach-Object { Get-BuildPipelineInfo($_) } | 
    Where-Object { $_ -ne $null }
}

function Get-FirstFailingBuild ($pipelineId) {
    $runs = Invoke-RestMethod `
        -Uri "$($uriPrefix)/build/builds?definitions=$($pipelineId)&`$top=20&queryOrder=queueTimeDescending&api-version=6.0" `
        -Method Get `
        -ContentType "application/json" `
        -Headers $header    

    $lastRun = $null
    foreach ($run in $runs.value) {
        if (($run.status -eq "completed") -and ($run.result -eq "succeeded")) {
            return $lastRun
        }

        $lastRun = $run
    }

    return $lastRun
}

function Get-AuthorOfBuild($run) {    
    if ($run.sourceBranch -like 'refs/*') {
        $commit = Invoke-RestMethod `
            -Uri "$($uriPrefix)/git/repositories/$($run.repository.id)/commits/$($run.sourceVersion)?`$top=1" `
            -Method Get `
            -ContentType "application/json" `
            -Headers $header
        return @{
            Display = $commit.author.name
            Unique  = $commit.author.email
            Image   = $commit.author.imageUrl
        }
    }
    else {
        $changeset = Invoke-RestMethod `
            -Uri "$($uriPrefix)/tfvc/changesets/$($run.sourceVersion)?`$top=1" `
            -Method Get `
            -ContentType "application/json" `
            -Headers $header
        return @{
            Display = $changeset.author.displayName
            Unique  = $changeset.author.uniqueName
            Image   = $changeset.author.imageUrl
        }
    }
}

function Get-FormattedDate($value) {
    [datetime]$date = New-Object DateTime
    if ([DateTime]::TryParse($value, [ref]$date)) {
        return $date.ToString("dd.MM.yyyy HH:mm")
    }

    return ""
}

function Get-BuildPipelineInfo($pipeline) {
    $result = @{
        Pipeline       = @{
            Id   = $pipeline.id
            Name = $pipeline.name
        }
        Succeeded      = $null
        LastRun        = $null
        FirstFailedRun = $null
    }

    $runs = Invoke-RestMethod `
        -Uri "$uriPrefix/build/builds?definitions=$($pipeline.id)&`$top=1&queryOrder=queueTimeDescending&api-version=6.0" `
        -Method Get `
        -ContentType "application/json" `
        -Headers $header
    if ($runs.value.Count -eq 0) {
        return $result
    }

    $run = $runs.value[0]
    $result.LastRun = @{
        Id     = $run.id
        Link   = $run._links.web.href
        Status = $run.status
        Result = $run.result
        Date   = $run.finishTime
        Author = @{
            Display = $run.requestedFor.displayName
            Unique  = $run.requestedFor.uniqueName
            Image   = $run.requestedFor.imageUrl
        }
    }

    if (($run.status -ne "completed") -or ($run.result -eq "succeeded") -or ($run.result -eq "partiallySucceeded")) {
        $result.Succeeded = $true
        return $result
    }

    $result.Succeeded = $false
    $firstFailingBuild = Get-FirstFailingBuild $pipeline.id
    $result.FirstFailedRun = @{
        Id     = $firstFailingBuild.id
        Link   = $firstFailingBuild._links.web.href
        Status = $firstFailingBuild.status
        Result = $firstFailingBuild.result
        Date   = $firstFailingBuild.finishTime
        Author = @{
            Display = $firstFailingBuild.requestedFor.displayName
            Unique  = $firstFailingBuild.requestedFor.uniqueName
            Image   = $firstFailingBuild.requestedFor.imageUrl
        }
    }

    if ($result.LastRun.Author.Display -eq "Microsoft.VisualStudio.Services.TFS") {
        $result.LastRun.Author = Get-AuthorOfBuild($run)
    }

    if ($result.FirstFailedRun.Author.Display -eq "Microsoft.VisualStudio.Services.TFS") {
        $result.FirstFailedRun.Author = Get-AuthorOfBuild($run)
    }
    
    return $result
}

function Send-MsTeamsMessages($webhook, $pipelines) {
    foreach ($grouped in $pipelines | 
        Where-Object { $_.Succeeded -eq $false } | 
        Group-Object -Property { $_.FirstFailedRun.Author.Unique }) {        
        $sb = [System.Text.StringBuilder]::new()
        foreach ($build in $grouped.Group | 
            Sort-Object -Property { $_.Pipeline.Name }) {
            $sb.AppendLine("-   **$($build.Pipeline.Name)**") >$null 2>&1
            if ($build.FirstFailedRun.Id -eq $build.LastRun.Id) {
                $sb.AppendLine("    -   Run: [$($build.FirstFailedRun.Result) at $(Get-FormattedDate($build.FirstFailedRun.Date))]($($build.FirstFailedRun.Link))") >$null 2>&1
            }
            else {
                $sb.AppendLine("    -   Breaking run: [$($build.FirstFailedRun.Result) at $(Get-FormattedDate($build.FirstFailedRun.Date))]($($build.FirstFailedRun.Link))") >$null 2>&1
                $sb.AppendLine("    -   Last run: [$($build.LastRun.Result) at $(Get-FormattedDate($build.LastRun.Date))]($($build.LastRun.Link))") >$null 2>&1
            }
        }   

        $firstPipeline = $grouped.Group[0]
        $card = @{
            "@type"    = "MessageCard"
            "@context" = "http://schema.org/extensions"
            themeColor = "C10000"
            summary    = "$($firstPipeline.FirstFailedRun.Author.Display) need to check some builds"
            sections   = @(
                @{
                    activityTitle    = $firstPipeline.FirstFailedRun.Author.Display
                    activitySubtitle = $firstPipeline.FirstFailedRun.Author.Unique
                    activityImage    = "https://eu.ui-avatars.com/api/?name=$($firstPipeline.FirstFailedRun.Author.Display)&background=random&rounded=true"
                }
                @{
                    type = "TextBlock"
                    text = "Hi <at>$($firstPipeline.FirstFailedRun.Author.Unique)</at> - please check this builds:"
                }
                @{
                    type = "TextBlock"
                    text = $sb.ToString()
                    wrap = $true
                }
            )
        }

        $json = ConvertTo-Json -InputObject $card -Depth 100
        Invoke-RestMethod -Method post -ContentType 'application/json' -Body $json -Uri $webhook >$null 2>&1
    }
}

function Update-WikiPage($wikiPage, $pipelines) {
    [System.Text.StringBuilder] $sb = [System.Text.StringBuilder]::new()
    $sb.AppendLine("# Failed Build Details") >$null 2>&1
    $sb.AppendLine("") >$null 2>&1
    $sb.AppendLine("") >$null 2>&1
    $sb.AppendLine("| **Pipeline** | **Breaking** | **Last** | **Responsible** |") >$null 2>&1
    $sb.AppendLine("| --- | --- | --- | --- |") >$null 2>&1
    foreach ($pipeline in $pipelines | 
        Where-Object { $_.Succeeded -eq $false } | 
        Sort-Object -Property { $_.Pipeline.Name }) {
        $sb.Append("| $([System.Web.HttpUtility]::HtmlEncode($pipeline.Pipeline.Name)) ") >$null 2>&1
        $sb.Append("| [$([System.Web.HttpUtility]::HtmlEncode($pipeline.FirstFailedRun.Result)) at $(Get-FormattedDate($pipeline.FirstFailedRun.Date))]($($pipeline.FirstFailedRun.Link)) ") >$null 2>&1
        $sb.Append("| [$([System.Web.HttpUtility]::HtmlEncode($pipeline.LastRun.Result)) at $(Get-FormattedDate($pipeline.LastRun.Date))]($($pipeline.LastRun.Link)) ") >$null 2>&1
        $sb.AppendLine("| $([System.Web.HttpUtility]::HtmlEncode($pipeline.FirstFailedRun.Author.Display)) ($([System.Web.HttpUtility]::HtmlEncode($pipeline.FirstFailedRun.Author.Unique))) |") >$null 2>&1
    }

    $uri = "$($uriPrefix)/wiki/wikis/$($wikiPage)?includeContent=True&api-version=6.0-preview"
    $proxyUri = [Uri]$null
    $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
    if ($proxy) {
        $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
        $proxyUri = $proxy.GetProxy($uri)
    }
    $response = Invoke-WebRequest `
        -Uri $uri `
        -Method Get `
        -ContentType "application/json" `
        -Proxy $proxyUri `
        -ProxyUseDefaultCredentials `
        -Headers $header `
        -UseBasicParsing
    $etag = $response.Headers["ETag"]
    $content = ConvertFrom-Json $([string]::new($response.content)) 
    $content.content = $sb.ToString()
    $patchHeader = @{
        Authorization = "Basic $base64AuthInfo"
        Accept        = "application/json"
        'If-Match'    = $etag
    }
    $json = ConvertTo-Json -InputObject $content -Depth 100
    $response = Invoke-RestMethod `
        -Uri "$($uriPrefix)/wiki/wikis/$($wikiPage)?api-version=6.0-preview" `
        -Method Patch `
        -Headers $patchHeader `
        -UseDefaultCredentials `
        -ContentType "application/json" `
        -Body $json >$null 2>&1
}

function Update-MarkdownWidget($dashboardId, $widgetId, $markdown) {
    $widgetData = Invoke-RestMethod `
        -Uri "$($uriPrefix)/dashboard/dashboards/$($dashboardId)/widgets/$($widgetId)?api-version=6.0-preview.2" `
        -Method Get `
        -ContentType "application/json" `
        -Headers $header
    $widgetData.settings = $markdown
    $json = ConvertTo-Json -InputObject $widgetData -Depth 100
    Invoke-RestMethod -Uri "$($uriPrefix)/dashboard/dashboards/$($dashboardId)/widgets/$($widgetId)?api-version=6.0-preview.2" -Method Patch -ContentType "application/json" -Headers $header -Body $json >$null 2>&1
}

function Update-MarkdownWidget-FailedBuildDetail($dashboardId, $widgetId, $pipelines) {
    
    [System.Text.StringBuilder] $sb = [System.Text.StringBuilder]::new()
    $sb.AppendLine("# Failed Build Details") >$null 2>&1
    $sb.AppendLine("") >$null 2>&1
    $sb.AppendLine("") >$null 2>&1
    $sb.AppendLine("| **Pipeline** | **Breaking** |") >$null 2>&1
    $sb.AppendLine("| --- | --- |") >$null 2>&1
    foreach ($pipeline in $pipelines | 
        Where-Object { $_.Succeeded -eq $false } | 
        Sort-Object -Property { $_.Pipeline.Name }) {
        $sb.Append("| $([System.Web.HttpUtility]::HtmlEncode($pipeline.Pipeline.Name)) ") >$null 2>&1
        $sb.AppendLine("| [$($pipeline.LastRun.Author.Display)]($($pipeline.LastRun.Link)) ") >$null 2>&1
    }

    Update-MarkdownWidget $dashboardId $detailWidgetId $sb.ToString()
}

# can be used to get dashboard details
function Write-Dashboard($dashboardId) {
    $response = Invoke-RestMethod `
        -Uri "$($uriPrefix)/dashboard/dashboards/$($dashboardId)?api-version=6.0-preview.2" `
        -Method Get `
        -ContentType "application/json" `
        -Headers $header
    $json = ConvertTo-Json -InputObject $response -Depth 100
    Write-Output $json
}

$pipelines = Get-BuildPipelines

if ($null -ne $dashboardId -and $null -ne $detailWidgetId) {
    Update-MarkdownWidget-FailedBuildDetail $dashboardId $detailWidgetId $pipelines
}

if ($null -ne $detailWikiPage) {
    Update-WikiPage $detailWikiPage $pipelines
}

if ($null -ne $teamsWebhook) {
    Send-MsTeamsMessages $teamsWebhook $pipelines
}
