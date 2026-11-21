<#
.SYNOPSIS
    统计 vllm-ascend 仓库中带指定 label（默认 documentation）的 open issue，生成可定期刷新的表格（CSV）。

.DESCRIPTION
    - 只统计 open 状态、且不是 PR 的 issue。
    - 数据源：GitHub 搜索接口取 issue 列表 + issue 页面内嵌数据取评论/关联 PR（无需 token）。
    - 结果增量缓存到 .cache/issues，只有 updated_at 变化的 issue 才会重新抓取，适合定时反复执行。
    - 输出 CSV 中的「备注」列在多次运行之间保留，供人工填写。

.EXAMPLE
    .\vllm-ascend-docs-issues.ps1

.EXAMPLE
    .\vllm-ascend-docs-issues.ps1 -RefreshAll

.EXAMPLE
    .\vllm-ascend-docs-issues.ps1 -InstallSchedule -ScheduleTime 09:00

.NOTES
    GitHub Actions 等可直连外网的场景直接运行即可；本机经过公司代理且存在 TLS 中间人时，
    需加 -SkipTlsCheck 跳过证书校验。
#>
[CmdletBinding()]
param(
    [string]$Repo = 'vllm-project/vllm-ascend',
    [string]$Label = 'documentation',
    [string]$OutputPath,
    [int]$DelayMs = 300,
    [switch]$RefreshAll,
    [switch]$SkipTlsCheck,
    [switch]$InstallSchedule,
    [string]$ScheduleTime = '09:00',
    [string]$TaskName = 'vllm-ascend-docs-issues'
)

$ErrorActionPreference = 'Stop'

if (-not $OutputPath) { $OutputPath = Join-Path $PSScriptRoot 'vllm-ascend-docs-issues.csv' }
$CacheRoot = Join-Path $PSScriptRoot '.cache'
$IssueCacheDir = Join-Path $CacheRoot 'issues'
$ListCache = Join-Path $CacheRoot 'issue-list.json'
New-Item -ItemType Directory -Force -Path $IssueCacheDir | Out-Null

# Windows 上是 curl.exe，Linux/macOS 上是 curl
$CurlExe = if (Get-Command 'curl.exe' -ErrorAction SilentlyContinue) { 'curl.exe' } else { 'curl' }

# 机器人账号：不计入“已回复”
$BotPattern = '(\[bot\]$)|(^github-actions)|(-bot$)|(^bot$)|(^renovate)'
# 视为“给出了具体解决方案/修复方案”的文本特征
$SolutionPatterns = @(
    'https?://github\.com/[\w.\-]+/[\w.\-]+/pull/\d+',
    '(?i)\b(fix(es|ed)?|close[sd]?|resolve[sd]?|addressed)\b[^\r\n]{0,24}#\d+',
    '(?i)\b(has|have|will)\s+been\s+(fixed|updated|added|changed|released|merged)\b',
    '(?i)\b(see|refer to)\b[^\r\n]{0,16}\b(pr|pull request)\b',
    '已(经)?(修复|更新|解决|补充|修改|添加|提交|修正)'
)
# 视为“已明确解决”的文本特征（用于判断是否应该关闭，口径更严格）
$ResolvedPatterns = @(
    '(?i)\b(fixed|resolved|addressed)\s+(by|in)\b',
    '(?i)\bhas\s+been\s+(fixed|resolved|addressed|updated)\b',
    '已(经)?(修复|解决|更新)'
)
# 维护者/协作者身份
$MaintainerAssociations = @('MEMBER', 'OWNER', 'COLLABORATOR')

function Invoke-CurlGet {
    param([string]$Url, [string]$OutFile)
    $curlArgs = @(
        '-s', '-L', '-m', '90', '--retry', '2', '--retry-delay', '2',
        '-A', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
        '-o', $OutFile, '-w', '%{http_code}', $Url
    )
    if ($SkipTlsCheck) { $curlArgs = @('-k') + $curlArgs }
    $code = (& $CurlExe @curlArgs) -join ''
    return [int]$code
}

function Get-OpenIssueList {
    $items = @()
    $page = 1
    $query = [uri]::EscapeDataString("repo:$Repo label:$Label is:issue is:open")
    while ($true) {
        $url = "https://api.github.com/search/issues?q=$query&per_page=100&page=$page"
        $tmp = Join-Path $CacheRoot 'search.tmp.json'
        $code = Invoke-CurlGet -Url $url -OutFile $tmp
        if ($code -ne 200) {
            if (Test-Path $ListCache) {
                Write-Warning "获取 issue 列表失败(HTTP $code)，改用上次缓存。"
                return @(Get-Content $ListCache -Raw -Encoding UTF8 | ConvertFrom-Json)
            }
            throw "获取 issue 列表失败，HTTP $code"
        }
        $r = Get-Content $tmp -Raw -Encoding UTF8 | ConvertFrom-Json
        $items += @($r.items)
        if (@($r.items).Count -lt 100) { break }
        $page++
        Start-Sleep -Milliseconds 1200
    }
    , @($items) | ConvertTo-Json -Depth 8 -Compress | Set-Content -Path $ListCache -Encoding UTF8
    return @($items)
}

function Get-IssueDetail {
    param([int]$Number)

    $url = "https://github.com/$Repo/issues/$Number"
    $tmp = Join-Path $CacheRoot ("page_" + $Number + ".html")
    $code = Invoke-CurlGet -Url $url -OutFile $tmp
    if ($code -ne 200) { Remove-Item $tmp -ErrorAction SilentlyContinue; return $null }

    $html = [System.IO.File]::ReadAllText($tmp)
    Remove-Item $tmp -ErrorAction SilentlyContinue
    $m = [regex]::Match($html, '(?s)<script type="application/json" data-target="react-app\.embeddedData">(.*?)</script>')
    if (-not $m.Success) { return $null }

    $data = $m.Groups[1].Value | ConvertFrom-Json
    $issue = $data.payload.issueViewerRoute.data.repository.issue
    if (-not $issue -or -not $issue.number) { return $null }
    if ($issue.__typename -eq 'PullRequest') { return $null }

    $labels = @($issue.labels.edges | ForEach-Object { $_.node.name })
    $comments = @()
    $crossRefPRs = @()
    $closeCommits = @()

    foreach ($edge in @($issue.timelineItems.edges)) {
        $node = $edge.node
        switch ($node.__typename) {
            'IssueComment' {
                $comments += [pscustomobject]@{
                    author      = $node.author.login
                    association = $node.authorAssociation
                    body        = $node.body
                    createdAt   = $node.createdAt
                    url         = $node.url
                }
            }
            'CrossReferencedEvent' {
                $src = $node.source
                if ($src -and $src.__typename -eq 'PullRequest') {
                    $crossRefPRs += [pscustomobject]@{
                        number = $src.number
                        title  = $src.title
                        state  = $src.state
                        url    = $src.url
                    }
                }
            }
            'ReferencedEvent' {
                if ($node.willCloseSubject) {
                    $closeCommits += [pscustomobject]@{
                        sha     = ($node.commit.id + '')
                        message = ($node.commit.message + '').Split("`n")[0]
                    }
                }
            }
        }
    }

    $result = [pscustomobject]@{
        number       = [int]$issue.number
        title        = $issue.title
        state        = $issue.state
        createdAt    = $issue.createdAt
        updatedAt    = $issue.updatedAt
        author       = $issue.author.login
        body         = $issue.body
        url          = "https://github.com/$Repo/issues/$($issue.number)"
        labels       = $labels
        comments     = $comments
        crossRefPRs  = $crossRefPRs
        linkedPRs    = @($issue.linkedPullRequests.nodes | ForEach-Object {
                [pscustomobject]@{ number = $_.number; title = $_.title; state = $_.state; url = $_.url }
            })
        closedByPRs  = @($issue.closedByPullRequestsReferences.nodes | ForEach-Object {
                [pscustomobject]@{ number = $_.number; title = $_.title; state = $_.state; url = $_.url }
            })
        closeCommits = $closeCommits
        fetchedAt    = (Get-Date).ToString('s')
    }
    return $result
}

function Test-SolutionText {
    param([string]$Text)
    if (-not $Text) { return $false }
    foreach ($p in $SolutionPatterns) {
        if ($Text -match $p) { return $true }
    }
    return $false
}

function Test-ResolvedText {
    param([string]$Text)
    if (-not $Text) { return $false }
    foreach ($p in $ResolvedPatterns) {
        if ($Text -match $p) { return $true }
    }
    return $false
}

function Get-IssueRow {
    param($Issue, [string]$Remark)

    $author = $Issue.author
    $humanReplies = @($Issue.comments | Where-Object {
            $_ -and $_.author -and $_.author -ne $author -and $_.author -notmatch $BotPattern
        })
    $replied = $humanReplies.Count -gt 0
    $replyText = ($humanReplies | ForEach-Object { $_.body }) -join "`n"

    $prRefs = @($Issue.linkedPRs) + @($Issue.closedByPRs) + @($Issue.crossRefPRs)
    $solvedByPR = $prRefs.Count -gt 0
    # 维护者给出了较长/带命令的技术性回复，也视为提供了具体方案
    $technicalReply = @($humanReplies | Where-Object {
            $_.association -in $MaintainerAssociations -and $_.body -and $_.body.Length -ge 200
        }).Count -gt 0
    $hasSolution = $solvedByPR -or $technicalReply -or (Test-SolutionText $replyText)

    # 只有真正 merged 的 PR 才算“已解决”；closed 但未合并不算
    $mergedPR = @($prRefs | Where-Object { $_.state -eq 'MERGED' }).Count -gt 0
    $dupLabel = @($Issue.labels) -contains 'duplicate'
    # 维护者明确给出结论/方案（排除仍在追问信息的回复）才算“已经明确回复”
    $maintainerSolved = @($humanReplies | Where-Object {
            if ($_.association -notin $MaintainerAssociations -or -not $_.body) { return $false }
            if (Test-ResolvedText $_.body) { return $true }
            if ($_.body -match 'github\.com/[\w.\-]+/[\w.\-]+/pull/\d+') { return $true }
            return ($_.body.Length -ge 200 -and $_.body -notmatch '[?？]')
        }).Count -gt 0

    $shouldClose = $mergedPR -or $dupLabel -or $maintainerSolved

    $statusLabels = @($Issue.labels | Where-Object { $_ -ne $Label })

    [pscustomobject][ordered]@{
        'Issue编号'          = $Issue.number
        'Issue状态'          = $Issue.state.ToLower()
        '是否应该关闭'       = if ($shouldClose) { '是' } else { '否' }
        '状态标签'           = ($statusLabels -join ', ')
        '是否已回复'         = if ($replied) { '是' } else { '否' }
        '是否已提供具体解决方案' = if ($hasSolution) { '是' } else { '否' }
        'Issue链接'          = $Issue.url
        '备注'               = $Remark
    }
}

# ---------- 主流程 ----------
Write-Host "仓库: $Repo | label: $Label | 输出: $OutputPath"

$list = Get-OpenIssueList
Write-Host "open issue 数量: $($list.Count)"

$details = New-Object System.Collections.Generic.List[object]
$i = 0
$fetched = 0
foreach ($item in $list) {
    $i++
    $n = [int]$item.number
    $cacheFile = Join-Path $IssueCacheDir "$n.json"
    $cached = $null
    if (-not $RefreshAll -and (Test-Path $cacheFile)) {
        try { $cached = Get-Content $cacheFile -Raw -Encoding UTF8 | ConvertFrom-Json } catch { $cached = $null }
        if ($cached -and $cached.updatedAt -ne $item.updated_at) { $cached = $null }
    }
    if (-not $cached) {
        $cached = Get-IssueDetail -Number $n
        if ($cached) {
            $cached | ConvertTo-Json -Depth 8 -Compress | Set-Content -Path $cacheFile -Encoding UTF8
            $fetched++
            Start-Sleep -Milliseconds $DelayMs
        }
    }
    if (-not $cached) {
        Write-Warning "issue #$n 抓取失败，已跳过"
        continue
    }
    $details.Add($cached)
    if ($i % 25 -eq 0) { Write-Host "进度 $i/$($list.Count)（本次抓取 $fetched）" }
}

# 保留历史「备注」
$remarkMap = @{}
if (Test-Path $OutputPath) {
    try {
        foreach ($old in (Import-Csv -Path $OutputPath -Encoding UTF8)) {
            if ($old.'Issue编号') { $remarkMap[$old.'Issue编号'] = $old.'备注' }
        }
    } catch {
        Write-Warning "读取旧表格失败，备注不会被保留: $($_.Exception.Message)"
    }
}

$rows = foreach ($d in ($details | Sort-Object number)) {
    Get-IssueRow -Issue $d -Remark $remarkMap[[string]$d.number]
}

# 显式写入 UTF-8 BOM，保证 Excel 直接打开中文不乱码（PS5 / PS7 行为一致）
$csv = ($rows | ConvertTo-Csv -NoTypeInformation) -join "`r`n"
[System.IO.File]::WriteAllText($OutputPath, $csv + "`r`n", (New-Object System.Text.UTF8Encoding($true)))

Write-Host ""
Write-Host "完成：$($rows.Count) 条 open issue，本次网络抓取 $fetched 条"
Write-Host ("是否应该关闭=是: {0} / 否: {1}" -f @($rows | Where-Object { $_.'是否应该关闭' -eq '是' }).Count, @($rows | Where-Object { $_.'是否应该关闭' -eq '否' }).Count)
Write-Host ("未回复: {0} / 未提供解决方案: {1}" -f @($rows | Where-Object { $_.'是否已回复' -eq '否' }).Count, @($rows | Where-Object { $_.'是否已提供具体解决方案' -eq '否' }).Count)
Write-Host "表格已写入: $OutputPath"

if ($InstallSchedule) {
    if (-not (Get-Command schtasks -ErrorAction SilentlyContinue)) {
        Write-Warning '当前系统不支持 schtasks，请改用 cron 或 GitHub Actions 定时执行。'
    } else {
        $tr = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' + $PSCommandPath + '"'
        $res = schtasks /Create /TN $TaskName /SC DAILY /ST $ScheduleTime /TR $tr /F 2>&1
        Write-Host "已注册每日定时任务 [$TaskName] $ScheduleTime : $res"
    }
}
