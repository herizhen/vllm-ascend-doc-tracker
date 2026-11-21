<#
.SYNOPSIS
    统计 vllm-ascend 仓库中带指定 label（默认 documentation）的 issue，生成 CSV / Excel / Markdown 表格。

.DESCRIPTION
    - 数据源：GitHub REST API，需要 token（GITHUB_TOKEN 环境变量或 -Token 参数）。
    - 增量更新：历史 issue 仅刷新「Issue状态」列，新 issue 才全量抓取评论/关联 PR 并计算所有列。
    - 已关闭 issue 保留在表格中，状态列显示 closed，便于追溯。
    - 输出 CSV（用于下次增量合并）、XLSX（Excel 查看）、MD（仓库内直接阅读）。

.EXAMPLE
    .\vllm-ascend-docs-issues.ps1 -Token ghp_xxx

.EXAMPLE
    .\vllm-ascend-docs-issues.ps1 -RefreshAll

.NOTES
    Excel 输出依赖 ImportExcel 模块：Install-Module ImportExcel -Scope CurrentUser
    本机经过公司代理且存在 TLS 中间人时，需加 -SkipTlsCheck。
#>
[CmdletBinding()]
param(
    [string]$Repo = 'vllm-project/vllm-ascend',
    [string]$Label = 'documentation',
    [string]$Token = $env:GITHUB_TOKEN,
    [string]$OutputDir = $PSScriptRoot,
    [int]$DelayMs = 100,
    [switch]$RefreshAll,
    [switch]$SkipTlsCheck
)

$ErrorActionPreference = 'Stop'

if (-not $Token) {
    throw '未提供 GitHub token。请设置 GITHUB_TOKEN 环境变量或使用 -Token 参数。'
}

$CsvPath = Join-Path $OutputDir 'vllm-ascend-docs-issues.csv'
$XlsxPath = Join-Path $OutputDir 'vllm-ascend-docs-issues.xlsx'
$MdPath = Join-Path $OutputDir 'vllm-ascend-docs-issues.md'
$CacheRoot = Join-Path $OutputDir '.cache'
New-Item -ItemType Directory -Force -Path $CacheRoot | Out-Null

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

function Invoke-GitHubApi {
    param([string]$Url)

    $tmp = Join-Path $CacheRoot ("api_" + [guid]::NewGuid().ToString('N') + ".json")
    $curlArgs = @(
        '-s', '-L', '-m', '60', '--retry', '2', '--retry-delay', '2',
        '-H', 'Accept: application/vnd.github+json',
        '-H', 'X-GitHub-Api-Version: 2022-11-28',
        '-H', 'User-Agent: vllm-ascend-docs-tracker',
        '-o', $tmp, '-w', '%{http_code}', $Url
    )
    $curlArgs += @('-H', "Authorization: Bearer $Token")
    if ($SkipTlsCheck) { $curlArgs = @('-k') + $curlArgs }

    $code = (& $CurlExe @curlArgs) -join ''
    if ($code -ne 200) {
        $body = if (Test-Path $tmp) { Get-Content $tmp -Raw -Encoding UTF8 } else { '' }
        Remove-Item $tmp -ErrorAction SilentlyContinue
        throw "GitHub API 请求失败 HTTP $code : $Url`n$body"
    }

    $json = Get-Content $tmp -Raw -Encoding UTF8 | ConvertFrom-Json
    Remove-Item $tmp -ErrorAction SilentlyContinue
    return $json
}

function Get-OpenIssueList {
    $items = @()
    $page = 1
    $query = [uri]::EscapeDataString("repo:$Repo label:$Label is:issue is:open")
    while ($true) {
        $url = "https://api.github.com/search/issues?q=$query&per_page=100&page=$page"
        $r = Invoke-GitHubApi -Url $url
        $items += @($r.items)
        if (@($r.items).Count -lt 100) { break }
        $page++
        Start-Sleep -Milliseconds 1000
    }
    return @($items)
}

function Get-IssueDetail {
    param([int]$Number)

    $issue = Invoke-GitHubApi -Url "https://api.github.com/repos/$Repo/issues/$Number"
    if ($issue.pull_request) { return $null }

    $comments = @()
    $crossRefPRs = @()
    $page = 1
    while ($true) {
        $timeline = Invoke-GitHubApi -Url "https://api.github.com/repos/$Repo/issues/$Number/timeline?per_page=100&page=$page"
        $events = @($timeline)
        if ($events.Count -eq 0) { break }

        foreach ($event in $events) {
            switch ($event.event) {
                'commented' {
                    $comments += [pscustomobject]@{
                        author      = $event.user.login
                        association = $event.author_association
                        body        = $event.body
                        createdAt   = $event.created_at
                    }
                }
                'cross-referenced' {
                    $src = $event.source.issue
                    if ($src -and $src.pull_request) {
                        $merged = [bool]$src.pull_request.merged_at
                        # timeline 里 merged_at 可能缺失，closed 状态的 PR 用 PR 接口确认一次是否真正 merged
                        if (-not $merged -and $src.state -eq 'closed' -and $src.pull_request.url) {
                            try {
                                $pr = Invoke-GitHubApi -Url $src.pull_request.url
                                $merged = [bool]$pr.merged_at
                            } catch { }
                        }
                        $crossRefPRs += [pscustomobject]@{
                            number = $src.number
                            title  = $src.title
                            state  = $src.state.ToUpper()
                            url    = $src.html_url
                            merged = $merged
                        }
                    }
                }
            }
        }
        if ($events.Count -lt 100) { break }
        $page++
    }

    [pscustomobject]@{
        number      = [int]$issue.number
        title       = $issue.title
        state       = $issue.state
        createdAt   = $issue.created_at
        updatedAt   = $issue.updated_at
        author      = $issue.user.login
        body        = $issue.body
        url         = $issue.html_url
        labels      = @($issue.labels | ForEach-Object { $_.name })
        comments    = $comments
        crossRefPRs = $crossRefPRs
    }
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

    $prRefs = @($Issue.crossRefPRs)
    $solvedByPR = $prRefs.Count -gt 0
    $technicalReply = @($humanReplies | Where-Object {
            $_.association -in $MaintainerAssociations -and $_.body -and $_.body.Length -ge 200
        }).Count -gt 0
    $hasSolution = $solvedByPR -or $technicalReply -or (Test-SolutionText $replyText)

    # 只有真正 merged 的 PR 才算“已解决”；closed 但未合并不算
    $mergedPR = @($prRefs | Where-Object { $_.merged }).Count -gt 0
    $dupLabel = @($Issue.labels) -contains 'duplicate'
    $maintainerSolved = @($humanReplies | Where-Object {
            if ($_.association -notin $MaintainerAssociations -or -not $_.body) { return $false }
            if (Test-ResolvedText $_.body) { return $true }
            if ($_.body -match 'github\.com/[\w.\-]+/[\w.\-]+/pull/\d+') { return $true }
            return ($_.body.Length -ge 200 -and $_.body -notmatch '[?？]')
        }).Count -gt 0

    $shouldClose = $mergedPR -or $dupLabel -or $maintainerSolved

    $statusLabels = @($Issue.labels | Where-Object { $_ -ne $Label })

    [pscustomobject][ordered]@{
        'Issue编号'             = $Issue.number
        'Issue状态'             = $Issue.state.ToLower()
        '是否应该关闭'          = if ($shouldClose) { '是' } else { '否' }
        '状态标签'              = ($statusLabels -join ', ')
        '是否已回复'            = if ($replied) { '是' } else { '否' }
        '是否已提供具体解决方案' = if ($hasSolution) { '是' } else { '否' }
        'Issue链接'             = $Issue.url
        '备注'                  = $Remark
    }
}

function Export-CsvFile {
    param($Rows, [string]$Path)
    $csv = ($Rows | ConvertTo-Csv -NoTypeInformation) -join "`r`n"
    [System.IO.File]::WriteAllText($Path, $csv + "`r`n", (New-Object System.Text.UTF8Encoding($true)))
}

function Export-MarkdownFile {
    param($Rows, [string]$Path)

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('# vLLM Ascend Documentation Issues')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("> 更新时间: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | 仓库: $Repo | Label: $Label")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('| Issue编号 | Issue状态 | 是否应该关闭 | 状态标签 | 是否已回复 | 是否已提供具体解决方案 | Issue链接 | 备注 |')
    [void]$sb.AppendLine('|---|---|---|---|---|---|---|---|')
    foreach ($r in $Rows) {
        $link = "[#$($r.'Issue编号')]($($r.'Issue链接'))"
        $labels = ($r.'状态标签' -replace '\|', '\|')
        $remark = ($r.'备注' -replace '\|', '\|')
        [void]$sb.AppendLine("| $($r.'Issue编号') | $($r.'Issue状态') | $($r.'是否应该关闭') | $labels | $($r.'是否已回复') | $($r.'是否已提供具体解决方案') | $link | $remark |")
    }
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
}

function Export-ExcelFile {
    param($Rows, [string]$Path)

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Warning '未安装 ImportExcel 模块，跳过 Excel 输出。安装命令: Install-Module ImportExcel -Scope CurrentUser'
        return
    }
    Import-Module ImportExcel -ErrorAction Stop
    $Rows | Export-Excel -Path $Path -WorksheetName 'Issues' -AutoSize -FreezeTopRow -BoldTopRow -TableName 'DocIssues' -TableStyle Medium2
}

# ---------- 主流程 ----------
Write-Host "仓库: $Repo | label: $Label | 输出目录: $OutputDir"

# 读取历史表格（CSV 作为增量数据源）
$history = @{}
if (Test-Path $CsvPath) {
    try {
        foreach ($row in (Import-Csv -Path $CsvPath -Encoding UTF8)) {
            if ($row.'Issue编号') { $history[[string]$row.'Issue编号'] = $row }
        }
        Write-Host "历史记录: $($history.Count) 条"
    } catch {
        Write-Warning "读取历史 CSV 失败，将全量重建: $($_.Exception.Message)"
    }
}

$openList = Get-OpenIssueList
Write-Host "当前 open issue 数量: $($openList.Count)"

$openNumbers = @{}
foreach ($item in $openList) { $openNumbers[[string]$item.number] = $true }

$rows = New-Object System.Collections.Generic.List[object]
$newCount = 0
$reusedCount = 0
$closedCount = 0

# 处理当前 open 的 issue
$i = 0
foreach ($item in $openList) {
    $i++
    $n = [string]$item.number

    if ($history.ContainsKey($n) -and -not $RefreshAll) {
        # 历史 issue：仅刷新状态列，其他字段保持历史值
        $old = $history[$n]
        $old.'Issue状态' = 'open'
        $rows.Add($old)
        $reusedCount++
        continue
    }

    # 新 issue 或强制刷新：全量抓取
    try {
        $detail = Get-IssueDetail -Number $item.number
        if ($detail) {
            $remark = if ($history.ContainsKey($n)) { $history[$n].'备注' } else { '' }
            $rows.Add((Get-IssueRow -Issue $detail -Remark $remark))
            $newCount++
        } else {
            Write-Warning "issue #$n 不是 issue 或抓取失败，已跳过"
        }
    } catch {
        Write-Warning "issue #$n 抓取失败: $($_.Exception.Message)"
    }
    Start-Sleep -Milliseconds $DelayMs

    if ($i % 20 -eq 0) { Write-Host "open 处理进度 $i/$($openList.Count)（新增/刷新 $newCount，复用 $reusedCount）" }
}

# 处理历史中存在但当前不在 open 列表的 issue（已关闭）
foreach ($n in @($history.Keys)) {
    if ($openNumbers.ContainsKey($n)) { continue }
    if ($RefreshAll) { continue }  # RefreshAll 时只保留当前 open 的

    $old = $history[$n]
    try {
        $issue = Invoke-GitHubApi -Url "https://api.github.com/repos/$Repo/issues/$n"
        $old.'Issue状态' = $issue.state.ToLower()
    } catch {
        $old.'Issue状态' = 'closed'
    }
    $rows.Add($old)
    $closedCount++
    Start-Sleep -Milliseconds $DelayMs
}

$sortedRows = @($rows | Sort-Object { [int]$_.'Issue编号' })

Export-CsvFile -Rows $sortedRows -Path $CsvPath
Export-MarkdownFile -Rows $sortedRows -Path $MdPath
Export-ExcelFile -Rows $sortedRows -Path $XlsxPath

Write-Host ""
Write-Host "完成：总计 $($sortedRows.Count) 条"
Write-Host "  新增/全量刷新: $newCount | 复用历史(仅状态): $reusedCount | 已关闭: $closedCount"
Write-Host ("是否应该关闭=是: {0} / 否: {1}" -f @($sortedRows | Where-Object { $_.'是否应该关闭' -eq '是' }).Count, @($sortedRows | Where-Object { $_.'是否应该关闭' -eq '否' }).Count)
Write-Host ("未回复: {0} / 未提供解决方案: {1}" -f @($sortedRows | Where-Object { $_.'是否已回复' -eq '否' }).Count, @($sortedRows | Where-Object { $_.'是否已提供具体解决方案' -eq '否' }).Count)
Write-Host "CSV : $CsvPath"
Write-Host "MD  : $MdPath"
if (Test-Path $XlsxPath) { Write-Host "XLSX: $XlsxPath" }
