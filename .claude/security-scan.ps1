# Security scan before git push
# Called from PreToolUse hook (stdin: JSON)

$raw = [Console]::In.ReadToEnd()
$input_json = $raw | ConvertFrom-Json -ErrorAction SilentlyContinue
$command = $input_json.tool_input.command

# Skip if not git push
if ($command -notmatch 'git push') {
    exit 0
}

# Get diff of commits to be pushed
$diff = git diff origin/main..HEAD -- . 2>$null
if (-not $diff) {
    $diff = git diff HEAD~1..HEAD -- . 2>$null
}
if (-not $diff) {
    exit 0
}

# Sensitive patterns
$patterns = @(
    @{ Name = "API Key";             Regex = '(?i)(api[_-]?key|apikey)\s*[=:]\s*["'']?[A-Za-z0-9\-_]{20,}' },
    @{ Name = "Google Sheet ID";     Regex = '\|\s*[0-9A-Za-z\-_]{44}\s*\|' },
    @{ Name = "Password";            Regex = '(?i)(password|passwd|pwd)\s*[=:]\s*["'']?.{6,}' },
    @{ Name = "Token/Secret";        Regex = '(?i)(token|secret|bearer)\s*[=:]\s*["'']?[A-Za-z0-9\-_.]{20,}' },
    @{ Name = "Private Key Header";  Regex = '-----BEGIN (RSA |EC |OPENSSH )?PRIVATE KEY-----' },
    @{ Name = "Anthropic API Key";   Regex = 'sk-ant-[A-Za-z0-9\-_]{20,}' },
    @{ Name = "OpenAI API Key";      Regex = 'sk-[A-Za-z0-9]{40,}' },
    @{ Name = "GCP Credential";      Regex = '"private_key"\s*:\s*"-----BEGIN' },
    @{ Name = "AWS Access Key";      Regex = 'AKIA[0-9A-Z]{16}' },
    @{ Name = "JWT Token";           Regex = 'eyJ[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]+' },
    @{ Name = "Hardcoded SpreadsheetId"; Regex = "'\|[0-9A-Za-z\-_]{40,}\|'" }
)

$findings = @()

foreach ($p in $patterns) {
    $lines = $diff -split "`n" | Where-Object { $_ -match '^\+' -and $_ -notmatch '^\+\+\+' }
    foreach ($line in $lines) {
        if ($line -match $p.Regex) {
            $preview = $line.Substring(0, [Math]::Min(100, $line.Length))
            $findings += "  [$($p.Name)] $preview"
        }
    }
}

if ($findings.Count -gt 0) {
    $findingsText = $findings -join "\n"
    $msg = "SECURITY: git push blocked. Sensitive data detected:\n$findingsText\n\nPlease remove secrets and commit again before pushing."
    $output = "{`"decision`":`"block`",`"reason`":`"Sensitive information detected in diff`",`"systemMessage`":`"$msg`"}"
    Write-Output $output
    exit 2
}

exit 0
