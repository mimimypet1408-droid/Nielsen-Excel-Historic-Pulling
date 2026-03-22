# Refreshes the PATH from the registry (Machine + User) then executes the
# remaining arguments.  Intended to be called from the Makefile so that
# newly-installed tools (fnm, uv, …) are visible without opening a new shell.
#
# Usage:
#   powershell -ExecutionPolicy bypass -File scripts\refreshpath-exec.ps1 uv sync
#   powershell -ExecutionPolicy bypass -File scripts\refreshpath-exec.ps1 fnm install 22

param(
    [Parameter(Position = 0, ValueFromRemainingArguments)]
    [string[]]$Command
)

if (-not $Command) {
    Write-Error "Usage: refreshpath-exec.ps1 <command> [args...]"
    exit 1
}

$env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + `
            [System.Environment]::GetEnvironmentVariable('Path', 'User')

# If the first token is "fnm" we also need to run `fnm env` so that the
# node/npm shims are on the PATH.
if ($Command[0] -eq 'fnm') {
    fnm env --use-on-cd --shell powershell | Out-String | Invoke-Expression
}

# When npm/npx/node are needed but the command itself isn't fnm, the caller
# can prefix with "fnm-env" as a sentinel to trigger the same setup.
if ($Command[0] -eq 'fnm-env') {
    fnm env --use-on-cd --shell powershell | Out-String | Invoke-Expression
    $Command = $Command[1..$Command.Length]
}

Invoke-Expression ($Command -join ' ')
exit $LASTEXITCODE
