# xae-cli.ps1 — thin wrapper for TwinCAT XAE DTE CLI
# Must delegate to Windows PowerShell 5.1 (PS7 lacks Marshal.GetActiveObject).
[CmdletBinding()]
param(
  [Parameter(Position=0)]
  [ValidateSet('status','fingerprint','compile','boot','config','activate','run','wait','verify','all')]
  [string]$Command='status',
  [string]$Solution='',
  [string]$NetId='',
  [switch]$Json,
  [int]$Timeout=180,
  [string]$LogPath='',
  [switch]$NoLaunch,
  [string]$ExpectedRepoRoot='',
  [string]$ExpectedProjectName='',
  [switch]$DryRun,
  [string]$ToolPath=''
)
$ErrorActionPreference='Stop'
if(-not $ToolPath){
  $ToolPath=Join-Path $PSScriptRoot '..\src\xae_cli.ps1'
}
if(-not $ToolPath){ throw 'xae_cli.ps1 not found; pass -ToolPath <path>' }
$ps=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$argsList=@('-NoProfile','-ExecutionPolicy','Bypass','-File',$ToolPath,$Command)
if($Solution){$argsList+=@('-Solution',$Solution)}
if($NetId){$argsList+=@('-NetId',$NetId)}
if($Json){$argsList+='-Json'}
if($Timeout){$argsList+=@('-Timeout',[string]$Timeout)}
if($LogPath){$argsList+=@('-LogPath',$LogPath)}
if($NoLaunch){$argsList+='-NoLaunch'}
if($ExpectedRepoRoot){$argsList+=@('-ExpectedRepoRoot',$ExpectedRepoRoot)}
if($ExpectedProjectName){$argsList+=@('-ExpectedProjectName',$ExpectedProjectName)}
if($DryRun){$argsList+='-DryRun'}
& $ps @argsList
exit $LASTEXITCODE
