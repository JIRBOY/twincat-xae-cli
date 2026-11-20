<#
xae-full.ps1 — TwinCAT XAE 全功能自动化（Windows PowerShell 5.1）

设计：默认只读；任何改变 XAE/目标/PLC 的动作必须显式 -Arm。
高层：info/projects/items/tree/tree-item/read/errors/license/dte-get/dte-call
# 部署：委托到 src/xae_cli.ps1 的 status/fingerprint/compile/boot/config/activate/run/wait/verify/all
写：write 必须 -Arm，并写审计日志。
#>
[CmdletBinding()]
param(
  [Parameter(Position=0)]
  [ValidateSet('info','capabilities','projects','items','tree','tree-item','read','write','errors','license','dte-get','dte-call','plc-props','solution-open','solution-close','solution-save','status','fingerprint','compile','boot','config','activate','run','wait','verify','all')]
  [string]$Command='info',
  [string]$Solution='',
  [string]$NetId='',
  [string]$Symbol='',
  [string]$Value='',
  [ValidateSet('bool','int','double','string')]
  [string]$Type='double',
  [string]$Path='TISC',
  [int]$Depth=3,
  [int]$Limit=200,
  [string]$Member='',
  [string]$ArgsJson='[]',
  [switch]$Arm,
  [switch]$Json,
  [int]$Timeout=180,
  [string]$LogPath=''
)
$ErrorActionPreference='Stop'
trap { Write-Host ("XAE-FULL-ERROR: " + $_.Exception.Message); exit 1 }
if($PSVersionTable.PSEdition -eq 'Core'){ throw 'xae-full.ps1 must run under Windows PowerShell 5.1 (use xae-full.cmd)' }
$script:Rows=New-Object System.Collections.ArrayList
$script:Audit=@()
function Add-Audit([string]$action,[string]$detail){ $script:Audit += [ordered]@{time=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); action=$action; detail=$detail; arm=[bool]$Arm} }
function Out-Result($obj,[string]$text){
  if($LogPath){ $obj | ConvertTo-Json -Depth 10 | Set-Content -Path $LogPath -Encoding utf8 }
  if($Json){ $obj | ConvertTo-Json -Depth 10; return }
  if($text){ Write-Output $text; return }
  $obj | ConvertTo-Json -Depth 10
}
function Get-Dte {
  try { $dte=[Runtime.InteropServices.Marshal]::GetActiveObject('TcXaeShell.DTE.17.0') } catch { throw "XAE DTE not available: $($_.Exception.Message)" }
  if($null -eq $dte){ throw 'XAE DTE is null' }
  try { if(-not $dte.Solution.IsOpen){ throw 'XAE has no open solution' } } catch { throw "XAE solution not usable: $($_.Exception.Message)" }
  return $dte
}
function Get-DteAny {
  try { $dte=[Runtime.InteropServices.Marshal]::GetActiveObject('TcXaeShell.DTE.17.0'); if($dte){ return $dte } } catch {}
  return New-Object -ComObject TcXaeShell.DTE.17.0
}
function Get-SolutionProject($dte){
  if($dte.Solution.Projects.Count -lt 1){ throw 'solution has no project' }
  return $dte.Solution.Projects.Item(1)
}
function Get-InteropAsm {
  $interop=Join-Path $env:TEMP 'interop-TCatSysManager.dll'
  $tlb='C:\Program Files (x86)\Beckhoff\TwinCAT\3.1\Components\Base\TypeLib\3.1\TCatSysManager.tlb'
  if(-not (Test-Path $interop)){
    $tlbimp='C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\TlbImp.exe'
    if(-not (Test-Path $tlbimp)){ throw 'TlbImp.exe not found' }
    & $tlbimp $tlb /out:$interop /silent | Out-Null
  }
  return [Reflection.Assembly]::LoadFrom($interop)
}
function Get-SysManager($dte){
  $obj=$null
  for($i=0;$i -lt 20;$i++){ try { $obj=(Get-SolutionProject $dte).Object; if($obj){break} } catch {} ; Start-Sleep -Milliseconds 500 }
  if($null -eq $obj){ throw 'System Manager object not available' }
  $asm=Get-InteropAsm
  $t=$asm.GetType('interop-TCatSysManager.ITcSysManager')
  $punk=[Runtime.InteropServices.Marshal]::GetComInterfaceForObject($obj,$t)
  return [Runtime.InteropServices.Marshal]::GetTypedObjectForIUnknown($punk,$t)
}
function Get-PlcProject($dte){
  $sm=Get-SysManager $dte
  $tpc=$sm.LookupTreeItem('TIPC'); $start=$tpc.Child(1)
  $asm=Get-InteropAsm
  $t=$asm.GetType('interop-TCatSysManager.ITcPlcProject')
  $punk=[Runtime.InteropServices.Marshal]::GetComInterfaceForObject($start,$t)
  $pp=[Runtime.InteropServices.Marshal]::GetTypedObjectForIUnknown($punk,$t)
  return ,$pp
}
function Get-DtePath($root,[string]$expr){
  $cur=$root
  foreach($seg in ($expr -split '\.')){
    if($seg -match '^(?<n>[A-Za-z_]\w*)\((?<a>\d+)\)$'){
      $n=$matches['n']; $i=[int]$matches['a']; $cur=$cur.$n($i)
    } elseif($seg -match '^(?<n>[A-Za-z_]\w*)$'){
      $cur=$cur.($matches['n'])
    } else { throw "unsupported DTE path segment: $seg" }
  }
  return $cur
}
function Read-Ads([string]$sym,[string]$t){
  Add-Type -Path 'C:\Program Files (x86)\Beckhoff\TwinCAT\3.1\Components\Base\v150\TwinCAT.Ads.dll'
  $c=New-Object TwinCAT.Ads.TcAdsClient
  try {
    $c.Connect($NetId,851)
    $mt=[type]@{bool=[bool];int=[int32];double=[double];string=[string]}[$t]
    $val=$c.ReadSymbolByName($sym,$mt,$true)
    return [ordered]@{ok=$true;symbol=$sym;type=$t;value=$val;netId=$NetId;port=851}
  } finally { try{$c.Dispose()}catch{} }
}
function Write-Ads([string]$sym,[string]$v,[string]$t){
  if(-not $Arm){ throw 'write requires -Arm' }
  Add-Type -Path 'C:\Program Files (x86)\Beckhoff\TwinCAT\3.1\Components\Base\v150\TwinCAT.Ads.dll'
  $c=New-Object TwinCAT.Ads.TcAdsClient
  try {
    $c.Connect($NetId,851)
    $info=$c.ReadSymbolInfo($sym)
    $val=switch($t){ 'bool'{[bool]::Parse($v)}; 'int'{[int32]$v}; 'double'{[double]$v}; 'string'{$v} }
    $c.WriteSymbol($info,$val)
    $back=$c.ReadSymbolByName($sym,[type]@{bool=[bool];int=[int32];double=[double];string=[string]}[$t],$true)
    return [ordered]@{ok=$true;symbol=$sym;written=$val;readback=$back;type=$t;netId=$NetId}
  } finally { try{$c.Dispose()}catch{} }
}
function Get-Tree($item,[int]$d,[int]$max,[System.Collections.ArrayList]$acc){
  if($acc.Count -ge $max){ return }
  $node=[ordered]@{name='';type='';subType='';children=@()}
  try{$node.name=[string]$item.Name}catch{}
  try{$node.type=[string]$item.ItemType}catch{}
  try{$node.subType=[string]$item.ItemSubType}catch{}
  [void]$acc.Add($node)
  if($d -le 0){ return }
  $n=0; try{$n=[int]$item.ChildCount}catch{return}
  for($i=1;$i -le $n;$i++){
    if($acc.Count -ge $max){break}
    try{ $ch=$item.Child($i); $null=Get-Tree $ch ($d-1) $max $acc }catch{}
  }
}
function Get-License {
  $rows=@()
  try {
    $dir=(Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\Beckhoff\TwinCAT3\3.1' -Name LicenseDir -ErrorAction Stop).LicenseDir
    if($dir -and (Test-Path $dir)){ $files=Get-ChildItem $dir -File | ForEach-Object { [ordered]@{name=$_.Name;bytes=$_.Length;modified=$_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')} }; return [ordered]@{ok=$true;licenseDir=$dir;files=$files;note='file list only; validity/activation needs TwinCAT license UI'} }
  } catch {}
  foreach($p in @('HKLM:\SOFTWARE\WOW6432Node\Beckhoff\TwinCAT3\License','HKLM:\SOFTWARE\Beckhoff\TwinCAT3\License')){
    if(Test-Path $p){ $rows += [ordered]@{path=$p;values=(Get-ItemProperty $p | Select-Object * -ExcludeProperty PS*)} }
  }
  if($rows.Count -eq 0){ return [ordered]@{ok=$false;reason='license registry key not found'} }
  return [ordered]@{ok=$true;keys=$rows}
}
function Get-Capabilities {
  [ordered]@{
    ok=$true
    readOnly=@('info','capabilities','projects','items','tree','tree-item','read','errors','license','dte-get','plc-props')
    writesRequireArm=@('write','dte-call (with args)','solution-open','solution-close','solution-save','compile','boot','config','activate','run','all')
    delegatedToXaeCli=@('status','fingerprint','compile','boot','config','activate','run','wait','verify','all')
    genericBridge=@('dte-get (read path)','dte-call (member/args, -Arm for methods)')
    notAutomatedHere=@('license wizard dialogs','EtherCAT IO scan wizard','ScopeView UI recording','XAE GUI-only dialogs')
    note='COM automation covers DTE/SysManager/ADS; GUI-only wizards need UI automation/manual.'
  }
}

$tool=$env:XAE_CLI_TOOL
if(-not $tool){
  $tool=Join-Path $PSScriptRoot '..\src\xae_cli.ps1'
}
$delegate=@('status','fingerprint','compile','boot','config','activate','run','wait','verify','all')
if($delegate -contains $Command){
  if(-not (Test-Path $tool)){ throw "xae_cli.ps1 not found: $tool" }
  if($Command -notin @('status','fingerprint','wait') -and -not $Arm){ throw "$Command requires -Arm" }
  $p=@{ Command=$Command }
  if($Solution){$p.Solution=$Solution}
  if($NetId){$p.NetId=$NetId}
  if($Timeout){$p.Timeout=$Timeout}
  & $tool @p
  exit $LASTEXITCODE
}

$result=$null
switch($Command){
  'plc-props' {
    $dte=Get-Dte; $pp=Get-PlcProject $dte
    $result=[ordered]@{ok=$true;tmcFileCopy=$pp.TmcFileCopy;bootProjectEncryption=$pp.BootProjectEncryption;bootProjectAutostart=$pp.BootProjectAutostart}
  }
  'solution-open' {
    if(-not $Arm){ throw 'solution-open requires -Arm' }; if(-not $Solution){ throw '-Solution required' }
    $dte=Get-DteAny; $dte.Solution.Open($Solution); for($i=0;$i -lt 30;$i++){ if($dte.Solution.IsOpen -and $dte.Solution.FullName){break}; Start-Sleep -Seconds 1 }
    $result=[ordered]@{ok=$true;solution=[string]$dte.Solution.FullName;isOpen=[bool]$dte.Solution.IsOpen}
  }
  'solution-close' {
    if(-not $Arm){ throw 'solution-close requires -Arm' }; $dte=Get-Dte; $dte.Solution.Close($false)
    $result=[ordered]@{ok=$true;closed=$true}
  }
  'solution-save' {
    if(-not $Arm){ throw 'solution-save requires -Arm' }; $dte=Get-Dte; $dte.Solution.Save()
    $result=[ordered]@{ok=$true;saved=$true;solution=[string]$dte.Solution.FullName}
  }
  'info' {
    $dte=Get-Dte; $p=Get-SolutionProject $dte
    $result=[ordered]@{ok=$true;dteVersion=[string]$dte.Version;solution=[string]$dte.Solution.FullName;projects=@();mainWindow=[string]$dte.MainWindow.Caption}
    for($i=1;$i -le $dte.Solution.Projects.Count;$i++){ $pr=$dte.Solution.Projects.Item($i); $result.projects += [ordered]@{name=[string]$pr.Name;fullName=[string]$pr.FullName;kind=[string]$pr.Kind} }
  }
  'capabilities' { $result=Get-Capabilities }
  'projects' {
    $dte=Get-Dte; $result=[ordered]@{ok=$true;projects=@()}
    for($i=1;$i -le $dte.Solution.Projects.Count;$i++){ $pr=$dte.Solution.Projects.Item($i); $result.projects += [ordered]@{index=$i;name=[string]$pr.Name;fullName=[string]$pr.FullName;kind=[string]$pr.Kind} }
  }
  'items' {
    if(-not $Arm -and $false){ }
    $dte=Get-Dte; $pr=Get-SolutionProject $dte
    $rows=New-Object System.Collections.ArrayList
    function Walk($items,[int]$d,[string]$prefix){
      if($d -lt 0 -or $rows.Count -ge $Limit){return}
      for($i=1;$i -le $items.Count;$i++){
        if($rows.Count -ge $Limit){break}
        $it=$items.Item($i)
        $nm=''; try{$nm=[string]$it.Name}catch{}
        [void]$rows.Add([ordered]@{depth=$d;name=$nm;kind='';fullName=''})
        $rows[$rows.Count-1].kind=''; try{$rows[$rows.Count-1].kind=[string]$it.Kind}catch{}
        try{$rows[$rows.Count-1].fullName=[string]$it.FileNames(1)}catch{}
        try{ Walk $it.ProjectItems ($d+1) ($prefix+$nm+'/') }catch{}
      }
    }
    Walk $pr.ProjectItems 0 ''
    $result=[ordered]@{ok=$true;count=$rows.Count;items=$rows;truncated=($rows.Count -ge $Limit);note=''}
  }
  'tree' {
    $dte=Get-Dte; $sm=Get-SysManager $dte
    $item=$sm.LookupTreeItem($Path)
    $acc=New-Object System.Collections.ArrayList
    Get-Tree $item $Depth $Limit $acc
    $result=[ordered]@{ok=$true;path=$Path;depth=$Depth;limit=$Limit;nodes=$acc;truncated=($acc.Count -ge $Limit)}
  }
  'tree-item' {
    $dte=Get-Dte; $sm=Get-SysManager $dte
    $item=$sm.LookupTreeItem($Path)
    $result=[ordered]@{ok=$true;path=$Path;name=[string]$item.Name;itemType=[string]$item.ItemType;itemSubType=[string]$item.ItemSubType;childCount=[int]$item.ChildCount}
  }
  'read' {
    if(-not $Symbol){ throw '-Symbol required' }
    $result=Read-Ads $Symbol $Type
  }
  'write' {
    if(-not $Symbol){ throw '-Symbol required' }
    $result=Write-Ads $Symbol $Value $Type
    Add-Audit 'write' ($Symbol+'='+$Value)
  }
  'errors' {
    $dte=Get-Dte; $rows=@()
    try{ $el=$dte.ToolWindows.ErrorList; for($i=1;$i -le $el.ErrorItems.Count;$i++){ $it=$el.ErrorItems.Item($i); $rows += [ordered]@{description=[string]$it.Description;file=[string]$it.FileName;line=[int]$it.Line;level=[string]$it.ErrorLevel} } }catch{}
    $result=[ordered]@{ok=$true;count=$rows.Count;errors=$rows}
  }
  'license' { $result=Get-License }
  'dte-get' {
    if(-not $Member){ throw '-Member or -Path required (use -Member)' }
    $dte=Get-Dte
    $val=Get-DtePath $dte $Member
    $result=[ordered]@{ok=$true;path=$Member;value=[string]$val;type=($val.GetType().FullName)}
  }
  'dte-call' {
    if(-not $Member){ throw '-Member required' }
    if($ArgsJson -ne '[]' -and -not $Arm){ throw 'dte-call with args requires -Arm' }
    $dte=Get-Dte
    $args=@(); if($ArgsJson -and $ArgsJson -ne '[]'){ $args=@($ArgsJson | ConvertFrom-Json) }
    $flags=[Reflection.BindingFlags]::InvokeMethod
    if($args.Count -eq 0){ $flags=[Reflection.BindingFlags]::GetProperty }
    $val=$dte.GetType().InvokeMember($Member,$flags,$null,$dte,$args)
    $result=[ordered]@{ok=$true;member=$Member;args=$args;value=[string]$val;type=($val.GetType().FullName)}
    Add-Audit 'dte-call' ($Member+' '+$ArgsJson)
  }
  default { throw "unsupported command: $Command" }
}
if($script:Audit.Count -gt 0 -and $Arm){
  $log=Join-Path $env:TEMP 'xae-full-audit.log'
  foreach($a in $script:Audit){ ([ordered]@{when=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss');action=$a.action;detail=$a.detail} | ConvertTo-Json -Compress) | Add-Content $log -Encoding utf8 }
}
Out-Result $result $null
