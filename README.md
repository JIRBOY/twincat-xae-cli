# twincat-xae-cli

TwinCAT XAE automation from a command line: DTE COM, SysManager tree, compile/boot/activate, ADS read/write, and a generic COM bridge.

It is designed for engineers and AI agents that need a **deterministic, scriptable** way to inspect or deploy a TwinCAT 3 project. Read-only by default; writes and deployment actions require `-Arm`.

> **Not a safety tool.** Activation/run is not HIL, not commissioning approval and not a production release. Keep the machine's hardwired safety chain and human sign-off.

## Why

TwinCAT XAE is GUI-first. Batch scripts often fail on:

- attaching to the right solution (same names in multiple repos);
- distinguishing "compiled" from "boot project generated" and "activated";
- modal license/error dialogs blocking COM;
- no serial control when two tools touch the same target.

This CLI makes those steps explicit and returns structured results with a stable exit-code contract.

## Requirements

- Windows 10/11 (interactive desktop; Session 0/locked desktop is not supported);
- TwinCAT 3.1 + TcXaeShell (`TcXaeShell.DTE.17.0`);
- **Windows PowerShell 5.1**;
- an authorized target and AMS Net ID.

## Install

```powershell
git clone https://github.com/JIRBOY/twincat-xae-cli.git
cd twincat-xae-cli
.\scripts\xae-full.cmd info
```

No build step. The wrappers call Windows PowerShell 5.1; `src/xae_cli.ps1` needs the TwinCAT automation assemblies already installed by TwinCAT XAE.

If your execution policy blocks scripts:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\xae-full.ps1 info
```

## Usage

```powershell
# Read-only
.\scripts\xae-full.cmd info
.\scripts\xae-full.cmd projects
.\scripts\xae-full.cmd tree -Path TISC -Depth 2 -Limit 100
.\scripts\xae-full.cmd tree-item -Path TIPC
.\scripts\xae-full.cmd read -Symbol GVL_System.SysState -Type int
.\scripts\xae-full.cmd errors
.\scripts\xae-full.cmd license
.\scripts\xae-full.cmd plc-props
.\scripts\xae-full.cmd dte-get -Member Solution.FullName

# Write (requires -Arm)
.\scripts\xae-full.cmd write -Symbol GVL_System.SysState -Value 0 -Type int -Arm

# Deployment chain
.\scripts\xae-cli.cmd status -Solution "C:\path\TC3-StartHere.sln" -NoLaunch
.\scripts\xae-cli.cmd compile -Solution "C:\path\TC3-StartHere.sln" -Arm
.\scripts\xae-cli.cmd boot -Solution "C:\path\TC3-StartHere.sln" -Arm
.\scripts\xae-cli.cmd activate -Solution "C:\path\TC3-StartHere.sln" -NetId 127.0.0.1.1.1 -Arm
.\scripts\xae-cli.cmd wait -NetId 127.0.0.1.1.1
.\scripts\xae-cli.cmd verify -NetId 127.0.0.1.1.1
```

`-NetId` may be omitted only when the `.tsproj` contains `TargetNetId`; otherwise the CLI exits 2 rather than guessing a target.

## Exit codes

| Code | Meaning |
|---:|---|
| 0 | success |
| 2 | argument/file error |
| 3 | XAE/DTE unavailable |
| 4 | solution/project not ready |
| 5 | compile failed |
| 6 | boot project failed |
| 7 | switch to Config failed |
| 8 | activate failed |
| 9 | license/modal dialog blocked |
| 10 | target :10000 timeout |
| 11 | PLC :851 timeout/error 6 |
| 12 | verify failed |
| 13 | serial guard busy |

`xae-full` returns exit 1 on error; writes/changes without `-Arm` are refused before touching XAE.

## Project-specific verification

`verify` ships **example** symbols (`GVL_Diag`, `GVL_Heartbeat`, `GVL_IoMap`, `GVL_Command`). They are not a universal contract. Adapt `Invoke-Verify` in `src/xae_cli.ps1` to your PLC project; use `-ExpectedAxes N` only if your project exposes that axis count. A missing example symbol fails closed.

## Safety model

- read-only default;
- `-Arm` required for writes and deployment;
- serial guard refuses concurrent XAE/deployment activity (exit 13);
- activation/run is not HIL; verify only proves symbols read and heartbeat increments;
- do not run unattended production/HIL; do not bypass the machine's safety chain;
- do not put credentials on command lines or in logs.

## Scope

Automated: DTE solution/project/tree, build, boot, Config/activate/run/wait/verify, ADS read/write, DTE generic calls.

Not automated: license wizard, EtherCAT IO scan wizard, ScopeView recording, GUI-only dialogs.

## Tested with

Windows PowerShell 5.1, TwinCAT 3.1.4026.25, TcXaeShell 17.0. Public CI does not have TwinCAT installed, so the repository's CI is limited to static checks; deployment must be validated on a real, authorized target.

## License

MIT. See [LICENSE](LICENSE).

No Beckhoff source code is included; the tool talks to public TwinCAT automation interfaces.
