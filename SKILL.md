---
name: xae-cli
description: TwinCAT XAE automation via DTE COM, SysManager, ADS and a generic COM bridge. Read-only by default; writes require -Arm. Use for TwinCAT XAE CLI, compile/activate, solution/project tree, ADS read/write and XAE troubleshooting.
---

# xae-cli — TwinCAT XAE automation CLI

`xae-cli` wraps the TwinCAT XAE automation surface into commands for:

- opening/attaching to a solution;
- compiling and generating a boot project;
- switching the target to Config, activating, restarting RUN and verifying;
- enumerating projects and the SysManager tree;
- ADS read/write (writes require `-Arm`);
- generic DTE property/method calls.

It is a **reference implementation**, not a safety-certified tool. It does not replace TwinCAT safety wiring, commissioning approval, or production HIL.

## Requirements

- Windows 10/11 or Windows Server with the interactive desktop;
- TwinCAT 3.1 + TcXaeShell installed;
- **Windows PowerShell 5.1** (PS7 lacks `Marshal.GetActiveObject`);
- the target AMS Net ID reachable, and authorization to touch it.

## Layout

```text
scripts/
  xae-cli.cmd / xae-cli.ps1    # deployment chain wrapper
  xae-full.cmd / xae-full.ps1  # full COM/ADS surface
src/
  xae_cli.ps1                  # deployment-chain implementation
```

All `.ps1` files are UTF-8 with BOM so Windows PowerShell 5.1 reads Chinese comments correctly.

## Quick start

```powershell
# Read-only
.\scripts\xae-full.cmd info
.\scripts\xae-full.cmd projects
.\scripts\xae-full.cmd tree -Path TISC -Depth 2 -Limit 100
.\scripts\xae-full.cmd read -Symbol GVL_System.SysState -Type int
.\scripts\xae-full.cmd errors
.\scripts\xae-full.cmd license
.\scripts\xae-full.cmd plc-props

# Deployment chain (status/verify are read-only; other commands require -Arm)
.\scripts\xae-cli.cmd status -Solution "C:\path\TC3-StartHere.sln" -NoLaunch
.\scripts\xae-cli.cmd compile -Solution "C:\path\TC3-StartHere.sln" -Arm
.\scripts\xae-cli.cmd all -Solution "C:\path\TC3-StartHere.sln" -NetId 127.0.0.1.1.1 -Arm
```

`-NetId` may be omitted only when the `.tsproj` contains `TargetNetId`; otherwise the CLI exits 2 instead of guessing a target.

## Command map

| Area | Commands |
|---|---|
| DTE / solution | `info`, `solution-open`, `solution-close`, `solution-save` |
| Projects / tree | `projects`, `items`, `tree`, `tree-item` |
| PLC project | `plc-props` |
| Build / deploy | `compile`, `boot`, `config`, `activate`, `run`, `wait`, `verify`, `all` |
| Diagnostics | `errors`, `license`, `status`, `fingerprint` |
| ADS | `read`, `write` (write requires `-Arm`) |
| Generic COM | `dte-get`, `dte-call` |

## Safety model

1. **Read-only by default.** `write`, `solution-*`, `dte-call` with arguments, and deploy commands require `-Arm`.
2. **One XAE per target.** Serial guard refuses concurrent deployment (exit 13).
3. **Activation is not HIL.** A successful activate/run/verify proves the PLC is running and symbols are readable; it does not prove the machine is safe or accepted.
4. **Do not run production/HIL** from an unattended agent. Use a human-approved window and the machine's safety chain.
5. **Keep secrets out of logs.** Do not pass tokens/passwords as command-line arguments; prefer `--stdin`-style facilities where available.

## Project-specific verification

The deployment chain's `verify` step uses example symbol names:

```text
GVL_Diag.nSelfTestTick
GVL_Diag.bSelfTestDone
GVL_IoMap.N_IO_AXES
GVL_Heartbeat.ControlTaskHeartbeat
GVL_Command.eCommand
```

These are **examples**, not a universal contract. Adapt `Invoke-Verify` in `src/xae_cli.ps1` to your PLC project, or use `-ExpectedAxes N` to make the axis-count check explicit. If your project does not expose `GVL_Diag`, the program-identity gate will fail closed; that is intentional.

## Exit codes

`xae-cli`: 0 success; 2 argument/file; 3 DTE unavailable; 4 solution/project not ready; 5 compile failed; 6 boot failed; 7 Config failed; 8 activate failed; 9 license/modal dialog; 10 :10000 timeout; 11 :851 timeout/error 6; 12 verify failed; 13 serial guard busy.

`xae-full`: exit 1 on error; writes/changes without `-Arm` are rejected before execution.

## Not automated

- license wizard/activation dialogs;
- EtherCAT IO scan wizard;
- ScopeView UI recording;
- GUI-only popups. Use human interaction or a separate UI automation tool.

## License

MIT. See `LICENSE`.

Inspired by public TwinCAT automation interfaces; no Beckhoff source code is included.
