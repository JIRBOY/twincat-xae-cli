# Security policy

## Scope

This project drives TwinCAT XAE and ADS on a Windows host. A mistake can compile, activate, restart or write to a real controller.

## Reporting

Open a private security advisory on GitHub, or contact the maintainer through the repository. Do not publish exploit details before a fix is available.

## Safety expectations

- Use only on targets you are authorized to operate.
- Read `README.md` and `SKILL.md` before using write/deploy commands.
- `-Arm` is an explicit acknowledgement, not a safety interlock.
- Activation/run is not HIL and not a production release.
- Keep the machine's hardwired emergency-stop and safety chain independent of this tool.
- Do not put credentials in command lines, logs or issue reports.

## Secrets

No credentials are stored in this repository. If you find a committed secret, report it privately; do not open a public issue with the value.
