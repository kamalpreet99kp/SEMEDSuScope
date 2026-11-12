# Authoritative AuAgTe handoff — 2026-11-12

The user explicitly designated this uploaded ZIP as the current authority for
all 16 supplied workflow scripts. Active repository paths contain exact supplied
bytes; no processing or mineralogical logic was edited during synchronization.

## Recovery files

- `AuAgTe-Code-Handoff.zip`: original upload, SHA-256
  `1d796ad74ceefc828121f4658481e84e5daf12744ad7a67de98848e2ad9e59b5`.
- `SHA256-MANIFEST.txt`: exact supplied manifest, also copied to repository root.
- `Summary_Each_Sample.pre-handoff.py`: separately preserved working version
  from before this synchronization. It contains the summary formatting changes
  tested in this conversation. It is not the active authoritative ZIP version.

The ZIP's Summary_Each_Sample.py is 399 lines and lacks those later formatting
changes. The user's latest instruction makes the ZIP authoritative; do not
silently reactivate or merge the preserved version.

## Verification

- All 16 source-file hashes match the supplied manifest.
- All 16 Python source files passed syntax compilation without executing GUIs.
- Every one of the app's seven fixed script references is supplied.
- All eight A2 dropdown scripts are supplied.
- The app is 1,167 lines; SHA-256:
  `7f643a5a1973494ff7e3c7fffdff39934f60d9d818209b3fe7361cc82d4ba786`.
- Updated-AllAgMinerals-41-19.py is 419 lines; SHA-256:
  `256d65756f308da710119eca203d35d542aa87ef1ff0de3f94c449f8d13527b6`.

Source availability and syntax were checked. This does not assert that all GUI
workflows were executed on Windows, or that optional assets such as the AMTEL
logo and third-party packages are included in this source-only ZIP.

## Durable recovery

GitHub repository: https://github.com/kamalpreet99kp/SEMEDSuScope

Authoritative recovery branch: `codex/auagte-authoritative-2026-11-12`.
Recover files from that branch if an environment is initialized from older main.
A temporary workspace can reset; future environment branch selection is not
controlled by these files. Confirm the selected revision before editing.

Keep this dated archive and manifest unchanged as provenance. Future authorized
script changes can legitimately differ from the handoff hashes. Compare and
investigate before restoring; never silently discard newer authorized work.

