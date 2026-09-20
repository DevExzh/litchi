# Setup-only failed launch

The first attempted full-run command used the incorrect relative freeze path:
`docs/report/spec-gap-validation-evidence/ods-formula-lookups/ods-formula-lookups/gates/freeze.json`.
It exited before candidate verification, build, preflight, or timing with `FileNotFoundError` for that path.
No measurements were produced by this attempt. The authorized run was relaunched from the repository root with the corrected path and retained under `baseline-635fd2e1348b621426b50909cbd5765c91837306/` and `candidate-final/`.
