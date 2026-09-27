# Execution notes

All six recorded build/preflight rows passed on their first execution. All 992
native/profile children terminated successfully, with no retries or discarded
samples. Both independent raw audits and the full pre-cleanup validator passed.

A preparation shell command initially used unavailable `python`; its Python
steps did not execute. Root reran those preparation steps using `python3`
before the build. Formatting was applied before frozen build inputs were made.

The first post-cleanup replay rejected the intentionally removed probe binary.
The analyzer's custody reader was updated to accept only the exact build-binary
identity recorded in the cleanup witness, with the owned target absent. Raw
captures, frozen inputs, timing policy, and generated numerical results were
unchanged by this replay fix.
