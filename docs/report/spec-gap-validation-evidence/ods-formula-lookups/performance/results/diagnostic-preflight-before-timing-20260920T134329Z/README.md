# Rejected profile preflight

The first authorized attempt exited before candidate preflight or timing because
`harness/Cargo.lock` still named the previous metadata harness package. The
retained build log records the locked-build refusal; no timing samples exist.
The repair changes only that package name to match the lookup harness manifest;
all third-party versions and checksums remain unchanged. The failed harness
lock is available in Git at commit 364a188246.

An isolated full metadata probe also found an uncached optional `cc` package.
That probe requests dependencies beyond the active build target; the retry's
locked offline build is the authoritative active-target check. The root and
frozen gate lockfiles were not modified.
