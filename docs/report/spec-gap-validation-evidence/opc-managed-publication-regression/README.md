# Managed publication memory regression

Commit `83e60e8c62d52d46d75e06f3ddda121692775dc3` avoids reparsing and
reserving an unchanged content-types manifest when a prepared OPC transaction
has no part additions or removals. Candidate content-type queries use the
validated source catalog. Empty relationship edit plans also avoid their
unused fixed metadata reservation.

`diagnosis.txt` is the investigating agent's retrospective diagnostic record,
retained byte-for-byte (SHA-256
`a0b4df1abde4ee9d70392952efe55f52aba0e944912ff557b7c654738650909b`).
It identifies the compared commits, lockfile, commands, observed failures,
and removed temporary paths. It is not raw process output, a sealed benchmark,
or independently replayable performance evidence; no latency or general memory
improvement is claimed.

The committed regression test denies reads of the content-types member during
replacement-only preparation and checks candidate metadata and publication.
Root validation passed the OPC suite (418 unit tests and all integration and
doctest groups, with one existing external-corpus test ignored), 29 prepared
topology tests, strict Clippy, and formatting. Independent review approved the
source-catalog fallback and retained source, graph, signature, and limit checks.
A later root XLSX rerun could not compile because the separate form-control
implementation was still in progress; this record does not claim a green
current full XLSX suite.

Both diagnostic baseline commits remain reachable in Git. The listed owned
worktrees and targets were removed after the diagnostic record was retained;
other workers' artifacts were left untouched.
