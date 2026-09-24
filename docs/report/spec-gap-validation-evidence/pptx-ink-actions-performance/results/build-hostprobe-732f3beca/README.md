# PPTX InkAction build and host-probe retention

This directory is a durable copy of the committed `732f3beca` release-build
and public `--host-probe` provenance. `raw/preflight/` and
`raw/build-hostprobe/` contain every regular file from their corresponding raw
receipt directories, including Cargo metadata, source manifests, command logs,
empty stdout/stderr files, build output, and the initial fail-closed probe
refusal. `retention-manifest.json` records the source path, retained path,
byte count, and SHA-256 for every copied raw file.

The copy is bound to the clean source checkout and the semantic/production dual
pins recorded in `retention-manifest.json`. The retained binary and its
external Cargo target are referenced by exact path and SHA-256 and remain
outside this copy for the root byte audit.

The successful probe used `PPTX_INK_ACTIONS_CAPTURE_HEAD` equal to the captured
HEAD. The earlier probe invocation is retained too: it failed closed with exit
1 because that receipt environment was absent. That refusal is not a feature
or performance result.

No matrix, lane, `/usr/bin/time`, allocation, RSS, or timing command was run.
This bundle makes no native PowerPoint acceptance claim and no speedup claim;
the executable was built only to prove the scaffold links and the public host
probe emits its typed provenance receipt. The external raw source directories,
target, and binary remain available until root completes a full raw Git byte
audit. A later matrix must use a fresh runner-owned target unless root
explicitly authorizes reuse.

After the audit, cleanup is limited to the external target-parent and raw source
receipt directories named in the manifest. Keep this complete retention copy.
