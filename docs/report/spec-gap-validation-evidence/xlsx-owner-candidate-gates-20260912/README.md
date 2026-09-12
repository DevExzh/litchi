# XLSX owner candidate checkpoint — approval pending

This records a tested intermediate source revision, not production approval.
The base is `913eec261e52618599f85ab8282a7012429eb56f`, which includes the
reviewed OPC encrypted-admission and bounded relationship-handle prerequisite.

Root gates passed against unchanged hashes for all fourteen captured files:

- 1,153 XLSX library tests.
- 34 cached-name, 48 server-format, 34 pivot-data, and 22 form-control owner
  integration tests (138 total).
- Strict Clippy for the library and those four integration targets.
- Rustdoc with warnings denied and formatting for the captured files.

The pivot unit tests include a returned-patch regression that compares its
private before/after workbook snapshots with the actual incoming and published
workbooks. The incoming workbook contains an unrelated part absent from the
original patch source. The inverse must restore that incoming workbook. This
checks source binding beyond merely replaying a patch on identical bytes.

## Remaining approval requirements

The pivot design's **Limits, inverse, and source preservation** section requires
separate authored-row, authored-cell, and retained-byte budgets, independently
of XML/event bounds and logical dimensions. These requirements remain under
resource review at this checkpoint. Green tests do not establish them. The
review also covers duplicate retained strings, temporary allocations, and
multirow traversal costs. Do not substitute an implicit finite input limit for
an explicitly required independent budget.

The form-control owner implementation is independently frozen with 22 passing
tests, including actual VML mirror differences and source relationship-target
retention. Final owner-contract review remains pending. This batch contains no
form-control mutation API and establishes no native editing interoperability.

No timing run, throughput improvement, allocation-count result, or peak-memory
measurement is reported. The complete audit goal remains unfinished.

## Reproduction

`root-gates.json` records the commands, toolchain, source hashes, exit statuses,
source stability, test summaries, and decompressed log hashes. The four gzip
logs retain complete output. `source.tar.gz` contains the fourteen exact source
and test files, with repository-relative paths and a hash in the manifest.

In a disposable checkout of the base commit, overlay that archive, verify every
file against `source_sha256`, and run the recorded commands with `TMPDIR=/var/tmp`
and `RUSTFLAGS` unset. Set `RUSTDOCFLAGS=-D warnings` for the rustdoc command.
Use the base checkout's pinned toolchain and lockfile. The checkpoint does not
require the current dirty workspace, later resource fixes, or unrelated owned
temporary directories. The archive is retained review evidence, not generated
production source to install into a working checkout automatically.
