# 0362k — ODraw bounded raw APIs and image scanners

- Date: 2026-10-21.
- Base: `61515148e1`.
- Status: source and consumer migration complete; performance qualification is
  pending coordinator scheduling.

ODraw raw entry points now require the caller's policy. `Record::parse`,
`Parser::new`, `Children::new`, `Container::find_recursive`, shape and
property parsing, rule parsing, gradient decoding, and all image record
parsers take explicit `Limits`. The removed `Rule` conversion cannot silently
select a default policy. Retained records and owned image files keep the
policy that admitted their bytes, including when an owned BLIP is reparsed.

The raw writer and builder functions use the same required policy. They check
record and child limits, checked header-plus-payload extents, and fallible
output reservation before emitting a header or partial record. Unknown
OfficeArt records and image payloads remain borrowed or copied byte-for-byte;
the migration does not add a canonicalization pass or drop extensions.

Image scanning now has separate typed ceilings for raw record bytes, BLIP
bytes, store entries, files, depth and visited records. Top-level store scans
and recursive scans enforce the record ceiling and report the observed count;
file collection checks the cumulative file count when it merges nested
stores. Image errors retain the observed value and configured maximum, and
the scanner reports `Files`, `Depth` and `Records` distinctly. Overflow paths
also report the attempted count rather than the pre-increment value.

All DOC, PPT, XLS and image-conversion consumers, tests and examples migrate
in this slice. The existing `odraw` glossary profile covers the removed
default-policy names and immutable projections; no compatibility aliases or
deprecated entry points were added. The picture owner remains on its existing
owner migration track and is not treated as an exception here.

Validation completed for the final source and consumer set:

* `cargo test --offline -p litchi-odraw --all-targets --all-features`: 99
  passed;
* `cargo test --offline -p litchi-ppt writer::table::tests -- --nocapture`:
  12 passed;
* `cargo test --offline -p litchi-doc -p litchi-ppt -p litchi-xls -p
  litchi-imgconv --all-targets --all-features`: all affected host targets
  passed (DOC 1,249 tests and XLS 1,314 tests in the final run; one existing
  fixture remains ignored);
* `cargo fmt --all -- --check` and `git diff --check` passed.

These checks used the capped one-job environment and a target below 20 GB.
No native timing or profiling result is claimed for this source migration;
the coordinator must rebuild matched binaries and qualify performance after
integration. Generated binaries and raw evidence remain outside the checkout.
