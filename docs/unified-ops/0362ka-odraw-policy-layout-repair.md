# 0362ka — ODraw policy-layout repair

- Date: 2026-10-21.
- Base: `61515148e1`.
- Parent source checkpoint: `08f586ca42`.
- Status: source repair and focused validation complete; native timing and the
  post-policy matched release/allocator rebuild remain pending coordinator
  scheduling.

The first bounded-policy candidate retained a complete raw `Limits` value in
every `Record`, duplicated the image policy in `Store`, and kept public
`Children` values on the recursive traversal stack.  The initial measured
regression was:

| route | baseline | first candidate | change |
| --- | ---: | ---: | ---: |
| recursive allocation bytes | 256 | 288 | +32 |
| recursive peak bytes | 2,113,611 | 2,113,643 | +32 |
| image scan allocation bytes | 704 | 808 | +104 |
| image scan peak bytes | 2,110,086 | 2,110,190 | +104 |

The repair packs the record header fields while retaining the full originating
policy, stores only image-specific fields in `Store`, and uses one internal
cursor frame with the effective policy passed through recursive traversal.
The public `Children` policy remains unchanged.  The final compact layouts are
`Record=32`, internal traversal `Frame=32`, `Store=48`, `File=168`, and
`Shape=312` bytes.  The matched preflight for the compact candidate returned
to the baseline for recursive allocation (`256` bytes and peak `2,113,611`)
and image scanning (`704` bytes and peak `2,110,086`), while preserving the
existing writer and image-record routes.

The final policy correction intersects a retained record's raw framing and
traversal limits with every narrower caller policy.  This applies to recursive
container searches, child views, image stores, BLIPs, and `File` views.  A
`File` rebinds its borrowed BLIP to that intersection before it can be
detached, so `into_owned().blip()` cannot regain a wider caller policy.  Tests
cover both permissive-origin/strict-caller and strict-origin/permissive-caller
paths, including detached image files and retained recursive depth/record
budgets.

Validation completed on the final source tree:

* `cargo test --offline -p litchi-odraw --all-targets --all-features`: 103
  passed;
* `cargo clippy --offline -p litchi-odraw --all-targets --all-features --
  -D warnings`: passed;
* `cargo fmt --all` and `git diff --check`: passed;
* the host consumer suites for `litchi-doc`, `litchi-ppt`, `litchi-xls`, and
  `litchi-imgconv` passed with all targets and features before the final
  source-only policy correction; the correction is confined to ODraw and is
  covered by the focused ODraw suite;
* the matched release and counting preflights for the preceding compact
  source checkpoint passed all 12 routes with matching checksums, `Store=48`,
  `Shape=312`, and the baseline allocation values above.  The post-policy
  source manifest is recorded below; its release and counting binaries must be
  rebuilt before performance acceptance.

The post-policy candidate source manifest is:

`947f989f88587c0d70378930f42683348617c7463ec2deb0f616280ad35fca0c`

The binaries and raw preflight captures are retained outside the checkout
under `/home/zhuhe/code/litchi-worktrees/.coord/odraw-limits-v2/`; none are
committed.  Earlier compact-candidate binaries remain historical evidence and
must not be presented as binaries for this final source.  No native timing or
profiling result is claimed by this record.  The coordinator must rebuild and
qualify the final candidate against the matched baseline in a quiet window
before accepting performance.
