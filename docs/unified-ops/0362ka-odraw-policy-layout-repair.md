# 0362ka — ODraw policy-layout repair

- Date: 2026-10-21.
- Base: `61515148e1`.
- Parent source checkpoint: `08f586ca42`.
- Candidate source: `33f4981030`.
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
container searches, container image scans, child views, image stores, BLIPs,
and `File` views.  A
`File` rebinds its borrowed BLIP to that intersection before it can be
detached, so `into_owned().blip()` cannot regain a wider caller policy.  Tests
cover both permissive-origin/strict-caller and strict-origin/permissive-caller
paths, including detached image files and retained recursive depth/record
budgets.

Validation completed on the final source tree:

* `cargo test --offline -p litchi-odraw --all-targets --all-features`: 104
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

The final source validation on candidate `a5935d50d5` passed 104 ODraw tests
and all-target/all-feature Clippy with `-D warnings`.  The matched external
workspace used baseline `61515148e1`, candidate code `33f4981030`, the same
driver (`probe-main.rs`, SHA-256
`5025821f9af1986d1caf7fb9c2f631875ca30ab725e33487681f4aeac1ef2efa`), and
the same lock/profile.  The candidate source manifest is
`57f627ce87c8a9b32a80a8c3f3f03d998b4f08b463c401585d8654eb239c4c3e`.

Twelve-route untimed fidelity preflight checksums matched.  The counting
allocator preflight matched all twelve routes exactly; recursive remained 256
allocated bytes with 2,113,611 peak live bytes, and scan remained 704 with
2,110,086 peak live bytes.  The protected release and allocator binaries and
raw captures are retained outside the checkout under
`/home/zhuhe/code/litchi-worktrees/.coord/odraw-policy-final-20261021/`; none
are committed.  Candidate release SHA-256 is
`b6a3496aaf8298dc2d9413429ad7fb0bdaf008015daa0a26ba3df3230d2e825c`, and
candidate allocator SHA-256 is
`6073d882dfc470e085a4bc3bc4ea7ab1965be1c390e0ff76038047f784425903`.
Native timing is still pending a coordinator quiet window; no performance
acceptance is claimed before that paired run.

After the final narrow policy correction, the affected consumer matrix was
rerun against the same source tree. A combined all-target/all-feature check
passed for `litchi-doc`, `litchi-ppt`, `litchi-xls`, and `litchi-imgconv`.
Their corresponding all-target/all-feature test commands all exited
successfully, as did the 104-test ODraw suite. The consumer rerun found no
additional migration or policy gap: every remaining `Limits::default()` in
these paths is an explicit host-selected policy or test/fixture setup, while
raw ODraw entry points continue to require a `Limits` argument. The target
used for this rerun was
`/home/zhuhe/code/litchi-worktrees/.coord/uo2-w8-odraw-limits-target` and was
7.0 GB after completion; no generated artifact was added to the repository.
