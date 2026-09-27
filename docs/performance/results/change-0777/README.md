# Change 0777 evidence

See the [integration report](../../0777-xml-attribute-integration.md).

The packet contains source integration, quality, architecture, review, paired
native capture, differential/equivalence, and raw instruction/allocation
observations. Offline validation, independent review, and owned-target cleanup
are complete. `seal.json` records the complete SHA-256 file inventory,
verified by final offline validation.

## Source and architecture custody

- `origin.json` records the 0777 base, the reviewed 0770 origin, and the three
  unrelated main-worktree files that must remain byte-for-byte untouched.
- `final-source.json` binds base `87e926fcc6`, candidate
  `1e334ffba9efb89e5badb8aa7a3e33052e96635c`, and 305 source/config hashes.
- `architecture-inputs.json` hashes `docs/GOAL.md`, the CRUD taxonomy, the
  accepted ADR set, and the ADR index used for the architecture review.
- `workspace-Cargo.lock` is the archived workspace resolution. The capture
  runner copies one standalone lock to both release legs and checks that
  `quick-xml` remains the reviewed 0.41.0 version.
- `probe-src/` contains the identical standalone attribute and MCE probes used
  for both legs. The probes are source-bound to their recorded origins.
- `migration-review.md`, `source-review.md`, and `review-snapshot.patch`
  preserve the independent migration and integration review. The migration
  survey covers 581 sites in 15 production crates; `litchi-imgconv` is the
  sixteenth lint-enforced owner and has no XML attribute sites. All 545
  checked migrations and all 26 exceptional checked callers are accounted for.
- `upstream-draft.md` is an unsent technical note. It is not an upstream
  submission and does not add an external compatibility claim.
- `cleanup.json` records removal of the three owned build targets after source,
  fixture, package-inventory, binary, and unrelated-file checks. The packet
  seal binds all retained evidence files.
- `relocation-validation.json` replays the packet on an independent packet
  path without native binaries or a source checkout and retains the same
  19-case, 114-native-process, 72-instruction-measurement, and 12-allocation-
  measurement counts with zero differential/equivalence changes.

## Quality evidence

- `quality.py` runs nine serial gates with offline, locked dependencies and
  separate target storage.
- `quality-0/` retains the command logs, source census, and manifest for the
  passing run. The run covers 16 owners and records 10,519 passing tests, 80
  ignored tests, and 407 suites.
- The gates include formatting, affected-owner all-target/all-feature checks,
  tests, warnings-denied owner Clippy, warnings-denied rustdoc, the OOXML
  facade check, the standalone harness check, and crate-boundary validation.
  The facade check retains one pre-existing unused-helper warning for
  `missing_ooxml_catalog_part_error`; it exits successfully. This is not
  presented as a warning-free facade build.
- The source census and `final-source.json` bind the passing gates to the
  candidate. A later source or capture change requires a new source-bound run.

## Capture protocol

`capture.py` refuses to reuse a capture directory, keeps before and after
release targets separate, and checks source, fixture, package, probe, lock, and
binary identities throughout the run. The intended serial process order is:

```text
before, after, after, before, before, after
```

The 19 native cases produce 114 timed processes with nine samples and two
warmups on CPU 12:

- MCE: benign worksheet, benign document, stream-count worksheet,
  stream-count document, and prefixed controls with 1, 2, 8, 9, and 32
  attributes.
- OPC: relationship packages with `n = 0, 8, 29, 30, 32, 33, 256, 1024,
  4096, 16384` extra namespace declarations. The tag carries three ordinary
  attributes, so the 29/30 pair exercises the reviewed 32-name transition.

The MCE timer includes parser, configured observer, and digest work after input
construction. The OPC timer includes `OpcPackage::from_bytes` and the
deterministic name-sorted part digest after package construction. Input
construction, process startup, and JSON writing are outside both intervals.
RSS is a whole-process `/usr/bin/time` peak and is not a portable allocator or
live-byte measurement.

The completed capture ran one candidate iterator-equivalence probe and two
differential commands, one for each source leg. The differential covered 2,088
packages, 2,037 mutated packages, 113,830 bounded mutations, and 228,716
outcome comparisons, with zero changed outcomes. The equivalence probe visited
7,598,414 real tags and 2,332,430 mutated tags. It found 105 real tags and
1,015,360 mutated tags over the 32-name boundary, with zero mismatches. The
comparator covers the canonical OPC adapter and synchronized OLE copy against
quick-xml; the synchronized production copies also have shared quality tests,
but this is not a runtime comparison of all five copies over the corpus. These
receipts check item/error identity through the first parser error on
parser-reached tags; they do not prove every post-error recovery path.

## Native result and acceptance rules

`analysis.json` contains all 19 absolute p50/RSS rows. The positive changes
above five percent are `n=8` latency (+5.23%), `n=30` latency (+6.09%), `n=32`
RSS (+8.86%), `n=33` latency (+6.30%), `n=256` latency (+54.10%), and `n=1024`
RSS (+5.95%). The accepted `n=256` row is 22.811 → 35.151 µs. The
`n=1024`, `n=4096`, and `n=16384` rows return the exact existing namespace
binding-limit refusal in both legs and must remain labeled refusal controls.

The raw instruction lane has 56 primary perf rows plus eight primary iterator
rows in `instructions-0/`, and eight secondary `n=256` perf rows in
`instructions-256/`: 72 instrumented process measurements total, plus two
supported qualification runs. The qualification receipts report
`running_percent: 100.0` for the user-instruction event. The raw allocation
lane has ten observations in `allocations-0/` plus two additional `n=256`
observations in `allocations-256/`. The ordinary worksheet allocation total is
12,455,831 → 12,455,829 bytes with 129,512 calls in both legs, and the ordinary
document total is 2,539,502 → 2,539,500 bytes with 16,056 calls in both legs.
These −2-byte whole-process differences do not establish a parser allocation
reduction. The `n=256` −11,666-byte change, with 34 additional allocation
calls, is also whole-process data and does not establish a parser allocation
reduction. The extra `n=256` observations are retained as a follow-up to its
native +54.10% flag; the native capture was not rerun or selected from them.
Instrumented timings and whole-process counters remain separate from the native
timing table. The consolidated `observations.json` receipt is complete and
offline validation and the final packet seal check have passed.

`analyze.py`, `validate.py`, and the allocation/instruction helpers describe
their checks; their replay outputs are bound by `observations.json`. The
contract-only CRUD coverage validation also passes for 15 categories and 33
mapped selectors, with no scenario promotion or full-run timing report.

The report retains every case's absolute before-to-after p50 and RSS values,
every positive change above five percent, all refusal outcomes, and the
raw-source/binary receipts. It must retain all failed attempts and their raw
logs, preserve source/fixture/probe/lock identity, and distinguish instrumented
instruction/allocation observations from the uninstrumented timing table. No
result may be sealed if offline validation, cleanup custody, or outcome
identity checks are incomplete.

Historical 0770 timing or equivalence observations are stale and are not
substitutes for this capture. The unsent upstream draft is retained for review
only. The broader non-iWork goal remains active; this packet does not claim
CRUD, cold-cache, range-source, or scaling completion.
