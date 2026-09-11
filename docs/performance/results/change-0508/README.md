# 0508: default text-export baseline evidence

This batch adds existing semantic text-output selectors for plain RTF, ODT,
ODS and ODP to the default matrix. It is a descriptive baseline expansion,
not a before/after optimization or native-producer claim. The timed runners,
corpus bytes and semantic checks are unchanged.

`protocol.json` fixes one identity preflight and two serial full runs with
three warmups and fifteen measured samples per row. `run.py` captures each lane
exclusively, binds Rust/TOML/lock sources before and after, records build flags,
checks the release binary digest before each capture, and retains raw logs.
`host.json` records the actual toolchain, CPU and environment. This is a shared
host without a core pin or isolated-host claim; measurement lanes are serial.

Reproduction from this committed source:

```sh
mkdir /tmp/litchi-goal-0508
python3 -B docs/performance/results/change-0508/run.py build
python3 -B docs/performance/results/change-0508/run.py preflight
python3 -B docs/performance/results/change-0508/promote.py
python3 -B docs/performance/results/change-0508/run.py r1
python3 -B docs/performance/results/change-0508/run.py r2
python3 -B docs/performance/results/change-0508/verify.py
```

The runner refuses existing lane outputs. Use a separate checkout with a fresh
copy of the scripts and protocol, retaining `inputs/` but removing prior lane
outputs, receipts and source manifests before recapture. Set the protocol's
`revision` to that checkout's HEAD and refresh `host.json` for its toolchain
and host before running. The new checkout's
revision and dirty state will be recorded, so its catalog hash will differ;
`promote.py` must regenerate the checked catalog from that capture. Corpus and
result-key identities must still match. All build output belongs to the named
scratch directory and may be removed after the child processes finish.

For read-only replay of the committed evidence, simply run `verify.py`.
It checks the retained sources, old identities, generated catalogs, all raw
sample statistics, sink controls, policy bindings and two deliberate invalid
report probes. It does not require the removed release executable or rebuild
old binaries. `SHA256SUMS` binds the retained evidence bytes; run `sha256sum -c
SHA256SUMS` from this directory. Historical receipts record what ran, while
replay validates the retained reports and source tree.

RTF uses a bounded pre-reserved retained output buffer and exact byte equality.
ODF uses a bounded hashing discard sink and checks digest, byte/object progress,
and zero retained output bytes. Document opening precedes the timed export; digest
updates during ODF writes remain inside timing. These scopes should not be
compared as a pure encoder speed or document open/export lifecycle.

The RTF default variant is plain. Transport variants remain opt-in; the
checked-in watermark fixture remains conservatively unmapped in generator
provenance. Unsupported, protected and native-producer scenarios are not
promoted by these synthetic fixtures.

The final gates pass: 483 Rust tests (one existing ignored opt-in security
fixture), 199 Python tests, full harness formatting, warnings-denied Clippy
and rustdoc, doctests, allocator-feature all-target checking, boundaries,
and the strict claim/report classifiers. `gates.json` binds their logs.
`cleanup.json` records removal of 1,348,038,656 unique-inode allocated bytes
from the owned scratch after all process references closed.

The schema-1 identity retains its validator-pinned historical `harness_source`
string ending in `src/main.rs:Case::DEFAULT`. The actual selector list is in
`src/lib.rs`; the Rust source manifest and coverage selector-source binding
identify that file. The legacy string is not an accurate current symbol
location and is not used as source custody.
