# Change 0482 evidence: bounded XML audit reader

This bundle measures the bounded XML audit reader primitive that the later
DOCX decoded splice will use. Each materialized and streaming route is run by
the same executable built from one source snapshot. The result is a route
comparison for the named deterministic XML generator; it is not a DOCX append
benchmark and does not establish a package-wide memory or latency claim.

The formal schedule is A1/B1/B2/A2. A is `materialized` and B is `streaming`;
each arm covers 64 KiB, 8 MiB, and 128 MiB in forward order for repeat 1 and
reverse order for repeat 2. The normal and allocator binaries are measured in
separate groups. There are 24 case reports and 720 measured samples in total
(12 process receipts per instrumentation; each receipt contains one case).
Every process is pinned to CPU 2 and wrapped with `/usr/bin/time -v`; the
retained RSS value is the process maximum resident set size, including setup
and teardown. Warmups are three and retained samples are thirty.

Both successful release builds must be retained before the protocol is frozen.
The accepted source rotation writes digest-bound
`builds/normal-accepted.json` and `builds/allocator-accepted.json`; the
corresponding `validation/build-normal-accepted.json` and
`validation/build-allocator-accepted.json` gate receipts retain the cargo
command and stdout/stderr. Build labels are
`build-normal-accepted` and `build-allocator-accepted`; the earlier build
labels and receipts remain retained as history. The build records bind the
expanded source snapshot, executable metadata, revision, environment, and
build interval. Formal capture receipts are accepted only after both build
intervals finish.

`machine.json` records the host, storage, toolchain and affinity observations.
The full source snapshot also includes a pre-existing, formatting-only Keynote
working-tree change. `unrelated-source-state.patch.txt` retains that delta for
exact source reconstruction; this batch neither applies nor commits it to the
Keynote source file. The implementation and measured routes exclude iWork.

`xml-assets.json` retains the explicit 77-file XML resource inventory and one
SHA-256 for every compiled `crates/*/src/*.xml` asset. Freeze checks it against
the current source tree and records its digest in `protocol.json`; the default
verifier checks the retained inventory against the content-addressed source
snapshot, while `verify.py --live` also compares it with the current checkout.

The accepted validation set also covers the ZIP package format/tests/clippy/docs
gates and two functional XML fuzzer gates. `fuzz.py` retains the accepted
ASan/libFuzzer build and 10,000-run seed-482 smoke records under `fuzz/accepted/`:
the prepared manifest and lockfile, the 11-seed inventory, and copied
post-run corpus and artifact directories are all digest-bound by the frozen
protocol. Earlier fuzzer attempts remain retained as developmental history.

The harness command is:

```text
xml_stream_audit --mode materialized|streaming --sizes BYTES --samples 30 --warmup 3 --json REPORT
```

`run-measurements.py` does not build or discover binaries. Root supplies the
explicit paths recorded by the build receipts under `/tmp`, then freezes the
protocol and runs each A/B arm while `/tmp/litchi-goal-0482/cpu.lock` is held:

```text
python3 -B run-measurements.py --freeze \
  --normal-binary /tmp/litchi-goal-0482/normal-accepted/xml_stream_audit \
  --allocator-binary /tmp/litchi-goal-0482/allocator-accepted/xml_stream_audit
python3 -B run-measurements.py --instrumentation normal --arm materialized --repeat 1
python3 -B run-measurements.py --instrumentation normal --arm streaming --repeat 1
```

The two run commands are illustrative; root repeats them for repeats 2 and
the allocator instrumentation. A failed pilot or formal
process is retained with its started receipt, terminal receipt, stdout,
stderr, GNU-time resource file, and report when one was produced. Existing
paths are never overwritten.

The freeze command may repeat `--required-validation-label` to bind the final
validation set into the protocol; both build labels are always added and are
mandatory. Root supplies every final label selected for this change before
captures begin.

The XML, OPC, and ZIP checks from earlier source states remain retained under
their original labels as developmental history. The post-rustdoc `-final`
receipts remain retained too. The accepted source rotation uses the
corresponding `-accepted` labels, so the frozen protocol cannot silently treat
an earlier receipt as validation of the accepted source. The earlier evidence
test receipts and the generator-oracle/pilot rerun remain retained; the
required accepted receipt is `evidence-tests-final-v2-accepted`.

`analyze.py` validates the `litchi.xml-stream-audit.v1` report, requires
successful sample-by-sample audit equality between routes, checks generated
byte identity and zero allocator net live bytes, and derives the mean, 95%
t interval (`t(29)=2.045229642`), nearest-rank p50/p95/p99, process RSS, and
all pair/repeat changes over five percent. Normal and allocator results remain
separate. Allocator rows retain raw absolute live peaks and derive requested
bytes, allocation/deallocation/reallocation calls, and region-peak increments
from operation entry live bytes. It writes `summary.json` once; `--data-only`
recomputes it without writing.

`verify.py` checks protocol/script/binary/source custody, exact A1/B1/B2/A2
ordering, both required build and validation gates, developmental pilot
retention, final successful pilot source bindings per lane, every retained log
and receipt digest, report/resource presence, and an independent data-only
summary recomputation. Its default mode uses retained binary metadata, so the
bundle remains verifiable after temporary executables are removed; `--live`
also reopens those paths and checks current executable/source identity.
`test_evidence.py` also proves that missing-capture, tampered-summary, and a
fresh copied bundle with unchanged receipts are handled correctly.

The evidence supports the scoped statement “bounded XML audit primitive
necessary for the decoded splice”: the streaming route consumes a fixed
generator source window and the auditor's finite token/state envelope while
preserving the same deterministic report. It does not support a DOCX
transaction, OPC preservation, cold-cache, remote-source, constant-RSS, or
program-wide speedup claim.

The [measurement tables](measurements.md), independent
[raw measurement review](measurement-review.md), and
[evidence review](evidence-review.md) record the accepted results. Streaming
uses a 65,587-byte operation heap increment at all three measured sizes.
Normal latency increases by 3.587–4.092% at 8 MiB and decreases by
10.592–10.809% at 128 MiB. The tables retain all outliers and review flags.

After live verification, owned temporary binaries and fuzzer workspaces were
removed while shared build caches were preserved. The actual copied-bundle
[replay checks](evidence-validation/replay-check.json) then accepted an intact
fresh copy with unchanged receipts and rejected four separate copies with a
missing capture, modified summary, missing required gate, or modified fuzzer
record. These checks used retained evidence after runtime cleanup.

`SHA256SUMS` seals every other file in this directory, including reviews and
replay receipts. From this directory, run `sha256sum --check SHA256SUMS` for
the file-integrity check and `python3 -B verify.py` for protocol and result
verification. The latter remains usable after temporary executables have
been removed; `--live` requires those original runtime paths.
