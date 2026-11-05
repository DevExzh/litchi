# 0362ugy — historical probe compatibility and diagnostic builds

Reviewed source `ccc536d5242cefcc6f58df200342c38faa5ec91e` merges at
`f6fe430a4d19d94aa360c0e9ed03ebd757265225`. Historical W4b and W5
qualification probes now compile against their actual baseline APIs. The
production API and measured workloads are unchanged.

The W4b overlay removes exactly three candidate-only lazy selector-resolution
chains, translates the single unified commit accessor to its historical
field, and retains the previously recorded CRUD trait import. The baseline
is commit `be644568846ad2fcf0622386f2f2438d0eac21c0`, whose synthetic tree
is `684a028b2ae30a3a110b87bf6743fc598c6691a2`; the tree is not a commit.

The W5 overlay uses the recorded historical route module and retains both
later DOC publication controls. It translates the unified File options entry
point and result accessor to the historical API. It does not drop the
complete-publication controls or move their measurement boundaries.
Baseline W5 remains `7afba5f146aa1a16eba73ebed2144dd041b33fb5`.

The adaptations reject unexpected route shapes/counts. Full source refs,
baseline identities, route hashes, clean candidate state, disjoint paths,
no-overwrite behavior and analyzer identities remain enforced. All 21
protocol tests pass, independently repeated by the coordinator and reviewer.
Three historical W4b consumer checks and the W5 source/ordinary probe check
also pass. They are diagnostic `cargo check` commands, not native builds or
measurements. [The reduced compatibility receipt](results/0362ugy/historical-probe-adapters.json)
retains the precise v4/v5 source, overlay, command, lock and log boundaries.

## Retained diagnostic evidence

[The reduced native receipt](results/0362ugy/native-diagnostics.json) binds
the external build/validation manifests, commands and binary hashes:

- iWork constructor: baseline `ddf2ac8ae2117cd78829ef94bcb610d6cbdfa0e2`
  and candidate `b6732d02b241705ba153c1a2cb9ff03bf0a13f93`;
  four binary variants each validate all 15 routes.
- iWork owner: baseline `a18593385cf3162911f00ebd3cbb1379855a8eba`
  and the same b673 candidate; four variants each validate all four routes.
  These content checks are operationally rejected for acceptance because the
  fixture generator ran from the gate checkout at e157 while another command
  owned it. The explicit rejection receipt remains attached.
- DOCX durable ink: baseline `d0406be01712496ca0d43354edf4d2fe80162553`
  and candidate `e1576fa6dcd9f856ddea0b0ec6b038cb6b076265`; native and
  allocation untimed validation outputs match across the pair. Source/project
  locks are matched. The source archive's omitted configuration files remain
  explicit in its provenance receipt.

Independent review verified 221 artifact hashes and 41 successful command
receipts with 82 stdout/stderr streams. The coordinator additionally checked
all 224 distinct absolute-path artifact identities in that verification
receipt, including its extra provenance artifacts; none was missing or
mismatched. The four iWork build/validation manifests exceed 1 MiB and stay
outside Git. Their names, sizes, hashes and regeneration commands are bound
by the reduced receipt. No binary is committed.

These results establish diagnostic compilation and content checks only.
There is no timing, ABBA or final performance acceptance claim. In particular,
the owner fixture-generator overlap is not waived. Final qualification uses
its own frozen endpoint and the existing quiet-host guards.

ODP baseline/candidate diagnostic executables also build, but their original
dependency resolutions differ and no probe is run. A fresh e4-backed build
projection now matches frozen f6fe: all 204 external lock records, including
versions, checksums and dependency lists, are equal. Traversing ODP's resolved
dependencies reaches the same 163 nodes and 157 external rows on both sides.
The coordinator independently checked those sets and all nine source/metadata
artifact identities. Corrected v2 provenance preserves earlier receipts and
distinguishes the 204-entry complete lock from the 157-entry ODP closure.
This is preparation for matched builds, not performance acceptance.

## Acceptance source freeze

Production and protocol source is frozen at
`f6fe430a4d19d94aa360c0e9ed03ebd757265225`. A fresh full 53-step gate
runs from the clean diagnostic checkout with state directory
`(deleted-coord-artifact)`, the shared
20 GiB Cargo cap, one build job and a target limit of 20,000,000,000 bytes.
It invokes the reviewed safe runner without step selection or a starting
offset. Later receipt-only commits do not relabel this measured source.
Any required production or protocol correction requires a new source freeze.

Full gate completion, the all-68 rustdoc inventory and dispositions, actual
corpus replay, glossary closure and the remaining native qualification lanes
are still required. Only comprehensive redaction audit 22 is deferred.

The f6fe attempt subsequently returned required repairs; see
[0362uma](0362uma-final-gate-findings.md). It remains diagnostic source,
not the final production acceptance pin.

## Durable ODG baseline capture

The immutable `ad52437f5e5e2c71ea476eedd5ddcf61aee2021d` baseline source
closure completed on 2026-10-24. The earlier orphaned capture produced no
usable output; this fresh attempt writes the complete result atomically to
an external file. It performs no Cargo build, probe validation or timing.

[The reduced receipt](results/0362ugy/odg-baseline-source-closure.json) binds
460,379 source entries and 6,263,042,584 content bytes. The complete artifact
is `(deleted-coord-artifact)/baseline-source-closure.json`,
182,412,099 bytes, SHA-256
`b1cc42d57d07baa4f02115cb9db82847e8079416f08cdf67e4421df3d906b8a5`.
The coordinator verified that file, its embedded closure binding, the clean
source ref, lock and actual builder identities. The capture uses builder
SHA-256 `4f1f556475107c72fd28846503b1807b0490ae0e11251aed40d6e97179c8b46e`;
later tooling edits do not relabel that provenance. A future build must
establish its own current source and protocol bindings.

The adjacent `run_capture.py` has SHA-256
`9facacd33b30def52eca310d33e25d527fe9e661992d4595bb09ac1072b1ab41`.
To regenerate with the same source, lock and builder, choose a fresh external
directory and invoke the preserved wrapper with only its output root changed:

```sh
python3 - /path/to/fresh-external-directory <<'PY'
import importlib.util
from pathlib import Path
import sys
script = Path('(deleted-coord-artifact)')
spec = importlib.util.spec_from_file_location('odg_capture', script)
capture = importlib.util.module_from_spec(spec)
spec.loader.exec_module(capture)
capture.OUT = Path(sys.argv[1])
assert not capture.OUT.exists(), 'preserve the existing evidence'
raise SystemExit(capture.main())
PY
```
