# 0362dvifb — DrawingML native raw six-lane diagnostics

This is a finite runtime-validation follow-up to 0362dvif. It covers the
current raw adapters only: chart, diagram, text/media, theme plus InkML,
geometry/color/formatcode2/transform, and the complete Theme-part family
lane. It records correctness, limit-boundary, typed output-refusal, and
fidelity observations; it does not make an ABBA, RSS, allocation, latency,
or performance claim.

The reduced receipt is
`docs/unified-ops/results/0362dvifb/drawingml-native-six-lane-diagnostics.json`.
The full reports, streams, build logs, binary, and target remain outside Git
under:

`/home/zhuhe/code/litchi-worktrees/.coord/drawingml-native-raw-six.yo47Ie`

## Inputs and command

All six final processes used source revision
`75afbd79cbac2cd4a7f7fe42d296d79c3439b586`, tree
`7bd75b5fb3220b2f13aa3d6566c0cd921d34ac48`, and the tracked fixture
descriptor SHA-256
`e72237ec2095535667fa34e962e37950c96ef558eef151839a4e8235a1256925`
(2,407 bytes). The probe SHA-256 is
`af7975d475b42c0766b6280829cbe03fb6a9154556d2a4d7939d894e35f61bb3`
(236,030 bytes); the validator is
`4454b2f6eb2cd78b30686613573464d25daaf37ab8dab35a3ef523914536c450`
(70,755 bytes). The nested `tools/perf-baseline/Cargo.lock` was the locked
input (57,227 bytes,
`eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`).

The current debug binary was built offline and locked with one Cargo job,
the 20-GiB cap, and the seven current DrawingML features. Its exact argv,
environment, build log, binary hash, target size, and toolchain are captured
in the receipt:

    cargo build --manifest-path tools/perf-baseline/Cargo.toml \
      --bin drawingml-raw-native \
      --features current-api,drawingml-owner-current,drawingml-raw-current,\
    drawingml-theme-family-current,drawingml-theme-family-owner-limits-current,\
    drawingml-ink-limits-current,drawingml-transform-limits-current \
      --offline --locked \
      --target-dir /home/zhuhe/code/litchi-worktrees/.coord/drawingml-native-raw-repair-target \
      -j 1

The resulting binary is 25,154,104 bytes with SHA-256
`82182f0ba0fb5ba174a0ecff5430d1f702a35eae4595e17a308e47ac09378430`;
`du -sb` reported 5,419,181,476 bytes for the target. The process
environment unset `LITCHI_BLESS` and `LITCHI_CAPPED`, set
`LITCHI_MEM_MAX=20G`, and used `RUST_TEST_THREADS=1`,
`CARGO_INCREMENTAL=0`, and `CARGO_BUILD_JOBS=1`.

Each final run was a fresh candidate normal process with
`iterations=1`, `warmups=3`, `blocks=5`, `samples=20`,
`target-ms=500`, and `minimum-ms=150`. The exact seven command lines
(including the initial duplicate-output attempt), environment, summary, and
every output/sidecar byte count and SHA-256 are in the receipt and the
external `runtime-*` files.

## Results

| raw lane | result | boundary and fidelity evidence |
| --- | --- | --- |
| chart | PASS | near/at/over input limits and malformed XML; typed output cap refusal with the production bounded-partial behavior preserved (2,260 partial bytes within cap, non-zero bounded writes allowed) |
| diagram | PASS | input limit and malformed data-model refusal; source-edit-and-serialize output cap preflight with zero writes |
| text/media | PASS | text, contextual SVG, Model3D, and scanner-only image-blip limits/malformed cases; inherited SVG namespace context, contextual source/changed checks, and semantic model checks |
| theme + InkML | PASS | theme and InkML source limits/malformed cases; InkML `write_profile_to` exact retained source for unchanged and changed fixtures, with semantic checks |
| geometry/color/formatcode2/transform | PASS | geometry structural path limit, XML limits, malformed cases, native typed output refusals; independent formatcode2 lexical-splice oracle and semantic checks |
| Theme-part family | PASS | XML limit and malformed cases; independent lexical-splice oracle plus add/replace/remove mutation identity checks for family name, id, and variant id |

The six reports all have `status=qualified`, every operation row is
qualified, every applicable over-limit case is a typed `limit` refusal,
every malformed case is `invalid_input`, and every output-cap preflight
checks typed refusal behavior. The reduced receipt retains each generated
source/changed/malformed fixture hash and byte length, fidelity hash and
semantic flags, boundary result, output cap, operation scope, and refusal
kind/typed limit.

SVG source output is intentionally checked through the inherited contextual
route: its raw source bytes and semantic model are preserved even when
namespace materialization changes the serialized bytes. The standalone
route remains an additional outside-clock check. Image-blip is reported as
a scanner-only helper, so its output-cap row is not applicable. Chart's
bounded partial output is reported rather than being incorrectly treated as
a zero-write preflight.

The first family invocation attempted to reuse the bulk report path and
exited 1 before execution because the JSON destination was not fresh. It is
retained as an external setup diagnostic; the final family run used a
distinct report path and exited 0. This does not change any of the six
final report results.

No report from this record is a performance comparison: the reports are
single current-candidate diagnostics from a debug binary and do not contain
a matching historical baseline or ABBA pair. The earlier b2/v7 diagnostics
and the 5c67 compile/preflight receipt remain preserved externally.
