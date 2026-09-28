# Candidate selection at b76786208d

The current source and 0822 exact-owner stacks select `shape::Scanner::scan`.
The independent `profile_basis.py` reader binds the retained compressed and
decoded frame hashes and counts exact symbols with the original executable
identity. The prior 179-path HEAD seal passed before this trial. These are
overlapping inclusive sample counts, not elapsed-time fractions or an Amdahl
prediction.

| Stack property | Repeat 0 | Repeat 1 |
|---|---:|---:|
| Exact public edit owner | 2,971 | 2,996 |
| Contains `Scene::read_with` | 636 | 624 |
| Scene stack contains `NsReader::process_event` | 128 | 157 |
| Scene stack contains `NamespaceResolver::resolve_event` | 124 | 116 |
| Scene stack contains attribute iteration | 62 | 62 |
| Scene stack contains commit `compaction_scene` | 301 | 304 |

The public real-file path opens the package outside the edit clock. Inside it,
initial snapshot capture validates slides and notes and fingerprints every
logical payload. One shape edit reads the source scene, rewrites the selected
text, and reads the staged scene. Commit compacts changed XML and compares
scenes when compaction changes bytes, constructs its patch, and captures the
result. Publication applies the commit; output serialization and exact reference
comparison are outside this particular edit clock. The synthetic lifecycle
cases additionally time serialization.

The scanner currently passes the borrowed input event through `NsReader`'s
namespace-processing `Result<Event>` return and then through `resolve_event`'s
`(ResolveResult, Event)` return. The candidate keeps quick-xml's `Reader` and
`NamespaceResolver`, handling the same scope transitions locally and resolving
element names directly. It changes no XML grammar, namespace search algorithm,
attribute checker, scene record, semantic projection, or retained cache. The
trial must establish whether removing those event wrappers helps public
workflows; source inspection alone does not establish a speedup.

Two alternatives were reviewed first:

- Initial complete fingerprinting necessarily materializes every logical
  payload under the existing exact revision contract. The allocation-identity
  digest memo already avoids repeated hashing when bytes are shared. Replacing
  logical payload digests with compressed bytes or CRCs is not an equivalent
  optimization.
- Opened capture parses the presentation catalog twice. Reusing the borrowed
  view's catalog is semantically possible if its unvalidated form and limit
  ordering are retained. Change 0649 already identified this small cost below
  its then-current noise floor. No catalog change is included in this trial.

## Contract review

| Accepted decision | Trial obligation |
|---|---|
| 0001, 0002, 0010, 0011, 0024 | Private format-owned scanner change; no dependency or public API change. |
| 0003 | Same scene values, spans, refusal ordering, source-checked patches and publication bytes. |
| 0005 | Fresh before/after public workflows, allocation and RSS guards; no new cache or retained state. |
| 0006 | Exact namespace stack timing, malformed-input errors, finite limits, BOM-relative offsets and unknown markup. |
| 0008 | Differential oracle plus applicable crate checks, tests, lint, docs and boundary gate. |
| 0030, 0032 | Materialization, fingerprints and snapshot memos unchanged. |

All 35 normative inputs previously read for the program were rehashed against
0822 and remain unchanged. The three unrelated working-tree files remain outside
the trial's ownership. No iWork work is included.
