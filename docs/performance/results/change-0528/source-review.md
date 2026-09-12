# 0528 XLSX scanner namespace-resolution source review

status: bounded read-only audit of the conditional namespace-resolution
mechanism at the exact 0528 baseline

production_change: none

This review covers `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs`.
The proposed mechanism moves `NamespaceResolver::resolve_event` out of the
unconditional event-loop edge and invokes equivalent element resolution only
for `Start` and `Empty`. The full value-only validator is a separate path and
is intentionally excluded because it consumes the `End` namespace. OLE2/OOXML
remains the active priority; ODF is deferred and iWork is excluded.

## Identity and method

| item | identity |
| --- | --- |
| baseline revision | `11a14d2043f65b1a6a935beaa1b9c16d924b749c` (`11a14d204`) |
| workspace `Cargo.toml` SHA-256 | `911a52cf6932b81550bc9ffd6e522c327dec178297ee9bc46d2ccedb693d4885` |
| perf harness `Cargo.toml` SHA-256 | `a04de024b9cbe9683cdb7307c3d7199daab7171bbdcb6bb8c30b361766857aca` |
| perf harness `Cargo.lock` SHA-256 | `13333f511914d8146c60282b5d6385693cf89c0f939fa52b12e4db821c9b8b36` |
| 0528 plan SHA-256 | `259454bb139493e82da98ad28c09332e65bcecfb0e1fea165a6b1ee6370f9b1e` |
| 0528 source-binding SHA-256 | `8a133620217daf78a9ed91500ebf551a7db56d341df3ab02127588430935361f` |
| 0528 baseline source manifest SHA-256 | `9af673c4c13f2abb3aeaf5c4e31df6a613de6297abc104bf472931ff7733529f` |

The source files governing the call edge and its consumers are pinned as
follows:

| source file | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs` | `137a317c696e65043027007fbd04edf47f8407563bf0799b852c9be0d54996f8` |
| `crates/litchi-xlsx/src/raw/namespace.rs` | `60a5bd73a3bb81fd5325418388964478bc66380558fb9d77efed27366af691b2` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/wire.rs` | `86418ecd5c0517d6c5dd000c1e126db51426b3231d0a96d1bf31bc80459c707f` |
| `crates/litchi-xlsx/src/cell_values/validation.rs` | `52e7d2d18f59e716c686f1c4c59b5835632fe6981dc6789ef7b06b555d136c0b` |
| `crates/litchi-xlsx/src/raw/worksheet/validation.rs` | `dc79e3e8dbb11943acdf68d0d68259cb4d05ff01616184a700e2f829b5f36bfa` |
| `crates/litchi-xlsx/src/cell_values/source.rs` | `10c9a99892dea1cf5c0dd529f628313126c4908c0fdc6e3694439c1c3889213b` |
| `rust-toolchain.toml` | `e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310` |

The standalone perf lock selects `quick-xml` `0.41.0` with registry checksum
`e660451e55124f798a69a5af3f49ccfbefbd41910eefd25caf2393e1f3473ec1`; the
local crate archive has the same SHA-256. The inspected local dependency is
`/home/zhuhe/.cargo/registry/src/index.crates.io-1949cf8c6b5b557f/quick-xml-0.41.0`.
Relevant source hashes are:

| dependency source | SHA-256 |
| --- | --- |
| `src/reader/ns_reader.rs` | `7ab687c3e93b896554f5754bd7c76042664f01d23cbfacb99605725c7ecc4a35` |
| `src/name.rs` | `f2ac828a48ca16d6d605949771d57a1eb00829a93a6556c682971cf5f28cb336` |
| `src/reader/mod.rs` | `6d0162e7d841d6156d0b452351078399b929c435346bcfcd6493138c04ba9399` |
| `src/reader/state.rs` | `74942da1414ad772fd8edff43a6ada6863c4d4aa0b9ba69740025c9db82f6072` |
| `Cargo.toml` | `0e7df0b5caa523509bb47a3ce3cb282e49e64dcf59e39cd12ab8e3aacd732700` |

The exact `name.rs` and `reader/ns_reader.rs` inputs are also retained under
`docs/performance/results/change-0528/inputs/`, with hashes matching the
dependency table. The review used source search and line-level reading of the
pinned dependency and current consumers. No Rust source was edited; no build,
test, benchmark, profile, allocator run, or capture was performed.

## Current scanner contract

At `scan.rs:200-265`, the scanner creates `NsReader<&[u8]>`, explicitly keeps
`check_end_names = true`, reads one event, and currently calls
`reader.resolver().resolve_event(event)` before matching it. The namespace is
passed to `Scanner::start` and `Scanner::empty`; `Event::End(_)` discards both
the event name and namespace and calls `Scanner::finish` with the local frame
and source positions only.

Every namespace consumer in this scanner is on a `Start` or `Empty` path:

* `is_spreadsheetml_name` gates worksheet, merge, defaults, columns, rows,
  cells, and primary `f`/`v`/`is` elements.
* `is_mce_name` and the SpreadsheetML check classify `AlternateContent`,
  `extLst`, and unknown direct cell children. This sets `mce_payload` and
  preserves the existing lossless-or-refuse writer behavior.
* `scan_guard` uses the resolved element namespace for `sheetProtection`,
  `dataValidation`, and the X14 extension namespace.
* `x14ac::attribute_name` and other attribute helpers receive the live
  `NamespaceResolver` while the start-element scope is active.

`Scanner::finish` at `scan.rs:715-910` uses only its own `FrameKind` stack and
the event positions. Formula close handling, scalar and formula spans, row and
cell publication, merge bookkeeping, and root-close tracking do not inspect an
`End` namespace. Positions are obtained from the reader before and after the
same `read_event` call, so a pure namespace lookup cannot change source spans.

## quick-xml sequencing and equivalence

The pinned `NsReader` implementation establishes the key ordering at
`reader/ns_reader.rs:64-100`:

1. `read_event_impl` first pops a pending scope.
2. The underlying `Reader` parses the next event.
3. `process_event` pushes namespace declarations for `Start` and `Empty`.
4. `Empty` and `End` set `pending_pop` for the next read.

Therefore, when the scanner receives a `Start` or `Empty`, the resolver already
contains that element's declarations. `Empty` must be resolved in the same
match arm before the next `read_event`, because its scope is pending-pop. A
`Start` scope remains active until the corresponding end has been processed.

`NamespaceResolver::resolve_event` at `name.rs:933-941` is pure. It resolves
`Start`, `Empty`, and `End` by calling `resolve_prefix(..., true)` and returns
the original event; all other event kinds return `Unbound` without a lookup.
`resolve_prefix` at `name.rs:951-967` scans current bindings, returns borrowed
`Bound`/`Unbound` results, and allocates a `Vec` only for an unknown explicit
prefix. There is no error return or resolver mutation in this call.

Moving the call into only the `Start` and `Empty` arms is semantically
equivalent for this scanner if the branch uses the same element rule
(`resolve_event` for that event, or `resolve_element(element.name())` /
`resolve_prefix(element.name().prefix(), true)`). It removes the match/call
overhead on text, CDATA, references, comments, declarations, processing
instructions, document types, EOF, and End events. It also removes the
`resolve_prefix` walk and possible unknown-prefix allocation for End events.
The total current `resolve_event` edge cannot be treated as removable: all
Start/Empty lookups remain required, and the End-only subset has not been
quantified here.

Skipping the End lookup does not skip namespace scope maintenance. The
`pending_pop` flag is set by `NsReader::process_event` before the scanner's
match and is consumed at the beginning of the next `read_event_impl`, whether
or not the caller resolves the End event. The End event itself is still passed
to the scanner unchanged.

## End-name checks, errors, and namespace behavior

The underlying `ReaderState::emit_end` at `reader/state.rs:189-253` pops and
compares the literal raw start-name stack when `check_end_names` is enabled.
It returns mismatched or unmatched-end errors before `NsReader::process_event`
can return an End event. This check is independent of namespace resolution;
it intentionally compares the qualified byte names, so skipping
`resolve_event(End)` cannot weaken it. The scanner's own unmatched-close and
unclosed-element errors still observe the same End/Eof event stream.

Namespace declaration failures remain unchanged because `NamespaceResolver::push`
at `name.rs:708-728` still runs inside `NsReader::process_event` for every
Start/Empty before the scanner branch. Reserved-prefix, declaration-count,
unknown-prefix classification, default-namespace rebinding, and MCE aliases
therefore retain their existing behavior. Unknown Start/Empty prefixes still
produce the same `ResolveResult::Unknown` and the same direct-cell refusal
classification. An unknown End prefix previously produced only an unused
`Unknown(Vec<u8>)`; removing that allocation changes no scanner result.

The full validator must remain separate. At
`crates/litchi-xlsx/src/cell_values/validation.rs:31-94`, its `Event::End`
branch checks both the local closing name and the resolved closing namespace
against the selected SpreadsheetML dialect. That loop must keep resolving
End events. Other codecs that consume End namespaces are outside this
`snapshot/scan.rs` candidate and must not be changed by a broad search-and-
replace.

## Risks and required candidate shape

No source-level correctness blocker was found for this narrowly scoped
scanner change. The following constraints are required before a candidate can
be compiled or measured:

1. Keep `reader.read_event()` and `NsReader::process_event` unchanged. Resolve
   `Start` and `Empty` only after that call, while their resolver scope is live.
2. Use element semantics (`use_default = true`). Resolving as an attribute,
   resolving before `read_event`, or resolving after the next read would change
   default aliases, MCE detection, or extension classification.
3. Leave `reader.config_mut().check_end_names = true` intact and do not replace
   the scanner frame stack with namespace or local-name comparisons.
4. Keep branch-local namespace borrows valid through `Scanner::start` or
   `Scanner::empty`; avoid cloning the resolver or the borrowed event. A direct
   `resolve_element(...).0` form avoids reconstructing an `Event` while
   preserving the `resolve_event` result for Start/Empty.
5. Do not apply this change to `cell_values/validation.rs` or other codecs whose
   End branch consumes a namespace.

The principal observable risk is implementation drift in branch restructuring:
an accidental early return, changed `position` ordering, or a delayed Empty
lookup could alter error precedence or spans. A second risk is an allocator
profile change for unknown End prefixes; it is an intended removal of unused
work, but it must not be presented as a typed-error or OOM guarantee.
