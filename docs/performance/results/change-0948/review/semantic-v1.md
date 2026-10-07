# 0948 independent semantic review

Status: **PASS / no blocking issue found.** This is a static review of
`candidate-source-v3`, compared with `candidate-source-v2`; no source files
were changed and no Cargo command was run.

## Source identity

Both candidates derive from source commit
`b6c5c80ffbcb60736d6c4c2f3b590838640008a9`. The v3 tree contains 10,789 files
and 260,076,537 bytes. Its deterministic content manifest is
`7c32313d91b30edd9ed34bba92a20aef3b82139a783cacad23ccc48f71d7d4c9`, computed
from sorted relative paths and, for each file, its byte length and SHA-256.

Reviewed implementation hashes:

| File | SHA-256 (v3) |
| --- | --- |
| `crates/xml-minifier/src/audit.rs` | `29efd9aa99c9491a1e86cfc25f1bbb31b3da7b6cbfc4438a1638c15b18b95cee` |
| `crates/xml-minifier/src/audit/namespaces.rs` | `2b24bf90377ba75db3a15eeaa53d010bf817b17b2741b8b20e6e377fd34bbdfe` |
| `crates/xml-minifier/tests/source_replacement.rs` | `600712ff53ad145fa82659989104a9b496e2b03afd1295f20605f29ca7a5615e` |
| `crates/litchi-opc/src/lib.rs` | `1a9c1cc6810b343aa5647aab9aa60e09adf14fc63b6cc0459196316ffac442a2` |
| `crates/litchi-opc/src/source_backed.rs` | `571ec3e2e08d11626b5ce489c5d7de498445c1dd72441692306772c1893c3498` |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `50a4ef8fd9f2e810d2b1315d8a389086134a28901a2b86a98cc8abeb84308726` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/package.rs` | `56572673773c85ed18fc0e6c75478c7958223a5bb8d054a01b3c7927a92c2e7b` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/mod.rs` | `2cc0c3797f9ba16d1e4c23782cded2789948362947339384de3396284dc17559` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs` | `8b2079cd3c5437d7e7113e5d4486bed9ed8eb677e87ace82c63b8ce1559f5084` |
| `crates/litchi-xlsx/src/cell_values/facts_oracle_tests.rs` | `026f466c5eae0c641276c00dcf32bd1e89875c9ee650977ce20fcbd452d41825` |

## Findings and invariants

- XML `verify_source_replacement_with_spans` audits the original first and
  uses the canonical complete replacement audit for every uncertain case
  (`audit.rs:1139-1180`). Span bounds, ordering, overlap, and unchanged gaps
  are checked before a proof is accepted (`audit.rs:1188-1220`). Each source
  range must close as one non-root element; replay captures post-close
  inherited namespace state and the surrounding depth (`audit.rs:2475-2530`).
- Window replay starts with the correct depth and namespace bindings, rejects
  unbalanced or non-markup boundaries, removes the synthetic EOF event, and
  re-totals event, attribute, text, and byte limits before success
  (`audit.rs:2621-2677`, `2688-2718`). The reviewed adversarial tests cover
  adjacent siblings, shifted replacement coordinates, namespace aliases,
  malformed windows, root/non-element hints, reversed/nested hints, EOF, and
  aggregate limits; the mixed-content boundary behavior follows from the
  markup-ending and unchanged-gap invariants.
- Namespace retention uses checked logical admission accounting, including
  vector capacities and conservative hash-table charge; it is explicitly not
  a physical RSS claim (`audit/namespaces.rs:153-169`). The aggregate clone
  admission cap is 8 MiB and the span/depth/open-stack bounds are finite
  (`audit.rs:1109-1113`, `2450-2567`). The `Window` inline array is a separate
  bounded allocation (about 3 MiB at 4096 entries); fallback may briefly retain
  prior search state, but that state is finite.
- OPC exposes the bounded hint API (`source_backed.rs:1934-1943`), carries hints
  only through topology publication (`8348-8353`), and dispatches them only to
  the source XML pair auditor (`9673-9678`, `13003-13027`). Invalid or stale
  hints therefore cannot authorize malformed output.
- XLSX stores only detached pointer/length metadata plus bounded spans
  (`snapshot.rs:1511-1565`). Publication uses the hint only when both source and
  replacement allocations match (`1734-1770`). Rewritten bytes were already
  represented as `SourcePayload::Owned` for managed input; `Managed` and
  `OwnedBudgeted` still use the existing fallible copy route
  (`1702-1731`). Production edit admission is capped at 256 spans, below the
  lower-level 4096 bound.

## v3 delta review

The v2-to-v3 delta is compile/test plumbing only: the facts oracle supplies an
explicit action map, the snapshot test uses the source-backed row-visibility
editor, `MAX_REPLACEMENT_SPANS` is re-exported within the raw edit facade, and
empty boxed slices/action types are made explicit. These changes do not alter
production XML proof, OPC publication, XLSX hint creation, or ownership
behavior. I found no new semantic issue in v3.

Advisories remain non-blocking: malformed hints can spend bounded clone
admission before gap validation triggers fallback; dropping that search state
before the fallback audit would reduce peak memory. The OPC quality fixture
export issue is test harness plumbing and does not affect production behavior.
