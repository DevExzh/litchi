# 0897 source audit: restored XLSX full-cell fallback allocations

This is a read-only source audit for the producer-shaped XLSX fallback scan.
It proposes one bounded source-level mechanism for a later isolated change. It
does not edit production code, run a workload, build a candidate, or make a
performance claim.

## Scope and evidence binding

The evaluated tree is `c6d7ba22f525fbf82bbe7a7bcc1beb495169763b`, with
`506aa2edae6666d9c76ae6827c2f5941a8e47896` as the retained production baseline.
`docs/performance/results/change-0897/source-identity.json` records a passing
source comparison for the relevant build, crate, producer, test, and ADR paths.
The current source hashes that matter here are:

| path | SHA-256 |
| --- | --- |
| `crates/litchi-ooxml-common/src/mce/codec.rs` | `9bc4e4a6476318dce1dae6ea198847490112d1d5119fb773e388ab5d0b65a996` |
| `crates/litchi-xlsx/src/raw/worksheet/x14ac.rs` | `9f766274df5cc0a6a80291644ceecbc4f9d8afc6dac556af86bd21fbde66d148` |

The allocation evidence is the retained 0895 baseline pair:

* `docs/performance/results/change-0895/export-v1/runs/00-baseline/stacks.txt`
  (`7b022a98e061da2e5d2a8bbfd3b9f6544d506187c468f91657a311caea8f172f`); and
* `docs/performance/results/change-0895/export-v1/runs/03-baseline/stacks.txt`
  (`b2c415979f972a182ea6afae6207e201985dadde9bbda40e1a5bf1814e81f739`).

Each retained baseline process has 673,663 whole-process allocation events and
518,660 events in the attributed scan closure over 20 scans. The stacks are
heaptrack stack-presence attribution, so inclusive frame groups overlap and
cannot be added. The binary is the retained 0894 `baseline-native`; the source
identity and prior-capture verification make this useful for prioritization of
the restored path, but it is not a fresh native trial of a proposed change.

The actual fallback route is visible in the current source. `SourceWorksheet::stream_selection`
returns `Selection::Stored(self.store()?)` after `NotEligible`
(`crates/litchi-xlsx/src/workbook/source.rs:1034-1068`). `SourceWorksheet::store`
then calls `raw::worksheet::parse` (`source.rs:1180-1206`), and
`parse_with_repairs` calls x14ac capture, MCE processing, and the worksheet
parser (`crates/litchi-xlsx/src/raw/worksheet/mod.rs:468-485`). This is the
route represented by the producer full-cell fallback evidence.

These bounds follow `docs/GOAL.md` and ADR 0001, ADR 0005, and ADR 0006:
typed errors and resource limits stay authoritative, lossless lexical content
is retained, and an allocation hypothesis is not accepted without a matched
measurement and differential correctness evidence.

## Measured allocation sites

The following are measured stack groups, with current-source bindings. A stack
group identifies a call path; it does not prove that every byte or call on the
path is removable.

| retained baseline stack group | count over 20 scans | current source binding and interpretation |
| --- | ---: | --- |
| Dominant exact `slice::to_vec` / `str::to_owned` stack signature below `litchi_ooxml_common::mce::codec::start` | 167,220 (8,361/scan; 32.24% of the scan-closure count) | Two rows with the same exported stack text carry 166,800 and 420 calls; both must be counted. Heaptrack resolves the owning function but reports `codec.rs:0`; the current `start` has `Mode::Emit(q.to_owned())` at `codec.rs:950`, `1116`, and `1132`. The profile cannot distinguish those three branches, so this is a measured `start`-level signal, not an exact-line claim. |
| `quick_xml::events::attributes::IterState::check_for_duplicates` below `codec::start` | 85,560 (4,278/scan; 16.50%) | The current call is `e.checked_attributes()` at `codec.rs:877`; the checked iterator is `litchi_opc::xml_attributes::CheckedAttributes::next` (`crates/litchi-opc/src/xml_attributes.rs:222-240`). This is a measured validation allocation path, not yet an avoidable allocation. |
| `Box::clone_from_ref` below `litchi_xlsx::cell::Number::new` | 73,720 (3,686/scan; 14.21%) | `raw/worksheet/semantic.rs:168-184` calls `Number::new(value)`, whose `impl Into<Box<str>>` ownership is intentional at `crates/litchi-xlsx/src/cell.rs:388-398`: the exact numeric lexical form must survive. Removing it would require a semantic ownership redesign and is outside this fallback audit. |

The retained whole-process stacks include 24,615 allocation events with an
x14ac frame; the original disjoint nearest-owner classification assigns 9,792
events to x14ac. These are different scopes and must not be added. The inclusive
aggregate includes setup and capture/state work and does not isolate
`copy_bytes`, `copy_name`, or `copy_str`. It is a route location signal only.

## Primary hypothesis: retain the emitted QName as an output span

The highest-impact source-supported candidate is to stop allocating a second
owned qualified name solely to close an emitted element. `Mode::Emit(String)`
is declared at `crates/litchi-ooxml-common/src/mce/codec.rs:686-696`; the three
construction sites are in `start` at lines 950, 1116, and 1132. The start name
has already been appended to the same output buffer in both writer paths:

* the source-copy fast path writes `<`, `tag`, `>` at `codec.rs:1584-1594`; and
* the rebuilt path writes `<`, `q`, attributes, `>` at `codec.rs:1597-1632`.

After the `<` write, the qualified name is contiguous in both paths. A later
isolated implementation can have `write_start` return a checked
`Range<usize>` for that name, with the range computed as
`[prewrite_output_len + 1, prewrite_output_len + 1 + q.len())` and published
only after the complete start write succeeds. The range is an offset into the
append-only output, so a `Vec` reallocation does not change it. The writer must
keep the invariant explicit: the current codec never truncates the output
buffer while frames are live.

The frame mode should retain an offset only for an active emitted element. An
inactive emitted frame still needs its control-mode distinction, but it must
retain no name or range; an enum variant such as `EmitInactive` can encode that
invariant. This removes the current `q.to_owned()` even when the parent has
already made the frame invisible. An active frame must be constructed only with
a successfully returned span, so an impossible missing-span state cannot
silently omit a closing tag.

Closing must remain three separately checked output operations. The current
end path is `close`/`process_markup_compatibility` at `codec.rs:777-792` and
`close` at `codec.rs:1315-1338`. The candidate shape is:

1. `BoundedOutput::extend_from_slice(b"</")`;
2. validate the stored range, call the existing fallible `reserve` for its
   length, then call safe `Vec::extend_from_within(range)`; and
3. `BoundedOutput::push(b'>')`.

The helper must check `start <= end <= bytes.len()` before the `Vec` method,
check the length with overflow-safe arithmetic, and map reservation failure to
the existing `Error::Allocation { resource: "MCE output", .. }`. Reserving
before `extend_from_within` prevents that standard-library call from becoming
an untracked allocation. Keeping the three calls separate preserves the
current output-bound checks and their ordering at the `</`, name, and `>`
boundaries. No unsafe code or source-reader lifetime is required.

The candidate must preserve the current behavior: the closing QName comes from the start
QName, empty elements are expanded and closed through `close`, inactive and
skipped branches emit no bytes, namespace hoisting and attribute filtering run
before the start write, and all existing parser, output-bound, and typed allocation errors must remain
on their current paths. The range is an
internal view of already emitted bytes; it is not a borrowed source slice.

The candidate needs byte-for-byte MCE differential coverage for ordinary,
preserved, opaque, ignored, unwrapped, alternate-content, namespace-hoisted,
empty, and inactive elements. It also needs exact-limit tests at one byte below,
at, and above the admitted output, malformed/duplicate-attribute error-order
tests, and a matched native fallback trial measuring allocation count and bytes.
No benefit should be claimed before those checks.

## Lower-priority observations and constraints

* The 84,640 duplicate-check group warrants a separate measured experiment.
  `CheckedAttributes` uses quick-xml for the first 32 names and switches to an
  ordered map (`crates/litchi-opc/src/xml_attributes.rs:81-143`). A bounded
  inline scratch set could be considered only with exact duplicate positions,
  lexical error precedence, attribute limits, and iteration order preserved.
  Switching `codec::start` to unchecked attributes would weaken those contracts
  and is not a candidate.
* The actual x14ac fallback capture enters the legacy stream route when the
  worksheet contains MCE, `dyDescent`, and `AlternateContent`
  (`crates/litchi-xlsx/src/raw/worksheet/x14ac.rs:74-87`). Its raw observer
  copies four identity/value fields into `PendingCandidate`
  (`x14ac.rs:311-390`), and the active observer checks sequence, kind, qualified
  name, and expanded name before consuming it (`x14ac.rs:394-507`). A borrowed
  callback-scoped view is a possible later experiment, but it must prove those
  lifecycle checks, raw-before-active ordering, duplicate-marker errors,
  cancellation, and no source-buffer escape. The retained whole-process
  aggregate does not justify prioritizing it above the MCE name span.
* `Number::new` lexical ownership is required by the lossless cell contract;
  it is not an avoidable fallback temporary on the present evidence.
* A reader-borrowing redesign would change ownership and error/lifetime
  boundaries across `parse_with_repairs`; it is much larger than the output-span
  candidate and is not needed to test the measured `start` signal.

The previously rejected stream QName and selected-metadata candidates are
intentionally outside this audit. They must not be bundled into the output-span
experiment.
