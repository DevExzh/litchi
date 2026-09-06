# Source and verification review

The root works locally after earlier delegated agents exhausted usage limits.
The previous goal turn made progress by committing 0442. Current baseline is
f283c2cd9; the accepted ADR tree is unchanged from the prior complete read.
This batch examines the source-fragment scanner's measured ownership costs.

The old scanner copied each resolved event namespace into a Vec, including end
events whose namespace was unused, and copied start/empty local names. Open
frames retained those bytes solely for later fixed-name comparisons. The
candidate resolves each start/empty event into one of eight exact element
kinds. A frame stores that kind and its byte offset. It retains no borrowed
namespace across a reader advance. Unknown and unbound namespaces map to Other;
root namespace declarations and source XML still retain their complete bytes.

All state ordering remains: root attribute collection before on_open; style
collection before/after on_open in the original empty/start order; unchanged
depth/page bounds; on_close using the matching start frame; and unchanged
finish, style-name and root-binding logic. The reader still checks end names.
Spans, BOM handling, automatic styles, modeled declarations and opaque extras
remain under the original contracts. The scanner is a preservation projection,
not a new XML validity oracle. No public API, dependency, unsafe code, executor,
ambient I/O, source authority or commit/publication change is introduced.

The entire original scanner is retained byte-for-byte as scanner_reference.rs
behind cfg(test), independently verified against retained baseline source.
Three differential tests compare Debug projections containing every private
span, source byte, root binding and style name, or exact error diagnostics.
They cover 432 fragment/alias/BOM pairs, six style-container forms, 18 malformed
or fallback cases, and six depth/page/byte boundary cases (462 comparisons).
Existing real-producer preservation and LibreOffice/odfpy settings fixtures run
in the complete ODP suite. No fresh native rendering process is claimed.
ODP has no fuzz target; existing ODF targets exercise detection/ODT and facade
fuzz targets cover iWork. Those unrelated targets do not validate this scanner.
Differential tests are not coverage-guided fuzzing. No new low-level ownership
or concurrent code calls for Miri/loom.

CPU jobs and source switches are serialized. An unrelated host Chromium build
was observed while release correctness tests compiled. The experiment does not
claim an exclusive host; timing flags and temporal grouping remain relevant.
Frozen protocol, capture/profile drivers and the independent report oracle
remain unchanged after the baseline build. All attempts will be retained.

The protocol's frozen_utc label was inherited from 0442 during scaffolding.
This metadata error was noticed after A1 and correctness validation. The frozen
file is retained unchanged. freeze-review.json discloses the stale label,
observed file mtime before the first build, and the actual first capture time.
Mtime is supporting observation, not immutable timestamp authority; build and
capture descriptors bind the unchanged protocol hash. No threshold, lane,
measurement or selection rule was altered.

All 358 ODP tests pass with zero ignored; strict owner all-targets Clippy passes
with warnings denied. Both candidate repeats and the final baseline repeat
pass all archive/semantic/preservation/report gates. The practical allocation
call gate passes at 14.886%/14.943% medium/large reductions. The normal latency
gate fails. The keep decision accepts the disclosed sub-5% normal p50 costs
for deterministic allocation-call reduction. Requested bytes fall only
2.557%/2.663%; peak and retained live bytes are unchanged.

Two matched flags (medium allocator R1 p95/p99 +13.310%/+12.987%) and all eight
repeat flags remain retained. The instrumented tail effect does not recur in
R2, whose p95/p99 fall 2.478%/2.414%; it remains an unresolved observation.
No causal attribution to the separately observed host build is made. The
baseline tiny R2 timing shift precludes a useful tiny normal speedup claim.
No timing measurement is excluded, replaced or selected as a confirmation.

All four profiles pass with 13 addr2line warnings per conversion in both roles.
Whole-process cycles rise 1.516%, instructions fall 0.739%, branches fall 0.803%
and branch misses fall 9.036%. These include setup, warmups and Rust oracle
work, not operation-only causal effects. Zero L1 values support no cache claim.
The 1.0375x geometric mean of six matched normal p50 ratios is dominated by the
baseline tiny repeat shift and is not used to establish a normal speedup.
The generated measurements independently verify all 60 allocator observations
per role/shape agree for calls, requested bytes, peak and retained delta.

The complete harness passes 368 tests with one existing ignored test. The
non-iWork workspace all-targets/no-default-features check with litchi/odf,
ODF all-targets/all-features check, warning-denied owner rustdoc and pinned
three-file formatting check all pass. The changed Markdown links resolve.

The boundary audit passed for 64 packages/240 internal dependency declarations
with the same 14 explicit migration debt items. Final source review confirms
that all three files match the measured candidate, the original scanner remains
byte-identical, all four retained executables match their build hashes, and
the accepted ADR tree, prior sealed bundle and pinned GOAL.md are unchanged.

Portable copied controls passed before and after cleanup. All 11 precleanup
and 12 postcleanup independently corrupted copies were rejected after inventory
refresh, including the cleanup GOAL-digest mutation. Cleanup removed only the
four bound executables under /tmp/litchi-goal-0443-binaries (1,830,421,600 bytes),
preserving both build-cache directory identities and user-owned GOAL.md. Probe
copies cleaned themselves up. Compression and the inventory are resealed
before final read-only verification. No tests, checks or measured attempts
failed; the inherited freeze-timestamp metadata error remains disclosed.
