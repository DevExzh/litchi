# 0797 source-only review

Review basis: the frozen 0797 plan and README, the before/after helper archive,
the direct probe source, the mirror-test driver, and the custody scripts. This
review was source-only; it did not build, test, capture, profile, or replay
anything.

## Disposition

The candidate is semantically suitable for the root-owned 0797 preflight. It
does not authorize production adoption or public-workflow claims. The
production helper remains the 0797 baseline, and the candidate archive is
source-bound to the packet.

## Semantic review

The `ReplayFirst` boundary is now correct. After the unchecked first item is
returned, `next_first` enters `ReplayFirst` only when bytes after the borrowed
value contain a non-whitespace byte. `next_after_quick_xml` handles that phase
by constructing a fresh checked iterator, consuming its first item to seed
quick-xml's duplicate state, setting `QuickXml(1)`, and returning
`self.next()`. The returned call therefore takes the checked fast path for the
second item; it cannot fall through to `OwnCheck` before the 32-name boundary.
The source has the required `return` in this branch.

This preserves early duplicate ordering. A repeated second name is checked by
quick-xml before its value is parsed, including an unquoted, empty,
unterminated, or otherwise malformed value. A non-duplicate malformed second
item retains quick-xml's lexical error and position. The first item has no
prior name, so reading it with duplicate checks disabled cannot alter its
result. The replay uses the same immutable tag and discards the internally
consumed first result, so it does not expose that item twice.

`end_of` uses the borrowed value slice and the same one-byte closing-quote
advance as quick-xml. It handles empty values because an empty borrowed value's
end is still the position immediately after its closing quote. The
`trailing_whitespace` predicate uses quick-xml's four whitespace bytes. A sole
attribute, including trailing XML whitespace, becomes `Done`; any non-whitespace
tail takes the replay path. The helper tests cover the whitespace and empty
value machinery, although an explicit one-attribute `a=""` case would make
the preflight boundary easier to audit and is a useful retained coverage note.

The 32-name transition remains ordered correctly. The replay starts the checked
counter at one; calls two through 32 remain in quick-xml's path. The next call
turns duplicate checking off and builds the existing ordered-map state for the
33rd name. Errors in either phase set `Done`, and `FusedIterator` plus repeated
`None` checks cover empty, terminal, and error paths. Derived `Clone` preserves
the phase and the underlying iterator state; the archived tests and probe cover
the initial, replay, quick-xml, and own-check transition points, including the
error-followed-valid-tail case.

## Archive and testing scope

The canonical OPC candidate and the four copied modules are code-identical
after the documented visibility, test-path, and module-documentation
normalization. The shared test archive therefore checks both semantic parity
with quick-xml and copy parity. The direct probe compiles the exact OPC helper
modules and deliberately excludes `cfg(test)`; the separate mirror workspace
owns the five helper copies and shared tests. This is an appropriate boundary
for a source-only preflight and does not stand in for full production-crate or
public-workflow tests.

The first candidate mirror test attempt failed at the test fixture's
`any(Result::is_err)` function-item type mismatch. Root corrected the archived
test source to use an item closure, retained the failed logs under
`quality-failed-0/` and the exact failed mirror under `test-src-failed-0/`, and
reran the candidate gates successfully. This is test-fixture evidence, not a
candidate semantic failure; the failed attempt remains part of the custody
trail.

## Protocol audit

The frozen schedule is internally consistent: 33 cases produce 792 native
children and 23,760 native samples, plus 264 Callgrind children/samples, for
1,056 reports and 24,024 measured samples. The independent literal fixtures,
binary case catalog, capture order, analyzer labels, native audit, and final
validator agree on the same case order and counts. Native p50s, paired ratios,
and bootstrap ranks are fixed in the plan. Callgrind data is explicitly
guest-counter diagnostics with conservation and owner-qualification checks;
the scripts do not turn it into a native-latency or allocator-call claim.

The preflight decision is correctly narrower than adoption: semantic equality
is required, distinct-one consume improvement and its confidence bound are
required for advancement to fresh workflow trials, and consume regressions are
diagnostic blockers for that advancement. Construct flags remain diagnostic.
Failure archives the candidate and retains the production baseline; success
still does not change production or make a public-workflow claim.

No source-level blocker remains for root's serialized capture and offline
analysis. Any result must retain the explicit opaque-construction and replay
overhead limitations described in the packet.
