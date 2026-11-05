# 0362ul94sa — Ribbon authored-publication hardening

This follow-up starts from Ribbon repair commit
`a9a167b5c2ec27896928abc0ba17e721c2e53b22` (base
`3774fe4374dfb66e74b98b1ae742672ed6cad37f`) and tightens the authored whole-
XML OPC seam used by Ribbon graph replay.

The two explicit methods remain publication-only provenance paths:

* `OpcPackage::try_add_authored_xml_part_bytes(..., SourceLimits)` checks the
  member name, part bytes, part count, aggregate part bytes, XML publication
  audit, XML depth/events and per-attribute bytes before reserving or mutating
  package state;
* `OpcPackage::try_replace_authored_xml_part_bytes(..., SourceLimits)` applies
  the same replacement and aggregate checks, exact stale guard, and bounded
  audit before installing the replacement allocation.

Both paths reserve all destination maps before `revoke_exact_source`, retain
the exact allocation only in `source_xml_parts` for source-preserving output,
and never manufacture an `OwnedXmlPart` or `SourceXmlPart` range-splice token.
The strict source-splice namespace proof therefore remains unchanged. An
equal replacement is fully audited and then returns an exact no-op, including
on a signed package; stale, malformed, limit and signature failures leave the
package bytes, provenance and signature state untouched.

The integration regression exercises a real `ribbon::Edit::commit()` and its
patch replay and inverse, saves and reopens the changed package, and asserts
the escaped namespace bytes after every leg. It also checks stale patch
rejection, malformed and tighter-limit staging failures, and signed edit
atomicity. OPC-level tests cover XML attribute, part-count and aggregate-byte
admission, stale replacement, malformed replacement and signed no-op behavior.

The OPC minimal/full glossary profiles register both methods with their
required `SourceLimits` argument. The human glossary records that these are
publication-only seams and explicitly preserves the `SourceXmlPart` proof
boundary; no alias, deprecation, `_with_limits` twin or gate exemption was
added.

No Cargo, compiler, build, rustdoc, native or performance command was run.

## Earlier apply-check receipt

The exact earlier read-only check was run from
`/home/zhuhe/code/litchi-worktrees/uo2-ooxml-ribbon-94s`:

```text
set -e
cand=/home/zhuhe/code/litchi-worktrees/uo2-pptx-combined-0362ul89
git -C "$cand" status --short --branch
git -C "$cand" apply --check (deleted-coord-artifact)
```

Its complete output was:

```text
## uo2/pptx-combined-0362ul89
error: patch failed: crates/litchi-ooxml-common/src/ribbon/package.rs:662
error: crates/litchi-ooxml-common/src/ribbon/package.rs: patch does not apply
```

The command did not print the candidate commit, so this receipt establishes
only the branch label and the exact failed check above. A later read-only
query of that worktree reported commit
`cc2ab3f79d29ce32cbd6abdd27460dd0fec3f7ff` (tree
`aa002f5a33d4ca9df2a93577e6ad9d6032c89870`); no cause is inferred from
either observation, and the output is not attributed to an unverified base
or context transition.
