# 0362ugk — PPTX lazy edit entry

This source-only repair moves only the primary PPTX edit entry point to the
edit-only lazy state required by the 0362 final sweep. `Package::edit` and
`edit_with_options` retain the cheap OPC graph, limits, owner policy, and the
package's payload-digest memo sink; they no longer capture the opened
presentation, scan the slide catalog and notes, build names, or fingerprint
the complete package at construction. Dropping an untouched edit therefore
does no presentation capture. The first source getter, semantic getter,
mutation, or commit materializes the existing validated capture path and
returns its ordinary typed error. A failed capture restores the deferred
state, so retries do not observe a partial edit.

`Package::opened_presentation*` remains the explicit eager, fallible read
projection. `Snapshot::edit` now also keeps its package graph, slide catalog,
and name index behind shared ownership: cloning the edit root is O(1), and the
first mutating operation obtains the private OPC graph with copy-on-write.
The edit's source, slides, and retained-MCE getters are fallible where
materialization is required; every direct consumer (including masters,
layouts, themes, SmartArt, notes, pictures, text, and hyperlinks) resolves
that state before reading it. Releasing retained MCE before materializing is
remembered and applied to the captured snapshot. The existing digest memo is
projected back to the package's own payload allocations, preserving the memo
fast path for later edits and replay. Signed admission, no-op publication,
patch/replay behavior, source provenance, owner policy, and output limits stay
on the existing capture/commit paths.

The constructor/drop regression uses the test fingerprint log to prove that
entry and drop do not capture, while the first `slides` use does. No Cargo
build was run in this isolated leaf; the coordinator must run the PPTX tests,
workspace Clippy/rustdoc, the API glossary, and the full gate after merge.
