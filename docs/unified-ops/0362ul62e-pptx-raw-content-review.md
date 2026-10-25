# 0362ul62e — PPTX authoring replay-index publication correction

`Package::new_authoring` creates the minimal package before the builder
materializes authored slides. The final `publish_authored` boundary now
refreshes the retained notes/media replay index after all slide, notes, and
optional font mutations complete. Ordinary edit and replay publication still
updates the index from affected-part deltas.

Builder regressions cover a changed notes replay while another notes slide is
unchanged, and a resource-only media payload replay whose slide XML remains
unchanged. Direct mutation seams were audited: font embedding is inside the
same authoring publication boundary; metadata and unsign edits do not change
notes/media relationships; staged and opened edits already carry delta-updated
indexes.

No generated evidence, fixture, executable, or `3rdparty/` path is included.
