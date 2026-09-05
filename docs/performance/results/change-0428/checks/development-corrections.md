# 0428 validation and review corrections

These corrections happened before formal capture. Failed commands remain in
their original receipts and losslessly compressed logs; they are not counted
as passing checks.

- `harness-unit-final` failed because one configured-limit constructor omitted
  the output limit. Supplying the image source's 8 GiB output limit fixed the
  build; `harness-unit-verified` passed five tests.
- `harness-strict` found four added diagnostics: one large enum variant and
  three needless borrows. Corpus variants are now boxed during setup, and the
  borrows are removed. Subsequent strict logs match the retained 0427 baseline:
  29 findings in 17 message/source-file groups, none in the new module. The
  strict commands still fail because that pre-existing debt remains.
- Source review corrected the image admission floor. Reservation retry can
  evict metadata, so the exact cap uses the pinned presentation root plus the
  payload, with primed metadata usage recorded separately.
- Source review removed leaked dynamically generated phase labels and moved
  corpus and payload-oracle construction before all observations. Repeated
  publication uses fixed labels. Pinning compares image A at position 0 and
  image B at the corpus's selected position 3, matching the report's selected
  image identity.
- `cli-smoke` passed both lifecycle controls and their mutation probes, then
  stopped because the validator treated the publication hash vector as a
  scalar. The vector now has explicit element and scenario checks.
- Independent review rejected fabricated zero pinning read fields. They now
  carry measured metadata-to-A/B and drop-to-reload source deltas. Read
  windows absent from a scenario are null; process and phase source counters
  remain separate, observable values.

The preliminary debug reports and failed preflight directory are development
history. Formal counts include only the frozen release matrix in `capture/`.
