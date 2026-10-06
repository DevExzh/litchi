# 0930 source and ADR review

Status: approved. This was a static read-only review; no build, test, capture, or source edit was performed. `git diff --check` was clean for the five reviewed files.

The publication helper captures the current main XML and complete OPC artifact fingerprint, and reuses an inventory only when the XML bytes and every `Limits` field match. A post-read source-version fence and the existing patch checks for exact bytes, lineage, revision, artifact fingerprint, and limits run before output. Changed commits retain the base inventory only within the checked cap; overflow, over-cap, zero-cap, detached sources, and explicit release fall back to parsing. A no-op uses its already-held projected snapshot and still passes normal authorization.

Inventory accounting includes descriptor/column/header/footer/string capacities, source-span bindings and nested byte capacities, Arc headers, and tolerance report/vector/location payloads. It is logical accounting rather than RSS. The owner-charge documentation condition was resolved in the current `layout.rs`: it calls the owner a logical charge and separately excludes the always-present inline `Option<Snapshot>` slot from the conditional estimate.

ADR 0003 is satisfied for immutable/full-validated commits, exact-source authorization, and no-op reuse. ADR 0005 is satisfied for freshness, finite Limits, observable retention, release, and recomputation fallback. ADR 0031 is respected because managed cache-backed section-layout sources remain refused. The authoritative formal-gate correction uses baseline p50 divided by candidate p50, with values above one meaning candidate speedup; it is prospective pending the fresh cohort.

The machine-readable receipt is `source-adr-review-v1.json` in this external directory.
