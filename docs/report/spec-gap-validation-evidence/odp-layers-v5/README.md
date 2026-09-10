# ODP layer metadata and reversible edits

This batch exposes bounded layer inventories and source-bound layer edits for presentation content and master pages. It preserves opaque markup and source bytes on exact no-ops, validates layer references and owner order, and publishes atomic reversible changes. Master-page identity uses validated `style:name`. Unknown unqualified children remain opaque; unresolved prefixes and malformed document boundaries are refused.

Root validation on `159ff704c` plus the exact ten source paths passed 404 tests across 29 targets, with zero ignored, strict all-feature/all-target Clippy, warning-denied rustdoc, and scoped formatting. Independent V5 review is CLEAR and verifies the earlier pending feature is retained. V3/V4 findings drove XML name, owner-order, root-closure and pre-growth quota corrections.

Limits on layer sets and layers apply per scanned XML part; no recoverable-OOM, rendering, or native-producer acceptance claim is made. The feature matrix describes the supported metadata/edit scope.

`publication.json` binds source hashes and exact gzip-preserved author/root/review artifacts. `primary-preimages.tar.gz` retains the pending bytes checked before integration. Earlier immutable candidate and review paths remain identified in the manifests.
