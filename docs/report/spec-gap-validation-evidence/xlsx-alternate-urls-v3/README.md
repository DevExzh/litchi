XLSX alternate external-link URLs
================================

This batch adds bounded inert alternateUrls metadata and source-backed editing with exact expanded-name owner selection, physical relationship target/type replacement and atomic publication. Input and relationship IDs are admitted before cloning/decoding; removing an owner with opaque content is refused. External URLs are not resolved or refreshed.

Root validation passed 1,430 tests across 63 targets including doctests, 0 ignored, strict all-feature/all-target Clippy, warning-denied rustdoc and scoped formatting. Independent V3 review exercised namespace decoys, oversized XML/IDs, opaque removal, physical relationship updates and collision atomicity. Independent integration review verified the existing refreshIntervals feature survives the shared exports/matrix merge.

Only the eight owned clean paths are staged. The exact pending MS-OREACTXML matrix row is restored to the working copy after staging; all other pending files remain untouched. The mixed snapshot's unrelated changes are excluded. The clean source is based on 52519031b; subsequent eff351997 changes only ODS. The retained root Cargo.lock is an explicit untracked build input, not committed source.
