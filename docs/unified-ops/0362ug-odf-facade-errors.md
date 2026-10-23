# 0362ug — ODF facade error consumers

Status: source repair complete; compiled validation and native performance
qualification remain with the coordinator's integration lane.

This slice starts at consumer-v6 source `a814baa1de47c0844eb1ea092805b73e7d7bb4ae`.
The compiler receipt showed that the ODF package owners had moved their public
results to `litchi_odf_common::Error`, while the litchi-ops ODF facade still
sent those results through `litchi_core::Error` mappers. The affected consumers
now map each branch at its actual boundary: ODF package, limit, refusal,
conflict and source-publication errors retain their typed ODF leaf; source
owners and private scene/codec helpers that still return a core error retain
the core mapper.

The repair covers the ODF object tree formula, database, chart and image
families; ODP eager, source-backed and flat owners; ODS eager, source-backed
and flat owners, worksheet features and edits; the path source owner; and the
ODP/ODS typed-error tests. The ODP scene edit adapter and ODS feature/write
helpers preserve caller-side reclassification while retaining typed ODF
causes. The malformed-package tests now exercise the ODF leaf error directly
and assert the facade keeps it as `Source::Odf`.

The v6 compiler receipt remains outside the checkout:

* `/home/zhuhe/code/litchi-worktrees/.coord/replay-integration/api-assembly-consumers-v6/compiler-errors.json`
* 299,345 bytes, SHA-256
  `77134b691f7c50fe98d18255fad4bc228d74bdcc6249f3414c9778e629e978c2`

Direct rustfmt and `git diff --check` pass. Cargo was not run in this isolated
source-only lane; the coordinator will run the affected ODF tests and the
workspace checks after integration. This record makes no performance claim.
