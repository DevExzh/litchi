"""Reviewed, repository-local inputs for the ODG joined qualification.

The qualification manifest is external and therefore cannot be allowed to
choose a look-alike adapter or fixture merely by supplying a matching digest.
These bindings are changed only when the reviewed source inputs themselves
change, and every consumer checks both the fixed repository-relative path and
the approved bytes.
"""

from __future__ import annotations

from pathlib import Path


REVIEWED_REPO_ROOT = Path(__file__).resolve().parents[1]
REVIEWED_PROBE_RELATIVES = {
    "common": Path("tools/perf-baseline/odg_joined_qualification_common.rs"),
    "allocator": Path("tools/perf-baseline/odg_joined_qualification_allocator.rs"),
    "baseline": Path("tools/perf-baseline/odg_joined_qualification_ad524_probe.rs"),
    "candidate": Path("tools/perf-baseline/odg_joined_qualification_probe.rs"),
}
REVIEWED_FIXTURE_RELATIVE = Path("test-data/odg/joined-content.xml")

# These values are deliberately source bindings, not values copied from a
# caller's manifest.  They are refreshed together with the reviewed inputs.
REVIEWED_PROBE_IDENTITIES = {
    "common": {"bytes": 17863, "sha256": "cdbe48f55e51ba6a35285c0b0a321e36c4d07c616053be685c9865d79fa3b9c5"},
    "allocator": {"bytes": 10357, "sha256": "8f657e87e65dffcd10dae5906212ccf1b4e68e12d8bc2008c261e189ef1e815d"},
    "baseline": {"bytes": 11351, "sha256": "8b3adac40dee59156fa667d6478779f7c63a28c60c1253f4e0187140a57a5a40"},
    "candidate": {"bytes": 13180, "sha256": "ba4dab289b8786c83882925d3afb50ff131805a575fced4ea6c735d1f0256fbc"},
}
REVIEWED_FIXTURE_IDENTITY = {
    "bytes": 738,
    "sha256": "88776e2ebef02c6b5da1237e03dcf6fd8f6aad159d233c59758f240c1e9fa6c0",
}
REVIEWED_FIXTURE_PROVENANCE = (
    "authored for 0362ux by litchi ODG PackageWriter; no third-party input; "
    "repository Apache-2.0 license"
)
