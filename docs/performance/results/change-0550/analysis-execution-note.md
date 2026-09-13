# Analysis execution custody

The profile agent created a preflight report before the root execution
receipts were recorded. It is preserved as `profile-analysis-preflight.json`,
SHA-256 `64ee681a71674696e65f3d36922a03dec950e8af5d2e49ecf00b1c65c5dfd0ae`.
Root runs both canonical analyzers after freezing their complete dependency
hashes; the canonical profile report must reproduce these bytes exactly.
The preflight is not substituted for a root execution receipt.

The profile agent also generated sixteen deterministic annotation sidecars
(eight inclusive trees and eight self trees) during preflight. Their commands,
environment and byte hashes are bound in the report. Root canonical replay
recomputed and checked every sidecar without changing it. Both canonical
analysis receipts passed, and the profile report exactly matches the retained
preflight bytes.
