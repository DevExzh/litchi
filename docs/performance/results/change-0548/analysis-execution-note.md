# Analysis execution scope

All build and measurement processes terminated before canonical analysis began.
The candidate quality command stream ran alongside deterministic analysis; no
native, guard, allocator or profiler capture overlapped either activity.

The analysis-runs/*/inputs.json files are broad before-command bundle inventories.
They record incidental quality logs and verifier drafts as observed at that instant;
those files are not numerical analyzer inputs and were not asserted immutable.
Actual consumed reports, raw captures, source manifests, binary identities, plans
and analyzer scripts are separately validated by each canonical analyzer and the
strict verifier. Per-command receipts retain output hashes and exit status.
No numerical report or raw capture was replaced or deleted.
