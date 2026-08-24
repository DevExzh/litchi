# Change 0270: OPC relationship-open timing boundary

Date: 2026-08-24

Status: corrected evidence boundary; no performance claim

## Scope

The `opc_relationship_open` evidence case now times only the production OPC
relationship-open operation. Source preparation, selector setup, and all
independent result construction stay outside the clock. The production
operation is the existing opened-package path; this change does not add a
public API, dependency edge, selector, or default case.

## Timing and oracle boundary

The production-open result is passed through a `black_box` fence before the
elapsed value is read. This keeps the returned package observable while
keeping the timer boundary around the production open itself. Relationship
and package correctness oracles run after timing and are not included in the
reported interval.

The post-timing checks remain the authority for correctness. They validate the
opened relationship state and the case's existing package/source identities;
the timer is not used as evidence for those checks.

## Claim disposition

This correction establishes a truthful timing boundary only. It does not
publish a latency, allocation, RSS, peak-memory, physical-I/O, decompression,
copy, or production-performance claim. No claim-registry entry or historical
classification table is changed.
