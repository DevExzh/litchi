## Performance smoke comparison

- Mode: `fetched_reference` — regression comparison against the last successful full run (11111111, schedule)
- Detects regressions: yes
- Reference: run `11111111` at revision `7082a1a3f480589c0025aad8925cae315a4cfd4b`
- Comparator status: `invalid`
- Outcome: `reference_environment_drift`
- Enforcement: `advisory` (advisory)

Detail:

- build identity mismatch for 'cpu_model': 'Intel(R) Xeon(R) Platinum 8370C CPU' != 'AMD EPYC 9R45'
