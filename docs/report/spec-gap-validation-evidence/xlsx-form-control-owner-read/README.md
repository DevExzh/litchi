# XLSX form-control owner read validation

The current candidate is the six-file capture in `source-v6-final.json`. It adds read-only worksheet ownership for form-control properties, drawing shapes and VML mirrors through the ordinary eager and source-backed worksheet APIs. Collections retain source readsets and report ambiguous or unsupported mirror relationships through typed diagnostics. Owner mutation and publication are outside this batch.

Both worksheet facades expose `form_controls()` and `form_controls_with_limits(OwnerLimits)`, plus checked single-control lookup. The public types are re-exported from `litchi_xlsx::form_control`; callers can inspect ownership, properties, mirror diagnostics and source readsets without manipulating OPC identifiers.

The final correction charges diagnostic strings and vector capacities, reconciles collection capacity, and holds eager inventory reservations through scanning. The scanner projection ledger includes the retained eager inventory charge, so one aggregate `max_projection_bytes` ceiling applies even without an execution context. The existing inventory reservation is not charged to the execution context a second time. Public tests check exact and one-byte-under projection limits separately for each facade; their live allocations differ, so equal absolute minima are not required.

The root built the exact captured source in an isolated checkout and target directory. It passed 1175 library tests, 39 owner integration tests and 18 properties integration tests (1232 total, none ignored), all-target Clippy with warnings denied, and rustdoc with warnings denied. This is the named test selection, not a claim that every XLSX integration target ran. The captured hashes matched the built checkout and shared source after the gates. Commands were:

```sh
cargo test --locked -p litchi-xlsx --lib --test form_control_owner_read --test form_control_properties --offline
cargo clippy --locked -p litchi-xlsx --all-targets --offline -- -D warnings
RUSTDOCFLAGS="-D warnings" cargo doc --locked -p litchi-xlsx --no-deps --offline
```

Logs and exit codes are retained in `v6-*.log.gz` and `v6-exits.json`. Independent final review approves the aggregate correction; see `v6-review.json`. Earlier manifests and receipts are historical captures explaining the review corrections, and do not approve this candidate. Historical measurements under `xlsx-form-control-owner-performance` describe their own source capture and establish no performance improvement for this revision.
