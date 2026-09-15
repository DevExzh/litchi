## Performance smoke comparison

- Mode: `fetched_reference` — regression comparison against the last successful full run (11111111, schedule)
- Detects regressions: yes
- Reference: run `11111111` at revision `7082a1a3f480589c0025aad8925cae315a4cfd4b`
- Comparator status: `regression`
- Outcome: `reference_regression`
- Enforcement: `advisory` (advisory)

Detail:

- opc_file_eager_open operation_metrics.allocation.allocated_bytes.values: 33824465.0 -> 37206912.0
- opc_file_eager_open operation_metrics.allocation.allocation_calls.values: 251.0 -> 276.0
- opc_file_eager_open operation_metrics.allocation.deallocated_bytes.values: 33824465.0 -> 37206912.0
- opc_file_eager_open operation_metrics.allocation.deallocation_calls.values: 232.0 -> 255.0
- opc_file_eager_open operation_metrics.allocation.live_bytes_after.values: 1337.0 -> 1473.0
- opc_file_eager_open operation_metrics.allocation.live_bytes_before.values: 1337.0 -> 1473.0
- opc_file_eager_open operation_metrics.allocation.peak_live_bytes_after.values: 33657344.0 -> 37023081.0
- opc_file_eager_open operation_metrics.allocation.peak_live_bytes_before.values: 4025.0 -> 4430.0
- opc_file_eager_open operation_metrics.allocation.reallocation_calls.values: 19.0 -> 21.0
- opc_file_eager_open operation_metrics.allocation.allocated_bytes.values: 33824465.0 -> 37206912.0
- opc_file_eager_open operation_metrics.allocation.allocation_calls.values: 251.0 -> 276.0
- opc_file_eager_open operation_metrics.allocation.deallocated_bytes.values: 33824465.0 -> 37206912.0
- opc_file_eager_open operation_metrics.allocation.deallocation_calls.values: 232.0 -> 255.0
- opc_file_eager_open operation_metrics.allocation.live_bytes_after.values: 1337.0 -> 1473.0
- opc_file_eager_open operation_metrics.allocation.live_bytes_before.values: 1337.0 -> 1473.0
- opc_file_eager_open operation_metrics.allocation.peak_live_bytes_after.values: 33657344.0 -> 37023081.0
- opc_file_eager_open operation_metrics.allocation.peak_live_bytes_before.values: 4025.0 -> 4430.0
- opc_file_eager_open operation_metrics.allocation.reallocation_calls.values: 19.0 -> 21.0
