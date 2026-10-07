# 0956 reducer v4 adapter review

Reviewed `tools/reduce_0956_v4.py` as a read-only diff against the frozen reducer v3. Exact file:

- path: `/home/zhuhe/litchi-performance-0956/tools/reduce_0956_v4.py`
- bytes: 21,168
- SHA-256: `a3589d6b83350284c2df3a640193f8bff940b27e1bdd909e98d7dce325d86d30`

The diff changes only the reduction schema/error label and the terminal deadline compatibility check. `start.json` must still contain the expected `deadline_seconds` value. For `terminal.json`, an absent `deadline_seconds` is accepted to match the sealed v3 capture packet; if the field is present, it must equal the expected value. All other v3 identity, artifact, validation, report-scope, metric, pairing, bootstrap, and ratio logic is unchanged.

This is the correct narrow adapter for the preserved packet and does not rewrite raw evidence. Approval is conditional on the root-side binding of this adapter SHA to the preserved packet manifest and final plan/source/build descriptors. No workload or reducer run was performed during this review.
