# Closure draft path correction

The first direct `close_v2.py verify` reached the decision binding and rejected
`decision analysis binding changed`. The initial decision descriptors stored
relative paths while the closure requires absolute paths. `decision-draft.json`
retains that original draft. Only the four descriptor paths in `decision.json`
were made absolute; byte counts and SHA-256 values are unchanged. No source,
measurement, reader, admitted file or workload was changed or rerun.
