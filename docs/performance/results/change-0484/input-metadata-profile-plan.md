# Input metadata profile

profile_input_metadata.py is a bounded diagnostic helper for the frozen
normal route executable. Its default inventory has four child processes:

- s131072-a64-short-c64 with the deterministic owned input;
- s131072-a64-short-c64 with the prepared file input;
- s64-a16384-short-c64 with the deterministic owned input;
- s64-a16384-short-c64 with the prepared file input.

Each child uses measure_routes._axis_argv with one sample and one warmup.
The driver validates the resulting report with
measure_routes._check_axis_report, binds the normal binary to the explicit
build attempt, and records the input path, byte count, and SHA-256 for the
prepared file arm. The output tree is
input-metadata-profiles/<attempt>/; prior attempts are never replaced.

The primary profiler is strace -f -c with metadata, positional read/write,
open/close, seek, and sync syscalls in its explicit filter. The summary and
the route's report.json and resource.txt are required to be nonempty
before a receipt is marked ok. The helper reuses the content-bound
profile_routes._run_command, so timeout process-group evidence is retained.
Profiler refusal or launch failure is recorded as unavailable; failed
artifacts remain in place.

--raw-authored adds two extra authored-heavy children, one per input mode.
Those traces use the same metadata filter with timestamps and path decoding.
The raw file is SHA-256 hashed before gzip compression, and both bindings are
retained. This optional mode exists for source-path attribution and does not
add a performance claim.

The coordinator supplies the external gate and CPU lock. A representative
invocation is:

~~~text
python3 -B gate.py --attempt INPUT_METADATA_GATE route-input-metadata \
  python3 -B profile_input_metadata.py \
  --binary /home/zhuhe/.cache/litchi-goal-0484/0484/routes/formal1/normal/docx_replayable_tail_append \
  --attempt metadata1 --build-attempt formal1
~~~

The profile output attempt and frozen build attempt are independent,
path-safe tokens. The helper does not build, freeze, modify the prepared
fixture, or compare elapsed times.
