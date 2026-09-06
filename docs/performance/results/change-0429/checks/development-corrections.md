# 0429 development corrections

Before execution, source review corrected provider baseline placement so
source handles are absent when their diagnostics are unavailable. The native
lifecycle replaced extra strong source references with Weak owner checks and
moved its selected phase after one timed metadata query. File staging moved
before the warmup loop; each measured iteration opens fresh file handles.
Bytes/file providers have no configured delay; an explicitly configured zero
range delay remains distinguishable. Request-size histograms were added to
meet the range-source distribution requirement.

The first `harness-range-tests` command failed compilation with seven emitted
diagnostics from three causes: private budget fields accessed by the sibling
observer, ambiguous sum types in a histogram test, and non-exhaustive guarded
native CLI matching. The original receipt and log are retained. The fixes
expose the two tool-only fields at sibling visibility, specify `sum::<u64>`,
and complete the valid provider match arm. Verified commands use distinct tags.

The original/LibreOffice shapes proposal was corrected before formal capture.
Source review found both inputs unsuitable for a positive image run under the
current refusal policy. Their static inputs and oracles remain, and focused
runtime refusal tests accompany the unmodified POI positive fixtures. No
native input was rewritten to satisfy the benchmark.

Native positive runtime testing failed with ZIP `BufferTooSmall`. Four focused
ZIP integration regressions then failed on the original implementation; the
bounded refill fix makes all four pass. The first test-only compile attempt
used two unavailable accessor methods and is retained separately from the
red runtime tests. Production and consumer checks follow the correction.

The first ZIP strict lint command found test-only `io::Error::other`, newer
than this crate’s Rust 1.70 MSRV. The test now uses `Error::new(InvalidInput, …)`;
the failed log remains. Expanded tests cover an exactly buffered empty-name
central record and ZIP32/ZIP64 extra fields under capped reads.

Driver source review added a pinned replay-verifier digest and explicit checks
of the complete tracked/untracked Rust/TOML/lock source manifest before build
binding and before/after formal capture. The source-custody driver is itself
pinned. CLI smoke reports are retained as preflight controls, not included in
formal sample counts or promoted into independently replayed baseline results.

The first release CLI preflight produced its plain-byte report successfully,
but the independent verifier rejected the generic corpus manifest as a source
archive identity. The existing cross-copy generator's generic manifest names
the destination archive; the report separately records both source and
destination hashes. The verifier is corrected to bind that manifest to the
destination, while preserving both explicit identities. The rejected report
and failed preflight receipt remain separate from formal captures.

The next CLI preflight reached mutation controls, where the histogram probe
selected the unavailable baseline's null histogram. That probe now selects an
available histogram before corrupting its count. This is a mutation-driver
correction; the report's unavailable baseline remains null as required.

After formal capture, supplementary CPU recording completed for media-rich
bytes, but the frozen verifier rejected the protocol's 100-sample profile
because its count policy admitted only controls and formal baseline rows.
The explicit validation amendment retains the original build/verifier/policy,
pins unchanged capture artifacts, and permits (100,3) only for the two declared
media-rich bytes/file profile roles. All 32 baseline reports are revalidated;
none is recaptured. `resume-profiles.py` processes the existing completed bytes
recording, retaining the failed wrapper receipt and marking reconstructed
command provenance, then records the file profile. The original recording's
start timestamp was not retained by the failed wrapper and remains unavailable.
