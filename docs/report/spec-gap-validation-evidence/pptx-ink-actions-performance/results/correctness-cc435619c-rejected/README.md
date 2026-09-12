# Rejected cc435 correctness capture

This retention bundle preserves every regular file from the fresh
cc435619cfe96264c40889e6012f139fb8dcff9e capture results and the one separately authorized
representative-lane diagnostic. It is deliberately REJECTED AT THE CORRECTNESS MATRIX: preflight, release build, binary hash, and host probe passed; the
matrix exited 1 after reporting failures for all 42 lanes; the fail-closed
operator skipped all subsequent individual-lane commands.

The matrix failure was systematic accounting evidence. The representative
package_read_tiny_shared diagnostic had semantic, preservation, inverse, and
baseline-reopenability checks true, while retained_baseline_balance_ok was
false: the retained baseline was 41,788 bytes and after-drop live bytes were
41,880 bytes. The 92-byte delta came from Prepared lane/recipe clone
allocations made after the baseline snapshot. This bundle makes no claim that
all 42 lanes are resolved.

raw/ contains the 38 capture receipt files
(4921714 bytes), and raw/diagnostic/
contains the two diagnostic stdout/stderr files. The complete file list and
SHA-256 values are in retention-manifest.json. The release binary and its external Cargo target were retained through the
root Git-byte audit; the target contained 1015 files and 356994119 bytes. The binary SHA-256 is
c68b101dd2e46080da30a25415ec1d72a63186f123e19c4989f30cc897fbd6a4.

The source manifests and Cargo metadata match before and after. No /usr/bin/time,
timing matrix, native PowerPoint acceptance, or speedup claim was produced.
Root verified all 40 raw files against committed Git bytes in `d2cc59bdd`,
then removed the duplicate external capture, diagnostic, and build target.
The clean source worktree remains available. A post-fix capture requires a
fresh target and results directory.
