# Independent audit correction

The initial reader stopped at its first historical identity comparison because it included `metrics.elapsed_ns`. Inspection found the tiny capture difference was 275332 versus 301901 nanoseconds. The protocol excludes elapsed values; the corrected reader omits only that field and continues comparing every semantic metric. The initial reader is retained in `root-audit-initial.py.txt`. No raw report, capture, or qualification requirement was changed.

While adding the final inheritance check, the validator initially prepended the `fn verify_output(` delimiter to a digest whose receipt covers the bytes after that delimiter. The check now uses the receipt’s original byte boundary; the 20,004-byte tail is identical to 0793. This checker correction changed no archived probe or build input.

The first post-commit replay rejected the documentation-only commit because `analyze.py` compared the complete current source manifest, including HEAD, to the baseline manifest. All production hashes matched. The final checker compares the complete production file map and requires the frozen revision to be an ancestor of HEAD; build and restoration receipts remain bound to the original revision. No output or source witness was rewritten.
