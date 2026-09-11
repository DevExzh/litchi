# Superseded source capture

These captures and passing reports/tests preceded the final two-line Clippy
cleanup in lib.rs. Strict Clippy found two existing `!is_none()` expressions.
The final source uses `is_some()` and is rebuilt/retested/remeasured in the
parent directory. These raw captures are retained for custody only and are
not the accepted final baseline. Original paths in receipts are historical.
The candidate patch and source manifest bind the earlier source.
