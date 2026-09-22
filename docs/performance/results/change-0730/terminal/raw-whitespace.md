The full staged `git diff --cached --check` reports the required single-space
context lines inside the archived unified patches and final blank lines in raw
Cargo output. Those evidence bytes are intentionally preserved. The same check
with only packet `.patch` and `.log` artifacts excluded passes; production,
tests, scripts, and authored documentation have no whitespace errors.
