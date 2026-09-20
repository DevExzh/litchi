# Superseded formatter grammar handoff

The retained candidate passed seven gates and 1,592 tests, but subsequent
independent semantic review found three TEXT grammar cases outside the
219-observation oracle: unquoted whitespace admission, missing internal
percent output, and repeated outer percent admission. The review is retained
here. These receipts describe the earlier source hashes only; they do not
establish acceptance of the corrected candidate. No timing conclusion is
drawn from these gate receipts.

`retained-files.json` hashes the archived gate and review receipts. Final
source custody and gates will be recorded in the top-level `gates/` directory.

`source/` retains all 25 selected inputs at their earlier frozen hashes.
Overlay them on baseline `8f09231e36982248eface4d599432143a67f6e49`
to reconstruct this superseded candidate, using its retained Cargo.lock.
