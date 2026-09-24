# Rejected profile preflight

The release harness compiled. Candidate preflight rejected reference-control-average: the broad reference-control prefix incorrectly routed statistical controls through the metadata validator, whose unknown-case result was NaN. The fixture-derived average is -0.03125. Commit 64e1aaf8fd restricts metadata controls to their five explicit cases without changing expected statistical values or production source. No timing samples ran.

Raw logs and target cleanup receipt are retained alongside this note.
