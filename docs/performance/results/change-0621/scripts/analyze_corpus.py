"""Differential over every XLS fixture: projection digest, I/O and observations."""
import pathlib, re, sys
S = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else ".")

def load(path):
    rows = {}
    for line in (S / path).read_text().splitlines():
        # Fixture names can contain spaces, so split on the operation token.
        match = re.match(r"^(.*?) (full-text|all-cells) (.*)$", line)
        assert match, line
        name, op, rest = match.groups()
        outcome = re.sub(r"\s+(reads|bytes|versions)=\d+", "", rest)
        numbers = dict(re.findall(r"(reads|bytes|versions)=(\d+)", rest))
        rows[(name, op)] = (outcome, numbers)
    return rows

before, after = load("corpus-before.txt"), load("corpus-after.txt")
assert before.keys() == after.keys(), "fixture set differs between legs"
same_outcome = same_io = counted = 0
observed_before = observed_after = 0
differences = []
for key, (outcome, numbers) in before.items():
    other_outcome, other_numbers = after[key]
    if outcome == other_outcome:
        same_outcome += 1
    else:
        differences.append(("outcome", key, outcome[:90], other_outcome[:90]))
    if (numbers.get("reads"), numbers.get("bytes")) == (
        other_numbers.get("reads"),
        other_numbers.get("bytes"),
    ):
        same_io += 1
    else:
        differences.append(("io", key, numbers, other_numbers))
    if "versions" in numbers and "versions" in other_numbers:
        observed_before += int(numbers["versions"])
        observed_after += int(other_numbers["versions"])
        counted += 1
print(f"cells compared: {len(before)} ({counted} reached the reader; the rest are typed refusals)")
print(f"identical outcome (projection digest or exact typed refusal): {same_outcome}")
print(f"identical read_calls and read_bytes: {same_io}")
print(
    f"observations over the corpus: {observed_before:,} -> {observed_after:,} "
    f"({(observed_after - observed_before) / observed_before * 100:.2f}%)"
)
for row in differences[:20]:
    print("DIFF", row)
