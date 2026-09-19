# Initial ASCII-first candidate

This candidate checked is_ascii before the 31-unit bound. It passed all quality
gates and improved warm queries, but 54016 open regressed 12–16%. Assembly
showed unnecessary classification of overlong names, +1,039 code bytes and
+128 stack bytes. The next candidate checks the bound first. Baseline captures
remain in the parent packet; all captures here belong to the archived source.
The initial cold profiles are diagnostics, not a final performance claim.
