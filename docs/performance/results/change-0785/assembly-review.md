# Release resolver code generation

Both retained ordinary release binaries contain the same mangled `notes::resolved`
symbol. `assembly/receipt.json` binds each disassembly to its exact executable.
The candidate compiler folds the constant-length guards into a length-indexed
jump table (range 41 through 67 bytes). Six destinations select the corresponding
static namespace address and byte length, then converge on one `bcmp` call.
An equal comparison branches directly to the static-string success path; a
length miss or unequal comparison enters the existing decoding path.

This is one byte comparison for a recognized length, rather than a sequence
of six complete comparisons. The six byte lengths are 58, 46, 53, 41, 67 and 55.
The resolver grows from the baseline assembly, and unknown values pay the
length dispatch and, when lengths match, an unsuccessful comparison. This is
why the public vendor and Unicode fallback controls remain adoption gates.

The assembly supports the intended generated-code mechanism. It is not an
operation-level cycle attribution, a speedup estimate, or evidence about
other architectures. Public paired measurements determine disposition.
