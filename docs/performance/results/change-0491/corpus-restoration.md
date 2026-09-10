# Reproducing the v1 corpus restoration

The existing producer changed in commit `a260174e4`: namespace-aware section
serialization added a local `xmlns:w` to a fresh final section. It repeats the
same namespace on the containing document root. This adds 71 logical XML bytes
and two compressed archive bytes. The benchmark's historical v1 pin was not
updated consistently with its generator's unit test.

The initial failed pilot is retained. The unpinned edit diagnostic's catalog
identifies the current generated source; it is not cold performance evidence.
`corpus-diagnostic.rs.txt` reproduces that producer using the actual compiled
DOCX dependency and writes a diagnostic archive. The recorded emitter build
and run gates retain exact commands and source manifests. Its archive hash
matches the current full harness output exactly.

`corpus-legacy-diagnostic.rs.txt` then reads that diagnostic archive through
`OpcPackage`, removes only the exact section-local declaration, and writes it
through the existing OPC writer. The resulting SHA-256 is exactly the original
v1 pin. `corpus-restoration-proof.json` compares all 20 members in source order,
checks their CRCs and bytes, and records the sole changed main-document member.
The non-media current XML and full current member hash inventory are retained.
Both diagnostic sources use explicit output arguments; their intermediate
executables and DOCX files are scratch and may be regenerated from the recipes.

The integrated benchmark generator admits only the exact historical archive
or this exact newer producer archive. It restores the known spelling and
requires the complete old archive hash afterward. Any further producer drift
fails closed. It does not change production serialization or ordinary saves.
The existing deterministic corpus/edit/reopen test now requires the original
v1 hash, size, logical bytes and member count; an altered archive is rejected.

For `aligned-corpus-oracle.json`, take the exact restored archive, verify its
final 22-byte EOCD has an empty comment, set its two-byte little-endian comment
length to `(-archive_length) % 4096`, and append that many zero bytes. On this
host this adds 564 bytes, yielding 16,793,600 bytes with SHA-256
`d1e6f59e6c6c698aa91463a2ed341351da3b549b818122bc0a2b9b9b8449dc88`.
All 20 decompressed members must remain equal and pass CRC verification. This
is an independent expected identity for the existing cold verifier's private
aligned copy; it does not claim that the copy was actually cold.

The first legacy diagnostic build used an incorrect rlib filename and failed;
its retained receipt is superseded by the successful `corpus-legacy-build2`
gate. No failed attempt is used as executable evidence.
