# Change 0741 legacy-byte fixture

This isolated packet generator captures a pre-change owned cross-slide copy
using the current baseline implementation. It creates a one-slide source with
one relationship-free PNG leaf, repacks the source through the public OPC
physical writer with every member Store-compressed, and records:

- `source.pptx`
- `destination.pptx`
- `forward.patch` (LPCP0003)
- `inverse.patch` (LPCP0003)
- `expected-target.pptx`
- `MANIFEST.tsv`

Run it from the repository root before production edits:

```text
cargo run --release --manifest-path \
  docs/performance/results/change-0741/legacy-fixture/Cargo.toml -- \
  docs/performance/results/change-0741/legacy-fixture/captured
```

The generator applies the serialized forward patch to a fresh destination,
then applies the serialized inverse patch and checks byte-for-byte restoration.
The source image is checked through `PhysPkgReader::blob_for_borrowed`, which
only succeeds for a validated Store member.
