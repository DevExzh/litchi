# Change 0457 native ODP fixture oracle

This directory contains a read-only fixture inventory and an independent
Python oracle for a source-backed ODP append. The oracle does not import the
Rust producer, invoke LibreOffice, rewrite either archive, or make a native
application claim. The separate Rust semantic reopen probe checks the result
through the existing library reader; it also makes no native application claim.

`staticinventory.json` records ten local LibreOffice ODP archives selected
from the ignored `3rdparty/libreoffice-core` checkout. The inventory was taken
at repository revision `48f11c7745f8f798a00e11a14f1f34dc1c9c7954`.
`configure.ac` identifies the checkout as LibreOffice `26.8.0.2`; it has no
nested `.git` metadata, so each archive is pinned by path, byte count, and
SHA-256. Where the archive carries a `meta:generator` value, its embedded
LibreOffice project commit is recorded too. The bounded search found 330 local
`.odp` paths below `3rdparty/libreoffice-core` and none below `3rdparty/poi`.

The retained upstream notices in `upstream-notices/` are byte-exact copies of
the vendor `COPYING`, `COPYING.LGPL`, `COPYING.MPL`, and `README.md` files
supplied with that checkout. The fixture bytes remain unchanged; any output is
derived only by the requested append operation. The notices document upstream
attribution without making a single license claim for every fixture.

The nine entries whose `office:presentation` starts with a contiguous direct
page run followed by `presentation:settings` are static scanner-shape leads.
`tdf169979.odp` has six direct header/date/footer declaration elements before
its pages and was marked as a scanner refusal in the initial inventory.
That classification predates support for leading declarations; the retained
probe receipts determine the actual outcome of the final implementation. All ten are
non-encrypted, have no signature-named member, begin with stored `mimetype`,
remain below the package/member limits, and contain extra members beyond the
canonical six-member set. These are package/XML facts only.

To validate a producer output, provide the untouched fixture as `--source`
and the producer-created archive as `--output`:

```text
python3 docs/performance/results/change-0457/native/verify-output.py \
  --source 3rdparty/libreoffice-core/sd/qa/unit/data/odp/trailing-paragraphs.odp \
  --output /tmp/native-output.odp \
  --title 'Native title' \
  --body 'Native body'
```

Use `--name page2` when the producer was asked for an explicit generated page
name. Exit status 0 prints a JSON `validated` result. Exit status 1 prints a
validation error; status 2 is a typed `refused` result for an archive shape
outside this oracle's safe envelope, such as ZIP64 records or an unsupported
generated data descriptor.

The oracle independently parses ZIP bytes and XML. For every untouched ZIP
member it checks the complete local span, compressed payload, compression
method, local metadata, central record with only the four-byte local-offset
field masked, central member order, local member order, and archive comment.
It independently decodes stored/Deflate payloads and checks CRC and sizes. It
then parses `content.xml` with Expat byte offsets and requires the output to
have exactly one additional direct `draw:page` at the source last-page byte
end, with the exact requested plain title/body roles and text. The source
content prefix and suffix are compared byte-for-byte, and source/output
archive and content hashes are emitted in the result. Producer-reported
hashes or positions are never trusted.

No fixture is copied back, rewritten, or mutated by the inventory or oracle.

## Retained-capture verifier

After the successful candidate build, retain and bind the probe before its
first invocation:

```text
python3 -B docs/performance/results/change-0457/native/bind.py \
  --build-receipt docs/performance/results/change-0457/checks/candidate-build.json \
  --binary tools/perf-baseline/target/release/odp_native_tail_append_probe \
  --retained-binary /tmp/litchi-goal-0457/candidate/odp_native_tail_append_probe
python3 -B docs/performance/results/change-0457/check.py --tag native -- \
  python3 -B docs/performance/results/change-0457/native/run.py
```

Both commands refuse to overwrite existing captures or bindings. Failed
attempts must be retained and any retry must use separately identified paths.

`verify.py` authenticates the runner capture and replays retained archives
without the original fixture, binary, or repository checkout:

```text
python3 -B docs/performance/results/change-0457/native/verify.py --precleanup
python3 -B docs/performance/results/change-0457/native/verify.py --portable
```

`--precleanup` additionally checks the bound native binary, every original
fixture, the runner argv/cwd, and the output files while `/tmp/litchi-goal-0457/native`
still exists. `--portable` copies the complete change bundle to a private
temporary directory, extracts each bounded retained `.odp.gz`, and runs the
copied `verify-output.py` against those extracted files. The two flags are
mutually exclusive. The ordinary no-flag pass is the post-cleanup bundle check
used by the copied verifier.

The frozen `binary-binding.json` carries the executable identity, the
`inventory_sha256`, `oracle_sha256`, and `runner_sha256` bindings, and refs to
the completed `build_receipt` and `source_manifest`. The build receipt and
source manifest use the standard `check.py` schema and are known before the
native capture starts. The capture itself is run through the fixed parent gate
`../checks/native.json`; its standard `check.py` receipt is authenticated by
the verifier's `driver_sha256`, command, canonical `cwd`, source-before and
source-after identities, unchanged-source flag, and hashed log. Its final
digest is intentionally not prebound because `runs/index.json` binds the
frozen binary binding before that receipt exists.

The final `../SHA256SUMS` at the change bundle root covers every retained
bundle file except itself, including the native run gate. The verifier checks every retained receipt's
compressed and decoded hashes, immutable source hash, probe JSON
(`source_slides`, `target_slides`, generated name, insertion offset, and
content byte counts) against the independent oracle JSON, then reruns the
oracle from a private temporary directory. Probe semantic reopen remains
separately identified as the native probe's recorded fact.
