# Native resave evidence helper

This isolated Cargo project creates small, deterministic Litchi changes and
reopens the resulting Office artifacts through their format APIs. Enable only
the formats needed for a run:

```sh
cargo run --manifest-path tools/native-resave/Cargo.toml --locked \
  --features doc,docx,xls,xlsx,ppt,pptx,rtf,odt,ods,odp,odf,odg,odb -- \
  generate FORMAT OUTPUT
cargo run --manifest-path tools/native-resave/Cargo.toml --locked \
  --features doc,docx,xls,xlsx,ppt,pptx,rtf,odt,ods,odp,odf,odg,odb -- \
  readback FORMAT INPUT
```

`tools/native_odf_resave.py` performs isolated-profile CLI resaves for formats
with export filters. LibreOffice's ODB filter is import-only, so ODB evidence
uses `uno/UnoStore.java` against a separately launched, isolated-profile UNO
endpoint and invokes `XStorable.store()` on the same package. The helper opens
the document hidden and forces `MacroExecMode.NEVER_EXECUTE`; it never opens a
database connection or executes a query, form, report, or macro.

Exact runtime versions, commands, hashes, and retained artifacts are recorded
in `test-data/office-interop/PROVENANCE.md`.
