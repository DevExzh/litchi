# 0921 shared-owner quality receipt

Final six-file candidate review at `2026-10-06T07:13:15.675654Z`.

- Base: `bb19ca74bd52044fbd2c1c7cc9875e975bb334c0`
- Final patch: `22350` bytes, SHA-256 `7705c8ec7f12a5a5083b2ff29408451b0ee2afbbf8a25bb827c01b885aece67d`
- Candidate source: `/tmp/litchi-performance-0921/source`
- Final scoped Clippy: pass (`clippy-cfb-common-final.log`, SHA-256 `6e7c1f8bbe8512e420aa68fa6fc14a8bf8f0fa38c755cd6da5bc081a37feb5f6`)
- `git diff --check`: pass
- Candidate-only rustfmt: pass

The final source delta after broad testing contains only Clippy-driven lifetime/auto-deref cleanup and a test initializer cleanup. Therefore the broad receipts remain valid as behavior evidence without rerunning all suites: **5059 passed, 0 failed** across CFB 479, OLE common 199 + 30 object, DOC 1347, XLS 1704, and PPT 1300. The source-bound corpus covered **215 OLE2 files / 9139 expectations**.

The comparer is bound to the exact `OleFile<Cursor<&[u8]>>` source. Missing staged paths still receive structural and physical validation; malformed chains error before fallback, late ranges remain checked after an early mismatch, and short-tail bytes only zero-fill. Candidate publication remains atomic through the cloned editor.

The public adapter lacks a direct synthetic malformed/short-tail test, but the private generic oracle covers those paths and the public corpus passes; this is a nonblocking coverage gap with no observed semantic defect. Workspace format drift remains a baseline failure in unrelated soapberry-zip/iWork files.
