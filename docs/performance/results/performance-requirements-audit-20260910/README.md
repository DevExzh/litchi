# Performance requirements audit snapshot: 2026-09-10

This directory retains the exact generated report and row-level matrix used by
the 2026-09-10 audit of the `cf4f383c4` allocator publication. The snapshot is
historical: later commits and the committed normal-control and paired-scaling
publications do not change its row counts or status labels. It is evidence for
the dated audit figures cited by
[change 0470](../../changes/0470-performance-authority-reconciliation.md),
not a claim that the current performance program is complete.

The audit evaluated 264 rows: 50 CRUD checklist rows and 214 `GOAL.md` or
program rows. Status counts are 10 complete, 99 incomplete, 116 weak, and 39
missing. The audited target was `cf4f383c4c203381880d4c2e8445d2616a53d283`;
the measured allocator source was
`fccbe6595f3a29561a8bfc6c192d8fa5e874ad05`.

The source inputs were the supplied untracked `docs/GOAL.md` with SHA-256
`bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1` and
`docs/CRUD_Scenario_Checklist.md` with SHA-256
`d3a63f3aa001ad6e3dda7627b6294230ae87b3b5aaa595c45c3327e026dbf02a`. The
exact original `docs/GOAL.md` bytes are retained as deterministic
`GOAL.md.gz`; its compressed SHA-256 is
`7f2788c9db151750116fb809145c8b5cef96d1e1f199e6e521628d1c8f048bd9` and
decompressing it reproduces the input hash above. The
uncompressed generated artifacts have these identities:

| artifact | bytes | SHA-256 |
|---|---:|---|
| `performance-requirements-audit-20260910.md` | 10,242 | `5df805c0da7825cfe3338caf93a418b451cd20d8d9269d6694853dba9288dc28` |
| `requirements-matrix.json` | 153,630 | `22f774393d5c507def65d2c91de60415c88482b8c9f7ca1f8f20aa167e57f2b1` |

The repository stores deterministic gzip copies of those exact bytes and a byte-identical uncompressed requirements-matrix.json so the historical report’s relative link resolves. Verify
the compressed file identities with `SHA256SUMS`, then inspect them without
rewriting the originals:

```sh
gzip -dc performance-requirements-audit-20260910.md.gz
gzip -dc requirements-matrix.json.gz | python3 -m json.tool >/dev/null
gzip -dc GOAL.md.gz | sha256sum
```

The audit report's `29-file bundle` wording counts the 29 payload files listed
by the allocator bundle manifest. That inner tar archive has 31 regular files:
the additional `bundle-manifest.json` and `bundle-manifest.sha256` are its
self-authentication metadata and are intentionally outside the manifest's
payload list. The separate allocator publication directory (not this audit snapshot) is a six-file outer wrapper.

The report describes the narrow allocator publication as complete while
leaving the overall program incomplete. The matrix remains the authoritative
row-level detail for this dated audit; normal-control and later scaling
publication status are recorded separately in the authority record.
