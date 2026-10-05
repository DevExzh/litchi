# Unified-ops commit hashes after the 2026-10-05 history rewrite

The 2026-10-05 rewrite (see `AGENTS.md` and
`docs/performance/history-rewrite-2026-10-05.commit-map`) also rewrote the
`feat/unified-ops-api` and `uo/*` branches. The repository-wide commit map does
not list their commits. Records written before the rewrite (`PROGRESS.md`,
`m0a-log.md`, records 0002–0007, `results/0003/before/`) cite the old hashes.
This table maps each cited hash to its replacement; the commit subjects are
unchanged.

| Old | New | Commit |
|---|---|---|
| `59f75c7f79` | `fb039e1f7a` | fork point on `feat/office-format-completeness` |
| `fb433660eb` | `da4503a1ef` | docs: add unified operations API design and implementation goal |
| `cc3c793e5d` | `10eae2ec8c` | require read and write for every claimed format; open progress log |
| `333d5db3ca` | `77ff2f8490` | tier formats as OOXML, legacy OLE, then delayed RTF/ODF/iWork |
| `af70966a6d` | `9abd42bd56` | 0004: draft ADR 0033 and scaffold `litchi-ops` |
| `14054ce372` | `1e54ac20ee` | 0006: accept ADR 0033 |
| `cfe7681627` | `80327dde94` | 0002: leaf write-capability inventory |
| `45a932ea2e` | `a2e49fac5a` | 0003: M0a migration plan |
| `da4c207478` | `a8f45144c2` | 0007 defaults; the pre-M0a revision S1 measured (the baseline binary's file name keeps the old hash) |
| `7010055311` | `2f94cb1a7d` | M0a S2: unified error |
| `b9951ae15f` | `ac00b8a4f5` | M0a S3: `OpenOptions`, `Locked`/`Signed` shells |
| `9f8e49b181` | `5c5c1ea340` | 0005 (M0b): path/selector grammar |
| `7b897d11a2` | `cceb1570b6` | M0a S4: detection moved into `litchi_ops::detect` |
| `6e871e46bb` | `20c7c0b1a4` | per-format handle module stubs for S5 |
| `5f7c0523f4` | `d209c85d17` | M0a S5a: `Docx` handle |
| `6c96efff21` | `abf6e22e00` | M0a S1: pre-migration baseline |

The M0a step commits (S1–S5a and the stubs commit) are on `uo/m0a` and reach
`feat/unified-ops-api` only with the single M0a merge; every hash in the "New"
column exists in `uo/m0a`'s history.
