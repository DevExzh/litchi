# 0362ul53 — merged constructors worktree cleanup

Date: 2026-10-24. Cleanup wave 55 removes only the inactive
`uo2-coordinator-constructors` checkout, without force. Its branch
`uo2/coordinator-constructors` remains at
`035595049a95da85b0dfeefb37ae4611ebf2851e`, an ancestor of the held integration
`267ea715266ffc3bd754bfbfa4896b139adad9fc`. The tracked tree and index were
clean, and no process held a working directory or open file in the checkout.
Ten bounded JSON receipts in the current IWA qualification packet contained
no executable binding to this checkout.

All 21,662 ignored/untracked regular files were moved to quarantine on the
same filesystem, preserving their bytes, modes and inodes. Their logical
size is 19,569,961,202 bytes: 21,645 target files and 17 other files. The
retained target is not discarded. The restore helper passed both a dry run
and an actual sparse hardlink restoration; the temporary restore worktree
was then removed without force.

The former checkout occupied 25,568,145,408 allocated bytes. The complete
quarantine, including its manifest and restore receipts, occupies
18,032,373,760 allocated bytes. Their difference is 7,535,771,648 bytes
(approximately 7.54 GB). This is checkout-allocation accounting, not an
isolated filesystem free-space measurement; other work was running.

## Evidence

External quarantine root:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave55-coordinator-constructors/`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `manifest.json` | 6,111,936 | `c2e04e67c42641c182c6304ba5ec8939f994dfdb1e5bf353389fb4558f413bf7` |
| `restore.py` | 4,371 | `5cce92ef0af4624b72184f88e177e755abc7f2fd9bfca92864bf41c5ee591325` |
| `restore-check.json` | 1,500 | `1f7be52f8d4891f20ea3ca791eea41cc45495e766b24ead7f1a3624fbe002743` |
| `receipt.json` | 4,381 | `038e46fd948bdb0367537150134acba507f2e2340d3c5e4e0519854473990e63` |
| `manifest-header.json` | 501,228 | `e4f68e65c7f2a46bd56e1b4781593b7672ad8e4bb7dcdb373746df1fd41daaf9` |
| `rebuild_manifest_from_payload.py` | 4,297 | `adb0f26d294540f74f5324eb751e792bda44c3b72549c3bc9cae27d4aafece8d` |

The manifest is larger than 1 MiB and remains outside the repository.
Recovery uses the retained branch and `restore.py`; its command-line usage
is available with `python3 restore.py --help`. The recorded dry-run and
actual-restore results bind the helper and all payload records.

Regenerate the large manifest from its retained metadata and payload:

```sh
python3 /home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave55-coordinator-constructors/rebuild_manifest_from_payload.py \
  --header /home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave55-coordinator-constructors/manifest-header.json \
  --payload-root /home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave55-coordinator-constructors/ignored/coordinator-constructors \
  --output /home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave55-coordinator-constructors/rebuilt-manifest.json
```

The helper walks and hashes the retained payload, rather than deriving file
records from the original manifest. Its small synthetic payload test passes.
Root verifies that the retained header plus original records reproduces the
immutable manifest byte for byte; the full 19.56-GB regeneration is not
claimed. The final allocation above includes these provenance supplements.

Root independently verifies the four artifact digests, every payload's
size/mode/inode, all 17 non-target content hashes, branch retention, removed
checkout, held integration HEAD and protected 15-byte trunk constructor
probe. Target content hashes were verified by the cleanup worker; root does
not claim a second 19.56-GB content pass. Root's 2,476-byte review is retained
beside the quarantine as `wave55-root-independent-review.json`, SHA-256
`197dc67b4166a3b5b11b8f75028e7935519284486262fa6d3e94d50306925a5e`.
Its initial relative-path lookup error and corrected check are explicit.

No active target, source capture, native evidence, protected source root or
foreign process was removed. The integration hold is released. This cleanup
does not claim API, gate or performance completion.
