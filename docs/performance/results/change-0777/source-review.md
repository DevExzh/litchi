# 0777 source review

Independent read-only review of production source 488453934567165e39eae5adf1134a863217e26e found no correctness blocker. The adapter disables quick-xml checking before the hash path, replays exactly the first 32 successful names, and preserves duplicate/error precedence, including malformed duplicate values. The four copies normalize to the canonical implementation.

The DOCX requested-storage envelope retains its previous vector/hash charge and adds a conservative ordered-index allowance above 32 names. Under the pinned standard-library B-tree layout (11 slots, five minimum entries in non-root nodes), four key/value slots per name cover node arrays and child links for n >= 33. The explicit map wrapper, next-position usize and adapter owner charge fixed state. This is requested storage under pinned layout assumptions, not portable RSS or global allocation recovery: BTreeMap and Box allocations remain infallible. Name comparisons are O(n log n); their byte cost still depends on name lengths.

Historical migration 54b03a912c was based on 2be9758ff9. Its 545 checked sites comprise 519 mechanically recognized fail-fast loops and 26 exceptional sites read by the implementer. The old independent review stopped at an API limit, so those 26 must not be described as independently hand-reviewed. The fresh migration is cherry-pick 7d38db4514 on base87e926fcc6. Current-source migration audit is recorded separately.

Root corrected two integration gaps: replace the stale broad MCE stream lint exemption with its unchecked-attribute adapter call, and charge the new ordered scratch in DOCX. The existing 0776 expanded-attribute check and 0775 shared namespace names are retained. No original external worktree was edited.

The iterator equivalence probe compares items and exact errors through the first error on parser-reached tags; it does not independently prove every reader's malformed-input recovery or lenient-path semantics. Native and allocation observations must stay scoped to their actual probe operations.
