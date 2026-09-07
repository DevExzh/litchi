# Source and measurement review

The 0459 diagnostic build adds explicit frame pointers and unwind tables; the
fp and 32 KiB DWARF recordings resolve phase ancestry for 94.62% and 91.60% of
sampled periods respectively. These are whole-process sampled estimates,
including warmup calls. They are separate from ordinary measurement binaries.

Candidate slide readback accounts for roughly 59–60% of the commit group's
sampled periods; writer work accounts for roughly 27–30%. The cached attribute
lookup is prominent in both initial slide parsing and candidate readback.
Checking the local name before the namespace URI can reject mismatches before
repeated longer comparisons. Both predicates are pure borrowed-slice checks;
scan order, namespace resolution, lazy decoding, first-match and error behavior
remain unchanged. The one-shot reference reader is unchanged.

The bounded setup audit rejects retaining Package/Presentation: reopening
accounts for at most about 3% of the transaction group, while staging metadata
accounts for about 64% and source-fragment scanning about 32%. Retaining the
large XML owner would increase retention for little measured benefit. The
family check is only MIME/size/body-marker admission; it does not perform the
full ContentDocumentValidator scan. The initial contrary hypothesis was
corrected from source before any production experiment.

A later fused staging/fragment traversal could share tokenization and namespace
maintenance while keeping all four state machines. It must preserve historical
settings/declarations/pages-first error precedence, source scanner errors,
limits, BOM-relative exact spans, lexical preservation and independent final
candidate readback. This is a separate hypothesis, not part of this patch.

Existing ElementAttrs tests compare the cache to the unchanged one-shot reader
and preserve malformed/duplicate error order, unknown prefixes, namespace
shadowing, normalized values and drawing-attribute harvest order. The full ODP
suite also exercises native fixture preservation, no-op and reversible patches.
No new test duplicates the commutativity of two pure Boolean comparisons.
