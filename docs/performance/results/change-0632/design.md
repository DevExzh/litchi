# Frozen design — change 0632: the central directory is read once, and the buffer it lands in is sized to it

Written before any production line was changed, against base `c7326f680`.

## The cost being attacked

After changes [0611](../../0611-zip-single-read-per-member.md) and
[0623](../../0623-zip-structural-span-accessor-and-prefetch.md) a source-backed
OOXML open of `ConditionalFormattingSamples.xlsx` (132 members) costs **10**
positional requests. Three of them are the locate:

| # | site | offset | bytes |
| --- | --- | --- | ---: |
| 1 | `ZipLocator::locate_in_reader`'s fixed-EOCD probe | `len - 22` | 22 |
| 2 | `ZipLocator::finish_locate_in_reader`'s first-central-record probe | `central_dir_offset` | 46 |
| 3 | `ZipEntries::next_entry`'s first refill, from `IndexedArchive::from_zip_archive_with_limits_and_policy` | `central_dir_offset` | 9,724 |

Requests 2 and 3 **begin at the same offset**: the 46 bytes of request 2 are the
first 46 bytes of request 3. Two `ReadAt` calls fetch one byte range.

`from_reader_with_limits_and_policy` also allocates and zero-fills **two**
64 KiB scratch buffers per open — one for the locator, one for the
central-directory scan — although the second one's exact requirement,
`central_directory_size`, is known the moment the EOCD is parsed. That is
[0587](../../0587-remaining-opportunity-survey.md)'s ZIP-8, first half.

## The change

**One bounded read of the head of the central directory serves both the probe
and the scan, and the buffer it lands in is the scan's buffer.**

`finish_locate_in_reader`, when the locator is configured with a directory
prefill bound `W`, replaces its 46-byte stack probe with one read of
`min(central_directory_size, W)` bytes at the same offset into a fallibly
reserved buffer it hands back to the caller. `IndexedArchive` uses that buffer,
already holding those bytes, as the central-directory scan buffer. The scan's
first refill is therefore free, and no second 64 KiB scratch is allocated.

`W = RECOMMENDED_BUFFER_SIZE` (64 KiB), the size of the buffer the scan uses
today and the per-request ceiling change 0623 already set for one speculative
read. Corpus: the largest central directory under `test-data` is **9,724 bytes**
across 533 containers (p50 1,217, p99 7,166), so the window covers the whole
directory on every container in the corpus and the buffer shrinks on every one
of them.

## Why the read grammar does not move

Let `n = min(central_directory_size, W)` and `d = directory_offset()`.

1. **The probe sees the same bytes.** Today: `read_exact_at(&mut [0u8; 46], d)`,
   then `ZipFileHeaderFixed::parse` of those 46 bytes. After:
   `try_read_at_least_at(&mut buf[..n], n, d)`, then the same parse of
   `buf[..46]`. When the source delivers, both hold the same 46 bytes. When it
   does not — an `Err`, or a zero-length read before 46 bytes — both reach the
   same `None`, because `read_exact_at` and `try_read_at_least_at` stop on
   exactly the same two conditions. `n >= ZipFileHeaderFixed::SIZE` is a
   precondition of enabling the prefill at all; a directory shorter than one
   fixed record keeps today's 46-byte probe.
2. **The base-offset fallback is unchanged.** When the probe at `d` does not
   parse, the same second probe runs at `head_eocd_offset - central_dir_size`,
   as the same kind of read, and updates `base_offset` and `central_dir_offset`
   under exactly today's condition. A prefill whose start is not the archive's
   final `directory_offset()` is discarded.
3. **The scan starts in the state its own first read would have left.** Today
   `ZipEntries::next_entry`'s first refill calls
   `read_at_least_at(&mut buffer[..max_read], 46, d)` with
   `max_read = min(central_dir_size, buffer.len())`, leaving `pos = 0`,
   `end = read`, `offset = d + read`. With the prefill the iterator starts at
   `pos = 0`, `end = n'`, `offset = d + n'` where `n'` is what the prefill read
   returned. For `buffer.len() = n` those are the same numbers. Every later
   refill, every limit charge, every parse and every error is reached from the
   same buffer contents at the same logical position.
4. **The oversized-record spill boundary is pinned, not moved.** `next_entry`
   spills a record whose variable part exceeds `buffer.len()` into an owned
   buffer, and that boundary is therefore a function of the caller's buffer
   size. Shrinking the scan buffer from 64 KiB to `central_dir_size` would move
   it: a record declaring `central_dir_size < variable_length <= 65,536` refuses
   with `BufferTooSmall` today (through the refill) and would refuse with `Eof`
   (through the spill) after. So `ZipEntries` gains an explicit
   `spill_threshold`, which every existing constructor sets to `buffer.len()`
   — byte-for-byte today's behaviour — and which the index scan sets to
   `max(buffer.len(), RECOMMENDED_BUFFER_SIZE)`. With the threshold pinned at
   64 KiB the shrunken buffer reproduces today's refusal identity in both
   directions: `variable_length > 65,536` spills as before, and anything
   smaller that does not fit refuses with `BufferTooSmall` as before.
5. **Nothing is trusted because it came from the buffer.** The prefill is a
   cache of source bytes. Every central record is parsed, every metadata byte
   charged and every limit checked by exactly the code that does so when the
   scan reads the same bytes itself.
6. **Bounded and fallible.** The prefill buffer is `try_reserve_exact`d, is at
   most `W`, is at most `central_directory_size`, and is released with the
   index construction. No new `unsafe`, no new dependency, no raised limit.

## What is deliberately not done

* The **locator's own** 64 KiB scratch stays. Its size is what keeps the
  backwards EOCD search to one request for a comment-bearing or trailing-byte
  archive, and the directory size is not known before the locate.
* The **EOCD search rules are untouched**: the `len - 22` fast path keeps its
  `comment_len == 0 && !is_zip64()` gate, and everything else keeps the
  backwards search with its `max_search_space`.
* `from_zip_archive*` — the public constructor for an archive the caller
  located itself — gets the right-sized scan buffer but no prefill, because
  its caller's locate has already happened.

## Falsification

If the extended differential of change 0611, run in full over change 0582's
22,875-input corpus, both limit profiles and both builds, reports any divergence
other than none, the change is withdrawn.
