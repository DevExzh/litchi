# 0359k — W5 COW allocation and retained-history evidence

- Date: 2026-10-21.
- Candidate: `ad9904197c3b1dc79aba6b35c2f7dbb42a9b2190`.
- Status: allocation evidence complete; native latency qualification remains
  pending and no performance exception is accepted.

The paired counting-allocator run used the identical external probe source
(`ad4c52c01333fbc13c7b3cca18e2c0a0d2ccf89d5c07972e2e453fa9bd8d713a`),
identical lock (`bc5a133f0668142d18d637b2b196097e855c8890697cb3bbf43aeb937145f04a`),
and one capped release build per side. The baseline probe SHA-256 is
`23a7088c3f9a6563adc3ae1de644a0d2b05fe91604824afd7517aa426c981fc0`; the
COW candidate is
`ec39b816664c7771711d0081e77bad795f40a9911e693094e5296b3b629fbbe6`.
The complete external receipt is
`/home/zhuhe/code/litchi-worktrees/.coord/w5-ole/counting-cow-ad990-20261021`;
its 256 KiB directory manifest has SHA-256
`0f7fdf44002e971cb955cd176b9785654ad8f645de3cafc1ff10027bfc16d546`.
The runner reports allocation counts and bytes only; it does not report native
latency.

Ordinary property read and edit have identical candidate and baseline
counters: 45 calls and 1,192 bytes for read, and 33 calls and 3,523 bytes for
edit. The full XLS cell and stage routes each reduce calls from 1,589 to 1,588
but increase requested bytes from 2,361,904 to 2,362,493 (+589, 0.025%).
Their retained live bytes are unchanged; absolute peak bytes rise from
1,332,635 to 1,332,729 (+94). This small byte increase remains attributed in
the native follow-up rather than being called unchanged.

The retained 128-step histories keep 129 packages and 128 patches alive after
forward/inverse exact-output validation. DOC history improves from 1,609,900
to 1,572,032 allocation calls and from 1,830,103,867 to 1,770,486,356
requested bytes. Its retained live and peak bytes are unchanged at 139,409,018
and 142,365,936. XLS history reduces calls from 387,769 to 387,511 but raises
requested bytes from 441,765,854 to 441,852,556 (+86,702, +0.0196%); retained
live and peak bytes are unchanged at 47,964,514 and 48,722,505. These results
are paired live-counter evidence, with decimal byte counts, not exit-leak or
RSS measurements.

The repaired native qualification still must run the owner clone and all
eleven selectors from 0359j, with every accepted batched leg at least 250 ms,
host-overlap checks, exact output, and forward/inverse replay. The counting
result does not close the remaining native timing residuals.
