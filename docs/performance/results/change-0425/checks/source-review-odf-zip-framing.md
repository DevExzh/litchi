# ODF ZIP framing test corrections

The locked baseline at `340cc91ae2bdec338dfe7682b4b5d8c219a2d288` reproduced the
two ODF failures recorded in `baseline-odp-locked.json` and
`baseline-odt-locked.json`. Both failures are test expectations made stale by
the strict local and descriptor validation introduced in `d18cd04a2`.

The ODP oversized-content fixture now records `content.xml`'s checked payload
range before changing its central uncompressed-size declaration. The range is
therefore obtained from a valid archive, while the source-backed open still
sees the deliberately oversized central metadata. The test continues to
require the family size refusal before any read overlaps the content payload;
the generic `member_range` validation path is unchanged.

The ODT malformed-descriptor test now expects refusal for CRC, compressed-size,
and uncompressed-size mutations. The strict ZIP reader must reject each
inconsistent descriptor before logical rebuild fallback. It also snapshots
`Document::to_bytes()` and verifies the failed definition edit leaves the
document byte-for-byte unchanged.

These are integration-test corrections only. No production validation or
fallback behavior is weakened. The edit lane ran targeted Rust 1.98.1
formatting and `git diff --check`; it ran no Cargo, build, or test commands.
