//! Byte offsets for the positions an XML pull reader reports.
//!
//! Every OOXML crate in this workspace parses XML with `quick-xml`, built
//! without its `encoding` feature. Before such a reader reports its first
//! event it removes one leading UTF-8 byte-order mark (`EF BB BF`) from its
//! input, and it does not count those three bytes: `buffer_position()` and
//! `error_position()` of a reader built over a marked input are three bytes
//! short of the byte offsets in that input, from the first event to the last.
//! A position is a byte offset only when the input has no mark.
//!
//! [`ReaderOrigin`] is the one place that knows this. Build it from the exact
//! bytes a reader was built over and convert every position with it before
//! the position addresses those bytes, is stored as a span, or is reported in
//! an error. A second mark right after the first is character data to the
//! reader, not part of the origin; a UTF-16 mark is not removed at all.
//!
//! The model is exact for readers over a byte slice or string. A reader over
//! a buffered stream removes the mark only when its first fill holds all
//! three bytes; streaming callers either present the whole mark in the first
//! fill or remove it themselves. `litchi-opc`'s `reader_origin_contract` test
//! pins this model to the linked `quick-xml`.

/// Where position zero of an XML reader lies in the reader's input.
///
/// `quick-xml` (0.41, without its `encoding` feature) removes one leading
/// UTF-8 byte-order mark before its first event and reports positions
/// relative to the byte after it. [`ReaderOrigin::of`] records how many input
/// bytes precede position zero — three for a marked input, otherwise none —
/// so [`offset`](Self::offset) and [`position`](Self::position) convert
/// between the reader's positions and byte offsets into the input.
///
/// ```
/// use litchi_core::xml::ReaderOrigin;
///
/// let marked = b"\xEF\xBB\xBF<root/>";
/// let origin = ReaderOrigin::of(marked);
/// assert_eq!(origin.skipped(), 3);
/// // The reader reports `<root/>` at positions 0..7.
/// let (start, end) = (origin.offset(0).unwrap(), origin.offset(7).unwrap());
/// assert_eq!(&marked[start..end], b"<root/>");
/// assert_eq!(origin.position(start), Some(0));
///
/// assert_eq!(ReaderOrigin::of(b"<root/>").skipped(), 0);
/// ```
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub struct ReaderOrigin {
    skipped: usize,
}

/// The UTF-8 byte-order mark a reader removes before its first event.
const UTF8_BYTE_ORDER_MARK: [u8; 3] = [0xEF, 0xBB, 0xBF];

impl ReaderOrigin {
    /// The origin of a reader whose input begins at position zero.
    pub const NONE: Self = Self { skipped: 0 };

    /// The origin of a reader built over exactly `input`.
    ///
    /// Only the first three bytes are read. Pass the bytes the reader was
    /// built over, not a larger buffer they were cut from: a reader over a
    /// fragment that starts at an element has no mark to remove.
    #[inline]
    #[must_use]
    pub const fn of(input: &[u8]) -> Self {
        if let [a, b, c, ..] = input
            && *a == UTF8_BYTE_ORDER_MARK[0]
            && *b == UTF8_BYTE_ORDER_MARK[1]
            && *c == UTF8_BYTE_ORDER_MARK[2]
        {
            return Self {
                skipped: UTF8_BYTE_ORDER_MARK.len(),
            };
        }
        Self::NONE
    }

    /// Input bytes before reader position zero: 3 when the input begins with
    /// a UTF-8 byte-order mark, 0 otherwise.
    #[inline]
    #[must_use]
    pub const fn skipped(self) -> usize {
        self.skipped
    }

    /// Whether the reader's input begins with a UTF-8 byte-order mark that
    /// the reader removes uncounted.
    #[inline]
    #[must_use]
    pub const fn has_byte_order_mark(self) -> bool {
        self.skipped != 0
    }

    /// The byte offset in the input of reader position `position`.
    ///
    /// Returns `None` when the offset does not fit `usize`.
    #[inline]
    #[must_use]
    pub fn offset(self, position: u64) -> Option<usize> {
        usize::try_from(position).ok()?.checked_add(self.skipped)
    }

    /// The reader position of input byte offset `offset`, the inverse of
    /// [`offset`](Self::offset).
    ///
    /// Returns `None` for an offset inside the removed mark, which no reader
    /// position addresses, and when the position does not fit `u64`.
    #[inline]
    #[must_use]
    pub fn position(self, offset: usize) -> Option<u64> {
        u64::try_from(offset.checked_sub(self.skipped)?).ok()
    }
}

#[cfg(test)]
mod tests {
    use super::ReaderOrigin;

    #[test]
    fn only_a_complete_leading_utf8_mark_moves_the_origin() {
        let cases: [(&[u8], usize); 11] = [
            (b"", 0),
            (b"<a/>", 0),
            (b"\xEF", 0),
            (b"\xEF\xBB", 0),
            (b"\xEF\xBB\xBF", 3),
            (b"\xEF\xBB\xBF<a/>", 3),
            // A second mark is character data to the reader.
            (b"\xEF\xBB\xBF\xEF\xBB\xBF<a/>", 3),
            // A mark after the first byte is not a leading mark.
            (b" \xEF\xBB\xBF<a/>", 0),
            // UTF-16 marks are not removed by a reader without `encoding`.
            (b"\xFF\xFE<\0a\0/\0>\0", 0),
            (b"\xFE\xFF\0<\0a\0/\0>", 0),
            (b"\xEF\xBB\xBE<a/>", 0),
        ];
        for (input, skipped) in cases {
            let origin = ReaderOrigin::of(input);
            assert_eq!(origin.skipped(), skipped, "{input:?}");
            assert_eq!(origin.has_byte_order_mark(), skipped != 0, "{input:?}");
        }
        assert_eq!(ReaderOrigin::default(), ReaderOrigin::NONE);
    }

    #[test]
    fn offsets_and_positions_are_inverse_and_checked() {
        let marked = ReaderOrigin::of(b"\xEF\xBB\xBF<a/>");
        assert_eq!(marked.offset(0), Some(3));
        assert_eq!(marked.offset(4), Some(7));
        assert_eq!(marked.position(3), Some(0));
        assert_eq!(marked.position(7), Some(4));
        for offset in 0..3 {
            assert_eq!(marked.position(offset), None, "inside the mark");
        }
        for position in [0_u64, 1, 2, 3, 1 << 20] {
            let offset = marked.offset(position).expect("fits");
            assert_eq!(marked.position(offset), Some(position));
        }

        let plain = ReaderOrigin::NONE;
        assert_eq!(plain.offset(0), Some(0));
        assert_eq!(plain.position(0), Some(0));
        assert_eq!(plain.offset(5), Some(5));

        // Overflow is refused rather than wrapped.
        let max = u64::try_from(usize::MAX).expect("usize fits u64 on supported targets");
        assert_eq!(plain.offset(max), Some(usize::MAX));
        assert_eq!(marked.offset(max), None);
        assert_eq!(marked.offset(u64::MAX), None);
    }
}
