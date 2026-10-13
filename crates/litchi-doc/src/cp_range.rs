//! Character-position ranges in a document's logical text.

/// A half-open range of character positions (CPs) in a document's logical
/// text.
///
/// A CP counts UTF-16 code units across the document's stories, which the
/// FIB lays out back to back in a fixed order: the main story from CP 0, then
/// the footnote, header, macro, comment, endnote, textbox and header-textbox
/// stories ([MS-DOC] §2.3.1). A range read from a paragraph, run, table, row
/// or cell is where that value's text sits in this space, so it can be
/// mapped to an editor's view of the same text without comparing strings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct CpRange {
    pub(crate) start: u32,
    pub(crate) end: u32,
}

impl CpRange {
    /// The range `start..end`, or `None` when `start > end`.
    #[must_use]
    pub const fn new(start: u32, end: u32) -> Option<Self> {
        if start > end {
            None
        } else {
            Some(Self { start, end })
        }
    }

    /// The first CP of the range.
    #[must_use]
    pub const fn start(self) -> u32 {
        self.start
    }

    /// The CP just past the range.
    #[must_use]
    pub const fn end(self) -> u32 {
        self.end
    }

    /// The number of CPs in the range.
    #[must_use]
    pub const fn len(self) -> u32 {
        self.end - self.start
    }

    /// Whether the range holds no CP.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.start == self.end
    }

    /// Whether `cp` lies in the range.
    #[must_use]
    pub const fn contains(self, cp: u32) -> bool {
        self.start <= cp && cp < self.end
    }

    /// Whether `other` lies entirely in this range.
    #[must_use]
    pub const fn encloses(self, other: Self) -> bool {
        self.start <= other.start && other.end <= self.end
    }

    /// The smallest range holding both ranges.
    #[must_use]
    pub const fn cover(self, other: Self) -> Self {
        Self {
            start: if self.start < other.start {
                self.start
            } else {
                other.start
            },
            end: if self.end > other.end {
                self.end
            } else {
                other.end
            },
        }
    }
}

#[cfg(test)]
mod tests {
    use super::CpRange;

    #[test]
    fn ranges_are_half_open_and_ordered() {
        assert!(CpRange::new(5, 4).is_none());
        let range = CpRange::new(2, 6).unwrap();
        assert_eq!((range.start(), range.end(), range.len()), (2, 6, 4));
        assert!(!range.is_empty());
        assert!(range.contains(2) && range.contains(5) && !range.contains(6));
        assert!(CpRange::new(3, 3).unwrap().is_empty());
        assert!(range.encloses(CpRange::new(3, 6).unwrap()));
        assert!(!range.encloses(CpRange::new(1, 3).unwrap()));
        assert_eq!(
            range.cover(CpRange::new(8, 9).unwrap()),
            CpRange::new(2, 9).unwrap()
        );
    }
}
