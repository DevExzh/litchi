//! Small format-owned selectors and borrowed SVG payload inputs.
//!
//! Package relationships, media-part identity, generated IDs, and OPC graph
//! planning remain private to the source-backed worksheet transaction.  These
//! values only express which existing semantic picture is selected and which
//! caller-owned bytes should be validated by that transaction.

/// A semantic worksheet drawing/picture selector.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[must_use]
pub struct PictureSelector {
    /// Zero-based worksheet drawing position.
    pub drawing: usize,
    /// Zero-based direct picture position within that drawing.
    pub picture: usize,
}

impl PictureSelector {
    /// Construct a zero-based drawing/picture selector.  The operation using
    /// it validates both positions against the selected worksheet drawing.
    pub const fn new(drawing: usize, picture: usize) -> Self {
        Self { drawing, picture }
    }
}

/// A borrowed SVG payload supplied to a source-backed attachment operation.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[must_use]
pub struct SvgInput<'a>(&'a [u8]);

impl<'a> SvgInput<'a> {
    /// Wrap caller-owned bytes without copying them.
    pub const fn new(bytes: &'a [u8]) -> Self {
        Self::borrowed(bytes)
    }

    /// Wrap caller-owned bytes without copying them.
    pub const fn borrowed(bytes: &'a [u8]) -> Self {
        Self(bytes)
    }

    /// Borrow the supplied SVG bytes.
    #[must_use]
    pub const fn as_bytes(self) -> &'a [u8] {
        self.0
    }
}

impl<'a> From<&'a [u8]> for SvgInput<'a> {
    fn from(value: &'a [u8]) -> Self {
        Self::borrowed(value)
    }
}

impl<'a> From<&'a Vec<u8>> for SvgInput<'a> {
    fn from(value: &'a Vec<u8>) -> Self {
        Self::borrowed(value.as_slice())
    }
}

impl AsRef<[u8]> for SvgInput<'_> {
    fn as_ref(&self) -> &[u8] {
        self.0
    }
}
