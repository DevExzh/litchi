/// Specifies which ZIP headers to place data.
///
/// The ZIP specification allows for different data in local file headers versus
/// central directory headers. This type provides control over where data is
/// placed.
///
/// The default value is to place data in both header locations.
///
/// For usage example, see
/// [`ZipFileBuilder::extra_field`](crate::ZipFileBuilder::extra_field)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Header(u8);

impl Header {
    /// Include data only in the local file header.
    pub const LOCAL: Self = Self(0b01);

    /// Include data only in the central directory.
    pub const CENTRAL: Self = Self(0b10);

    #[inline]
    pub(crate) const fn new(value: u8) -> Self {
        Self(value)
    }

    #[inline]
    pub(crate) const fn includes_local(self) -> bool {
        self.0 & Self::LOCAL.0 != 0
    }

    #[inline]
    pub(crate) const fn includes_central(self) -> bool {
        self.0 & Self::CENTRAL.0 != 0
    }

    #[inline]
    pub(crate) const fn intersects(self, other: Self) -> bool {
        (self.0 & other.0) != 0
    }
}

impl Default for Header {
    fn default() -> Self {
        Self(Self::LOCAL.0 | Self::CENTRAL.0)
    }
}

impl std::ops::BitOr for Header {
    type Output = Self;

    #[inline]
    fn bitor(self, rhs: Self) -> Self::Output {
        Self(self.0 | rhs.0)
    }
}

impl std::ops::BitOrAssign for Header {
    #[inline]
    fn bitor_assign(&mut self, rhs: Self) {
        self.0 |= rhs.0;
    }
}

impl std::ops::BitAnd for Header {
    type Output = Self;

    #[inline]
    fn bitand(self, rhs: Self) -> Self::Output {
        Self(self.0 & rhs.0)
    }
}

impl std::ops::BitAndAssign for Header {
    #[inline]
    fn bitand_assign(&mut self, rhs: Self) {
        self.0 &= rhs.0;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_header_bitflags_behavior() {
        // Test that default equals LOCAL | CENTRAL
        assert_eq!(Header::LOCAL | Header::CENTRAL, Header::default());

        // Test includes methods
        assert!(Header::LOCAL.includes_local());
        assert!(!Header::LOCAL.includes_central());

        assert!(!Header::CENTRAL.includes_local());
        assert!(Header::CENTRAL.includes_central());

        assert!(Header::default().includes_local());
        assert!(Header::default().includes_central());

        // Test bitwise operations
        let mut header = Header::LOCAL;
        header |= Header::CENTRAL;
        assert_eq!(header, Header::default());

        let intersection = Header::default() & Header::LOCAL;
        assert_eq!(intersection, Header::LOCAL);

        // Test intersects method
        assert!(Header::default().intersects(Header::LOCAL));
        assert!(Header::default().intersects(Header::CENTRAL));
        assert!(Header::LOCAL.intersects(Header::default()));
        assert!(!Header::LOCAL.intersects(Header::CENTRAL));
    }

    #[test]
    fn test_header_default() {
        assert_eq!(Header::default(), Header::LOCAL | Header::CENTRAL);
    }
}
