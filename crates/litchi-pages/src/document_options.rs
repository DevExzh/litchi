//! Compact, archive-free Pages document formatter options.

const BODY: u8 = 1;
const HEADERS: u8 = 2;
const FOOTERS: u8 = 4;
const FACING_PAGES: u8 = 8;
const HYPHENATION: u8 = 16;
const LIGATURES: u8 = 32;

/// Lossless options shown by Pages' Document formatter.
///
/// The two bytes retain protobuf presence and value bits without carrying any
/// archive or codec state. Convenience methods expose the native effective
/// defaults separately from the lossless optional values.
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub struct Options {
    present: u8,
    values: u8,
}

impl Options {
    /// Construct options while retaining each native field's presence.
    #[must_use]
    pub const fn new(
        body_enabled: Option<bool>,
        headers_enabled: Option<bool>,
        footers_enabled: Option<bool>,
        facing_pages: Option<bool>,
        automatic_hyphenation: Option<bool>,
        ligatures_enabled: Option<bool>,
    ) -> Self {
        let mut options = Self {
            present: 0,
            values: 0,
        };
        options = options.with(BODY, body_enabled);
        options = options.with(HEADERS, headers_enabled);
        options = options.with(FOOTERS, footers_enabled);
        options = options.with(FACING_PAGES, facing_pages);
        options = options.with(HYPHENATION, automatic_hyphenation);
        options.with(LIGATURES, ligatures_enabled)
    }

    /// Return the optional body-enabled value.
    #[must_use]
    pub const fn body_enabled(self) -> Option<bool> {
        self.get(BODY)
    }

    /// Return the optional headers-enabled value.
    #[must_use]
    pub const fn headers_enabled(self) -> Option<bool> {
        self.get(HEADERS)
    }

    /// Return the optional footers-enabled value.
    #[must_use]
    pub const fn footers_enabled(self) -> Option<bool> {
        self.get(FOOTERS)
    }

    /// Return the optional facing-pages value.
    #[must_use]
    pub const fn facing_pages(self) -> Option<bool> {
        self.get(FACING_PAGES)
    }

    /// Return the optional automatic-hyphenation value.
    #[must_use]
    pub const fn automatic_hyphenation(self) -> Option<bool> {
        self.get(HYPHENATION)
    }

    /// Return the optional ligature value.
    #[must_use]
    pub const fn ligatures_enabled(self) -> Option<bool> {
        self.get(LIGATURES)
    }

    /// Return whether the body is effectively enabled.
    #[must_use]
    pub const fn body_is_enabled(self) -> bool {
        match self.body_enabled() {
            Some(value) => value,
            None => true,
        }
    }

    /// Return whether headers are effectively enabled.
    #[must_use]
    pub const fn headers_are_enabled(self) -> bool {
        match self.headers_enabled() {
            Some(value) => value,
            None => true,
        }
    }

    /// Return whether footers are effectively enabled.
    #[must_use]
    pub const fn footers_are_enabled(self) -> bool {
        match self.footers_enabled() {
            Some(value) => value,
            None => true,
        }
    }

    /// Return whether facing-page layout is effectively enabled.
    #[must_use]
    pub const fn uses_facing_pages(self) -> bool {
        match self.facing_pages() {
            Some(value) => value,
            None => false,
        }
    }

    /// Return whether automatic hyphenation is effectively enabled.
    #[must_use]
    pub const fn uses_automatic_hyphenation(self) -> bool {
        match self.automatic_hyphenation() {
            Some(value) => value,
            None => false,
        }
    }

    /// Return whether typographic ligatures are effectively enabled.
    #[must_use]
    pub const fn uses_ligatures(self) -> bool {
        match self.ligatures_enabled() {
            Some(value) => value,
            None => false,
        }
    }

    /// Set or clear the optional body-enabled value.
    pub fn set_body_enabled(&mut self, value: Option<bool>) {
        self.set(BODY, value);
    }

    /// Set or clear the optional headers-enabled value.
    pub fn set_headers_enabled(&mut self, value: Option<bool>) {
        self.set(HEADERS, value);
    }

    /// Set or clear the optional footers-enabled value.
    pub fn set_footers_enabled(&mut self, value: Option<bool>) {
        self.set(FOOTERS, value);
    }

    /// Set or clear the optional facing-pages value.
    pub fn set_facing_pages(&mut self, value: Option<bool>) {
        self.set(FACING_PAGES, value);
    }

    /// Set or clear the optional automatic-hyphenation value.
    pub fn set_automatic_hyphenation(&mut self, value: Option<bool>) {
        self.set(HYPHENATION, value);
    }

    /// Set or clear the optional ligature value.
    pub fn set_ligatures_enabled(&mut self, value: Option<bool>) {
        self.set(LIGATURES, value);
    }

    const fn with(mut self, bit: u8, value: Option<bool>) -> Self {
        self.set(bit, value);
        self
    }

    const fn get(self, bit: u8) -> Option<bool> {
        if self.present & bit != 0 {
            Some(self.values & bit != 0)
        } else {
            None
        }
    }

    const fn set(&mut self, bit: u8, option: Option<bool>) {
        if let Some(value) = option {
            self.present |= bit;
            if value {
                self.values |= bit;
            } else {
                self.values &= !bit;
            }
        } else {
            self.present &= !bit;
            self.values &= !bit;
        }
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::Options;

    #[test]
    fn options_pack_presence_and_retain_effective_defaults() {
        assert_eq!(size_of::<Options>(), 2);
        let options = Options::new(Some(false), None, Some(true), Some(true), None, Some(false));
        assert_eq!(options.body_enabled(), Some(false));
        assert_eq!(options.headers_enabled(), None);
        assert_eq!(options.footers_enabled(), Some(true));
        assert_eq!(options.facing_pages(), Some(true));
        assert_eq!(options.ligatures_enabled(), Some(false));
        assert!(!options.body_is_enabled());
        assert!(options.headers_are_enabled());
        assert!(options.uses_facing_pages());
        assert!(!options.uses_ligatures());
    }
}
