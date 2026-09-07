//! Archive-free properties shared by Keynote slide media.
//!
//! The package adapter translates native drawable fields into this value and
//! applies replacements while keeping the native records and their unknown
//! fields private.  Optional values intentionally retain the distinction
//! between an omitted field and an explicitly encoded default, including an
//! empty string and `false`.

/// User-facing properties attached to one Keynote media drawable.
///
/// The value owns its strings so it can outlive the package that produced it.
/// Its fields are private to keep the semantic API stable while native archive
/// representation remains an implementation detail of the package adapter.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct MediaProperties {
    hyperlink_url: Option<String>,
    locked: Option<bool>,
    aspect_ratio_locked: Option<bool>,
    accessibility_description: Option<String>,
}

impl MediaProperties {
    /// Construct properties with every native field omitted.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            hyperlink_url: None,
            locked: None,
            aspect_ratio_locked: None,
            accessibility_description: None,
        }
    }

    /// Construct properties from their detached semantic values.
    #[must_use]
    pub fn from_parts(
        hyperlink_url: Option<String>,
        locked: Option<bool>,
        aspect_ratio_locked: Option<bool>,
        accessibility_description: Option<String>,
    ) -> Self {
        Self {
            hyperlink_url,
            locked,
            aspect_ratio_locked,
            accessibility_description,
        }
    }

    /// Return the optional hyperlink target.
    #[must_use]
    pub fn hyperlink_url(&self) -> Option<&str> {
        self.hyperlink_url.as_deref()
    }

    /// Return the optional native lock state.
    #[must_use]
    pub const fn locked(&self) -> Option<bool> {
        self.locked
    }

    /// Return the optional aspect-ratio lock state.
    #[must_use]
    pub const fn aspect_ratio_locked(&self) -> Option<bool> {
        self.aspect_ratio_locked
    }

    /// Return the optional accessibility description.
    #[must_use]
    pub fn accessibility_description(&self) -> Option<&str> {
        self.accessibility_description.as_deref()
    }

    /// Return a copy with the supplied optional hyperlink target.
    #[must_use]
    pub fn with_hyperlink_url(mut self, hyperlink_url: Option<String>) -> Self {
        self.hyperlink_url = hyperlink_url;
        self
    }

    /// Return a copy with the supplied optional native lock state.
    #[must_use]
    pub const fn with_locked(mut self, locked: Option<bool>) -> Self {
        self.locked = locked;
        self
    }

    /// Return a copy with the supplied optional aspect-ratio lock state.
    #[must_use]
    pub const fn with_aspect_ratio_locked(mut self, aspect_ratio_locked: Option<bool>) -> Self {
        self.aspect_ratio_locked = aspect_ratio_locked;
        self
    }

    /// Return a copy with the supplied optional accessibility description.
    #[must_use]
    pub fn with_accessibility_description(
        mut self,
        accessibility_description: Option<String>,
    ) -> Self {
        self.accessibility_description = accessibility_description;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::MediaProperties;

    #[test]
    fn omitted_values_are_distinct_from_explicit_defaults() {
        let omitted = MediaProperties::new();
        let explicit = MediaProperties::new()
            .with_hyperlink_url(Some(String::new()))
            .with_locked(Some(false))
            .with_aspect_ratio_locked(Some(false))
            .with_accessibility_description(Some(String::new()));

        assert_eq!(omitted, MediaProperties::default());
        assert_eq!(omitted.hyperlink_url(), None);
        assert_eq!(omitted.locked(), None);
        assert_eq!(omitted.aspect_ratio_locked(), None);
        assert_eq!(omitted.accessibility_description(), None);
        assert_eq!(explicit.hyperlink_url(), Some(""));
        assert_eq!(explicit.locked(), Some(false));
        assert_eq!(explicit.aspect_ratio_locked(), Some(false));
        assert_eq!(explicit.accessibility_description(), Some(""));
        assert_ne!(omitted, explicit);
    }

    #[test]
    fn strings_are_owned_and_borrowed_without_losing_unicode() {
        let properties = MediaProperties::from_parts(
            Some("https://例.example/音声".to_owned()),
            Some(true),
            Some(true),
            Some("説明 🎵".to_owned()),
        );

        assert_eq!(properties.hyperlink_url(), Some("https://例.example/音声"));
        assert_eq!(properties.locked(), Some(true));
        assert_eq!(properties.aspect_ratio_locked(), Some(true));
        assert_eq!(properties.accessibility_description(), Some("説明 🎵"));
    }
}
