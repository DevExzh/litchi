//! Archive-free combined Pages document formatter settings.
//!
//! This value combines the independent Document and Footnotes formatter
//! settings that share one Pages document-settings transaction. Native
//! package identifiers and protobuf representations remain private to the
//! package adapter.
//!
//! The transaction surface is deliberately scoped here: use
//! [`Edit`][crate::document_settings::Edit] to stage one replacement,
//! [`Commit`][crate::document_settings::Commit] to access the verified package
//! snapshot, and [`Patch`][crate::document_settings::Patch] to apply that
//! replacement to the exact source artifact. Both [`Error`][crate::document_settings::Error]
//! `Display` and `Debug` output redact package bytes, member names, native
//! identifiers, field values, and retained patch artifacts.

use crate::{document_options, footnote};

pub use crate::package::document_settings::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};

/// Validated combined settings for Pages document and footnote formatters.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Settings {
    options: document_options::Options,
    footnotes: footnote::Settings,
}

impl Settings {
    /// Construct validated document formatter settings.
    ///
    /// # Errors
    ///
    /// Returns an error when the footnote settings contain a non-canonical
    /// representation of a known Pages value.
    pub fn new(
        options: document_options::Options,
        footnotes: footnote::Settings,
    ) -> footnote::Result<Self> {
        let settings = Self { options, footnotes };
        settings.validate()?;
        Ok(settings)
    }

    /// Return the document formatter options.
    #[must_use]
    pub const fn options(self) -> document_options::Options {
        self.options
    }

    /// Return the footnote formatter settings.
    #[must_use]
    pub const fn footnotes(self) -> footnote::Settings {
        self.footnotes
    }

    /// Replace the document formatter options.
    pub fn set_options(&mut self, options: document_options::Options) {
        self.options = options;
    }

    /// Replace the footnote formatter settings after validating them.
    ///
    /// # Errors
    ///
    /// Returns an error when the replacement contains a non-canonical
    /// representation of a known Pages value. On error this setting remains
    /// unchanged.
    pub fn set_footnotes(&mut self, footnotes: footnote::Settings) -> footnote::Result<()> {
        footnotes.validate()?;
        self.footnotes = footnotes;
        Ok(())
    }

    /// Validate settings before an archive adapter publishes them.
    ///
    /// # Errors
    ///
    /// Returns an error when the footnote settings contain a non-canonical
    /// representation of a known Pages value.
    pub fn validate(self) -> footnote::Result<()> {
        self.footnotes.validate()
    }
}

#[cfg(test)]
mod tests {
    use super::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Settings};
    use crate::{document_options::Options, footnote};

    #[test]
    fn settings_are_copyable_and_validate_replacements_before_publication() {
        fn assert_traits<T: Clone + Copy + Eq + std::fmt::Debug>() {}

        assert_traits::<Settings>();
        let options = Options::new(Some(true), None, None, None, None, None);
        let mut settings = Settings::new(options, footnote::Settings::default()).unwrap();
        assert_eq!(settings.options(), options);
        assert_eq!(settings.footnotes(), footnote::Settings::default());

        let invalid = footnote::Settings {
            kind: Some(footnote::Kind::Unknown(0)),
            ..footnote::Settings::default()
        };
        assert!(settings.set_footnotes(invalid).is_err());
        assert_eq!(settings.footnotes(), footnote::Settings::default());
    }

    #[test]
    fn transaction_surface_is_available_from_the_semantic_module() {
        fn assert_type<T>() {}

        assert_type::<Edit<'static>>();
        assert_type::<Patch>();
        assert_type::<Commit>();
        assert_type::<Diagnostics>();
        assert_type::<Error>();
        assert_type::<LimitKind>();
    }
}
