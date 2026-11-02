//! Option-bearing fresh Keynote authoring.
//!
//! [`Builder`] starts from the same embedded blank presentation as
//! [`Package::new`] — one title-and-subtitle slide at 1920×1080 points —
//! stamps the caller's required [`Identity`] on it, writes the requested
//! values (language, locale, slide size with its placeholder layout, title,
//! subtitle and speaker notes) in one publication through the same wire
//! rewrites the crate's transactions use, and shows the slide number through
//! the slide placeholder-visibility transaction. [`Builder::build`] returns the
//! resulting [`Package`]; publication belongs to [`Package::write_to`] and
//! [`Package::save`].
//!
//! The result is a pure function of the builder's values: no clock,
//! randomness or locale is consulted, so two builds with the same fixed
//! identity are byte-identical.
//!
//! ```
//! use litchi_keynote::{Builder, Identity, SlideSelector, show::Size};
//!
//! let package = Builder::new()
//!     .identity(Identity::generate())
//!     .title("Quarterly review")
//!     .subtitle("October 2026")
//!     .notes("Open with the headline number.")
//!     .slide_size(Size::new(1024.0, 768.0)?)
//!     .build()?;
//! assert_eq!(
//!     package.slide_title(SlideSelector::index(0))?.as_deref(),
//!     Some("Quarterly review")
//! );
//! assert_eq!(package.show_settings()?.size(), Size::new(1024.0, 768.0)?);
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use thiserror::Error as ThisError;

use crate::identity::Identity;
use crate::package::ReadError;
use crate::show::Size;
use crate::{Package, SlideSelector, slide};

/// An error raised by [`Builder::build`].
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// No [`Identity`] was supplied; fresh authoring never invents one.
    #[error("a Keynote builder requires an explicit document identity")]
    MissingIdentity,
    /// The embedded template could not carry the requested identity.
    #[error("could not stamp the Keynote document identity: {0}")]
    Identity(#[source] ReadError),
    /// The title or subtitle could not be written.
    #[error(transparent)]
    Text(#[from] slide::text::Error),
    /// The speaker notes could not be written.
    #[error(transparent)]
    Notes(#[from] slide::notes::Error),
    /// The slide-number placeholder's visibility could not be written.
    #[error(transparent)]
    PlaceholderVisibility(#[from] slide::placeholder::Error),
    /// The document language is not a valid language tag.
    #[error("invalid Keynote document language: {0}")]
    InvalidLanguage(#[source] litchi_iwa_text::language::Error),
    /// The document locale is empty or contains a NUL character.
    #[error("a Keynote document locale must be non-empty text without NUL")]
    InvalidLocale,
    /// The language, locale or slide size (with the placeholder layout)
    /// could not be written, or the one publication of the fresh values
    /// failed.
    #[error("could not write the Keynote document settings: {0}")]
    Settings(#[source] ReadError),
}

/// Builder for a fresh one-slide Keynote presentation.
///
/// Unset values keep the blank presentation's defaults: title
/// "Presentation Title", subtitle "Presentation Subtitle", no speaker notes,
/// a 1920×1080-point slide, language `en`, locale `en_US` and a hidden
/// slide number. The [`Identity`] has no default.
#[derive(Debug, Clone, Default, PartialEq)]
#[must_use = "a Keynote builder does nothing until `build` is called"]
pub struct Builder {
    identity: Option<Identity>,
    title: Option<String>,
    subtitle: Option<String>,
    notes: Option<String>,
    slide_size: Option<Size>,
    language: Option<String>,
    locale: Option<String>,
    slide_number_visible: Option<bool>,
}

impl Builder {
    /// Start a builder with every value at its blank default and no identity.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the required native document identity.
    pub const fn identity(mut self, identity: Identity) -> Self {
        self.identity = Some(identity);
        self
    }

    /// Set the first slide's title text.
    pub fn title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Set the first slide's subtitle (body placeholder) text.
    pub fn subtitle(mut self, subtitle: impl Into<String>) -> Self {
        self.subtitle = Some(subtitle.into());
        self
    }

    /// Set the first slide's speaker notes.
    pub fn notes(mut self, notes: impl Into<String>) -> Self {
        self.notes = Some(notes.into());
        self
    }

    /// Set the presentation's slide size in points.
    ///
    /// The title and body placeholders of the layout and the first slide are
    /// laid out for the size: the title at 10 % × 24 % of the slide,
    /// 80 % × 20 % of it, the body at 15 % × 52 %, 70 % × 14 %.
    /// [`Size`] is validated at construction (finite, strictly positive).
    pub const fn slide_size(mut self, size: Size) -> Self {
        self.slide_size = Some(size);
        self
    }

    /// Set the document language (a BCP 47 tag such as `"fr"`), written to
    /// the document and to every text storage's language, as Keynote does.
    /// Unset, the blank presentation's `"en"` stays.
    pub fn language(mut self, language: impl Into<String>) -> Self {
        self.language = Some(language.into());
        self
    }

    /// Set the document locale (for example `"fr_FR"`), written to the
    /// document's current and creation locale and the calculation engine.
    /// Unset, the blank presentation's `"en_US"` stays.
    pub fn locale(mut self, locale: impl Into<String>) -> Self {
        self.locale = Some(locale.into());
        self
    }

    /// Show (`true`) or hide (`false`) the first slide's slide-number
    /// placeholder, the state [`Package::edit_slide_placeholder_visibility`]
    /// with [`slide::placeholder::Kind::SlideNumber`] edits. Unset, it stays
    /// hidden. The show-wide slide-number preference is not changed.
    pub const fn slide_number_visible(mut self, visible: bool) -> Self {
        self.slide_number_visible = Some(visible);
        self
    }

    /// Build the presentation.
    ///
    /// # Costs
    ///
    /// The language, locale, slide size, title, subtitle and notes are
    /// written together in one publication and one reopen; a visible slide
    /// number is one more exact-source transaction. Values left at their
    /// defaults cost nothing.
    ///
    /// # Errors
    ///
    /// Returns [`Error::MissingIdentity`] when no identity was supplied, or
    /// the typed error of the transaction that failed.
    pub fn build(self) -> Result<Package, Error> {
        let identity = self.identity.ok_or(Error::MissingIdentity)?;
        if let Some(language) = self.language.as_deref() {
            litchi_iwa_text::TextLanguageTag::new(language).map_err(Error::InvalidLanguage)?;
        }
        if self
            .locale
            .as_deref()
            .is_some_and(|locale| locale.trim().is_empty() || locale.contains('\0'))
        {
            return Err(Error::InvalidLocale);
        }
        let mut package = Package::new()
            .with_identity(&identity)
            .map_err(Error::Identity)?;
        package = package
            .with_fresh_settings(crate::package::FreshSettings {
                language: self.language.as_deref(),
                locale: self.locale.as_deref(),
                size: self.slide_size,
                title: self.title.as_deref(),
                subtitle: self.subtitle.as_deref(),
                notes: self.notes.as_deref(),
            })
            .map_err(|error| match error {
                crate::package::FreshError::Read(error) => Error::Settings(error),
                crate::package::FreshError::Text(error) => Error::Text(error),
                crate::package::FreshError::Notes(error) => Error::Notes(error),
            })?;
        if let Some(visible) = self.slide_number_visible {
            let state = if visible {
                slide::placeholder::State::Visible
            } else {
                slide::placeholder::State::Hidden
            };
            package = package
                .edit_slide_placeholder_visibility(
                    SlideSelector::index(0),
                    slide::placeholder::Kind::SlideNumber,
                )?
                .set(state)
                .commit()?
                .into_package();
        }
        Ok(package)
    }
}
