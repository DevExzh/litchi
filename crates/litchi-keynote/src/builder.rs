//! Option-bearing fresh Keynote authoring.
//!
//! [`Builder`] starts from the same embedded blank presentation as
//! [`Package::new`] — one title-and-subtitle slide at 1920×1080 points —
//! stamps the caller's required [`Identity`] on it, and applies each
//! requested value through the crate's ordinary exact-source transactions
//! (slide text, speaker notes, show settings). [`Builder::build`] returns the
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
use crate::{Package, SlideSelector, show, slide};

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
    /// The slide size could not be written.
    #[error(transparent)]
    ShowSettings(#[from] show::Error),
    /// The title and body placeholder frames could not be laid out for the
    /// slide size.
    #[error("could not lay out the Keynote placeholders for the slide size: {0}")]
    Layout(#[source] ReadError),
    /// The slide-number placeholder's visibility could not be written.
    #[error(transparent)]
    PlaceholderVisibility(#[from] slide::placeholder::Error),
    /// The document language is not a valid language tag.
    #[error("invalid Keynote document language: {0}")]
    InvalidLanguage(#[source] litchi_iwa_text::language::Error),
    /// The document locale is empty or contains a NUL character.
    #[error("a Keynote document locale must be non-empty text without NUL")]
    InvalidLocale,
    /// The language or locale could not be written.
    #[error("could not write the Keynote document language or locale: {0}")]
    Locale(#[source] ReadError),
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
    /// placeholder, as [`Package::edit_slide_placeholder_visibility`] with
    /// [`slide::placeholder::Kind::SlideNumber`] does. Unset, it stays
    /// hidden. The show-wide slide-number preference is not changed.
    pub const fn slide_number_visible(mut self, visible: bool) -> Self {
        self.slide_number_visible = Some(visible);
        self
    }

    /// Build the presentation.
    ///
    /// # Costs
    ///
    /// Each requested value is one exact-source transaction that rewrites
    /// one component and reopens the candidate; values left at their
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
        if self.language.is_some() || self.locale.is_some() {
            package = package
                .with_language_and_locale(self.language.as_deref(), self.locale.as_deref())
                .map_err(Error::Locale)?;
        }
        let first = SlideSelector::index(0);
        if let Some(title) = self.title.as_deref() {
            let mut edit = package.edit_slide_title(first)?;
            edit.set(title)?;
            package = edit.commit()?.into_package();
        }
        if let Some(subtitle) = self.subtitle.as_deref() {
            let mut edit = package.edit_slide_body(first)?;
            edit.set(subtitle)?;
            package = edit.commit()?.into_package();
        }
        if let Some(notes) = self.notes.as_deref() {
            let mut edit = package.edit_slide_notes(first)?;
            edit.set(notes)?;
            package = edit.commit()?.into_package();
        }
        if let Some(size) = self.slide_size {
            let edit = package.edit_show_settings()?;
            let mut settings = edit.settings();
            settings.set_size(size);
            package = edit.set(settings).commit()?.into_package();
            package = package
                .with_placeholder_frames(size)
                .map_err(Error::Layout)?;
        }
        if let Some(visible) = self.slide_number_visible {
            let state = if visible {
                slide::placeholder::State::Visible
            } else {
                slide::placeholder::State::Hidden
            };
            package = package
                .edit_slide_placeholder_visibility(first, slide::placeholder::Kind::SlideNumber)?
                .set(state)
                .commit()?
                .into_package();
        }
        Ok(package)
    }
}
