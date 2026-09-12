//! Checked, additive ODF data-style vocabulary.
//!
//! The legacy automatic-style values in [`crate::advanced`] intentionally
//! remain source compatible.  This module owns the additional number-style
//! particles and common data-style metadata needed to describe scientific,
//! fraction, embedded-text, and transliteration values.  It is a semantic
//! value layer: package scanning and source splicing stay with the ODS
//! package owner.

use std::{borrow::Cow, collections::BTreeSet, fmt::Write as _, num::IntErrorKind, str::FromStr};

use litchi_core::{Error, Result};

pub(crate) mod source;

/// Maximum size of one style name or text value.
pub const MAX_STYLE_TEXT_BYTES: usize = 65_536;
/// Maximum number of child particles in one authored number body.
pub const MAX_STYLE_BODY_ELEMENTS: usize = 4_096;
/// Maximum aggregate body text emitted by one typed style.
pub const MAX_STYLE_BODY_BYTES: usize = 1_048_576;
/// Maximum number of authored graph nodes.
pub const MAX_STYLE_GRAPH_NODES: usize = 4_096;
/// Maximum canonical output of one authored graph.  Individual style bodies
/// remain bounded by [`MAX_STYLE_BODY_BYTES`]; this separate graph budget keeps
/// those per-style limits composable without imposing a one-style limit on the
/// whole graph.
pub const MAX_STYLE_GRAPH_BYTES: usize = MAX_STYLE_BODY_BYTES.saturating_mul(MAX_STYLE_GRAPH_NODES);
/// Maximum number of entries in a source style catalog.
pub const MAX_STYLE_CATALOG_ENTRIES: usize = 65_536;
/// Maximum aggregate common metadata bytes for one style.
pub const MAX_STYLE_METADATA_BYTES: usize = 64 * 1024;
/// Hard byte ceiling for one scanned XML style owner.
pub const MAX_STYLE_OWNER_BYTES: usize = 256 * 1024 * 1024;
/// Hard element ceiling for one scanned XML style owner.
pub const MAX_STYLE_OWNER_ELEMENTS: usize = 1_048_576;
/// Hard nesting ceiling for one scanned XML style owner.
pub const MAX_STYLE_OWNER_DEPTH: usize = 1_024;

const NUMBER_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

/// The source owner of a data-style definition.
#[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
#[non_exhaustive]
pub enum Owner {
    /// Automatic styles in the mutable `content.xml` owner.
    ContentAutomatic,
    /// Direct common styles in `styles.xml`, which are read-only in this batch.
    CommonStyles,
    /// Automatic styles in `styles.xml`, which are read-only in this batch.
    StylesAutomatic,
}

impl Owner {
    /// Whether this owner is the mutable automatic-style owner in this batch.
    #[must_use]
    pub const fn is_mutable(self) -> bool {
        matches!(self, Self::ContentAutomatic)
    }
}

/// The ODF family selected by a data-style lookup.
#[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
#[non_exhaustive]
pub enum Family {
    Number,
    Date,
    Time,
    Currency,
    Percentage,
    Boolean,
    Text,
}

impl Family {
    /// Return the ODF root local name for this family.
    #[must_use]
    pub const fn root_element(self) -> &'static str {
        match self {
            Self::Number => "number-style",
            Self::Date => "date-style",
            Self::Time => "time-style",
            Self::Currency => "currency-style",
            Self::Percentage => "percentage-style",
            Self::Boolean => "boolean-style",
            Self::Text => "text-style",
        }
    }
}

/// A source-qualified data-style selector.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct Selector<'a> {
    pub owner: Owner,
    pub family: Family,
    pub name: &'a str,
}

impl<'a> Selector<'a> {
    /// Select an automatic style in `content.xml`.
    #[must_use]
    pub const fn automatic(name: &'a str, family: Family) -> Self {
        Self {
            owner: Owner::ContentAutomatic,
            family,
            name,
        }
    }

    /// Select a common style in `styles.xml`.
    #[must_use]
    pub const fn common(name: &'a str, family: Family) -> Self {
        Self {
            owner: Owner::CommonStyles,
            family,
            name,
        }
    }

    /// Select an automatic style in the read-only `styles.xml` owner.
    #[must_use]
    pub const fn styles_automatic(name: &'a str, family: Family) -> Self {
        Self {
            owner: Owner::StylesAutomatic,
            family,
            name,
        }
    }

    /// Validate the selector name before it reaches a source lookup.
    pub fn validate(self) -> Result<()> {
        validate_name(self.name, "data style selector name")
    }
}

/// The three ODF transliteration display styles.
#[derive(Clone, Copy, Debug, Default, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub enum TransliterationStyle {
    #[default]
    Short,
    Medium,
    Long,
}

impl TransliterationStyle {
    /// Return the ODF lexical spelling.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Short => "short",
            Self::Medium => "medium",
            Self::Long => "long",
        }
    }
}

impl FromStr for TransliterationStyle {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        match value {
            "short" => Ok(Self::Short),
            "medium" => Ok(Self::Medium),
            "long" => Ok(Self::Long),
            _ => invalid(format!("invalid ODF transliteration style '{value}'")),
        }
    }
}

/// Transliteration metadata shared by every ODF data-style root.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct Transliteration {
    /// One Unicode `Nd` character whose numeric value is one.
    pub format: Option<String>,
    pub language: Option<String>,
    pub country: Option<String>,
    pub style: Option<TransliterationStyle>,
}

impl Transliteration {
    /// Construct and validate transliteration metadata.
    pub fn new(
        format: Option<String>,
        language: Option<String>,
        country: Option<String>,
        style: Option<TransliterationStyle>,
    ) -> Result<Self> {
        let value = Self {
            format,
            language,
            country,
            style,
        };
        value.validate()?;
        Ok(value)
    }

    /// Build from borrowed source fields.  Each field is preflighted before
    /// any destination string is reserved or copied.
    pub fn try_from_borrowed(
        format: Option<&str>,
        language: Option<&str>,
        country: Option<&str>,
        style: Option<TransliterationStyle>,
    ) -> Result<Self> {
        validate_optional_text(format, "transliteration format")?;
        validate_optional_code(language, "transliteration language", validate_country_code)?;
        validate_optional_code(country, "transliteration country", validate_country_code)?;
        if let Some(format) = format {
            validate_transliteration_format(format)?;
        }
        Ok(Self {
            format: clone_optional_text(format, "transliteration format")?,
            language: clone_optional_code_text(
                language,
                "transliteration language",
                validate_country_code,
            )?,
            country: clone_optional_code_text(
                country,
                "transliteration country",
                validate_country_code,
            )?,
            style,
        })
    }

    /// Validate the normative Unicode decimal-digit constraint.
    pub fn validate(&self) -> Result<()> {
        validate_optional_text(self.format.as_deref(), "transliteration format")?;
        validate_optional_code(
            self.language.as_deref(),
            "transliteration language",
            validate_country_code,
        )?;
        validate_optional_code(
            self.country.as_deref(),
            "transliteration country",
            validate_country_code,
        )?;
        if let Some(format) = self.format.as_deref() {
            validate_transliteration_format(format)?;
        }
        Ok(())
    }

    /// The effective format.  ODF uses ASCII digits when the attribute is omitted.
    #[must_use]
    pub fn effective_format(&self) -> &str {
        self.format.as_deref().unwrap_or("1")
    }

    /// Return the locale language only when transliteration is active.
    #[must_use]
    pub fn effective_language(&self) -> Option<&str> {
        self.format.as_ref().and(self.language.as_deref())
    }

    /// Return the locale country only when transliteration is active.
    #[must_use]
    pub fn effective_country(&self) -> Option<&str> {
        self.format.as_ref().and(self.country.as_deref())
    }

    /// The effective style.  ODF defaults to `short`.
    #[must_use]
    pub const fn effective_style(&self) -> TransliterationStyle {
        match (self.format.is_some(), self.style) {
            (true, Some(style)) => style,
            _ => TransliterationStyle::Short,
        }
    }

    /// Atomically set the format after validating it.
    pub fn set_format(&mut self, format: Option<String>) -> Result<()> {
        let previous = self.format.clone();
        self.format = format;
        if let Err(error) = self.validate() {
            self.format = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically set the transliteration language.
    pub fn set_language(&mut self, language: Option<String>) -> Result<()> {
        let previous = self.language.clone();
        self.language = language;
        if let Err(error) = self.validate() {
            self.language = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically set the transliteration country.
    pub fn set_country(&mut self, country: Option<String>) -> Result<()> {
        let previous = self.country.clone();
        self.country = country;
        if let Err(error) = self.validate() {
            self.country = previous;
            return Err(error);
        }
        Ok(())
    }
}

/// Common data-style root attributes.  `None` means that the source omitted
/// the attribute; accessors apply an ODF default without changing this value.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct Attributes {
    pub display_name: Option<String>,
    pub language: Option<String>,
    pub country: Option<String>,
    pub script: Option<String>,
    pub rfc_language_tag: Option<String>,
    pub title: Option<String>,
    pub volatile: Option<bool>,
    pub transliteration: Transliteration,
}

impl Attributes {
    /// Construct empty metadata with all optional attributes omitted.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            display_name: None,
            language: None,
            country: None,
            script: None,
            rfc_language_tag: None,
            title: None,
            volatile: None,
            transliteration: Transliteration {
                format: None,
                language: None,
                country: None,
                style: None,
            },
        }
    }

    /// Validate all metadata fields and their aggregate bound.
    pub fn validate(&self) -> Result<()> {
        validate_optional_text(self.display_name.as_deref(), "data-style display name")?;
        validate_optional_code(
            self.language.as_deref(),
            "data-style language",
            validate_language_code,
        )?;
        validate_optional_code(
            self.country.as_deref(),
            "data-style country",
            validate_country_code,
        )?;
        validate_optional_code(
            self.script.as_deref(),
            "data-style script",
            validate_script_code,
        )?;
        validate_optional_code(
            self.rfc_language_tag.as_deref(),
            "data-style RFC language tag",
            validate_language_tag,
        )?;
        validate_optional_text(self.title.as_deref(), "data-style title")?;
        self.transliteration.validate()?;
        validate_metadata_size([
            self.display_name.as_deref(),
            self.language.as_deref(),
            self.country.as_deref(),
            self.script.as_deref(),
            self.rfc_language_tag.as_deref(),
            self.title.as_deref(),
            self.transliteration.format.as_deref(),
            self.transliteration.language.as_deref(),
            self.transliteration.country.as_deref(),
        ])?;
        Ok(())
    }

    /// Start a checked metadata builder.
    #[must_use]
    pub fn builder() -> AttributesBuilder {
        AttributesBuilder::default()
    }

    /// Build common metadata from borrowed source values.  Validation and
    /// aggregate sizing occur before any string is copied.
    pub fn try_from_borrowed(
        display_name: Option<&str>,
        language: Option<&str>,
        country: Option<&str>,
        script: Option<&str>,
        rfc_language_tag: Option<&str>,
        title: Option<&str>,
        volatile: Option<bool>,
        transliteration: (
            Option<&str>,
            Option<&str>,
            Option<&str>,
            Option<TransliterationStyle>,
        ),
    ) -> Result<Self> {
        validate_optional_text(display_name, "data-style display name")?;
        validate_optional_code(language, "data-style language", validate_language_code)?;
        validate_optional_code(country, "data-style country", validate_country_code)?;
        validate_optional_code(script, "data-style script", validate_script_code)?;
        validate_optional_code(
            rfc_language_tag,
            "data-style RFC language tag",
            validate_language_tag,
        )?;
        validate_optional_text(title, "data-style title")?;
        validate_metadata_size([
            display_name,
            language,
            country,
            script,
            rfc_language_tag,
            title,
            transliteration.0,
            transliteration.1,
            transliteration.2,
        ])?;
        let transliteration = Transliteration::try_from_borrowed(
            transliteration.0,
            transliteration.1,
            transliteration.2,
            transliteration.3,
        )?;
        let value = Self {
            display_name: clone_optional_text(display_name, "data-style display name")?,
            language: clone_optional_code_text(
                language,
                "data-style language",
                validate_language_code,
            )?,
            country: clone_optional_code_text(
                country,
                "data-style country",
                validate_country_code,
            )?,
            script: clone_optional_code_text(script, "data-style script", validate_script_code)?,
            rfc_language_tag: clone_optional_code_text(
                rfc_language_tag,
                "data-style RFC language tag",
                validate_language_tag,
            )?,
            title: clone_optional_text(title, "data-style title")?,
            volatile,
            transliteration,
        };
        value.validate()?;
        Ok(value)
    }

    /// Apply a metadata patch atomically.
    pub fn apply_patch(&mut self, patch: &Patch) -> Result<()> {
        patch.apply(self)
    }
}

/// Checked builder for [`Attributes`].
#[derive(Clone, Debug, Default)]
pub struct AttributesBuilder {
    value: Attributes,
}

impl AttributesBuilder {
    pub fn display_name(self, value: &str) -> Result<Self> {
        let mut this = self;
        this.value.display_name = Some(clone_text(value, "data-style display name")?);
        Ok(this)
    }

    pub fn language(self, value: &str) -> Result<Self> {
        let mut this = self;
        this.value.language = Some(clone_code_text(
            value,
            "data-style language",
            validate_language_code,
        )?);
        Ok(this)
    }

    pub fn country(self, value: &str) -> Result<Self> {
        let mut this = self;
        this.value.country = Some(clone_code_text(
            value,
            "data-style country",
            validate_country_code,
        )?);
        Ok(this)
    }

    pub fn script(self, value: &str) -> Result<Self> {
        let mut this = self;
        this.value.script = Some(clone_code_text(
            value,
            "data-style script",
            validate_script_code,
        )?);
        Ok(this)
    }

    pub fn rfc_language_tag(self, value: &str) -> Result<Self> {
        let mut this = self;
        this.value.rfc_language_tag = Some(clone_code_text(
            value,
            "data-style RFC language tag",
            validate_language_tag,
        )?);
        Ok(this)
    }

    pub fn title(self, value: &str) -> Result<Self> {
        let mut this = self;
        this.value.title = Some(clone_text(value, "data-style title")?);
        Ok(this)
    }

    pub fn volatile(mut self, value: bool) -> Self {
        self.value.volatile = Some(value);
        self
    }

    pub fn transliteration(mut self, value: Transliteration) -> Self {
        self.value.transliteration = value;
        self
    }

    pub fn build(self) -> Result<Attributes> {
        self.value.validate()?;
        Ok(self.value)
    }
}

/// A validated ODF `double` preserving its source lexical spelling.
#[derive(Clone, Debug, PartialEq)]
pub struct Double {
    lexical: String,
    value: f64,
}

impl Double {
    /// Construct a value from an IEEE double and retain its canonical Rust spelling.
    pub fn new(value: f64) -> Result<Self> {
        let lexical = if value.is_nan() {
            "NaN".to_string()
        } else if value == f64::INFINITY {
            "INF".to_string()
        } else if value == f64::NEG_INFINITY {
            "-INF".to_string()
        } else {
            value.to_string()
        };
        Self::from_lexical(&lexical)
    }

    /// Parse an XML Schema double lexical form.
    pub fn from_lexical(value: &str) -> Result<Self> {
        let parsed = parse_double_lexical(value)?;
        let lexical = clone_double_lexical(value)?;
        Ok(Self {
            lexical,
            value: parsed,
        })
    }

    /// Parse borrowed XML Schema double text before allocating retained
    /// lexical storage.
    pub fn try_from_lexical(value: &str) -> Result<Self> {
        let parsed = parse_double_lexical(value)?;
        let lexical = clone_double_lexical(value)?;
        Ok(Self {
            lexical,
            value: parsed,
        })
    }

    /// Return the retained lexical spelling.
    #[must_use]
    pub fn lexical(&self) -> &str {
        &self.lexical
    }

    /// Return the parsed numeric value.
    #[must_use]
    pub const fn value(&self) -> f64 {
        self.value
    }

    /// Validate the retained value.
    pub fn validate(&self) -> Result<()> {
        let reparsed = parse_double_lexical(&self.lexical)?;
        if reparsed.is_nan() != self.value.is_nan()
            || (!reparsed.is_nan() && reparsed != self.value)
        {
            return invalid("ODF double lexical and numeric values disagree");
        }
        Ok(())
    }
}

impl TryFrom<&str> for Double {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::from_lexical(value)
    }
}

impl TryFrom<String> for Double {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        Self::from_lexical(&value)
    }
}

/// A bounded leading or trailing `number:text-with-fillchar` particle.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct Affix {
    pub text: Option<String>,
    pub fill_character: Option<String>,
    pub text_after_fill: Option<String>,
}

impl Affix {
    /// Construct and validate an affix.
    pub fn new(
        text: Option<String>,
        fill_character: Option<String>,
        text_after_fill: Option<String>,
    ) -> Result<Self> {
        let value = Self {
            text,
            fill_character,
            text_after_fill,
        };
        value.validate()?;
        Ok(value)
    }

    /// Construct an affix from borrowed source text after particle preflight.
    pub fn try_from_borrowed(
        text: Option<&str>,
        fill_character: Option<&str>,
        text_after_fill: Option<&str>,
    ) -> Result<Self> {
        validate_optional_text(text, "number-style leading text")?;
        validate_optional_text(fill_character, "number-style fill character")?;
        validate_optional_text(text_after_fill, "number-style trailing text")?;
        if let Some(fill) = fill_character {
            if fill.chars().count() != 1 {
                return invalid("ODF number:fill-character must contain exactly one character");
            }
        }
        if text_after_fill.is_some() && fill_character.is_none() {
            return invalid("ODF number:text after fill requires a fill-character particle");
        }
        if fill_character.is_none() && text.is_some() && text_after_fill.is_some() {
            return invalid("ODF number-style affix cannot contain adjacent number:text particles");
        }
        validate_affix_payload(text, fill_character, text_after_fill)?;
        Self::new(
            clone_optional_text(text, "number-style leading text")?,
            clone_optional_text(fill_character, "number-style fill character")?,
            clone_optional_text(text_after_fill, "number-style trailing text")?,
        )
    }

    /// Construct a plain text affix.
    pub fn text(value: &str) -> Result<Self> {
        Self::try_from_borrowed(Some(value), None, None)
    }

    /// Construct a fill-character affix.
    pub fn fill(value: &str) -> Result<Self> {
        Self::try_from_borrowed(None, Some(value), None)
    }

    /// Validate particle ordering and text bounds.
    pub fn validate(&self) -> Result<()> {
        if self.text.is_none() && self.fill_character.is_none() && self.text_after_fill.is_none() {
            return invalid("ODF number-style affix cannot be empty");
        }
        validate_optional_text(self.text.as_deref(), "number-style leading text")?;
        validate_optional_text(
            self.fill_character.as_deref(),
            "number-style fill character",
        )?;
        validate_optional_text(
            self.text_after_fill.as_deref(),
            "number-style trailing text",
        )?;
        if let Some(fill) = self.fill_character.as_deref() {
            if fill.chars().count() != 1 {
                return invalid("ODF number:fill-character must contain exactly one character");
            }
        }
        if self.text_after_fill.is_some() && self.fill_character.is_none() {
            return invalid("ODF number:text after fill requires a fill-character particle");
        }
        if self.fill_character.is_none() && self.text.is_some() && self.text_after_fill.is_some() {
            return invalid("ODF number-style affix cannot contain adjacent number:text particles");
        }
        validate_affix_payload(
            self.text.as_deref(),
            self.fill_character.as_deref(),
            self.text_after_fill.as_deref(),
        )?;
        Ok(())
    }
}

/// A checked `number:embedded-text` particle.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct EmbeddedText {
    /// One-based position from the right of the integer portion.
    pub position: i64,
    pub text: String,
}

impl EmbeddedText {
    /// Construct an embedded-text particle.
    pub fn new(position: i64, text: &str) -> Result<Self> {
        Self::try_from_borrowed(position, text)
    }

    /// Construct an embedded-text particle from borrowed source text.
    pub fn try_from_borrowed(position: i64, text: &str) -> Result<Self> {
        if position <= 0 {
            return invalid("ODF number:position must be a positive one-based position");
        }
        Ok(Self {
            position,
            text: clone_text(text, "number:embedded-text")?,
        })
    }

    /// Validate the positive display position and bounded text.
    pub fn validate(&self) -> Result<()> {
        if self.position <= 0 {
            return invalid("ODF number:position must be a positive one-based position");
        }
        validate_text(&self.text, "number:embedded-text")
    }
}

/// The decimal `number:number` body shared by number, currency, and percentage roots.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Decimal {
    pub decimal_places: Option<i64>,
    pub min_decimal_places: Option<i64>,
    pub min_integer_digits: Option<i64>,
    pub grouping: Option<bool>,
    pub decimal_replacement: Option<String>,
    pub display_factor: Option<Double>,
    pub embedded_text: Vec<EmbeddedText>,
}

impl Decimal {
    /// Return a default decimal body with all source attributes omitted.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            decimal_places: None,
            min_decimal_places: None,
            min_integer_digits: None,
            grouping: None,
            decimal_replacement: None,
            display_factor: None,
            embedded_text: Vec::new(),
        }
    }

    /// Validate scalar relationships, text values, and bounded children.
    pub fn validate(&self) -> Result<()> {
        if let (Some(minimum), Some(decimal_places)) =
            (self.min_decimal_places, self.decimal_places)
            && minimum > decimal_places
        {
            return invalid("ODF number:min-decimal-places exceeds decimal-places");
        }
        validate_optional_text(
            self.decimal_replacement.as_deref(),
            "number:decimal-replacement",
        )?;
        if let Some(display_factor) = &self.display_factor {
            display_factor.validate()?;
        }
        validate_embedded_texts(&self.embedded_text)
    }

    /// Resolve decimal places against the default table-cell style.
    #[must_use]
    pub fn resolve_decimal_places(&self, inherited: Option<i64>) -> Resolution {
        match self.decimal_places {
            Some(value) => Resolution::Explicit(value),
            None => inherited.map_or(Resolution::Unresolved, Resolution::Inherited),
        }
    }

    /// Resolve minimum decimal places while retaining omitted/inherited state.
    #[must_use]
    pub fn resolve_min_decimal_places(&self, inherited: Option<i64>) -> Resolution {
        if let Some(value) = self.min_decimal_places {
            return Resolution::Explicit(value);
        }
        if decimal_replacement_is_empty(self.decimal_replacement.as_deref()) {
            Resolution::Explicit(0)
        } else {
            self.resolve_decimal_places(inherited)
        }
    }

    /// Validate this body against a table-cell decimal-places default.
    pub fn validate_with_inherited(&self, inherited: Option<i64>) -> Result<()> {
        self.validate()?;
        if self.decimal_places.is_none()
            && let (Some(minimum), Some(default)) = (self.min_decimal_places, inherited)
            && minimum > default
        {
            return invalid("ODF number:min-decimal-places exceeds inherited decimal-places");
        }
        Ok(())
    }

    /// Resolve minimum decimal places and refuse an inherited cross-field
    /// violation before exposing a typed result.
    pub fn resolve_min_decimal_places_checked(&self, inherited: Option<i64>) -> Result<Resolution> {
        self.validate_with_inherited(inherited)?;
        Ok(self.resolve_min_decimal_places(inherited))
    }

    /// ODF's effective grouping default.
    #[must_use]
    pub const fn effective_grouping(&self) -> bool {
        match self.grouping {
            Some(value) => value,
            None => false,
        }
    }

    /// ODF's effective display-factor.  The lexical wrapper retains `1` as a
    /// deliberate default without changing the omitted source field.
    pub fn effective_display_factor(&self) -> Result<Double> {
        self.display_factor
            .clone()
            .map_or_else(|| Double::new(1.0), Ok)
    }

    /// Fallibly append one bounded embedded-text child.
    pub fn try_push_embedded_text(&mut self, value: EmbeddedText) -> Result<()> {
        value.validate()?;
        if self.embedded_text.len() >= MAX_STYLE_BODY_ELEMENTS {
            return invalid("ODF number:embedded-text count exceeds its limit");
        }
        let current_bytes = embedded_text_payload_bytes(&self.embedded_text)?;
        let aggregate = current_bytes
            .checked_add(value.text.len())
            .ok_or_else(|| invalid_error("ODF embedded-text payload size overflow"))?;
        if aggregate > MAX_STYLE_BODY_BYTES {
            return invalid("ODF embedded-text payload exceeds its body byte limit");
        }
        self.embedded_text
            .try_reserve(1)
            .map_err(|source| allocation("ODS embedded-text children", source))?;
        self.embedded_text.push(value);
        Ok(())
    }

    /// Atomically replace decimal places.
    pub fn set_decimal_places(&mut self, value: Option<i64>) -> Result<()> {
        let previous = self.decimal_places;
        self.decimal_places = value;
        if let Err(error) = self.validate() {
            self.decimal_places = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace minimum decimal places.
    pub fn set_min_decimal_places(&mut self, value: Option<i64>) -> Result<()> {
        let previous = self.min_decimal_places;
        self.min_decimal_places = value;
        if let Err(error) = self.validate() {
            self.min_decimal_places = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace minimum integer digits.
    pub fn set_min_integer_digits(&mut self, value: Option<i64>) -> Result<()> {
        let previous = self.min_integer_digits;
        self.min_integer_digits = value;
        if let Err(error) = self.validate() {
            self.min_integer_digits = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace grouping.
    pub fn set_grouping(&mut self, value: Option<bool>) -> Result<()> {
        let previous = self.grouping;
        self.grouping = value;
        if let Err(error) = self.validate() {
            self.grouping = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace decimal-replacement text.
    pub fn set_decimal_replacement(&mut self, value: Option<String>) -> Result<()> {
        let previous = self.decimal_replacement.clone();
        self.decimal_replacement = value;
        if let Err(error) = self.validate() {
            self.decimal_replacement = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace display-factor.
    pub fn set_display_factor(&mut self, value: Option<Double>) -> Result<()> {
        let previous = self.display_factor.clone();
        self.display_factor = value;
        if let Err(error) = self.validate() {
            self.display_factor = previous;
            return Err(error);
        }
        Ok(())
    }
}

/// A scientific `number:scientific-number` body.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Scientific {
    pub decimal_places: Option<i64>,
    pub min_decimal_places: Option<i64>,
    pub min_integer_digits: Option<i64>,
    pub grouping: Option<bool>,
    pub min_exponent_digits: Option<i64>,
    pub exponent_interval: Option<u64>,
    pub forced_exponent_sign: Option<bool>,
}

impl Scientific {
    /// Construct an omitted-attribute scientific body.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            decimal_places: None,
            min_decimal_places: None,
            min_integer_digits: None,
            grouping: None,
            min_exponent_digits: None,
            exponent_interval: None,
            forced_exponent_sign: None,
        }
    }

    /// Validate scientific scalar relationships and positive interval.
    pub fn validate(&self) -> Result<()> {
        if let (Some(minimum), Some(decimal_places)) =
            (self.min_decimal_places, self.decimal_places)
            && minimum > decimal_places
        {
            return invalid("ODF scientific min-decimal-places exceeds decimal-places");
        }
        if self.exponent_interval == Some(0) {
            return invalid("ODF scientific exponent-interval must be positive");
        }
        Ok(())
    }

    /// Resolve decimal places against the default table-cell style.
    #[must_use]
    pub fn resolve_decimal_places(&self, inherited: Option<i64>) -> Resolution {
        match self.decimal_places {
            Some(value) => Resolution::Explicit(value),
            None => inherited.map_or(Resolution::Unresolved, Resolution::Inherited),
        }
    }

    /// Resolve the omitted minimum decimal places according to ODF defaults.
    #[must_use]
    pub fn resolve_min_decimal_places(&self, inherited: Option<i64>) -> Resolution {
        self.min_decimal_places.map_or_else(
            || self.resolve_decimal_places(inherited),
            Resolution::Explicit,
        )
    }

    /// Validate this body against a table-cell decimal-places default.
    pub fn validate_with_inherited(&self, inherited: Option<i64>) -> Result<()> {
        self.validate()?;
        if self.decimal_places.is_none()
            && let (Some(minimum), Some(default)) = (self.min_decimal_places, inherited)
            && minimum > default
        {
            return invalid("ODF scientific min-decimal-places exceeds inherited decimal-places");
        }
        Ok(())
    }

    /// Resolve minimum decimal places and refuse an inherited cross-field
    /// violation before exposing a typed result.
    pub fn resolve_min_decimal_places_checked(&self, inherited: Option<i64>) -> Result<Resolution> {
        self.validate_with_inherited(inherited)?;
        Ok(self.resolve_min_decimal_places(inherited))
    }

    /// ODF's effective grouping default.
    #[must_use]
    pub const fn effective_grouping(&self) -> bool {
        match self.grouping {
            Some(value) => value,
            None => false,
        }
    }

    /// ODF's effective exponent interval.
    #[must_use]
    pub const fn effective_exponent_interval(&self) -> u64 {
        match self.exponent_interval {
            Some(value) => value,
            None => 1,
        }
    }

    /// ODF's effective forced-sign default.
    #[must_use]
    pub const fn effective_forced_exponent_sign(&self) -> bool {
        match self.forced_exponent_sign {
            Some(value) => value,
            None => true,
        }
    }

    /// Atomically set exponent interval.
    pub fn set_exponent_interval(&mut self, value: Option<u64>) -> Result<()> {
        let previous = self.exponent_interval;
        self.exponent_interval = value;
        if let Err(error) = self.validate() {
            self.exponent_interval = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace decimal places.
    pub fn set_decimal_places(&mut self, value: Option<i64>) -> Result<()> {
        let previous = self.decimal_places;
        self.decimal_places = value;
        if let Err(error) = self.validate() {
            self.decimal_places = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace minimum decimal places.
    pub fn set_min_decimal_places(&mut self, value: Option<i64>) -> Result<()> {
        let previous = self.min_decimal_places;
        self.min_decimal_places = value;
        if let Err(error) = self.validate() {
            self.min_decimal_places = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace minimum integer digits.
    pub fn set_min_integer_digits(&mut self, value: Option<i64>) -> Result<()> {
        let previous = self.min_integer_digits;
        self.min_integer_digits = value;
        if let Err(error) = self.validate() {
            self.min_integer_digits = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace grouping.
    pub fn set_grouping(&mut self, value: Option<bool>) -> Result<()> {
        let previous = self.grouping;
        self.grouping = value;
        if let Err(error) = self.validate() {
            self.grouping = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace minimum exponent digits.
    pub fn set_min_exponent_digits(&mut self, value: Option<i64>) -> Result<()> {
        let previous = self.min_exponent_digits;
        self.min_exponent_digits = value;
        if let Err(error) = self.validate() {
            self.min_exponent_digits = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace forced exponent sign.
    pub fn set_forced_exponent_sign(&mut self, value: Option<bool>) -> Result<()> {
        let previous = self.forced_exponent_sign;
        self.forced_exponent_sign = value;
        if let Err(error) = self.validate() {
            self.forced_exponent_sign = previous;
            return Err(error);
        }
        Ok(())
    }
}

/// A fraction `number:fraction` body.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Fraction {
    pub min_numerator_digits: Option<i64>,
    pub min_denominator_digits: Option<i64>,
    pub denominator_value: Option<i64>,
    pub max_denominator_value: Option<u64>,
    pub min_integer_digits: Option<i64>,
    pub grouping: Option<bool>,
}

impl Fraction {
    /// Construct an omitted-attribute fraction body.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            min_numerator_digits: None,
            min_denominator_digits: None,
            denominator_value: None,
            max_denominator_value: None,
            min_integer_digits: None,
            grouping: None,
        }
    }

    /// Validate positive max-denominator values.
    pub fn validate(&self) -> Result<()> {
        if self.max_denominator_value == Some(0) {
            return invalid("ODF fraction max-denominator-value must be positive");
        }
        Ok(())
    }

    /// Return the optional denominator cap without inventing a value.
    #[must_use]
    pub const fn effective_max_denominator_value(&self) -> Option<u64> {
        if self.denominator_value.is_some() {
            None
        } else {
            self.max_denominator_value
        }
    }

    /// ODF's effective grouping default.
    #[must_use]
    pub const fn effective_grouping(&self) -> bool {
        match self.grouping {
            Some(value) => value,
            None => false,
        }
    }

    /// Atomically set max denominator value.
    pub fn set_max_denominator_value(&mut self, value: Option<u64>) -> Result<()> {
        let previous = self.max_denominator_value;
        self.max_denominator_value = value;
        if let Err(error) = self.validate() {
            self.max_denominator_value = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace minimum numerator digits.
    pub fn set_min_numerator_digits(&mut self, value: Option<i64>) -> Result<()> {
        let previous = self.min_numerator_digits;
        self.min_numerator_digits = value;
        if let Err(error) = self.validate() {
            self.min_numerator_digits = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace minimum denominator digits.
    pub fn set_min_denominator_digits(&mut self, value: Option<i64>) -> Result<()> {
        let previous = self.min_denominator_digits;
        self.min_denominator_digits = value;
        if let Err(error) = self.validate() {
            self.min_denominator_digits = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace denominator value.
    pub fn set_denominator_value(&mut self, value: Option<i64>) -> Result<()> {
        let previous = self.denominator_value;
        self.denominator_value = value;
        if let Err(error) = self.validate() {
            self.denominator_value = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace minimum integer digits.
    pub fn set_min_integer_digits(&mut self, value: Option<i64>) -> Result<()> {
        let previous = self.min_integer_digits;
        self.min_integer_digits = value;
        if let Err(error) = self.validate() {
            self.min_integer_digits = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Atomically replace grouping.
    pub fn set_grouping(&mut self, value: Option<bool>) -> Result<()> {
        let previous = self.grouping;
        self.grouping = value;
        if let Err(error) = self.validate() {
            self.grouping = previous;
            return Err(error);
        }
        Ok(())
    }
}

/// One typed number-format particle.
#[derive(Clone, Debug, PartialEq)]
pub enum Format {
    Decimal(Decimal),
    Scientific(Scientific),
    Fraction(Fraction),
}

impl Format {
    /// Return the corresponding number-style family.
    #[must_use]
    pub const fn family(&self) -> Family {
        Family::Number
    }

    /// Validate the selected body.
    pub fn validate(&self) -> Result<()> {
        match self {
            Self::Decimal(value) => value.validate(),
            Self::Scientific(value) => value.validate(),
            Self::Fraction(value) => value.validate(),
        }
    }
}

/// A checked number-style node supporting decimal, scientific, and fraction bodies.
#[derive(Clone, Debug, PartialEq)]
pub struct Number {
    pub name: String,
    pub attributes: Attributes,
    pub leading: Option<Affix>,
    pub format: Option<Format>,
    pub trailing: Option<Affix>,
}

impl Number {
    /// Construct an empty number-style node.
    pub fn new(name: &str) -> Result<Self> {
        validate_name(name, "number style name")?;
        Ok(Self {
            name: clone_text(name, "number style name")?,
            attributes: Attributes::new(),
            leading: None,
            format: None,
            trailing: None,
        })
    }

    /// Construct a named node after preflighting the borrowed name.
    pub fn try_new(name: &str) -> Result<Self> {
        validate_name(name, "number style name")?;
        Self::new(name)
    }

    /// Return the number family.
    #[must_use]
    pub const fn family(&self) -> Family {
        Family::Number
    }

    /// Return the mutable automatic-style selector used by authored graphs.
    #[must_use]
    pub fn selector(&self) -> Selector<'_> {
        Selector::automatic(&self.name, Family::Number)
    }

    /// Return a source-qualified selector for a projected owner.
    #[must_use]
    pub fn selector_in(&self, owner: Owner) -> Selector<'_> {
        Selector {
            owner,
            family: Family::Number,
            name: &self.name,
        }
    }

    /// Validate metadata, particles, and body bounds.
    pub fn validate(&self) -> Result<()> {
        validate_name(&self.name, "number style name")?;
        self.attributes.validate()?;
        if let Some(leading) = &self.leading {
            leading.validate()?;
        }
        if let Some(trailing) = &self.trailing {
            trailing.validate()?;
        }
        if self
            .leading
            .as_ref()
            .and_then(|value| value.fill_character.as_ref())
            .is_some()
            && self
                .trailing
                .as_ref()
                .and_then(|value| value.fill_character.as_ref())
                .is_some()
        {
            return invalid("ODF number-style permits at most one fill-character particle");
        }
        if number_has_adjacent_text(
            self.leading.as_ref(),
            self.format.as_ref(),
            self.trailing.as_ref(),
        ) {
            return invalid("ODF number-style cannot contain adjacent number:text particles");
        }
        if self.format.is_none() && self.trailing.is_some() {
            return invalid("ODF number-style trailing text requires a number body");
        }
        if let Some(format) = &self.format {
            format.validate()?
        }
        if number_particle_count(
            self.leading.as_ref(),
            self.format.as_ref(),
            self.trailing.as_ref(),
        )? > MAX_STYLE_BODY_ELEMENTS
        {
            return invalid("ODS number-style body exceeds its element limit");
        }
        let payload = number_payload_bytes(
            self.leading.as_ref(),
            self.format.as_ref(),
            self.trailing.as_ref(),
        )?;
        if payload > MAX_STYLE_BODY_BYTES {
            return invalid("ODS number-style body exceeds its aggregate byte limit");
        }
        Ok(())
    }

    /// Produce canonical compact XML for this typed node.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let estimated = self.markup_size()?;
        enforce_markup_size(estimated)?;
        let mut output = String::new();
        output
            .try_reserve(estimated)
            .map_err(|source| allocation("ODS number-style markup", source))?;
        write!(
            output,
            "<number:number-style xmlns:number=\"{NUMBER_NS}\" xmlns:style=\"{STYLE_NS}\" style:name=\"{}\"",
            escape_xml_checked(&self.name)?
        )
        .map_err(|_| invalid_error("ODS number-style markup formatting failed"))?;
        append_common_attributes(&mut output, &self.attributes)?;
        output.push('>');
        if let Some(leading) = &self.leading {
            append_affix(&mut output, leading)?;
        }
        if let Some(format) = &self.format {
            append_number_format(&mut output, format)?;
        }
        if let Some(trailing) = &self.trailing {
            append_affix(&mut output, trailing)?;
        }
        output.push_str("</number:number-style>");
        enforce_markup_limit(&output)?;
        Ok(output)
    }

    /// Return the exact canonical markup size after validation without
    /// allocating the serialized fragment.  Package owners use this for
    /// caller-limit preflight before constructing a replacement buffer.
    pub(crate) fn serialized_markup_size(&self) -> Result<usize> {
        self.validate()?;
        self.markup_size()
    }

    /// Atomically replace the number body.
    pub fn set_format(&mut self, format: Option<Format>) -> Result<()> {
        let previous = self.format.clone();
        self.format = format;
        if let Err(error) = self.validate() {
            self.format = previous;
            return Err(error);
        }
        Ok(())
    }

    fn markup_size(&self) -> Result<usize> {
        let mut size = root_markup_size("number-style", &self.name)?;
        size = checked_add_size(size, common_attributes_markup_size(&self.attributes)?)?;
        size = checked_add_size(size, 1)?; // root close bracket
        if let Some(leading) = &self.leading {
            size = checked_add_size(size, affix_markup_size(leading)?)?;
        }
        if let Some(format) = &self.format {
            size = checked_add_size(size, number_format_markup_size(format)?)?;
        }
        if let Some(trailing) = &self.trailing {
            size = checked_add_size(size, affix_markup_size(trailing)?)?;
        }
        checked_add_size(size, "</number:number-style>".len())
    }
}

/// Checked builder for [`Number`].
#[derive(Clone, Debug)]
pub struct NumberBuilder {
    node: Number,
}

impl NumberBuilder {
    /// Start a decimal number style.
    pub fn decimal(name: &str) -> Result<Self> {
        Self::with_format(name, Format::Decimal(Decimal::new()))
    }

    /// Borrowed, fallible decimal constructor.
    pub fn try_decimal(name: &str) -> Result<Self> {
        validate_name(name, "number style name")?;
        Self::decimal(name)
    }

    /// Start a scientific number style.
    pub fn scientific(name: &str) -> Result<Self> {
        Self::with_format(name, Format::Scientific(Scientific::new()))
    }

    /// Borrowed, fallible scientific constructor.
    pub fn try_scientific(name: &str) -> Result<Self> {
        validate_name(name, "number style name")?;
        Self::scientific(name)
    }

    /// Start a fraction number style.
    pub fn fraction(name: &str) -> Result<Self> {
        Self::with_format(name, Format::Fraction(Fraction::new()))
    }

    /// Borrowed, fallible fraction constructor.
    pub fn try_fraction(name: &str) -> Result<Self> {
        validate_name(name, "number style name")?;
        Self::fraction(name)
    }

    /// Start a style with an explicitly selected body.
    pub fn with_format(name: &str, format: Format) -> Result<Self> {
        validate_name(name, "number style name")?;
        format.validate()?;
        Ok(Self {
            node: Number {
                name: clone_text(name, "number style name")?,
                attributes: Attributes::new(),
                leading: None,
                format: Some(format),
                trailing: None,
            },
        })
    }

    pub fn attributes(mut self, value: Attributes) -> Self {
        self.node.attributes = value;
        self
    }

    pub fn leading(mut self, value: Affix) -> Self {
        self.node.leading = Some(value);
        self
    }

    pub fn trailing(mut self, value: Affix) -> Self {
        self.node.trailing = Some(value);
        self
    }

    pub fn decimal_places(mut self, value: i64) -> Result<Self> {
        match self.node.format.as_mut() {
            Some(Format::Decimal(number)) => number.set_decimal_places(Some(value))?,
            Some(Format::Scientific(number)) => number.set_decimal_places(Some(value))?,
            Some(Format::Fraction(_)) => {
                return invalid("number:decimal-places is not valid on a fraction body");
            },
            None => return invalid("number:decimal-places requires a number body"),
        }
        Ok(self)
    }

    pub fn min_decimal_places(mut self, value: i64) -> Result<Self> {
        match self.node.format.as_mut() {
            Some(Format::Decimal(number)) => number.set_min_decimal_places(Some(value))?,
            Some(Format::Scientific(number)) => number.set_min_decimal_places(Some(value))?,
            Some(Format::Fraction(_)) => {
                return invalid("number:min-decimal-places is not valid on a fraction body");
            },
            None => return invalid("number:min-decimal-places requires a number body"),
        }
        Ok(self)
    }

    pub fn min_integer_digits(mut self, value: i64) -> Result<Self> {
        match self.node.format.as_mut() {
            Some(Format::Decimal(number)) => number.set_min_integer_digits(Some(value))?,
            Some(Format::Scientific(number)) => number.set_min_integer_digits(Some(value))?,
            Some(Format::Fraction(number)) => number.set_min_integer_digits(Some(value))?,
            None => return invalid("number:min-integer-digits requires a number body"),
        }
        Ok(self)
    }

    pub fn grouping(mut self, value: bool) -> Result<Self> {
        match self.node.format.as_mut() {
            Some(Format::Decimal(number)) => number.set_grouping(Some(value))?,
            Some(Format::Scientific(number)) => number.set_grouping(Some(value))?,
            Some(Format::Fraction(number)) => number.set_grouping(Some(value))?,
            None => return invalid("number:grouping requires a number body"),
        }
        Ok(self)
    }

    pub fn decimal_replacement(self, value: &str) -> Result<Self> {
        validate_text(value, "number:decimal-replacement")?;
        let mut this = self;
        if let Some(Format::Decimal(number)) = this.node.format.as_mut() {
            number
                .set_decimal_replacement(Some(clone_text(value, "number:decimal-replacement")?))?;
        } else {
            return invalid("number:decimal-replacement requires a decimal body");
        }
        Ok(this)
    }

    pub fn display_factor(mut self, value: Double) -> Result<Self> {
        if let Some(Format::Decimal(number)) = self.node.format.as_mut() {
            number.set_display_factor(Some(value))?;
        } else {
            return invalid("number:display-factor requires a decimal body");
        }
        Ok(self)
    }

    pub fn min_exponent_digits(mut self, value: i64) -> Result<Self> {
        if let Some(Format::Scientific(number)) = self.node.format.as_mut() {
            number.set_min_exponent_digits(Some(value))?;
        } else {
            return invalid("number:min-exponent-digits requires a scientific body");
        }
        Ok(self)
    }

    pub fn exponent_interval(mut self, value: u64) -> Result<Self> {
        if let Some(Format::Scientific(number)) = self.node.format.as_mut() {
            number.set_exponent_interval(Some(value))?;
        } else {
            return invalid("number:exponent-interval requires a scientific body");
        }
        Ok(self)
    }

    pub fn forced_exponent_sign(mut self, value: bool) -> Result<Self> {
        if let Some(Format::Scientific(number)) = self.node.format.as_mut() {
            number.set_forced_exponent_sign(Some(value))?;
        } else {
            return invalid("number:forced-exponent-sign requires a scientific body");
        }
        Ok(self)
    }

    pub fn min_numerator_digits(mut self, value: i64) -> Result<Self> {
        if let Some(Format::Fraction(number)) = self.node.format.as_mut() {
            number.set_min_numerator_digits(Some(value))?;
        } else {
            return invalid("number:min-numerator-digits requires a fraction body");
        }
        Ok(self)
    }

    pub fn min_denominator_digits(mut self, value: i64) -> Result<Self> {
        if let Some(Format::Fraction(number)) = self.node.format.as_mut() {
            number.set_min_denominator_digits(Some(value))?;
        } else {
            return invalid("number:min-denominator-digits requires a fraction body");
        }
        Ok(self)
    }

    pub fn denominator_value(mut self, value: i64) -> Result<Self> {
        if let Some(Format::Fraction(number)) = self.node.format.as_mut() {
            number.set_denominator_value(Some(value))?;
        } else {
            return invalid("number:denominator-value requires a fraction body");
        }
        Ok(self)
    }

    pub fn max_denominator_value(mut self, value: u64) -> Result<Self> {
        if let Some(Format::Fraction(number)) = self.node.format.as_mut() {
            number.set_max_denominator_value(Some(value))?;
        } else {
            return invalid("number:max-denominator-value requires a fraction body");
        }
        Ok(self)
    }

    /// Fallibly append an embedded-text particle to a decimal body.
    pub fn embedded_text(mut self, value: EmbeddedText) -> Result<Self> {
        match self.node.format.as_mut() {
            Some(Format::Decimal(number)) => number.try_push_embedded_text(value)?,
            Some(Format::Scientific(_)) | Some(Format::Fraction(_)) | None => {
                return invalid("number:embedded-text requires a decimal number body");
            },
        }
        Ok(self)
    }

    /// Build and validate the node.
    pub fn build(self) -> Result<Number> {
        self.node.validate()?;
        Ok(self.node)
    }
}

/// The supported typed data-style root bodies.
#[derive(Clone, Debug, PartialEq)]
pub enum Body {
    Date,
    Time { decimal_places: Option<i64> },
    Currency { symbol: String, number: Decimal },
    Percentage { number: Decimal },
    Boolean,
}

impl Body {
    /// Return the root family represented by this body.
    #[must_use]
    pub const fn family(&self) -> Family {
        match self {
            Self::Date => Family::Date,
            Self::Time { .. } => Family::Time,
            Self::Currency { .. } => Family::Currency,
            Self::Percentage { .. } => Family::Percentage,
            Self::Boolean => Family::Boolean,
        }
    }

    /// Validate body fields.
    pub fn validate(&self) -> Result<()> {
        match self {
            Self::Date | Self::Boolean => Ok(()),
            Self::Time { .. } => Ok(()),
            Self::Currency { symbol, number } => {
                validate_text(symbol, "currency symbol")?;
                number.validate()?;
                if decimal_particle_count(number)?
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("ODS currency body element count overflow"))?
                    > MAX_STYLE_BODY_ELEMENTS
                {
                    return invalid("ODS currency body exceeds its element limit");
                }
                let payload = symbol
                    .len()
                    .checked_add(decimal_payload_bytes(number)?)
                    .and_then(|value| value.checked_add(1))
                    .ok_or_else(|| invalid_error("ODS currency body size overflow"))?;
                if payload > MAX_STYLE_BODY_BYTES {
                    return invalid("ODS currency body exceeds its aggregate byte limit");
                }
                Ok(())
            },
            Self::Percentage { number } => {
                number.validate()?;
                if decimal_particle_count(number)?
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("ODS percentage body element count overflow"))?
                    > MAX_STYLE_BODY_ELEMENTS
                {
                    return invalid("ODS percentage body exceeds its element limit");
                }
                let payload = decimal_payload_bytes(number)?
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("ODS percentage body size overflow"))?;
                if payload > MAX_STYLE_BODY_BYTES {
                    return invalid("ODS percentage body exceeds its aggregate byte limit");
                }
                Ok(())
            },
        }
    }
}

/// A typed date/time/currency/percentage/boolean data-style node.
#[derive(Clone, Debug, PartialEq)]
pub struct Data {
    pub name: String,
    pub attributes: Attributes,
    pub body: Body,
}

impl Data {
    /// Construct a typed root.
    pub fn new(name: &str, body: Body) -> Result<Self> {
        validate_name(name, "data style name")?;
        body.validate()?;
        Ok(Self {
            name: clone_text(name, "data style name")?,
            attributes: Attributes::new(),
            body,
        })
    }

    /// Construct a named typed root after preflighting the borrowed name.
    pub fn try_new(name: &str, body: Body) -> Result<Self> {
        validate_name(name, "data style name")?;
        Self::new(name, body)
    }

    /// Return the family derived from the body variant.
    #[must_use]
    pub const fn family(&self) -> Family {
        self.body.family()
    }

    /// Return the mutable automatic-style selector used by authored graphs.
    #[must_use]
    pub fn selector(&self) -> Selector<'_> {
        Selector::automatic(&self.name, self.family())
    }

    /// Return a source-qualified selector for a projected owner.
    #[must_use]
    pub fn selector_in(&self, owner: Owner) -> Selector<'_> {
        Selector {
            owner,
            family: self.family(),
            name: &self.name,
        }
    }

    /// Validate metadata and body.
    pub fn validate(&self) -> Result<()> {
        validate_name(&self.name, "data style name")?;
        self.attributes.validate()?;
        self.body.validate()
    }

    /// Produce canonical compact XML for the closed typed root.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let estimated = self.markup_size()?;
        enforce_markup_size(estimated)?;
        let mut output = String::new();
        output
            .try_reserve(estimated)
            .map_err(|source| allocation("ODS data-style markup", source))?;
        let root = match self.family() {
            Family::Date => "date-style",
            Family::Time => "time-style",
            Family::Currency => "currency-style",
            Family::Percentage => "percentage-style",
            Family::Boolean => "boolean-style",
            Family::Number | Family::Text => {
                return invalid("typed data-style body has no number/text root");
            },
        };
        write!(
            output,
            "<number:{root} xmlns:number=\"{NUMBER_NS}\" xmlns:style=\"{STYLE_NS}\" style:name=\"{}\"",
            escape_xml_checked(&self.name)?
        )
        .map_err(|_| invalid_error("ODS data-style markup formatting failed"))?;
        append_common_attributes(&mut output, &self.attributes)?;
        output.push('>');
        match &self.body {
            Body::Date => append_legacy_date_body(&mut output),
            Body::Time { decimal_places } => {
                append_legacy_time_body(&mut output, *decimal_places)?;
            },
            Body::Currency { symbol, number } => {
                output.push_str("<number:currency-symbol>");
                output.push_str(&escape_xml_checked(symbol)?);
                output.push_str("</number:currency-symbol>");
                append_number_body(&mut output, "number", number)?;
            },
            Body::Percentage { number } => {
                append_number_body(&mut output, "number", number)?;
                output.push_str("<number:text>%</number:text>");
            },
            Body::Boolean => output.push_str("<number:boolean/>"),
        }
        output.push_str("</number:");
        output.push_str(root);
        output.push('>');
        enforce_markup_limit(&output)?;
        Ok(output)
    }

    /// Return the exact canonical markup size after validation without
    /// allocating the serialized fragment.  Package owners use this for
    /// caller-limit preflight before constructing a replacement buffer.
    pub(crate) fn serialized_markup_size(&self) -> Result<usize> {
        self.validate()?;
        self.markup_size()
    }

    fn markup_size(&self) -> Result<usize> {
        let root = match self.family() {
            Family::Date => "date-style",
            Family::Time => "time-style",
            Family::Currency => "currency-style",
            Family::Percentage => "percentage-style",
            Family::Boolean => "boolean-style",
            Family::Number | Family::Text => {
                return invalid("typed data-style body has no number/text root");
            },
        };
        let mut size = root_markup_size(root, &self.name)?;
        size = checked_add_size(size, common_attributes_markup_size(&self.attributes)?)?;
        size = checked_add_size(size, 1)?; // root close bracket
        let body = match &self.body {
            Body::Date => legacy_date_markup_size(),
            Body::Time { decimal_places } => legacy_time_markup_size(*decimal_places)?,
            Body::Currency { symbol, number } => {
                let mut body_size = tagged_text_markup_size("currency-symbol", symbol)?;
                body_size =
                    checked_add_size(body_size, number_body_markup_size("number", number)?)?;
                body_size
            },
            Body::Percentage { number } => {
                let mut body_size = number_body_markup_size("number", number)?;
                body_size = checked_add_size(body_size, tagged_text_markup_size("text", "%")?)?;
                body_size
            },
            Body::Boolean => "<number:boolean/>".len(),
        };
        size = checked_add_size(size, body)?;
        let closing = root
            .len()
            .checked_add("</number:>".len())
            .ok_or_else(|| invalid_error("ODS data-style closing tag size overflow"))?;
        checked_add_size(size, closing)
    }

    /// Atomically replace root metadata.
    pub fn set_attributes(&mut self, value: Attributes) -> Result<()> {
        let previous = self.attributes.clone();
        self.attributes = value;
        if let Err(error) = self.validate() {
            self.attributes = previous;
            return Err(error);
        }
        Ok(())
    }
}

/// Checked builder for existing-family data-style roots.
#[derive(Clone, Debug)]
pub struct DataBuilder {
    node: Data,
}

impl DataBuilder {
    pub fn date(name: &str) -> Result<Self> {
        Ok(Self {
            node: Data::new(name, Body::Date)?,
        })
    }

    pub fn try_date(name: &str) -> Result<Self> {
        validate_name(name, "data style name")?;
        Self::date(name)
    }

    pub fn time(name: &str) -> Result<Self> {
        Ok(Self {
            node: Data::new(
                name,
                Body::Time {
                    decimal_places: None,
                },
            )?,
        })
    }

    pub fn try_time(name: &str) -> Result<Self> {
        validate_name(name, "data style name")?;
        Self::time(name)
    }

    pub fn currency(name: &str, symbol: &str) -> Result<Self> {
        validate_name(name, "data style name")?;
        validate_text(symbol, "currency symbol")?;
        Ok(Self {
            node: Data::new(
                name,
                Body::Currency {
                    symbol: clone_text(symbol, "currency symbol")?,
                    number: Decimal::new(),
                },
            )?,
        })
    }

    pub fn try_currency(name: &str, symbol: &str) -> Result<Self> {
        validate_name(name, "data style name")?;
        validate_text(symbol, "currency symbol")?;
        Self::currency(name, symbol)
    }

    pub fn percentage(name: &str) -> Result<Self> {
        Ok(Self {
            node: Data::new(
                name,
                Body::Percentage {
                    number: Decimal::new(),
                },
            )?,
        })
    }

    pub fn try_percentage(name: &str) -> Result<Self> {
        validate_name(name, "data style name")?;
        Self::percentage(name)
    }

    pub fn boolean(name: &str) -> Result<Self> {
        Ok(Self {
            node: Data::new(name, Body::Boolean)?,
        })
    }

    pub fn try_boolean(name: &str) -> Result<Self> {
        validate_name(name, "data style name")?;
        Self::boolean(name)
    }

    pub fn attributes(mut self, value: Attributes) -> Self {
        self.node.attributes = value;
        self
    }

    pub fn decimal_places(mut self, value: i64) -> Result<Self> {
        match &mut self.node.body {
            Body::Time { decimal_places } => *decimal_places = Some(value),
            Body::Currency { number, .. } | Body::Percentage { number } => {
                number.set_decimal_places(Some(value))?
            },
            Body::Date | Body::Boolean => {
                return invalid("number:decimal-places is not valid on this data-style body");
            },
        }
        Ok(self)
    }

    pub fn min_decimal_places(mut self, value: i64) -> Result<Self> {
        match &mut self.node.body {
            Body::Currency { number, .. } | Body::Percentage { number } => {
                number.set_min_decimal_places(Some(value))?;
            },
            Body::Date | Body::Time { .. } | Body::Boolean => {
                return invalid("number:min-decimal-places requires a decimal data-style body");
            },
        }
        Ok(self)
    }

    pub fn min_integer_digits(mut self, value: i64) -> Result<Self> {
        match &mut self.node.body {
            Body::Currency { number, .. } | Body::Percentage { number } => {
                number.set_min_integer_digits(Some(value))?;
            },
            Body::Date | Body::Time { .. } | Body::Boolean => {
                return invalid("number:min-integer-digits requires a decimal data-style body");
            },
        }
        Ok(self)
    }

    pub fn grouping(mut self, value: bool) -> Result<Self> {
        match &mut self.node.body {
            Body::Currency { number, .. } | Body::Percentage { number } => {
                number.set_grouping(Some(value))?
            },
            Body::Date | Body::Time { .. } | Body::Boolean => {
                return invalid("number:grouping requires a decimal data-style body");
            },
        }
        Ok(self)
    }

    pub fn decimal_replacement(self, value: &str) -> Result<Self> {
        validate_text(value, "number:decimal-replacement")?;
        let mut this = self;
        match &mut this.node.body {
            Body::Currency { number, .. } | Body::Percentage { number } => {
                number.set_decimal_replacement(Some(clone_text(
                    value,
                    "number:decimal-replacement",
                )?))?;
            },
            Body::Date | Body::Time { .. } | Body::Boolean => {
                return invalid("number:decimal-replacement requires a decimal data-style body");
            },
        }
        Ok(this)
    }

    pub fn display_factor(mut self, value: Double) -> Result<Self> {
        match &mut self.node.body {
            Body::Currency { number, .. } | Body::Percentage { number } => {
                number.set_display_factor(Some(value))?;
            },
            Body::Date | Body::Time { .. } | Body::Boolean => {
                return invalid("number:display-factor requires a decimal data-style body");
            },
        }
        Ok(self)
    }

    /// Fallibly append one embedded-text child to currency/percentage decimal bodies.
    pub fn embedded_text(mut self, value: EmbeddedText) -> Result<Self> {
        match &mut self.node.body {
            Body::Currency { number, .. } | Body::Percentage { number } => {
                number.try_push_embedded_text(value)?
            },
            Body::Date | Body::Time { .. } | Body::Boolean => {
                return invalid("number:embedded-text requires a decimal data-style body");
            },
        }
        Ok(self)
    }

    pub fn build(self) -> Result<Data> {
        self.node.validate()?;
        Ok(self.node)
    }
}

/// Metadata-only representation of an unsupported body retained by a catalog.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum OpaqueBody {
    Preserved,
}

/// A source-catalog entry whose body is intentionally opaque.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Opaque {
    pub owner: Owner,
    pub name: String,
    pub family: Family,
    pub attributes: Attributes,
    pub body: OpaqueBody,
}

impl Opaque {
    /// Validate metadata and selector identity.
    pub fn validate(&self) -> Result<()> {
        validate_name(&self.name, "opaque data style name")?;
        self.attributes.validate()
    }

    /// Return the source-qualified selector.
    #[must_use]
    pub fn selector(&self) -> Selector<'_> {
        Selector {
            owner: self.owner,
            family: self.family,
            name: &self.name,
        }
    }
}

/// Metadata-only representation of an opaque `number:text-style` root.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TextEntry {
    pub owner: Owner,
    pub name: String,
    pub attributes: Attributes,
    pub body: OpaqueBody,
}

impl TextEntry {
    /// Validate metadata and the source-qualified text-style identity.
    pub fn validate(&self) -> Result<()> {
        validate_name(&self.name, "text data style name")?;
        self.attributes.validate()
    }

    /// Return the source-qualified text-style selector.
    #[must_use]
    pub fn selector(&self) -> Selector<'_> {
        Selector {
            owner: self.owner,
            family: Family::Text,
            name: &self.name,
        }
    }
}

/// A typed or opaque catalog entry.
#[derive(Clone, Debug, PartialEq)]
pub enum Entry {
    /// A typed number style paired with its source owner.
    Number(Owned<Number>),
    /// A typed existing-family style paired with its source owner.
    Existing(Owned<Data>),
    Opaque(Opaque),
    Text(TextEntry),
}

/// A typed value paired with the source owner that supplied it.
#[derive(Clone, Debug, PartialEq)]
pub struct Owned<T> {
    pub owner: Owner,
    pub value: T,
}

impl<T> Owned<T> {
    /// Pair a typed value with its source owner.
    #[must_use]
    pub const fn new(owner: Owner, value: T) -> Self {
        Self { owner, value }
    }
}

impl Entry {
    /// Return the source family.
    #[must_use]
    pub fn family(&self) -> Family {
        match self {
            Self::Number(_) => Family::Number,
            Self::Existing(value) => value.value.family(),
            Self::Opaque(value) => value.family,
            Self::Text(_) => Family::Text,
        }
    }

    /// Return the source owner.
    #[must_use]
    pub fn owner(&self) -> Owner {
        match self {
            Self::Number(value) => value.owner,
            Self::Existing(value) => value.owner,
            Self::Opaque(value) => value.owner,
            Self::Text(value) => value.owner,
        }
    }

    /// Return the entry name.
    #[must_use]
    pub fn name(&self) -> &str {
        match self {
            Self::Number(value) => &value.value.name,
            Self::Existing(value) => &value.value.name,
            Self::Opaque(value) => &value.name,
            Self::Text(value) => &value.name,
        }
    }

    /// Return a source-qualified selector for this entry.
    #[must_use]
    pub fn selector(&self) -> Selector<'_> {
        Selector {
            owner: self.owner(),
            family: self.family(),
            name: self.name(),
        }
    }

    /// Validate the selected typed or opaque entry before a source operation.
    pub fn validate(&self) -> Result<()> {
        match self {
            Self::Number(value) => value.value.validate(),
            Self::Existing(value) => value.value.validate(),
            Self::Opaque(value) => value.validate(),
            Self::Text(value) => value.validate(),
        }
    }

    /// Pair a typed number style with a source owner.
    #[must_use]
    pub const fn number_at(owner: Owner, value: Number) -> Self {
        Self::Number(Owned::new(owner, value))
    }

    /// Pair an existing-family typed style with a source owner.
    #[must_use]
    pub const fn existing_at(owner: Owner, value: Data) -> Self {
        Self::Existing(Owned::new(owner, value))
    }
}

/// A decimal-places lookup that retains explicit, inherited, and unresolved states.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Resolution {
    Explicit(i64),
    Inherited(i64),
    Unresolved,
}

impl Resolution {
    /// Return the resolved number, if one exists.
    #[must_use]
    pub const fn value(self) -> Option<i64> {
        match self {
            Self::Explicit(value) | Self::Inherited(value) => Some(value),
            Self::Unresolved => None,
        }
    }

    /// Whether the value came from the source body.
    #[must_use]
    pub const fn is_explicit(self) -> bool {
        matches!(self, Self::Explicit(_))
    }

    /// Whether the value came from a table-cell default.
    #[must_use]
    pub const fn is_inherited(self) -> bool {
        matches!(self, Self::Inherited(_))
    }
}

/// The operation applied to one optional metadata attribute.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub enum Op<T> {
    #[default]
    Keep,
    Set(T),
    Clear,
}

/// A source-preserving patch for common data-style attributes.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct Patch {
    pub display_name: Op<String>,
    pub language: Op<String>,
    pub country: Op<String>,
    pub script: Op<String>,
    pub rfc_language_tag: Op<String>,
    pub title: Op<String>,
    pub volatile: Op<bool>,
    pub transliteration_format: Op<String>,
    pub transliteration_language: Op<String>,
    pub transliteration_country: Op<String>,
    pub transliteration_style: Op<TransliterationStyle>,
}

impl Patch {
    /// Validate only patch values, without allocating a candidate metadata value.
    pub fn validate(&self) -> Result<()> {
        validate_patch_text(&self.display_name, "data-style display name")?;
        validate_patch_code(
            &self.language,
            "data-style language",
            validate_language_code,
        )?;
        validate_patch_code(&self.country, "data-style country", validate_country_code)?;
        validate_patch_code(&self.script, "data-style script", validate_script_code)?;
        validate_patch_code(
            &self.rfc_language_tag,
            "data-style RFC language tag",
            validate_language_tag,
        )?;
        validate_patch_text(&self.title, "data-style title")?;
        validate_patch_text(&self.transliteration_format, "transliteration format")?;
        validate_patch_code(
            &self.transliteration_language,
            "transliteration language",
            validate_country_code,
        )?;
        validate_patch_code(
            &self.transliteration_country,
            "transliteration country",
            validate_country_code,
        )?;
        if let Op::Set(value) = &self.transliteration_format {
            validate_transliteration_format(value)?;
        }
        validate_patch_metadata_size(self)?;
        Ok(())
    }

    /// Apply the patch after borrowed preflight, leaving the target untouched on error.
    pub fn apply(&self, target: &mut Attributes) -> Result<()> {
        self.validate()?;
        validate_patched_metadata_size(target, self)?;
        let mut candidate = target.clone();
        apply_patch(&mut candidate.display_name, &self.display_name);
        apply_patch(&mut candidate.language, &self.language);
        apply_patch(&mut candidate.country, &self.country);
        apply_patch(&mut candidate.script, &self.script);
        apply_patch(&mut candidate.rfc_language_tag, &self.rfc_language_tag);
        apply_patch(&mut candidate.title, &self.title);
        apply_patch(&mut candidate.volatile, &self.volatile);
        apply_patch(
            &mut candidate.transliteration.format,
            &self.transliteration_format,
        );
        apply_patch(
            &mut candidate.transliteration.language,
            &self.transliteration_language,
        );
        apply_patch(
            &mut candidate.transliteration.country,
            &self.transliteration_country,
        );
        apply_patch(
            &mut candidate.transliteration.style,
            &self.transliteration_style,
        );
        candidate.validate()?;
        *target = candidate;
        Ok(())
    }

    pub fn set_display_name(mut self, value: &str) -> Result<Self> {
        self.display_name = Op::Set(clone_text(value, "data-style display name")?);
        Ok(self)
    }

    pub fn clear_display_name(mut self) -> Self {
        self.display_name = Op::Clear;
        self
    }

    pub fn set_language(mut self, value: &str) -> Result<Self> {
        self.language = Op::Set(clone_code_text(
            value,
            "data-style language",
            validate_language_code,
        )?);
        Ok(self)
    }

    pub fn clear_language(mut self) -> Self {
        self.language = Op::Clear;
        self
    }

    pub fn set_country(mut self, value: &str) -> Result<Self> {
        self.country = Op::Set(clone_code_text(
            value,
            "data-style country",
            validate_country_code,
        )?);
        Ok(self)
    }

    pub fn clear_country(mut self) -> Self {
        self.country = Op::Clear;
        self
    }

    pub fn set_script(mut self, value: &str) -> Result<Self> {
        self.script = Op::Set(clone_code_text(
            value,
            "data-style script",
            validate_script_code,
        )?);
        Ok(self)
    }

    pub fn clear_script(mut self) -> Self {
        self.script = Op::Clear;
        self
    }

    pub fn set_rfc_language_tag(mut self, value: &str) -> Result<Self> {
        self.rfc_language_tag = Op::Set(clone_code_text(
            value,
            "data-style RFC language tag",
            validate_language_tag,
        )?);
        Ok(self)
    }

    pub fn clear_rfc_language_tag(mut self) -> Self {
        self.rfc_language_tag = Op::Clear;
        self
    }

    pub fn set_title(mut self, value: &str) -> Result<Self> {
        self.title = Op::Set(clone_text(value, "data-style title")?);
        Ok(self)
    }

    pub fn clear_title(mut self) -> Self {
        self.title = Op::Clear;
        self
    }

    pub fn set_volatile(mut self, value: bool) -> Self {
        self.volatile = Op::Set(value);
        self
    }

    pub fn clear_volatile(mut self) -> Self {
        self.volatile = Op::Clear;
        self
    }

    pub fn set_transliteration_format(mut self, value: &str) -> Result<Self> {
        validate_transliteration_format(value)?;
        self.transliteration_format = Op::Set(clone_text(value, "transliteration format")?);
        Ok(self)
    }

    pub fn clear_transliteration_format(mut self) -> Self {
        self.transliteration_format = Op::Clear;
        self
    }

    pub fn set_transliteration_language(mut self, value: &str) -> Result<Self> {
        self.transliteration_language = Op::Set(clone_code_text(
            value,
            "transliteration language",
            validate_country_code,
        )?);
        Ok(self)
    }

    pub fn clear_transliteration_language(mut self) -> Self {
        self.transliteration_language = Op::Clear;
        self
    }

    pub fn set_transliteration_country(mut self, value: &str) -> Result<Self> {
        self.transliteration_country = Op::Set(clone_code_text(
            value,
            "transliteration country",
            validate_country_code,
        )?);
        Ok(self)
    }

    pub fn clear_transliteration_country(mut self) -> Self {
        self.transliteration_country = Op::Clear;
        self
    }

    pub fn set_transliteration_style(mut self, value: TransliterationStyle) -> Self {
        self.transliteration_style = Op::Set(value);
        self
    }

    pub fn clear_transliteration_style(mut self) -> Self {
        self.transliteration_style = Op::Clear;
        self
    }
}

/// An authored graph of additive number/data styles.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Graph {
    pub number_styles: Vec<Number>,
    pub data_styles: Vec<Data>,
}

impl Graph {
    /// Start a checked graph builder.
    #[must_use]
    pub fn builder() -> GraphBuilder {
        GraphBuilder::default()
    }

    /// Validate node bounds, names, families, and all child particles.
    pub fn validate(&self) -> Result<()> {
        let total = self
            .number_styles
            .len()
            .checked_add(self.data_styles.len())
            .ok_or_else(|| invalid_error("ODS style graph node count overflow"))?;
        if total > MAX_STYLE_GRAPH_NODES {
            return invalid("ODS extended style graph exceeds its node limit");
        }
        let mut number_names = BTreeSet::new();
        for style in &self.number_styles {
            style.validate()?;
            if !number_names.insert(style.name.as_str()) {
                return invalid(format!("duplicate ODS number style '{}'", style.name));
            }
        }
        let mut data_names = BTreeSet::new();
        for style in &self.data_styles {
            style.validate()?;
            let key = (style.family(), style.name.as_str());
            if !data_names.insert(key) {
                return invalid(format!("duplicate ODS data style '{}'", style.name));
            }
        }
        Ok(())
    }

    /// Produce concatenated canonical XML for the graph.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let estimated = self.estimated_markup_size()?;
        enforce_graph_markup_size(estimated)?;
        let mut output = String::new();
        output
            .try_reserve(estimated)
            .map_err(|source| allocation("ODS extended style graph markup", source))?;
        for style in &self.number_styles {
            output.push_str(&style.to_xml()?);
        }
        for style in &self.data_styles {
            output.push_str(&style.to_xml()?);
        }
        enforce_graph_markup_size(output.len())?;
        Ok(output)
    }

    fn estimated_markup_size(&self) -> Result<usize> {
        let mut size = 0usize;
        for style in &self.number_styles {
            size = checked_add_size(size, style.markup_size()?)?;
        }
        for style in &self.data_styles {
            size = checked_add_size(size, style.markup_size()?)?;
        }
        Ok(size)
    }
}

/// Fallible builder for [`Graph`].
#[derive(Clone, Debug, Default)]
pub struct GraphBuilder {
    number_styles: Vec<Number>,
    data_styles: Vec<Data>,
}

impl GraphBuilder {
    /// Fallibly append a number style.  Validation happens again at `build` so
    /// direct struct construction remains safe and source compatible.
    pub fn number_style(&mut self, style: Number) -> Result<&mut Self> {
        self.try_number_style(style)
    }

    /// Fallibly append a builder-produced number style.
    pub fn number_style_builder(&mut self, style: NumberBuilder) -> Result<&mut Self> {
        self.try_number_style(style.build()?)
    }

    /// Fallibly append a data style.
    pub fn data_style(&mut self, style: Data) -> Result<&mut Self> {
        self.try_data_style(style)
    }

    /// Fallibly append a builder-produced data style.
    pub fn data_style_builder(&mut self, style: DataBuilder) -> Result<&mut Self> {
        self.try_data_style(style.build()?)
    }

    fn try_number_style(&mut self, style: Number) -> Result<&mut Self> {
        style.validate()?;
        if self
            .number_styles
            .iter()
            .any(|existing| existing.name == style.name)
        {
            return invalid(format!("duplicate ODS number style '{}'", style.name));
        }
        if self.number_styles.len() + self.data_styles.len() >= MAX_STYLE_GRAPH_NODES {
            return invalid("ODS extended style graph exceeds its node limit");
        }
        self.number_styles
            .try_reserve(1)
            .map_err(|source| allocation("ODS extended number styles", source))?;
        self.number_styles.push(style);
        Ok(self)
    }

    fn try_data_style(&mut self, style: Data) -> Result<&mut Self> {
        style.validate()?;
        if self
            .data_styles
            .iter()
            .any(|existing| existing.family() == style.family() && existing.name == style.name)
        {
            return invalid(format!("duplicate ODS data style '{}'", style.name));
        }
        if self.number_styles.len() + self.data_styles.len() >= MAX_STYLE_GRAPH_NODES {
            return invalid("ODS extended style graph exceeds its node limit");
        }
        self.data_styles
            .try_reserve(1)
            .map_err(|source| allocation("ODS extended data styles", source))?;
        self.data_styles.push(style);
        Ok(self)
    }

    /// Finish the graph after a final dependency/name validation pass.
    pub fn build(self) -> Result<Graph> {
        let graph = Graph {
            number_styles: self.number_styles,
            data_styles: self.data_styles,
        };
        graph.validate()?;
        Ok(graph)
    }
}

fn validate_name(value: &str, label: &str) -> Result<()> {
    if value.len() > MAX_STYLE_TEXT_BYTES || value.chars().any(char::is_control) {
        invalid(format!("{label} is invalid"))
    } else if !is_ncname(value) {
        invalid(format!("{label} is not an XML NCName"))
    } else {
        Ok(())
    }
}

fn is_ncname(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(is_ncname_start) && characters.all(is_ncname_char)
}

// XML 1.0 NameStartChar/NameChar ranges, with ':' removed for NCName.
fn is_ncname_start(character: char) -> bool {
    let code = character as u32;
    matches!(
        code,
        0x0041..=0x005A
            | 0x005F
            | 0x0061..=0x007A
            | 0x00C0..=0x00D6
            | 0x00D8..=0x00F6
            | 0x00F8..=0x02FF
            | 0x0370..=0x037D
            | 0x037F..=0x1FFF
            | 0x200C..=0x200D
            | 0x2070..=0x218F
            | 0x2C00..=0x2FEF
            | 0x3001..=0xD7FF
            | 0xF900..=0xFDCF
            | 0xFDF0..=0xFFFD
            | 0x10000..=0xEFFFF
    )
}

fn is_ncname_char(character: char) -> bool {
    is_ncname_start(character)
        || matches!(
            character as u32,
            0x002D | 0x002E | 0x0030..=0x0039 | 0x00B7 | 0x0300..=0x036F | 0x203F..=0x2040
        )
}

/// Parse a schema `integer` into the bounded public scalar.
pub fn parse_integer(value: &str, label: &str) -> Result<i64> {
    let normalized = collapse_xml_whitespace(value, label)?;
    normalized
        .parse::<i64>()
        .map_err(|error| match error.kind() {
            IntErrorKind::PosOverflow | IntErrorKind::NegOverflow => Error::Unsupported(format!(
                "ODF {label} integer is outside the bounded i64 model"
            )),
            _ => invalid_error(format!("invalid ODF integer for {label}")),
        })
}

/// Parse a schema `positiveInteger` into a bounded public scalar.
pub fn parse_positive(value: &str, label: &str) -> Result<u64> {
    let normalized = collapse_xml_whitespace(value, label)?;
    let parsed = normalized
        .parse::<u64>()
        .map_err(|error| match error.kind() {
            IntErrorKind::PosOverflow => Error::Unsupported(format!(
                "ODF {label} positive integer is outside the bounded u64 model"
            )),
            _ => invalid_error(format!("invalid ODF positive integer for {label}")),
        })?;
    if parsed == 0 {
        return invalid(format!("ODF {label} must be positive"));
    }
    Ok(parsed)
}

/// Parse the ODF boolean lexical domain without normalizing an omitted field.
pub fn parse_boolean(value: &str, label: &str) -> Result<bool> {
    match collapse_xml_whitespace(value, label)?.as_ref() {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => invalid(format!("invalid ODF boolean for {label}")),
    }
}

fn validate_text(value: &str, label: &str) -> Result<()> {
    if value.len() > MAX_STYLE_TEXT_BYTES || value.chars().any(|character| !is_xml_char(character))
    {
        invalid(format!(
            "{label} exceeds its text limit or contains an XML-illegal character"
        ))
    } else {
        Ok(())
    }
}

/// XML 1.0 fifth-edition character domain used by ODF text content.
const fn is_xml_char(character: char) -> bool {
    matches!(
        character as u32,
        0x0009 | 0x000A | 0x000D | 0x0020..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
    )
}

fn validate_optional_code(
    value: Option<&str>,
    label: &str,
    validator: fn(&str, &str) -> Result<()>,
) -> Result<()> {
    value.map_or(Ok(()), |value| validator(value, label))
}

fn validate_language_code(value: &str, label: &str) -> Result<()> {
    let normalized = collapse_xml_whitespace(value, label)?;
    if normalized.is_empty()
        || normalized.len() > 8
        || !normalized.bytes().all(|byte| byte.is_ascii_alphabetic())
    {
        return invalid(format!("{label} is not an XML Schema language code"));
    }
    Ok(())
}

fn validate_country_code(value: &str, label: &str) -> Result<()> {
    let normalized = collapse_xml_whitespace(value, label)?;
    if normalized.is_empty()
        || normalized.len() > 8
        || !normalized.bytes().all(|byte| byte.is_ascii_alphanumeric())
    {
        return invalid(format!("{label} is not an ODF country code"));
    }
    Ok(())
}

fn validate_script_code(value: &str, label: &str) -> Result<()> {
    let normalized = collapse_xml_whitespace(value, label)?;
    if normalized.is_empty()
        || normalized.len() > 8
        || !normalized.bytes().all(|byte| byte.is_ascii_alphanumeric())
    {
        return invalid(format!("{label} is not an ODF script code"));
    }
    Ok(())
}

fn validate_language_tag(value: &str, label: &str) -> Result<()> {
    let normalized = collapse_xml_whitespace(value, label)?;
    let mut parts = normalized.split('-');
    let Some(primary) = parts.next() else {
        return invalid(format!("{label} is not an XML Schema language tag"));
    };
    if primary.is_empty()
        || primary.len() > 8
        || !primary.bytes().all(|byte| byte.is_ascii_alphabetic())
        || parts.any(|part| {
            part.is_empty()
                || part.len() > 8
                || !part.bytes().all(|byte| byte.is_ascii_alphanumeric())
        })
    {
        return invalid(format!("{label} is not an XML Schema language tag"));
    }
    Ok(())
}

fn collapse_xml_whitespace<'a>(value: &'a str, label: &str) -> Result<Cow<'a, str>> {
    if value.len() > MAX_STYLE_TEXT_BYTES {
        return invalid(format!("{label} exceeds its text limit"));
    }
    if !value.chars().any(is_xml_whitespace) {
        return Ok(Cow::Borrowed(value));
    }
    let mut normalized = String::new();
    normalized
        .try_reserve_exact(value.len())
        .map_err(|source| allocation("ODS XML Schema whitespace", source))?;
    let mut pending_space = false;
    for character in value.chars() {
        if is_xml_whitespace(character) {
            if !normalized.is_empty() {
                pending_space = true;
            }
        } else {
            if pending_space {
                normalized.push(' ');
                pending_space = false;
            }
            normalized.push(character);
        }
    }
    Ok(Cow::Owned(normalized))
}

fn clone_text(value: &str, label: &str) -> Result<String> {
    validate_text(value, label)?;
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| allocation("ODS data-style text", source))?;
    owned.push_str(value);
    Ok(owned)
}

fn clone_optional_text(value: Option<&str>, label: &str) -> Result<Option<String>> {
    value.map_or(Ok(None), |value| clone_text(value, label).map(Some))
}

fn clone_code_text(
    value: &str,
    label: &str,
    validator: fn(&str, &str) -> Result<()>,
) -> Result<String> {
    validator(value, label)?;
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| allocation("ODS data-style code", source))?;
    owned.push_str(value);
    Ok(owned)
}

fn clone_optional_code_text(
    value: Option<&str>,
    label: &str,
    validator: fn(&str, &str) -> Result<()>,
) -> Result<Option<String>> {
    value
        .map(|value| clone_code_text(value, label, validator).map(Some))
        .unwrap_or(Ok(None))
}

fn clone_double_lexical(value: &str) -> Result<String> {
    let _normalized = normalize_double_lexical(value)?;
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| allocation("ODS double lexical form", source))?;
    owned.push_str(value);
    Ok(owned)
}

fn parse_double_lexical(value: &str) -> Result<f64> {
    let normalized = normalize_double_lexical(value)?;
    match normalized.as_ref() {
        "INF" => Ok(f64::INFINITY),
        "-INF" => Ok(f64::NEG_INFINITY),
        "NaN" => Ok(f64::NAN),
        _ => normalized
            .parse::<f64>()
            .map_err(|_| invalid_error(format!("invalid ODF double '{value}'"))),
    }
}

fn normalize_double_lexical(value: &str) -> Result<Cow<'_, str>> {
    let normalized = collapse_xml_whitespace(value, "ODF double lexical form")?;
    if normalized.is_empty() || normalized.chars().any(is_xml_whitespace) {
        return invalid("ODF double lexical form contains invalid whitespace");
    }
    if matches!(normalized.as_ref(), "INF" | "-INF" | "NaN") {
        return Ok(normalized);
    }
    let bytes = normalized.as_bytes();
    let mut index = usize::from(matches!(bytes.first(), Some(b'+' | b'-')));
    if index == bytes.len() {
        return invalid("ODF double lexical form has no mantissa");
    }
    let integer_digits_start = index;
    while index < bytes.len() && bytes[index].is_ascii_digit() {
        index += 1;
    }
    let integer_digits = index != integer_digits_start;
    let mut fractional_digits = false;
    if bytes.get(index) == Some(&b'.') {
        index += 1;
        let fraction_start = index;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
        }
        fractional_digits = index != fraction_start;
    }
    if !integer_digits && !fractional_digits {
        return invalid("ODF double lexical form has no decimal digits");
    }
    if matches!(bytes.get(index), Some(b'e' | b'E')) {
        index += 1;
        if matches!(bytes.get(index), Some(b'+' | b'-')) {
            index += 1;
        }
        let exponent_start = index;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
        }
        if index == exponent_start {
            return invalid("ODF double lexical form has an incomplete exponent");
        }
    }
    if index != bytes.len() {
        return invalid(format!("invalid ODF double '{value}'"));
    }
    Ok(normalized)
}

const fn is_xml_whitespace(value: char) -> bool {
    matches!(value, ' ' | '\t' | '\r' | '\n')
}

fn validate_optional_text(value: Option<&str>, label: &str) -> Result<()> {
    value.map_or(Ok(()), |value| validate_text(value, label))
}

fn validate_metadata_size(values: [Option<&str>; 9]) -> Result<()> {
    let mut bytes = 0usize;
    for value in values.into_iter().flatten() {
        bytes = bytes
            .checked_add(value.len())
            .ok_or_else(|| invalid_error("ODF data-style metadata size overflow"))?;
    }
    if bytes > MAX_STYLE_METADATA_BYTES {
        return invalid("ODF data-style metadata exceeds its aggregate byte limit");
    }
    Ok(())
}

fn validate_patch_metadata_size(patch: &Patch) -> Result<()> {
    let values = [
        patch_text_bytes(&patch.display_name),
        patch_text_bytes(&patch.language),
        patch_text_bytes(&patch.country),
        patch_text_bytes(&patch.script),
        patch_text_bytes(&patch.rfc_language_tag),
        patch_text_bytes(&patch.title),
        patch_text_bytes(&patch.transliteration_format),
        patch_text_bytes(&patch.transliteration_language),
        patch_text_bytes(&patch.transliteration_country),
    ];
    let bytes = values.into_iter().try_fold(0usize, |total, value| {
        total
            .checked_add(value)
            .ok_or_else(|| invalid_error("ODF data-style patch metadata size overflow"))
    })?;
    if bytes > MAX_STYLE_METADATA_BYTES {
        return invalid("ODF data-style patch metadata exceeds its aggregate byte limit");
    }
    Ok(())
}

fn validate_patched_metadata_size(target: &Attributes, patch: &Patch) -> Result<()> {
    let values = [
        patched_text_bytes(target.display_name.as_deref(), &patch.display_name),
        patched_text_bytes(target.language.as_deref(), &patch.language),
        patched_text_bytes(target.country.as_deref(), &patch.country),
        patched_text_bytes(target.script.as_deref(), &patch.script),
        patched_text_bytes(target.rfc_language_tag.as_deref(), &patch.rfc_language_tag),
        patched_text_bytes(target.title.as_deref(), &patch.title),
        patched_text_bytes(
            target.transliteration.format.as_deref(),
            &patch.transliteration_format,
        ),
        patched_text_bytes(
            target.transliteration.language.as_deref(),
            &patch.transliteration_language,
        ),
        patched_text_bytes(
            target.transliteration.country.as_deref(),
            &patch.transliteration_country,
        ),
    ];
    let bytes = values.into_iter().try_fold(0usize, |total, value| {
        total
            .checked_add(value)
            .ok_or_else(|| invalid_error("ODF patched metadata size overflow"))
    })?;
    if bytes > MAX_STYLE_METADATA_BYTES {
        return invalid("ODF patched data-style metadata exceeds its aggregate byte limit");
    }
    Ok(())
}

fn patch_text_bytes(value: &Op<String>) -> usize {
    match value {
        Op::Set(value) => value.len(),
        Op::Keep | Op::Clear => 0,
    }
}

fn patched_text_bytes(current: Option<&str>, patch: &Op<String>) -> usize {
    match patch {
        Op::Keep => current.map_or(0, str::len),
        Op::Set(value) => value.len(),
        Op::Clear => 0,
    }
}

fn decimal_replacement_is_empty(value: Option<&str>) -> bool {
    value.is_some_and(str::is_empty)
}

fn validate_transliteration_format(value: &str) -> Result<()> {
    let mut chars = value.chars();
    let Some(character) = chars.next() else {
        return invalid("ODF transliteration format must be one Unicode Nd digit");
    };
    if chars.next().is_some() || !unicode_nd_digit_one(character) {
        return invalid(
            "ODF transliteration format must be exactly one Unicode Nd digit with value 1",
        );
    }
    Ok(())
}

fn validate_patch_text<T: TextPatchValue>(value: &Op<T>, label: &str) -> Result<()> {
    if let Op::Set(value) = value {
        let value = value.as_text();
        validate_text(value, label)?;
    }
    Ok(())
}

fn validate_patch_code(
    value: &Op<String>,
    label: &str,
    validator: fn(&str, &str) -> Result<()>,
) -> Result<()> {
    if let Op::Set(value) = value {
        validator(value, label)?;
    }
    Ok(())
}

trait TextPatchValue {
    fn as_text(&self) -> &str;
}

impl TextPatchValue for String {
    fn as_text(&self) -> &str {
        self
    }
}

impl TextPatchValue for bool {
    fn as_text(&self) -> &str {
        ""
    }
}

impl TextPatchValue for TransliterationStyle {
    fn as_text(&self) -> &str {
        self.as_str()
    }
}

fn apply_patch<T: Clone>(target: &mut Option<T>, patch: &Op<T>) {
    match patch {
        Op::Keep => {},
        Op::Set(value) => *target = Some(value.clone()),
        Op::Clear => *target = None,
    }
}

fn validate_embedded_texts(values: &[EmbeddedText]) -> Result<()> {
    if values.len() > MAX_STYLE_BODY_ELEMENTS {
        return invalid("ODF embedded-text particles exceed their element limit");
    }
    for value in values {
        value.validate()?;
    }
    if embedded_text_payload_bytes(values)? > MAX_STYLE_BODY_BYTES {
        return invalid("ODF embedded-text payload exceeds its body byte limit");
    }
    Ok(())
}

fn decimal_particle_count(value: &Decimal) -> Result<usize> {
    1usize
        .checked_add(value.embedded_text.len())
        .ok_or_else(|| invalid_error("ODS decimal body element count overflow"))
}

fn affix_particle_count(value: &Affix) -> usize {
    [
        value.text.as_ref(),
        value.fill_character.as_ref(),
        value.text_after_fill.as_ref(),
    ]
    .into_iter()
    .flatten()
    .count()
}

fn number_particle_count(
    leading: Option<&Affix>,
    format: Option<&Format>,
    trailing: Option<&Affix>,
) -> Result<usize> {
    let mut count = 0usize;
    if let Some(value) = leading {
        count = count
            .checked_add(affix_particle_count(value))
            .ok_or_else(|| invalid_error("ODS number-style element count overflow"))?;
    }
    if let Some(value) = format {
        let format_count = match value {
            Format::Decimal(value) => decimal_particle_count(value)?,
            Format::Scientific(_) | Format::Fraction(_) => 1,
        };
        count = count
            .checked_add(format_count)
            .ok_or_else(|| invalid_error("ODS number-style element count overflow"))?;
    }
    if let Some(value) = trailing {
        count = count
            .checked_add(affix_particle_count(value))
            .ok_or_else(|| invalid_error("ODS number-style element count overflow"))?;
    }
    Ok(count)
}

fn number_has_adjacent_text(
    leading: Option<&Affix>,
    format: Option<&Format>,
    trailing: Option<&Affix>,
) -> bool {
    let mut previous_text = false;
    if let Some(affix) = leading {
        if affix_has_adjacent_text(affix, &mut previous_text) {
            return true;
        }
    }
    if format.is_some() {
        previous_text = false;
    }
    if let Some(affix) = trailing
        && affix_has_adjacent_text(affix, &mut previous_text)
    {
        return true;
    }
    false
}

fn affix_has_adjacent_text(value: &Affix, previous_text: &mut bool) -> bool {
    if value.text.is_some() {
        if *previous_text {
            return true;
        }
        *previous_text = true;
    }
    if value.fill_character.is_some() {
        *previous_text = false;
    }
    if value.text_after_fill.is_some() {
        if *previous_text {
            return true;
        }
        *previous_text = true;
    }
    false
}

fn embedded_text_payload_bytes(values: &[EmbeddedText]) -> Result<usize> {
    values.iter().try_fold(0usize, |total, value| {
        total
            .checked_add(value.text.len())
            .ok_or_else(|| invalid_error("ODF embedded-text payload size overflow"))
    })
}

fn validate_affix_payload(
    text: Option<&str>,
    fill_character: Option<&str>,
    text_after_fill: Option<&str>,
) -> Result<()> {
    let payload = [text, fill_character, text_after_fill]
        .into_iter()
        .flatten()
        .try_fold(0usize, |total, value| {
            total
                .checked_add(value.len())
                .ok_or_else(|| invalid_error("ODS number-style affix size overflow"))
        })?;
    if payload > MAX_STYLE_BODY_BYTES {
        return invalid("ODS number-style affix exceeds its body byte limit");
    }
    Ok(())
}

fn decimal_payload_bytes(value: &Decimal) -> Result<usize> {
    let mut payload = value.decimal_replacement.as_ref().map_or(0, String::len);
    payload = payload
        .checked_add(embedded_text_payload_bytes(&value.embedded_text)?)
        .ok_or_else(|| invalid_error("ODS decimal body size overflow"))?;
    Ok(payload)
}

fn number_payload_bytes(
    leading: Option<&Affix>,
    format: Option<&Format>,
    trailing: Option<&Affix>,
) -> Result<usize> {
    let mut payload = 0usize;
    if let Some(affix) = leading {
        payload = payload
            .checked_add(affix_payload_bytes(affix)?)
            .ok_or_else(|| invalid_error("ODS number-style body size overflow"))?;
    }
    if let Some(format) = format {
        let format_payload = match format {
            Format::Decimal(value) => decimal_payload_bytes(value)?,
            Format::Scientific(_) | Format::Fraction(_) => 0,
        };
        payload = payload
            .checked_add(format_payload)
            .ok_or_else(|| invalid_error("ODS number-style body size overflow"))?;
    }
    if let Some(affix) = trailing {
        payload = payload
            .checked_add(affix_payload_bytes(affix)?)
            .ok_or_else(|| invalid_error("ODS number-style body size overflow"))?;
    }
    Ok(payload)
}

fn affix_payload_bytes(value: &Affix) -> Result<usize> {
    [
        value.text.as_deref(),
        value.fill_character.as_deref(),
        value.text_after_fill.as_deref(),
    ]
    .into_iter()
    .flatten()
    .try_fold(0usize, |total, text| {
        total
            .checked_add(text.len())
            .ok_or_else(|| invalid_error("ODS number-style affix size overflow"))
    })
}

fn append_common_attributes(output: &mut String, attributes: &Attributes) -> Result<()> {
    append_optional_text_attr(
        output,
        "style:display-name",
        attributes.display_name.as_deref(),
    )?;
    append_optional_text_attr(output, "number:language", attributes.language.as_deref())?;
    append_optional_text_attr(output, "number:country", attributes.country.as_deref())?;
    append_optional_text_attr(output, "number:script", attributes.script.as_deref())?;
    append_optional_text_attr(
        output,
        "number:rfc-language-tag",
        attributes.rfc_language_tag.as_deref(),
    )?;
    append_optional_text_attr(output, "number:title", attributes.title.as_deref())?;
    if let Some(value) = attributes.volatile {
        append_attr(
            output,
            "style:volatile",
            if value { "true" } else { "false" },
        )?;
    }
    let transliteration = &attributes.transliteration;
    append_optional_text_attr(
        output,
        "number:transliteration-format",
        transliteration.format.as_deref(),
    )?;
    append_optional_text_attr(
        output,
        "number:transliteration-language",
        transliteration.language.as_deref(),
    )?;
    append_optional_text_attr(
        output,
        "number:transliteration-country",
        transliteration.country.as_deref(),
    )?;
    if let Some(style) = transliteration.style {
        append_attr(output, "number:transliteration-style", style.as_str())?;
    }
    Ok(())
}

fn append_number_format(output: &mut String, format: &Format) -> Result<()> {
    match format {
        Format::Decimal(value) => append_number_body(output, "number", value),
        Format::Scientific(value) => {
            output.push_str("<number:scientific-number");
            append_optional_i64_attr(output, "decimal-places", value.decimal_places)?;
            append_optional_i64_attr(output, "min-decimal-places", value.min_decimal_places)?;
            append_optional_i64_attr(output, "min-integer-digits", value.min_integer_digits)?;
            append_optional_bool_attr(output, "grouping", value.grouping)?;
            append_optional_i64_attr(output, "min-exponent-digits", value.min_exponent_digits)?;
            append_optional_u64_attr(output, "exponent-interval", value.exponent_interval)?;
            append_optional_bool_attr(output, "forced-exponent-sign", value.forced_exponent_sign)?;
            output.push_str("/>");
            Ok(())
        },
        Format::Fraction(value) => {
            output.push_str("<number:fraction");
            append_optional_i64_attr(output, "min-numerator-digits", value.min_numerator_digits)?;
            append_optional_i64_attr(
                output,
                "min-denominator-digits",
                value.min_denominator_digits,
            )?;
            append_optional_i64_attr(output, "denominator-value", value.denominator_value)?;
            append_optional_u64_attr(
                output,
                "max-denominator-value",
                value.effective_max_denominator_value(),
            )?;
            append_optional_i64_attr(output, "min-integer-digits", value.min_integer_digits)?;
            append_optional_bool_attr(output, "grouping", value.grouping)?;
            output.push_str("/>");
            Ok(())
        },
    }
}

fn append_legacy_date_body(output: &mut String) {
    output.push_str(
        "<number:year number:style=\"long\"/><number:text>-</number:text><number:month number:style=\"long\"/><number:text>-</number:text><number:day number:style=\"long\"/>",
    );
}

fn append_legacy_time_body(output: &mut String, decimal_places: Option<i64>) -> Result<()> {
    output.push_str(
        "<number:hours number:style=\"long\"/><number:text>:</number:text><number:minutes number:style=\"long\"/><number:text>:</number:text><number:seconds number:style=\"long\"",
    );
    append_optional_i64_attr(output, "decimal-places", decimal_places)?;
    output.push_str("/>");
    Ok(())
}

fn checked_add_size(left: usize, right: usize) -> Result<usize> {
    left.checked_add(right)
        .ok_or_else(|| invalid_error("ODS data-style markup size overflow"))
}

fn enforce_markup_size(size: usize) -> Result<()> {
    if size > MAX_STYLE_BODY_BYTES {
        invalid("ODS data-style markup exceeds its body byte limit")
    } else {
        Ok(())
    }
}

fn enforce_graph_markup_size(size: usize) -> Result<()> {
    if size > MAX_STYLE_GRAPH_BYTES {
        invalid("ODS extended style graph exceeds its aggregate byte limit")
    } else {
        Ok(())
    }
}

fn escaped_markup_size(value: &str) -> Result<usize> {
    value.bytes().try_fold(0usize, |size, byte| {
        let escaped = match byte {
            b'&' => 5,
            b'<' | b'>' => 4,
            b'"' | b'\'' => 6,
            _ => 1,
        };
        checked_add_size(size, escaped)
    })
}

fn escape_xml_checked(value: &str) -> Result<String> {
    let capacity = escaped_markup_size(value)?;
    let mut escaped = String::new();
    escaped
        .try_reserve_exact(capacity)
        .map_err(|source| allocation("ODS escaped data-style text", source))?;
    let bytes = value.as_bytes();
    let mut cursor = 0usize;
    for (index, byte) in bytes.iter().copied().enumerate() {
        let replacement = match byte {
            b'&' => Some("&amp;"),
            b'<' => Some("&lt;"),
            b'>' => Some("&gt;"),
            b'"' => Some("&quot;"),
            b'\'' => Some("&apos;"),
            _ => None,
        };
        if let Some(replacement) = replacement {
            escaped.push_str(&value[cursor..index]);
            escaped.push_str(replacement);
            cursor = index + 1;
        }
    }
    escaped.push_str(&value[cursor..]);
    Ok(escaped)
}

fn attribute_markup_size(name: &str, value: &str) -> Result<usize> {
    let mut size = name.len();
    size = checked_add_size(size, 4)?; // leading space, equals, and quotes
    checked_add_size(size, escaped_markup_size(value)?)
}

fn root_markup_size(root: &str, name: &str) -> Result<usize> {
    let mut size = "<number:".len();
    size = checked_add_size(size, root.len())?;
    size = checked_add_size(size, " xmlns:number=\"".len())?;
    size = checked_add_size(size, NUMBER_NS.len())?;
    size = checked_add_size(size, "\" xmlns:style=\"".len())?;
    size = checked_add_size(size, STYLE_NS.len())?;
    size = checked_add_size(size, "\" style:name=\"".len())?;
    size = checked_add_size(size, escaped_markup_size(name)?)?;
    checked_add_size(size, 1) // closing style:name quote
}

fn common_attributes_markup_size(attributes: &Attributes) -> Result<usize> {
    let mut size = 0usize;
    for (name, value) in [
        ("style:display-name", attributes.display_name.as_deref()),
        ("number:language", attributes.language.as_deref()),
        ("number:country", attributes.country.as_deref()),
        ("number:script", attributes.script.as_deref()),
        (
            "number:rfc-language-tag",
            attributes.rfc_language_tag.as_deref(),
        ),
        ("number:title", attributes.title.as_deref()),
        (
            "number:transliteration-format",
            attributes.transliteration.format.as_deref(),
        ),
        (
            "number:transliteration-language",
            attributes.transliteration.language.as_deref(),
        ),
        (
            "number:transliteration-country",
            attributes.transliteration.country.as_deref(),
        ),
    ] {
        if let Some(value) = value {
            size = checked_add_size(size, attribute_markup_size(name, value)?)?;
        }
    }
    if let Some(value) = attributes.volatile {
        size = checked_add_size(
            size,
            attribute_markup_size("style:volatile", if value { "true" } else { "false" })?,
        )?;
    }
    if let Some(style) = attributes.transliteration.style {
        size = checked_add_size(
            size,
            attribute_markup_size("number:transliteration-style", style.as_str())?,
        )?;
    }
    Ok(size)
}

fn tagged_text_markup_size(element: &str, value: &str) -> Result<usize> {
    let mut size = "<number:".len();
    size = checked_add_size(size, element.len())?;
    size = checked_add_size(size, 1)?;
    size = checked_add_size(size, escaped_markup_size(value)?)?;
    size = checked_add_size(size, "</number:".len())?;
    size = checked_add_size(size, element.len())?;
    checked_add_size(size, 1)
}

fn affix_markup_size(value: &Affix) -> Result<usize> {
    let mut size = 0usize;
    if let Some(text) = value.text.as_deref() {
        size = checked_add_size(size, tagged_text_markup_size("text", text)?)?;
    }
    if let Some(fill) = value.fill_character.as_deref() {
        size = checked_add_size(size, tagged_text_markup_size("fill-character", fill)?)?;
    }
    if let Some(text) = value.text_after_fill.as_deref() {
        size = checked_add_size(size, tagged_text_markup_size("text", text)?)?;
    }
    Ok(size)
}

fn optional_number_i64_markup_size(name: &str, value: Option<i64>) -> Result<usize> {
    value.map_or(Ok(0), |value| {
        attribute_markup_size(&format!("number:{name}"), &value.to_string())
    })
}

fn optional_number_u64_markup_size(name: &str, value: Option<u64>) -> Result<usize> {
    value.map_or(Ok(0), |value| {
        attribute_markup_size(&format!("number:{name}"), &value.to_string())
    })
}

fn optional_number_bool_markup_size(name: &str, value: Option<bool>) -> Result<usize> {
    value.map_or(Ok(0), |value| {
        attribute_markup_size(
            &format!("number:{name}"),
            if value { "true" } else { "false" },
        )
    })
}

fn embedded_text_markup_size(value: &EmbeddedText) -> Result<usize> {
    let mut size = "<number:embedded-text number:position=\"".len();
    size = checked_add_size(size, value.position.to_string().len())?;
    size = checked_add_size(size, 2)?; // quote and opening bracket
    size = checked_add_size(size, escaped_markup_size(&value.text)?)?;
    checked_add_size(size, "</number:embedded-text>".len())
}

fn number_body_markup_size(element: &str, value: &Decimal) -> Result<usize> {
    let mut size = "<number:".len();
    size = checked_add_size(size, element.len())?;
    size = checked_add_size(
        size,
        optional_number_i64_markup_size("decimal-places", value.decimal_places)?,
    )?;
    size = checked_add_size(
        size,
        optional_number_i64_markup_size("min-decimal-places", value.min_decimal_places)?,
    )?;
    size = checked_add_size(
        size,
        optional_number_i64_markup_size("min-integer-digits", value.min_integer_digits)?,
    )?;
    size = checked_add_size(
        size,
        optional_number_bool_markup_size("grouping", value.grouping)?,
    )?;
    if let Some(replacement) = value.decimal_replacement.as_deref() {
        size = checked_add_size(
            size,
            attribute_markup_size("number:decimal-replacement", replacement)?,
        )?;
    }
    if let Some(display_factor) = &value.display_factor {
        size = checked_add_size(
            size,
            attribute_markup_size("number:display-factor", display_factor.lexical())?,
        )?;
    }
    if value.embedded_text.is_empty() {
        return checked_add_size(size, 2);
    }
    size = checked_add_size(size, 1)?;
    for embedded in &value.embedded_text {
        size = checked_add_size(size, embedded_text_markup_size(embedded)?)?;
    }
    size = checked_add_size(size, "</number:".len())?;
    size = checked_add_size(size, element.len())?;
    checked_add_size(size, 1)
}

fn number_format_markup_size(value: &Format) -> Result<usize> {
    match value {
        Format::Decimal(value) => number_body_markup_size("number", value),
        Format::Scientific(value) => {
            let mut size = "<number:scientific-number".len();
            size = checked_add_size(
                size,
                optional_number_i64_markup_size("decimal-places", value.decimal_places)?,
            )?;
            size = checked_add_size(
                size,
                optional_number_i64_markup_size("min-decimal-places", value.min_decimal_places)?,
            )?;
            size = checked_add_size(
                size,
                optional_number_i64_markup_size("min-integer-digits", value.min_integer_digits)?,
            )?;
            size = checked_add_size(
                size,
                optional_number_bool_markup_size("grouping", value.grouping)?,
            )?;
            size = checked_add_size(
                size,
                optional_number_i64_markup_size("min-exponent-digits", value.min_exponent_digits)?,
            )?;
            size = checked_add_size(
                size,
                optional_number_u64_markup_size("exponent-interval", value.exponent_interval)?,
            )?;
            size = checked_add_size(
                size,
                optional_number_bool_markup_size(
                    "forced-exponent-sign",
                    value.forced_exponent_sign,
                )?,
            )?;
            checked_add_size(size, 2)
        },
        Format::Fraction(value) => {
            let mut size = "<number:fraction".len();
            size = checked_add_size(
                size,
                optional_number_i64_markup_size(
                    "min-numerator-digits",
                    value.min_numerator_digits,
                )?,
            )?;
            size = checked_add_size(
                size,
                optional_number_i64_markup_size(
                    "min-denominator-digits",
                    value.min_denominator_digits,
                )?,
            )?;
            size = checked_add_size(
                size,
                optional_number_i64_markup_size("denominator-value", value.denominator_value)?,
            )?;
            size = checked_add_size(
                size,
                optional_number_u64_markup_size(
                    "max-denominator-value",
                    value.effective_max_denominator_value(),
                )?,
            )?;
            size = checked_add_size(
                size,
                optional_number_i64_markup_size("min-integer-digits", value.min_integer_digits)?,
            )?;
            size = checked_add_size(
                size,
                optional_number_bool_markup_size("grouping", value.grouping)?,
            )?;
            checked_add_size(size, 2)
        },
    }
}

fn legacy_date_markup_size() -> usize {
    "<number:year number:style=\"long\"/><number:text>-</number:text><number:month number:style=\"long\"/><number:text>-</number:text><number:day number:style=\"long\"/>".len()
}

fn legacy_time_markup_size(decimal_places: Option<i64>) -> Result<usize> {
    let mut size = "<number:hours number:style=\"long\"/><number:text>:</number:text><number:minutes number:style=\"long\"/><number:text>:</number:text><number:seconds number:style=\"long\"".len();
    size = checked_add_size(
        size,
        optional_number_i64_markup_size("decimal-places", decimal_places)?,
    )?;
    checked_add_size(size, 2)
}

fn append_number_body(output: &mut String, element: &str, value: &Decimal) -> Result<()> {
    output.push('<');
    output.push_str("number:");
    output.push_str(element);
    append_optional_i64_attr(output, "decimal-places", value.decimal_places)?;
    append_optional_i64_attr(output, "min-decimal-places", value.min_decimal_places)?;
    append_optional_i64_attr(output, "min-integer-digits", value.min_integer_digits)?;
    append_optional_bool_attr(output, "grouping", value.grouping)?;
    append_optional_text_attr(
        output,
        "number:decimal-replacement",
        value.decimal_replacement.as_deref(),
    )?;
    if let Some(display_factor) = &value.display_factor {
        append_attr(output, "number:display-factor", display_factor.lexical())?;
    }
    if value.embedded_text.is_empty() {
        output.push_str("/>");
    } else {
        output.push('>');
        for embedded in &value.embedded_text {
            write!(
                output,
                "<number:embedded-text number:position=\"{}\">{}</number:embedded-text>",
                embedded.position,
                escape_xml_checked(&embedded.text)?
            )
            .map_err(|_| invalid_error("ODS embedded-text markup formatting failed"))?;
        }
        output.push_str("</number:");
        output.push_str(element);
        output.push('>');
    }
    Ok(())
}

fn append_affix(output: &mut String, affix: &Affix) -> Result<()> {
    if let Some(text) = affix.text.as_deref() {
        output.push_str("<number:text>");
        output.push_str(&escape_xml_checked(text)?);
        output.push_str("</number:text>");
    }
    if let Some(fill) = affix.fill_character.as_deref() {
        output.push_str("<number:fill-character>");
        output.push_str(&escape_xml_checked(fill)?);
        output.push_str("</number:fill-character>");
    }
    if let Some(text) = affix.text_after_fill.as_deref() {
        output.push_str("<number:text>");
        output.push_str(&escape_xml_checked(text)?);
        output.push_str("</number:text>");
    }
    Ok(())
}

fn append_optional_text_attr(output: &mut String, name: &str, value: Option<&str>) -> Result<()> {
    if let Some(value) = value {
        append_attr(output, name, value)?;
    }
    Ok(())
}

fn append_optional_i64_attr(output: &mut String, name: &str, value: Option<i64>) -> Result<()> {
    if let Some(value) = value {
        append_attr(output, &format!("number:{name}"), &value.to_string())?;
    }
    Ok(())
}

fn append_optional_u64_attr(output: &mut String, name: &str, value: Option<u64>) -> Result<()> {
    if let Some(value) = value {
        append_attr(output, &format!("number:{name}"), &value.to_string())?;
    }
    Ok(())
}

fn append_optional_bool_attr(output: &mut String, name: &str, value: Option<bool>) -> Result<()> {
    if let Some(value) = value {
        append_attr(
            output,
            &format!("number:{name}"),
            if value { "true" } else { "false" },
        )?;
    }
    Ok(())
}

fn append_attr(output: &mut String, name: &str, value: &str) -> Result<()> {
    write!(output, " {name}=\"{}\"", escape_xml_checked(value)?)
        .map_err(|_| invalid_error("ODS data-style attribute formatting failed"))
}

fn enforce_markup_limit(output: &str) -> Result<()> {
    if output.len() > MAX_STYLE_BODY_BYTES {
        invalid("ODS data-style markup exceeds its body byte limit")
    } else {
        Ok(())
    }
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn unicode_nd_digit_one(value: char) -> bool {
    matches!(
        value as u32,
        0x0031
            | 0x0661
            | 0x06F1
            | 0x07C1
            | 0x0967
            | 0x09E7
            | 0x0A67
            | 0x0AE7
            | 0x0B67
            | 0x0BE7
            | 0x0C67
            | 0x0CE7
            | 0x0D67
            | 0x0DE7
            | 0x0E51
            | 0x0ED1
            | 0x0F21
            | 0x1041
            | 0x1091
            | 0x17E1
            | 0x1811
            | 0x1947
            | 0x19D1
            | 0x1A81
            | 0x1A91
            | 0x1B51
            | 0x1BB1
            | 0x1C41
            | 0x1C51
            | 0xA621
            | 0xA8D1
            | 0xA901
            | 0xA9D1
            | 0xA9F1
            | 0xAA51
            | 0xABF1
            | 0xFF11
            | 0x104A1
            | 0x10D31
            | 0x10D41
            | 0x11067
            | 0x110F1
            | 0x11137
            | 0x111D1
            | 0x112F1
            | 0x11451
            | 0x114D1
            | 0x11651
            | 0x116C1
            | 0x116D1
            | 0x116DB
            | 0x11731
            | 0x118E1
            | 0x11951
            | 0x11BF1
            | 0x11C51
            | 0x11D51
            | 0x11DA1
            | 0x11F51
            | 0x16131
            | 0x16A61
            | 0x16AC1
            | 0x16B51
            | 0x16D71
            | 0x1CCF1
            | 0x1D7CF
            | 0x1D7D9
            | 0x1D7E3
            | 0x1D7ED
            | 0x1D7F7
            | 0x1E141
            | 0x1E2F1
            | 0x1E4F1
            | 0x1E5F2
            | 0x1E951
            | 0x1FBF1
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn transliteration_accepts_unicode_nd_digit_one_only() {
        for value in ["1", "١", "१", "１", "𐒡"] {
            assert!(Transliteration::new(Some(value.to_string()), None, None, None).is_ok());
        }
        for value in ["0", "2", "A", "¹", "12", ""] {
            assert!(Transliteration::new(Some(value.to_string()), None, None, None).is_err());
        }
    }

    #[test]
    fn transliteration_style_ignores_style_without_format() {
        let value = Transliteration {
            format: None,
            language: Some("ar".to_string()),
            country: Some("EG".to_string()),
            style: Some(TransliterationStyle::Long),
        };
        assert_eq!(value.effective_format(), "1");
        assert_eq!(value.effective_style(), TransliterationStyle::Short);
        assert_eq!(value.effective_language(), None);
        assert_eq!(value.effective_country(), None);
    }

    #[test]
    fn style_names_follow_xml_ncname_boundaries() {
        for value in ["N", "_N", "é", "a.b-2", "𐐀"] {
            assert!(Number::new(value).is_ok(), "{value}");
        }
        for value in ["", "1N", "-N", ":N", "a:b", "N name", "N/"] {
            assert!(Number::new(value).is_err(), "{value}");
        }
    }

    #[test]
    fn typed_text_uses_the_xml_character_domain() {
        for value in ["\t", "\n", "\r", "label\tline\n"] {
            assert!(Affix::text(value).is_ok(), "valid XML text {value:?}");
            assert!(
                EmbeddedText::new(1, value).is_ok(),
                "valid XML text {value:?}"
            );
        }
        for value in ["\0", "\u{000B}", "\u{000C}", "\u{000E}"] {
            assert!(Affix::text(value).is_err(), "illegal XML text {value:?}");
            assert!(
                EmbeddedText::new(1, value).is_err(),
                "illegal XML text {value:?}"
            );
        }
    }

    #[test]
    fn metadata_code_domains_match_odf_schema_types() {
        let attributes = Attributes::try_from_borrowed(
            None,
            Some(" en\t"),
            Some("US"),
            Some("Latn"),
            Some("en-US"),
            None,
            None,
            (
                Some("١"),
                Some("ar"),
                Some("EG"),
                Some(TransliterationStyle::Short),
            ),
        )
        .expect("schema metadata");
        assert_eq!(attributes.language.as_deref(), Some(" en\t"));
        for (field, value) in [
            ("language", "en-US"),
            ("country", "UnitedStates"),
            ("script", "Latn-Latn"),
            ("rfc", "en--US"),
        ] {
            let result = match field {
                "language" => Attributes::try_from_borrowed(
                    None,
                    Some(value),
                    None,
                    None,
                    None,
                    None,
                    None,
                    (None, None, None, None),
                ),
                "country" => Attributes::try_from_borrowed(
                    None,
                    None,
                    Some(value),
                    None,
                    None,
                    None,
                    None,
                    (None, None, None, None),
                ),
                "script" => Attributes::try_from_borrowed(
                    None,
                    None,
                    None,
                    Some(value),
                    None,
                    None,
                    None,
                    (None, None, None, None),
                ),
                _ => Attributes::try_from_borrowed(
                    None,
                    None,
                    None,
                    None,
                    Some(value),
                    None,
                    None,
                    (None, None, None, None),
                ),
            };
            assert!(result.is_err(), "invalid {field} value {value}");
        }
        assert!(Transliteration::try_from_borrowed(None, Some("ar-EG"), None, None).is_err());
    }

    #[test]
    fn double_uses_xml_schema_lexical_rules_and_retains_raw_text() {
        for value in [
            "0", "+1", "-1.25", ".5", "1.", "1e+2", "1E-2", " 1 ", "\t-2\n",
        ] {
            assert!(Double::try_from_lexical(value).is_ok(), "{value}");
        }
        let retained = Double::try_from_lexical(" 1e+2 ").expect("schema double");
        assert_eq!(retained.lexical(), " 1e+2 ");
        assert_eq!(retained.value(), 100.0);
        for value in [
            "inf", "infinity", "nan", "+NaN", "+INF", ".", "1e", "1e+", "1 2", "١",
        ] {
            assert!(Double::try_from_lexical(value).is_err(), "{value}");
        }
        assert!(Double::try_from_lexical("1e9999").is_ok());
    }

    #[test]
    fn omitted_and_inherited_decimal_places_remain_distinct() {
        let decimal = Decimal::new();
        assert_eq!(decimal.resolve_decimal_places(None), Resolution::Unresolved);
        assert_eq!(
            decimal.resolve_decimal_places(Some(4)),
            Resolution::Inherited(4)
        );
        assert_eq!(
            decimal.resolve_min_decimal_places(Some(4)),
            Resolution::Inherited(4)
        );
        assert_eq!(
            decimal.resolve_min_decimal_places(None),
            Resolution::Unresolved
        );
        let empty_replacement = Decimal {
            decimal_replacement: Some(String::new()),
            ..Decimal::new()
        };
        assert_eq!(
            empty_replacement.resolve_min_decimal_places(Some(4)),
            Resolution::Explicit(0)
        );
        let replacement = Decimal {
            decimal_replacement: Some("-".to_string()),
            ..Decimal::new()
        };
        assert_eq!(
            replacement.resolve_min_decimal_places(Some(4)),
            Resolution::Inherited(4)
        );
        let explicit = Decimal {
            decimal_places: Some(2),
            ..Decimal::new()
        };
        assert_eq!(
            explicit.resolve_decimal_places(Some(4)),
            Resolution::Explicit(2)
        );
        let invalid_inherited = Decimal {
            min_decimal_places: Some(5),
            ..Decimal::new()
        };
        assert!(
            invalid_inherited
                .resolve_min_decimal_places_checked(Some(4))
                .is_err()
        );
        let scientific = Scientific {
            min_decimal_places: Some(5),
            ..Scientific::new()
        };
        assert!(
            scientific
                .resolve_min_decimal_places_checked(Some(4))
                .is_err()
        );
    }

    #[test]
    fn fraction_and_scientific_boundaries_are_checked() {
        let bad_fraction = Fraction {
            max_denominator_value: Some(0),
            ..Fraction::new()
        };
        assert!(bad_fraction.validate().is_err());
        let bad_scientific = Scientific {
            exponent_interval: Some(0),
            ..Scientific::new()
        };
        assert!(bad_scientific.validate().is_err());
        let bad_minimum = Scientific {
            decimal_places: Some(2),
            min_decimal_places: Some(3),
            ..Scientific::new()
        };
        assert!(bad_minimum.validate().is_err());
        let fixed = Fraction {
            denominator_value: Some(16),
            max_denominator_value: Some(128),
            ..Fraction::new()
        };
        assert_eq!(fixed.effective_max_denominator_value(), None);
        assert!(!fixed.effective_grouping());
    }

    #[test]
    fn wrong_kind_builder_setters_refuse_without_silent_drop() {
        let fraction = NumberBuilder::fraction("Fraction")
            .expect("fraction builder")
            .min_exponent_digits(2);
        assert!(fraction.is_err());
        let scientific = NumberBuilder::scientific("Scientific")
            .expect("scientific builder")
            .max_denominator_value(8);
        assert!(scientific.is_err());
        let date = DataBuilder::date("Date")
            .expect("date builder")
            .min_integer_digits(2);
        assert!(date.is_err());
    }

    #[test]
    fn decimal_embedded_text_is_shared_by_currency_and_percentage() {
        let embedded = EmbeddedText::new(8, "-").expect("valid embedded text");
        let currency = DataBuilder::currency("Currency", "$")
            .expect("currency builder")
            .decimal_replacement(",")
            .expect("currency decimal replacement")
            .display_factor(Double::try_from_lexical("1.25").expect("factor"))
            .expect("currency display factor")
            .embedded_text(embedded.clone())
            .expect("currency embedded text")
            .build()
            .expect("currency style");
        let percentage = DataBuilder::percentage("Percentage")
            .expect("percentage builder")
            .embedded_text(embedded)
            .expect("percentage embedded text")
            .build()
            .expect("percentage style");
        assert!(matches!(currency.body, Body::Currency { .. }));
        assert!(matches!(percentage.body, Body::Percentage { .. }));
        let currency_xml = currency.to_xml().expect("currency XML");
        assert!(currency_xml.contains("number:embedded-text"));
    }

    #[test]
    fn affix_particles_reject_adjacent_text_and_multiple_fill() {
        assert!(Affix::new(Some("a".to_string()), None, Some("b".to_string())).is_err());
        assert!(Affix::new(None, None, Some("after-fill".to_string())).is_err());
        let leading = Affix::fill(".").expect("fill");
        let trailing = Affix::fill(".").expect("fill");
        let style = Number {
            name: "N".to_string(),
            attributes: Attributes::new(),
            leading: Some(leading),
            format: Some(Format::Decimal(Decimal::new())),
            trailing: Some(trailing),
        };
        assert!(style.validate().is_err());
        let trailing_without_body = Number {
            name: "Trailing".to_string(),
            attributes: Attributes::new(),
            leading: None,
            format: None,
            trailing: Some(Affix::text("suffix").expect("suffix")),
        };
        assert!(trailing_without_body.validate().is_err());
        let adjacent_without_body = Number {
            name: "Adjacent".to_string(),
            attributes: Attributes::new(),
            leading: Some(
                Affix::new(None, Some(".".to_string()), Some("a".to_string()))
                    .expect("filled affix"),
            ),
            format: None,
            trailing: Some(Affix::text("b").expect("trailing")),
        };
        assert!(adjacent_without_body.validate().is_err());
    }

    #[test]
    fn date_and_time_emit_required_legacy_particles() {
        let date = DataBuilder::date("Date")
            .expect("date")
            .build()
            .expect("date style");
        let date_xml = date.to_xml().expect("date XML");
        assert!(date_xml.contains("number:year"));
        assert!(date_xml.contains("number:month"));
        assert!(date_xml.contains("number:day"));
        assert!(date_xml.contains("</number:date-style>"));
        let time = DataBuilder::time("Time")
            .expect("time")
            .decimal_places(3)
            .expect("seconds precision")
            .build()
            .expect("time style");
        let time_xml = time.to_xml().expect("time XML");
        assert!(time_xml.contains("number:hours"));
        assert!(time_xml.contains("number:minutes"));
        assert!(time_xml.contains("number:seconds"));
        assert!(!time_xml.contains("<number:time/>"));
    }

    #[test]
    fn scalar_schema_whitespace_and_overflow_are_distinct() {
        assert_eq!(parse_integer("\t +12\n", "integer").expect("integer"), 12);
        assert_eq!(parse_positive("\n+12\r", "positive").expect("positive"), 12);
        assert!(parse_boolean("  true \t", "boolean").expect("boolean"));
        assert!(parse_boolean("1", "boolean").is_err());
        assert!(parse_boolean("0", "boolean").is_err());
        assert!(matches!(
            parse_integer("9223372036854775808", "integer"),
            Err(Error::Unsupported(_))
        ));
        assert!(matches!(
            parse_positive("18446744073709551616", "positive"),
            Err(Error::Unsupported(_))
        ));
        assert!(matches!(
            parse_integer("1.0", "integer"),
            Err(Error::InvalidFormat(_))
        ));
    }

    #[test]
    fn typed_entries_carry_their_source_owner() {
        let number = Number::new("CommonNumber").expect("number");
        let entry = Entry::number_at(Owner::CommonStyles, number);
        assert_eq!(entry.owner(), Owner::CommonStyles);
        assert_eq!(entry.selector().family, Family::Number);
        let automatic = Entry::number_at(
            Owner::StylesAutomatic,
            Number::new("StylesAutomaticNumber").expect("number"),
        );
        assert_eq!(automatic.owner(), Owner::StylesAutomatic);
        assert!(!automatic.owner().is_mutable());
        assert_eq!(
            Selector::styles_automatic("StylesAutomaticNumber", Family::Number).owner,
            Owner::StylesAutomatic
        );
        let automatic = Entry::number_at(
            Owner::ContentAutomatic,
            Number::new("AutomaticNumber").expect("number"),
        );
        assert_eq!(automatic.owner(), Owner::ContentAutomatic);
    }

    #[test]
    fn atomic_setters_restore_previous_value_on_failure() {
        let mut scientific = Scientific {
            decimal_places: Some(2),
            ..Scientific::new()
        };
        assert!(scientific.set_exponent_interval(Some(0)).is_err());
        assert_eq!(scientific.exponent_interval, None);

        let mut attrs = Attributes::new();
        attrs.transliteration.format = Some("1".to_string());
        assert!(
            attrs
                .transliteration
                .set_format(Some("0".to_string()))
                .is_err()
        );
        assert_eq!(attrs.transliteration.format.as_deref(), Some("1"));
    }

    #[test]
    fn builder_scalar_setters_refuse_invalid_values_immediately() {
        assert!(
            NumberBuilder::scientific("Scientific")
                .expect("scientific")
                .exponent_interval(0)
                .is_err()
        );
        assert!(
            NumberBuilder::fraction("Fraction")
                .expect("fraction")
                .max_denominator_value(0)
                .is_err()
        );
        assert!(
            NumberBuilder::decimal("Decimal")
                .expect("decimal")
                .decimal_places(2)
                .expect("decimal places")
                .min_decimal_places(3)
                .is_err()
        );
    }

    #[test]
    fn graph_builder_is_bounded_and_canonical() {
        let style = NumberBuilder::fraction("Fraction")
            .expect("fraction builder")
            .min_numerator_digits(2)
            .expect("numerator digits")
            .max_denominator_value(999)
            .expect("denominator limit")
            .build()
            .expect("fraction style");
        let mut builder = Graph::builder();
        builder.number_style(style).expect("graph reserve");
        let graph = builder.build().expect("graph");
        let xml = graph.to_xml().expect("graph XML");
        assert!(xml.contains("number:fraction"));
        assert!(xml.contains("number:max-denominator-value=\"999\""));
    }

    #[test]
    fn canonical_xml_covers_all_number_fields_and_particles() {
        let decimal = Decimal {
            decimal_places: Some(4),
            min_decimal_places: Some(2),
            min_integer_digits: Some(3),
            grouping: Some(true),
            decimal_replacement: Some(",".to_string()),
            display_factor: Some(Double::try_from_lexical("1.25").expect("double")),
            embedded_text: vec![
                EmbeddedText::new(1, "-").expect("embedded"),
                EmbeddedText::new(8, "+").expect("embedded"),
            ],
        };
        let attrs = Attributes {
            display_name: Some("Display".to_string()),
            language: Some("en".to_string()),
            country: Some("US".to_string()),
            script: Some("Latn".to_string()),
            rfc_language_tag: Some("en-US".to_string()),
            title: Some("Title".to_string()),
            volatile: Some(true),
            transliteration: Transliteration::new(
                Some("١".to_string()),
                Some("ar".to_string()),
                Some("EG".to_string()),
                Some(TransliterationStyle::Long),
            )
            .expect("transliteration"),
        };
        let number = Number {
            name: "AllFields".to_string(),
            attributes: attrs,
            leading: Some(Affix::text("$").expect("leading")),
            format: Some(Format::Decimal(decimal)),
            trailing: Some(Affix::text(" USD").expect("trailing")),
        };
        let xml = number.to_xml().expect("number XML");
        for attribute in [
            "number:decimal-places=\"4\"",
            "number:min-decimal-places=\"2\"",
            "number:min-integer-digits=\"3\"",
            "number:grouping=\"true\"",
            "number:decimal-replacement=\",\"",
            "number:display-factor=\"1.25\"",
            "number:transliteration-format=\"١\"",
            "number:transliteration-style=\"long\"",
        ] {
            assert!(xml.contains(attribute), "missing {attribute} in {xml}");
        }
        assert_eq!(xml.matches("number:embedded-text").count(), 4);

        let scientific = Number {
            name: "Scientific".to_string(),
            attributes: Attributes::new(),
            leading: None,
            format: Some(Format::Scientific(Scientific {
                decimal_places: Some(3),
                min_decimal_places: Some(1),
                min_integer_digits: Some(2),
                grouping: Some(false),
                min_exponent_digits: Some(2),
                exponent_interval: Some(3),
                forced_exponent_sign: Some(false),
            })),
            trailing: None,
        };
        let xml = scientific.to_xml().expect("scientific XML");
        assert!(xml.contains("number:exponent-interval=\"3\""));
        assert!(xml.contains("number:forced-exponent-sign=\"false\""));

        let fraction = Number {
            name: "Fraction".to_string(),
            attributes: Attributes::new(),
            leading: None,
            format: Some(Format::Fraction(Fraction {
                min_numerator_digits: Some(2),
                min_denominator_digits: Some(3),
                denominator_value: Some(16),
                max_denominator_value: Some(128),
                min_integer_digits: Some(1),
                grouping: Some(true),
            })),
            trailing: None,
        };
        let xml = fraction.to_xml().expect("fraction XML");
        assert!(xml.contains("number:denominator-value=\"16\""));
        assert!(!xml.contains("number:max-denominator-value"));
        let capped = Number {
            name: "Capped".to_string(),
            attributes: Attributes::new(),
            leading: None,
            format: Some(Format::Fraction(Fraction {
                max_denominator_value: Some(128),
                ..Fraction::new()
            })),
            trailing: None,
        };
        assert!(
            capped
                .to_xml()
                .expect("capped fraction XML")
                .contains("number:max-denominator-value=\"128\"")
        );
    }

    #[test]
    fn graph_and_embedded_text_limits_fail_before_extra_push() {
        let mut body = Decimal::new();
        for position in 1..=MAX_STYLE_BODY_ELEMENTS {
            body.try_push_embedded_text(
                EmbeddedText::new(i64::try_from(position).expect("bounded"), "x")
                    .expect("embedded"),
            )
            .expect("within limit");
        }
        assert!(
            body.try_push_embedded_text(EmbeddedText::new(1, "x").expect("embedded"))
                .is_err()
        );

        let mut builder = Graph::builder();
        for index in 0..MAX_STYLE_GRAPH_NODES {
            let name = format!("N{index}");
            let number = Number {
                name,
                attributes: Attributes::new(),
                leading: None,
                format: None,
                trailing: None,
            };
            builder.number_style(number).expect("within graph limit");
        }
        let extra = Number {
            name: "extra".to_string(),
            attributes: Attributes::new(),
            leading: None,
            format: None,
            trailing: None,
        };
        assert!(builder.number_style(extra).is_err());
    }

    #[test]
    fn aggregate_body_limits_cover_embedded_text_and_canonical_output() {
        let text = "x".repeat(256);
        let body = Decimal {
            embedded_text: (1..=MAX_STYLE_BODY_ELEMENTS)
                .map(|position| EmbeddedText {
                    position: i64::try_from(position).expect("bounded position"),
                    text: text.clone(),
                })
                .collect(),
            ..Decimal::new()
        };
        assert!(body.validate().is_ok(), "payload is exactly one MiB");
        let number = Number {
            name: "LargeMarkup".to_string(),
            attributes: Attributes::new(),
            leading: None,
            format: Some(Format::Decimal(body)),
            trailing: None,
        };
        assert!(
            number.to_xml().is_err(),
            "markup overhead is bounded before output allocation"
        );

        let overfull = Decimal {
            embedded_text: (1_i64..=16)
                .map(|position| EmbeddedText {
                    position,
                    text: "x".repeat(MAX_STYLE_TEXT_BYTES),
                })
                .collect(),
            ..Decimal::new()
        };
        assert!(overfull.validate().is_ok());
        let currency = Body::Currency {
            symbol: "y".to_string(),
            number: overfull,
        };
        assert!(
            currency.validate().is_err(),
            "currency symbol participates in aggregate body bound"
        );
    }

    #[test]
    fn metadata_patch_is_atomic_and_preserves_omitted_state() {
        let mut attrs = Attributes::new();
        let patch = Patch::default()
            .set_display_name("Shown")
            .expect("display name patch");
        attrs.apply_patch(&patch).expect("metadata patch");
        assert_eq!(attrs.display_name.as_deref(), Some("Shown"));
        assert_eq!(attrs.volatile, None);
        let invalid_patch = Patch {
            transliteration_format: Op::Set("0".to_string()),
            ..Patch::default()
        };
        assert!(attrs.apply_patch(&invalid_patch).is_err());
        assert_eq!(attrs.display_name.as_deref(), Some("Shown"));

        let empty = Patch::default()
            .set_display_name("")
            .expect("empty display name is a distinct set");
        attrs.apply_patch(&empty).expect("empty display name patch");
        assert_eq!(attrs.display_name.as_deref(), Some(""));
        attrs
            .apply_patch(&Patch::default().clear_display_name())
            .expect("clear display name patch");
        assert_eq!(attrs.display_name, None);
    }

    #[test]
    fn metadata_patch_precharges_aggregate_set_values_and_target() {
        let half = "x".repeat(MAX_STYLE_METADATA_BYTES / 2 + 1);
        let oversized_patch = Patch {
            display_name: Op::Set(half.clone()),
            title: Op::Set(half),
            ..Patch::default()
        };
        assert!(oversized_patch.validate().is_err());

        let mut attrs = Attributes::new();
        attrs.display_name = Some("x".repeat(MAX_STYLE_METADATA_BYTES));
        let target_before = attrs.clone();
        let patch = Patch::default().set_title("x").expect("small title patch");
        assert!(attrs.apply_patch(&patch).is_err());
        assert_eq!(attrs, target_before);
    }
}
