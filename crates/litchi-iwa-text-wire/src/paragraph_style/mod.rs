//! Storage-scoped paragraph-style property engine (record 0374).
//!
//! A uniformly styled `TSWP.StorageArchive` names one
//! `TSWP.ParagraphStyleArchive` in its paragraph-style table (field 5).
//! Its character and paragraph properties are the effective values of that
//! style's parent chain. Setting or resetting one property never edits a
//! named style: the engine either rewrites the storage's own managed
//! variation in place, derives a new child variation of the current style,
//! or collapses the variation back to its parent when no override is left.
//! This is the behaviour of the `litchi-iwa` host's `text::Draft`
//! `set_text_*`/`reset_text_*`/`set_paragraph_*`/`reset_paragraph_*`
//! families, collapsed into one property verb ([`Change`]).
//!
//! The engine owns only wire payloads. The caller (a format package):
//!
//! - resolves paragraph-style payloads through [`Styles`];
//! - says whether the storage's style is exclusive to it ([`Target`]): the
//!   only storage reference and nobody's parent;
//! - validates a [`ParagraphFollowingStyle::Named`] target against its
//!   stylesheet's named styles;
//! - publishes the [`Plan`]: it allocates the identifier of a derived
//!   variation, stores the variation payload (message type
//!   [`MESSAGE_TYPE`], version [`MESSAGE_VERSION`]) with its object
//!   references and [`ManagedFields`], repoints the storage with
//!   [`repoint`], registers or unregisters the variation in its stylesheet
//!   ([`stylesheet_with_variation`], [`stylesheet_without_variation`]) and
//!   keeps its package metadata.
//!
//! Every function takes the caller's [`WireLimits`]; scan work is charged
//! across the whole parent chain.
//!
//! [`ParagraphFollowingStyle::Named`]: litchi_iwa_text::paragraph::style::ParagraphFollowingStyle::Named

mod convert;
mod inherit;
pub mod named;
mod overrides;
mod storage;

use std::collections::{BTreeMap, HashMap};
use std::hash::BuildHasher;

use litchi_iwa_common::color::Rgba;
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_text::appearance::{Background, Outline, ParagraphBackground, Shadow};
use litchi_iwa_text::character::{
    TextBaselineShift, TextCapitalization, TextCharacterSpacing, TextDecorations, TextLigatures,
    TextScript, TextStyle,
};
use litchi_iwa_text::font::Font;
use litchi_iwa_text::paragraph::direction::WritingDirection;
use litchi_iwa_text::paragraph::flow::Flow;
use litchi_iwa_text::paragraph::format::{Alignment, Borders, Indents, LineSpacing, Spacing};
use litchi_iwa_text::paragraph::style::ParagraphFollowingStyle;
use litchi_iwa_text::paragraph::tabs::{DecimalCharacter, DefaultInterval, Stops};
use thiserror::Error as ThisError;

use crate::native::{Budget, NativeError};
use overrides::Overrides;

/// `TSWP.ParagraphStyleArchive`.
pub const MESSAGE_TYPE: u32 = 2_022;
/// The message version a new variation is stored with.
pub const MESSAGE_VERSION: [u32; 3] = [1, 0, 5];

/// One storage-scoped character or paragraph property with its value.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Property {
    /// Point size, bold and italic.
    TextStyle(TextStyle),
    /// Font family.
    TextFont(Font),
    /// Underline and strikethrough.
    TextDecorations(TextDecorations),
    /// Text color.
    TextColor(Rgba),
    /// Capitalization.
    TextCapitalization(TextCapitalization),
    /// Superscript or subscript.
    TextScript(TextScript),
    /// Baseline shift.
    TextBaselineShift(TextBaselineShift),
    /// Character spacing (tracking).
    TextCharacterSpacing(TextCharacterSpacing),
    /// Ligatures.
    TextLigatures(TextLigatures),
    /// Glyph outline.
    TextOutline(Outline),
    /// Glyph shadow.
    TextShadow(Shadow),
    /// Glyph background.
    TextBackground(Background),
    /// Paragraph fill.
    ParagraphBackground(ParagraphBackground),
    /// Paragraph borders.
    Borders(Borders),
    /// Hyphenation and pagination.
    Flow(Flow),
    /// Writing direction.
    WritingDirection(WritingDirection),
    /// Style applied to the following paragraph.
    FollowingStyle(ParagraphFollowingStyle),
    /// Horizontal alignment.
    Alignment(Alignment),
    /// Line spacing.
    LineSpacing(LineSpacing),
    /// Space before and after.
    Spacing(Spacing),
    /// First-line, left and right indents.
    Indents(Indents),
    /// Decimal tab character.
    DecimalTabCharacter(DecimalCharacter),
    /// Default tab interval.
    DefaultTabInterval(DefaultInterval),
    /// Explicit tab stops.
    TabStops(Stops),
}

/// The property a [`Property`] value sets, for reads and resets.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PropertyKind {
    /// [`Property::TextStyle`].
    TextStyle,
    /// [`Property::TextFont`].
    TextFont,
    /// [`Property::TextDecorations`].
    TextDecorations,
    /// [`Property::TextColor`].
    TextColor,
    /// [`Property::TextCapitalization`].
    TextCapitalization,
    /// [`Property::TextScript`].
    TextScript,
    /// [`Property::TextBaselineShift`].
    TextBaselineShift,
    /// [`Property::TextCharacterSpacing`].
    TextCharacterSpacing,
    /// [`Property::TextLigatures`].
    TextLigatures,
    /// [`Property::TextOutline`].
    TextOutline,
    /// [`Property::TextShadow`].
    TextShadow,
    /// [`Property::TextBackground`].
    TextBackground,
    /// [`Property::ParagraphBackground`].
    ParagraphBackground,
    /// [`Property::Borders`].
    Borders,
    /// [`Property::Flow`].
    Flow,
    /// [`Property::WritingDirection`].
    WritingDirection,
    /// [`Property::FollowingStyle`].
    FollowingStyle,
    /// [`Property::Alignment`].
    Alignment,
    /// [`Property::LineSpacing`].
    LineSpacing,
    /// [`Property::Spacing`].
    Spacing,
    /// [`Property::Indents`].
    Indents,
    /// [`Property::DecimalTabCharacter`].
    DecimalTabCharacter,
    /// [`Property::DefaultTabInterval`].
    DefaultTabInterval,
    /// [`Property::TabStops`].
    TabStops,
}

impl PropertyKind {
    /// Every property kind, in declaration order.
    pub const ALL: [Self; 24] = [
        Self::TextStyle,
        Self::TextFont,
        Self::TextDecorations,
        Self::TextColor,
        Self::TextCapitalization,
        Self::TextScript,
        Self::TextBaselineShift,
        Self::TextCharacterSpacing,
        Self::TextLigatures,
        Self::TextOutline,
        Self::TextShadow,
        Self::TextBackground,
        Self::ParagraphBackground,
        Self::Borders,
        Self::Flow,
        Self::WritingDirection,
        Self::FollowingStyle,
        Self::Alignment,
        Self::LineSpacing,
        Self::Spacing,
        Self::Indents,
        Self::DecimalTabCharacter,
        Self::DefaultTabInterval,
        Self::TabStops,
    ];

    const fn is_character(self) -> bool {
        matches!(
            self,
            Self::TextStyle
                | Self::TextFont
                | Self::TextDecorations
                | Self::TextColor
                | Self::TextCapitalization
                | Self::TextScript
                | Self::TextBaselineShift
                | Self::TextCharacterSpacing
                | Self::TextLigatures
                | Self::TextOutline
                | Self::TextShadow
                | Self::TextBackground
        )
    }
}

impl Property {
    /// The property this value sets.
    #[must_use]
    pub const fn kind(&self) -> PropertyKind {
        match self {
            Self::TextStyle(_) => PropertyKind::TextStyle,
            Self::TextFont(_) => PropertyKind::TextFont,
            Self::TextDecorations(_) => PropertyKind::TextDecorations,
            Self::TextColor(_) => PropertyKind::TextColor,
            Self::TextCapitalization(_) => PropertyKind::TextCapitalization,
            Self::TextScript(_) => PropertyKind::TextScript,
            Self::TextBaselineShift(_) => PropertyKind::TextBaselineShift,
            Self::TextCharacterSpacing(_) => PropertyKind::TextCharacterSpacing,
            Self::TextLigatures(_) => PropertyKind::TextLigatures,
            Self::TextOutline(_) => PropertyKind::TextOutline,
            Self::TextShadow(_) => PropertyKind::TextShadow,
            Self::TextBackground(_) => PropertyKind::TextBackground,
            Self::ParagraphBackground(_) => PropertyKind::ParagraphBackground,
            Self::Borders(_) => PropertyKind::Borders,
            Self::Flow(_) => PropertyKind::Flow,
            Self::WritingDirection(_) => PropertyKind::WritingDirection,
            Self::FollowingStyle(_) => PropertyKind::FollowingStyle,
            Self::Alignment(_) => PropertyKind::Alignment,
            Self::LineSpacing(_) => PropertyKind::LineSpacing,
            Self::Spacing(_) => PropertyKind::Spacing,
            Self::Indents(_) => PropertyKind::Indents,
            Self::DecimalTabCharacter(_) => PropertyKind::DecimalTabCharacter,
            Self::DefaultTabInterval(_) => PropertyKind::DefaultTabInterval,
            Self::TabStops(_) => PropertyKind::TabStops,
        }
    }
}

/// A property change on a uniformly styled storage.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Change {
    /// Set the effective value of a property.
    Set(Property),
    /// Remove the storage variation's own override of a property, so the
    /// value is inherited again.
    Reset(PropertyKind),
}

/// The storage's paragraph style and whether only that storage uses it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Target {
    style: u64,
    exclusive: bool,
}

impl Target {
    /// `style` is the storage's uniform paragraph style (see
    /// [`uniform_style`]); `exclusive` holds when the storage is its only
    /// referrer and no style names it as parent.
    #[must_use]
    pub const fn new(style: u64, exclusive: bool) -> Self {
        Self { style, exclusive }
    }

    /// The storage's paragraph style.
    #[must_use]
    pub const fn style(self) -> u64 {
        self.style
    }

    /// Whether only the storage uses the style.
    #[must_use]
    pub const fn exclusive(self) -> bool {
        self.exclusive
    }
}

/// The native field-path metadata a variation's message must carry
/// (`TSP.FieldInfo` with an "ignore and preserve" unknown-field rule): the
/// IWA archive records which nested messages the applications may extend.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "each flag is one independent field path"
)]
pub struct ManagedFields {
    text_fill: bool,
    text_stroke: bool,
    text_shadow: bool,
    text_background: bool,
}

impl ManagedFields {
    /// Path `[11, 46]` (character properties, drawing fill).
    #[must_use]
    pub const fn text_fill(self) -> bool {
        self.text_fill
    }

    /// Path `[11, 44]` (character properties, drawing stroke).
    #[must_use]
    pub const fn text_stroke(self) -> bool {
        self.text_stroke
    }

    /// Path `[11, 21]` (character properties, shadow).
    #[must_use]
    pub const fn text_shadow(self) -> bool {
        self.text_shadow
    }

    /// Path `[11, 26]` (character properties, background color).
    #[must_use]
    pub const fn text_background(self) -> bool {
        self.text_background
    }
}

/// A canonical paragraph-style variation payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Variation {
    payload: Vec<u8>,
    parent: u64,
    stylesheet: u64,
    references: Vec<u64>,
    managed: ManagedFields,
}

impl Variation {
    /// The `TSWP.ParagraphStyleArchive` payload.
    #[must_use]
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }

    /// Take the payload.
    #[must_use]
    pub fn into_payload(self) -> Vec<u8> {
        self.payload
    }

    /// The parent style.
    #[must_use]
    pub const fn parent(&self) -> u64 {
        self.parent
    }

    /// The stylesheet that registers the variation.
    #[must_use]
    pub const fn stylesheet(&self) -> u64 {
        self.stylesheet
    }

    /// The message's object references, in archive-metadata order (the
    /// parent, then a named following style).
    #[must_use]
    pub fn references(&self) -> &[u64] {
        &self.references
    }

    /// The field paths the message's metadata must declare.
    #[must_use]
    pub const fn managed_fields(&self) -> ManagedFields {
        self.managed
    }
}

/// What a [`Change`] does to the storage's style.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Plan {
    /// The effective value is already the requested one, or there is no
    /// managed override to reset.
    Unchanged,
    /// Replace the payload of the storage's own variation in place (its
    /// identifier, parent and stylesheet stay).
    Rewrite(Variation),
    /// Store a new child variation of the storage's style, repoint the
    /// storage to it and register it in its stylesheet.
    Derive(Variation),
    /// Remove the storage's variation, which no longer overrides anything:
    /// repoint the storage to `parent` and unregister the variation from
    /// `stylesheet`.
    Collapse {
        /// The variation's parent.
        parent: u64,
        /// The stylesheet that registers the variation.
        stylesheet: u64,
    },
}

impl Plan {
    /// Whether the plan changes anything.
    #[must_use]
    pub const fn changed(&self) -> bool {
        !matches!(self, Self::Unchanged)
    }
}

/// Why an engine operation failed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A payload is malformed, a referenced style is missing, or a native
    /// value cannot be represented.
    #[error("invalid paragraph-style source: {reason}")]
    InvalidSource {
        /// What is wrong.
        reason: &'static str,
    },
    /// A configured wire resource was exhausted.
    #[error("paragraph-style {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The exhausted resource.
        kind: LimitKind,
        /// The attempted aggregate consumption.
        observed: usize,
        /// The caller's configured bound.
        maximum: usize,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} paragraph-style items or bytes")]
    Allocation {
        /// The attempted item or byte capacity.
        amount: usize,
    },
    /// The planned payload did not read back as the requested value (an
    /// engine invariant).
    #[error("paragraph-style plan failed its read-back check")]
    Internal,
    /// The request is well formed but the named-style catalog refuses it.
    #[error("paragraph-style request refused: {0:?}")]
    Refused(Refusal),
}

/// Why a named-style request is refused ([`named`]).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The style is not a named style of the storage's stylesheet.
    NotNamed,
    /// Another named style already has the requested name.
    NameTaken,
    /// A theme must keep at least one named paragraph style.
    LastStyle,
    /// The storage has no overrides on its applied named style.
    NoOverrides,
    /// The storage does not currently use the named style.
    NotApplied,
    /// A deleted style needs a different replacement.
    SameReplacement,
}

impl From<NativeError> for Error {
    fn from(error: NativeError) -> Self {
        match error {
            NativeError::Invalid(reason) => Self::InvalidSource { reason },
            NativeError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind,
                observed,
                maximum,
            },
            NativeError::Allocation { amount } => Self::Allocation { amount },
        }
    }
}

/// The result of an engine operation.
pub type Result<T> = std::result::Result<T, Error>;

/// Paragraph-style payloads by object identifier.
pub trait Styles {
    /// The `TSWP.ParagraphStyleArchive` payload of `identifier`, or `None`
    /// when the package has no such paragraph style.
    fn paragraph_style(&self, identifier: u64) -> Option<&[u8]>;
}

impl<S: BuildHasher> Styles for HashMap<u64, Vec<u8>, S> {
    fn paragraph_style(&self, identifier: u64) -> Option<&[u8]> {
        self.get(&identifier).map(Vec::as_slice)
    }
}

impl Styles for BTreeMap<u64, Vec<u8>> {
    fn paragraph_style(&self, identifier: u64) -> Option<&[u8]> {
        self.get(&identifier).map(Vec::as_slice)
    }
}

/// `styles` with one identifier's payload replaced or added.
struct Overlay<'a, S: ?Sized> {
    styles: &'a S,
    identifier: u64,
    payload: &'a [u8],
}

impl<S: Styles + ?Sized> Styles for Overlay<'_, S> {
    fn paragraph_style(&self, identifier: u64) -> Option<&[u8]> {
        if identifier == self.identifier {
            Some(self.payload)
        } else {
            self.styles.paragraph_style(identifier)
        }
    }
}

/// The storage's uniform paragraph style: the only entry of its
/// paragraph-style table, at UTF-16 index zero.
///
/// # Errors
///
/// [`Error::InvalidSource`] when the storage has no table, several
/// boundaries or no style; resource errors as configured.
pub fn uniform_style(storage: &[u8], limits: WireLimits) -> Result<u64> {
    Ok(storage::uniform_style(storage, &mut Budget::new(limits))?)
}

/// Rewrite the storage's uniform paragraph-style reference from `from` to
/// `to`, keeping every other byte.
///
/// # Errors
///
/// [`Error::InvalidSource`] when the storage is not uniformly styled by
/// `from`; resource errors as configured.
pub fn repoint(storage: &[u8], from: u64, to: u64, limits: WireLimits) -> Result<Vec<u8>> {
    Ok(storage::repoint(
        storage,
        from,
        to,
        &mut Budget::new(limits),
    )?)
}

/// Register `style` in a `TSS.StylesheetArchive`: appended to its styles
/// and to `parent`'s children entry (created when the parent has none).
///
/// # Errors
///
/// [`Error::InvalidSource`] when the stylesheet is malformed, already
/// registers `style` or repeats `parent`; resource errors as configured.
pub fn stylesheet_with_variation(
    stylesheet: &[u8],
    parent: u64,
    style: u64,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    Ok(storage::stylesheet_with_variation(
        stylesheet,
        parent,
        style,
        &mut Budget::new(limits),
    )?)
}

/// Unregister `style` (a child of `parent`) from a `TSS.StylesheetArchive`;
/// a children entry left empty is dropped.
///
/// # Errors
///
/// [`Error::InvalidSource`] when the stylesheet does not register `style`
/// exactly once under `parent`; resource errors as configured.
pub fn stylesheet_without_variation(
    stylesheet: &[u8],
    parent: u64,
    style: u64,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    Ok(storage::stylesheet_without_variation(
        stylesheet,
        parent,
        style,
        &mut Budget::new(limits),
    )?)
}

/// The effective value of `kind` for a storage styled by `style`.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a missing or malformed style on the chain,
/// an inheritance cycle or a chain deeper than 64; resource errors as
/// configured.
pub fn effective(
    styles: &(impl Styles + ?Sized),
    style: u64,
    kind: PropertyKind,
    limits: WireLimits,
) -> Result<Property> {
    Ok(inherit::effective(
        styles,
        style,
        kind,
        &mut Budget::new(limits),
    )?)
}

/// Plan a [`Change`] of the storage styled by `target`.
///
/// # Errors
///
/// As [`effective`], plus [`Error::InvalidSource`] when the style has no
/// stylesheet or a managed variation has no parent.
pub fn plan(
    styles: &(impl Styles + ?Sized),
    target: Target,
    change: &Change,
    limits: WireLimits,
) -> Result<Plan> {
    let mut budget = Budget::new(limits);
    match change {
        Change::Set(property) => plan_set(styles, target, property, &mut budget),
        Change::Reset(kind) => plan_reset(styles, target, *kind, &mut budget),
    }
}

fn parent_of(style: &crate::native::ParagraphStyle) -> Result<u64> {
    style
        .base
        .parent
        .map(|parent| parent.identifier)
        .ok_or(Error::InvalidSource {
            reason: "iWork paragraph-style variation has no parent",
        })
}

fn stylesheet_of(style: &crate::native::ParagraphStyle) -> Result<u64> {
    style
        .base
        .stylesheet
        .map(|stylesheet| stylesheet.identifier)
        .ok_or(Error::InvalidSource {
            reason: "iWork paragraph style has no stylesheet",
        })
}

fn variation(
    parent: u64,
    stylesheet: u64,
    overrides: &Overrides,
    limits: WireLimits,
) -> Result<Variation> {
    let (payload, references, managed) =
        overrides::variation(parent, stylesheet, overrides, limits)?;
    Ok(Variation {
        payload,
        parent,
        stylesheet,
        references,
        managed,
    })
}

/// The font subset of a variation's overrides: font identity, bold,
/// italic and point size (the chart-font properties, record 0375a).
#[derive(Debug, Clone, PartialEq, Default)]
#[non_exhaustive]
pub struct FontOverrides {
    /// Font identity.
    pub font: Option<Font>,
    /// Bold.
    pub bold: Option<bool>,
    /// Italic.
    pub italic: Option<bool>,
    /// Point size.
    pub point_size: Option<litchi_iwa_text::character::TextPointSize>,
}

impl FontOverrides {
    /// Whether no font property is overridden.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.font.is_none()
            && self.bold.is_none()
            && self.italic.is_none()
            && self.point_size.is_none()
    }
}

/// The font overrides of `style` when it is a managed variation that
/// overrides only font properties (`None` otherwise).
///
/// # Errors
///
/// [`Error::InvalidSource`] for a missing or malformed style; resource
/// errors as configured.
pub fn font_overrides(
    styles: &(impl Styles + ?Sized),
    style: u64,
    limits: WireLimits,
) -> Result<Option<FontOverrides>> {
    let mut budget = Budget::new(limits);
    let (parsed, raw) = inherit::locate(styles, style, &mut budget)?;
    let Some(overrides) = overrides::direct_overrides(&parsed, raw, &mut budget)? else {
        return Ok(None);
    };
    let font = FontOverrides {
        font: overrides.font.clone(),
        bold: overrides.bold,
        italic: overrides.italic,
        point_size: overrides.point_size,
    };
    let font_count = u32::from(font.font.is_some())
        + u32::from(font.bold.is_some())
        + u32::from(font.italic.is_some())
        + u32::from(font.point_size.is_some());
    Ok((font_count > 0 && overrides.count() == font_count).then_some(font))
}

/// A canonical variation of `parent` in `stylesheet` holding exactly the
/// font overrides `font`.
///
/// # Errors
///
/// [`Error::InvalidSource`] for empty overrides; resource errors as
/// configured.
pub fn font_variation(
    parent: u64,
    stylesheet: u64,
    font: &FontOverrides,
    limits: WireLimits,
) -> Result<Variation> {
    let overrides = Overrides {
        font: font.font.clone(),
        bold: font.bold,
        italic: font.italic,
        point_size: font.point_size,
        ..Overrides::default()
    };
    variation(parent, stylesheet, &overrides, limits)
}

/// Every override a managed variation holds (opaque), with its font subset
/// editable: the chart legend typography keeps the variation's other
/// overrides when it rewrites the font ones (record 0414).
#[derive(Debug, Clone, PartialEq, Default)]
pub struct VariationOverrides(Overrides);

impl VariationOverrides {
    /// The font subset.
    #[must_use]
    pub fn font(&self) -> FontOverrides {
        FontOverrides {
            font: self.0.font.clone(),
            bold: self.0.bold,
            italic: self.0.italic,
            point_size: self.0.point_size,
        }
    }

    /// Replace the font subset, keeping every other override.
    pub fn set_font(&mut self, font: FontOverrides) {
        self.0.font = font.font;
        self.0.bold = font.bold;
        self.0.italic = font.italic;
        self.0.point_size = font.point_size;
    }

    /// How many properties the variation overrides (its stored override
    /// count).
    #[must_use]
    pub fn count(&self) -> u32 {
        self.0.count()
    }

    /// Whether nothing is overridden.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }
}

/// The overrides of `style` when it is a managed variation (`None`
/// otherwise).
///
/// # Errors
///
/// [`Error::InvalidSource`] for a missing or malformed style; resource
/// errors as configured.
pub fn variation_overrides(
    styles: &(impl Styles + ?Sized),
    style: u64,
    limits: WireLimits,
) -> Result<Option<VariationOverrides>> {
    let mut budget = Budget::new(limits);
    let (parsed, raw) = inherit::locate(styles, style, &mut budget)?;
    Ok(overrides::direct_overrides(&parsed, raw, &mut budget)?.map(VariationOverrides))
}

/// A canonical variation of `parent` in `stylesheet` holding exactly
/// `overrides`.
///
/// # Errors
///
/// [`Error::InvalidSource`] for empty overrides; resource errors as
/// configured.
pub fn variation_with(
    parent: u64,
    stylesheet: u64,
    overrides: &VariationOverrides,
    limits: WireLimits,
) -> Result<Variation> {
    variation(parent, stylesheet, &overrides.0, limits)
}

/// An identifier no package object uses, for reading a derived variation
/// back before the caller allocates its real identifier.
const PROVISIONAL_IDENTIFIER: u64 = u64::MAX;

fn plan_set(
    styles: &(impl Styles + ?Sized),
    target: Target,
    property: &Property,
    budget: &mut Budget,
) -> Result<Plan> {
    let kind = property.kind();
    if &inherit::effective(styles, target.style, kind, budget)? == property {
        return Ok(Plan::Unchanged);
    }
    let (style, raw) = inherit::locate(styles, target.style, budget)?;
    let stylesheet = stylesheet_of(&style)?;
    // iWork normalizes tab defaults as a child variation of the current
    // style: the applications ignore those fields inside a variation that
    // also holds other overrides.
    let requires_child_variation = matches!(
        kind,
        PropertyKind::DecimalTabCharacter | PropertyKind::DefaultTabInterval
    );
    // A shared style is never rewritten, so its own overrides are decoded
    // only for an exclusive variation (record 0412: native Keynote table
    // base styles carry a zero default tab interval the override decoder
    // refuses, and a derived child does not inherit that field as an
    // override).
    if target.exclusive
        && let Some(mut overrides) = overrides::direct_overrides(&style, raw, budget)?
        && (!requires_child_variation || overrides.is_tab_defaults_only())
    {
        let parent = parent_of(&style)?;
        let inherited = inherited_character(styles, parent, kind, budget)?;
        apply(&mut overrides, property, inherited.as_ref())?;
        if overrides.is_empty() {
            verify(&inherit::effective(styles, parent, kind, budget)?, property)?;
            return Ok(Plan::Collapse { parent, stylesheet });
        }
        let rewritten = variation(parent, stylesheet, &overrides, budget.limits())?;
        let overlay = Overlay {
            styles,
            identifier: target.style,
            payload: &rewritten.payload,
        };
        verify(
            &inherit::effective(&overlay, target.style, kind, budget)?,
            property,
        )?;
        return Ok(Plan::Rewrite(rewritten));
    }
    let mut overrides = Overrides::default();
    let inherited = inherited_character(styles, target.style, kind, budget)?;
    apply(&mut overrides, property, inherited.as_ref())?;
    let derived = variation(target.style, stylesheet, &overrides, budget.limits())?;
    let overlay = Overlay {
        styles,
        identifier: PROVISIONAL_IDENTIFIER,
        payload: &derived.payload,
    };
    verify(
        &inherit::effective(&overlay, PROVISIONAL_IDENTIFIER, kind, budget)?,
        property,
    )?;
    Ok(Plan::Derive(derived))
}

fn plan_reset(
    styles: &(impl Styles + ?Sized),
    target: Target,
    kind: PropertyKind,
    budget: &mut Budget,
) -> Result<Plan> {
    let (style, raw) = inherit::locate(styles, target.style, budget)?;
    let Some(mut overrides) = overrides::direct_overrides(&style, raw, budget)? else {
        return Ok(Plan::Unchanged);
    };
    if !has(&overrides, kind) || !target.exclusive {
        return Ok(Plan::Unchanged);
    }
    clear(&mut overrides, kind);
    let parent = parent_of(&style)?;
    let stylesheet = stylesheet_of(&style)?;
    let expected = inherit::effective(styles, parent, kind, budget)?;
    if overrides.is_empty() {
        return Ok(Plan::Collapse { parent, stylesheet });
    }
    let rewritten = variation(parent, stylesheet, &overrides, budget.limits())?;
    let overlay = Overlay {
        styles,
        identifier: target.style,
        payload: &rewritten.payload,
    };
    verify(
        &inherit::effective(&overlay, target.style, kind, budget)?,
        &expected,
    )?;
    Ok(Plan::Rewrite(rewritten))
}

fn verify(actual: &Property, expected: &Property) -> Result<()> {
    if actual == expected {
        Ok(())
    } else {
        Err(Error::Internal)
    }
}

/// The inherited value a character override is compared against: a
/// character property equal to the inherited value is not overridden.
fn inherited_character(
    styles: &(impl Styles + ?Sized),
    parent: u64,
    kind: PropertyKind,
    budget: &mut Budget,
) -> Result<Option<Property>> {
    if !kind.is_character() {
        return Ok(None);
    }
    Ok(Some(inherit::effective(styles, parent, kind, budget)?))
}

fn differs<T: PartialEq + Clone>(value: &T, inherited: &T) -> Option<T> {
    (value != inherited).then(|| value.clone())
}

#[allow(clippy::too_many_lines, reason = "one arm per property")]
fn apply(
    overrides: &mut Overrides,
    property: &Property,
    inherited: Option<&Property>,
) -> Result<()> {
    let mismatch = Error::Internal;
    match (property, inherited) {
        (Property::TextStyle(style), Some(Property::TextStyle(inherited))) => {
            overrides.point_size = differs(&style.point_size, &inherited.point_size);
            overrides.bold = differs(&style.bold, &inherited.bold);
            overrides.italic = differs(&style.italic, &inherited.italic);
        },
        (Property::TextFont(font), Some(Property::TextFont(inherited))) => {
            overrides.font = differs(font, inherited);
        },
        (Property::TextDecorations(decorations), Some(Property::TextDecorations(inherited))) => {
            overrides.underline = differs(&decorations.underline, &inherited.underline);
            overrides.strikethrough = differs(&decorations.strikethrough, &inherited.strikethrough);
        },
        (Property::TextColor(color), Some(Property::TextColor(inherited))) => {
            overrides.font_color = differs(color, inherited);
        },
        (Property::TextCapitalization(value), Some(Property::TextCapitalization(inherited))) => {
            overrides.capitalization = differs(value, inherited);
        },
        (Property::TextScript(value), Some(Property::TextScript(inherited))) => {
            overrides.script = differs(value, inherited);
        },
        (Property::TextBaselineShift(value), Some(Property::TextBaselineShift(inherited))) => {
            overrides.baseline_shift = differs(value, inherited);
        },
        (
            Property::TextCharacterSpacing(value),
            Some(Property::TextCharacterSpacing(inherited)),
        ) => overrides.character_spacing = differs(value, inherited),
        (Property::TextLigatures(value), Some(Property::TextLigatures(inherited))) => {
            overrides.ligatures = differs(value, inherited);
        },
        (Property::TextOutline(value), Some(Property::TextOutline(inherited))) => {
            overrides.outline = differs(value, inherited);
        },
        (Property::TextShadow(value), Some(Property::TextShadow(inherited))) => {
            overrides.shadow = differs(value, inherited);
        },
        (Property::TextBackground(value), Some(Property::TextBackground(inherited))) => {
            overrides.background = differs(value, inherited);
        },
        (Property::ParagraphBackground(value), None) => {
            overrides.paragraph_background = Some(*value);
        },
        (Property::Borders(value), None) => overrides.paragraph_borders = Some(*value),
        (Property::Flow(flow), None) => {
            overrides.hyphenation = Some(flow.hyphenation());
            overrides.keep_lines_together = Some(flow.keeps_lines_together());
            overrides.keep_with_next = Some(flow.keeps_with_next());
            overrides.start_on_new_page = Some(flow.starts_on_new_page());
            overrides.prevent_widow_orphan_lines = Some(flow.prevents_widow_orphan_lines());
        },
        (Property::WritingDirection(value), None) => overrides.writing_direction = Some(*value),
        (Property::FollowingStyle(value), None) => overrides.following_style = Some(*value),
        (Property::Alignment(value), None) => overrides.alignment = Some(*value),
        (Property::LineSpacing(value), None) => overrides.line_spacing = Some(*value),
        (Property::Spacing(spacing), None) => {
            overrides.space_before = Some(spacing.before);
            overrides.space_after = Some(spacing.after);
        },
        (Property::Indents(indents), None) => {
            overrides.first_line_indent = Some(indents.first_line);
            overrides.left_indent = Some(indents.left);
            overrides.right_indent = Some(indents.right);
        },
        (Property::DecimalTabCharacter(value), None) => {
            overrides.decimal_tab_character = Some(*value);
        },
        (Property::DefaultTabInterval(value), None) => {
            overrides.default_tab_interval = Some(*value);
        },
        (Property::TabStops(stops), None) => overrides.tab_stops = Some(stops.clone()),
        _ => return Err(mismatch),
    }
    Ok(())
}

const fn has(overrides: &Overrides, kind: PropertyKind) -> bool {
    match kind {
        PropertyKind::TextStyle => {
            overrides.point_size.is_some() || overrides.bold.is_some() || overrides.italic.is_some()
        },
        PropertyKind::TextFont => overrides.font.is_some(),
        PropertyKind::TextDecorations => {
            overrides.underline.is_some() || overrides.strikethrough.is_some()
        },
        PropertyKind::TextColor => overrides.font_color.is_some(),
        PropertyKind::TextCapitalization => overrides.capitalization.is_some(),
        PropertyKind::TextScript => overrides.script.is_some(),
        PropertyKind::TextBaselineShift => overrides.baseline_shift.is_some(),
        PropertyKind::TextCharacterSpacing => overrides.character_spacing.is_some(),
        PropertyKind::TextLigatures => overrides.ligatures.is_some(),
        PropertyKind::TextOutline => overrides.outline.is_some(),
        PropertyKind::TextShadow => overrides.shadow.is_some(),
        PropertyKind::TextBackground => overrides.background.is_some(),
        PropertyKind::ParagraphBackground => overrides.paragraph_background.is_some(),
        PropertyKind::Borders => overrides.paragraph_borders.is_some(),
        PropertyKind::Flow => {
            overrides.hyphenation.is_some()
                || overrides.keep_lines_together.is_some()
                || overrides.keep_with_next.is_some()
                || overrides.start_on_new_page.is_some()
                || overrides.prevent_widow_orphan_lines.is_some()
        },
        PropertyKind::WritingDirection => overrides.writing_direction.is_some(),
        PropertyKind::FollowingStyle => overrides.following_style.is_some(),
        PropertyKind::Alignment => overrides.alignment.is_some(),
        PropertyKind::LineSpacing => overrides.line_spacing.is_some(),
        PropertyKind::Spacing => {
            overrides.space_before.is_some() || overrides.space_after.is_some()
        },
        PropertyKind::Indents => {
            overrides.first_line_indent.is_some()
                || overrides.left_indent.is_some()
                || overrides.right_indent.is_some()
        },
        PropertyKind::DecimalTabCharacter => overrides.decimal_tab_character.is_some(),
        PropertyKind::DefaultTabInterval => overrides.default_tab_interval.is_some(),
        PropertyKind::TabStops => overrides.tab_stops.is_some(),
    }
}

fn clear(overrides: &mut Overrides, kind: PropertyKind) {
    match kind {
        PropertyKind::TextStyle => {
            overrides.point_size = None;
            overrides.bold = None;
            overrides.italic = None;
        },
        PropertyKind::TextFont => overrides.font = None,
        PropertyKind::TextDecorations => {
            overrides.underline = None;
            overrides.strikethrough = None;
        },
        PropertyKind::TextColor => overrides.font_color = None,
        PropertyKind::TextCapitalization => overrides.capitalization = None,
        PropertyKind::TextScript => overrides.script = None,
        PropertyKind::TextBaselineShift => overrides.baseline_shift = None,
        PropertyKind::TextCharacterSpacing => overrides.character_spacing = None,
        PropertyKind::TextLigatures => overrides.ligatures = None,
        PropertyKind::TextOutline => overrides.outline = None,
        PropertyKind::TextShadow => overrides.shadow = None,
        PropertyKind::TextBackground => overrides.background = None,
        PropertyKind::ParagraphBackground => overrides.paragraph_background = None,
        PropertyKind::Borders => overrides.paragraph_borders = None,
        PropertyKind::Flow => {
            overrides.hyphenation = None;
            overrides.keep_lines_together = None;
            overrides.keep_with_next = None;
            overrides.start_on_new_page = None;
            overrides.prevent_widow_orphan_lines = None;
        },
        PropertyKind::WritingDirection => overrides.writing_direction = None,
        PropertyKind::FollowingStyle => overrides.following_style = None,
        PropertyKind::Alignment => overrides.alignment = None,
        PropertyKind::LineSpacing => overrides.line_spacing = None,
        PropertyKind::Spacing => {
            overrides.space_before = None;
            overrides.space_after = None;
        },
        PropertyKind::Indents => {
            overrides.first_line_indent = None;
            overrides.left_indent = None;
            overrides.right_indent = None;
        },
        PropertyKind::DecimalTabCharacter => overrides.decimal_tab_character = None,
        PropertyKind::DefaultTabInterval => overrides.default_tab_interval = None,
        PropertyKind::TabStops => overrides.tab_stops = None,
    }
}

#[cfg(test)]
mod tests;
