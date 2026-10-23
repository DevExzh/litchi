//! Shared cell-format handles without native `SpreadsheetML` identifiers.

pub mod cell_format;
/// Applied cell-format views used by worksheet-facing APIs.
pub mod format;
/// The parsed `styles.xml` resource graph.
pub mod stylesheet;
/// The DrawingML theme resource associated with a workbook.
pub mod theme;

use std::fmt;
use std::hash::{Hash, Hasher};
use std::mem::size_of;
use std::ops::Range;
use std::sync::Arc;

/// Typed pattern-fill colour used by [`FillChange`].
pub use cell_format::Color as FillColor;
pub use cell_format::Format;
/// Typed SpreadsheetML pattern token used by [`FillChange`].
pub use cell_format::Pattern as FillPattern;

use crate::error::{Result, invalid};
use crate::workbook::Inner;
pub use stylesheet::alignment::{Horizontal, Indent, Reading, Rotation, Vertical};
pub use stylesheet::border::{Border, Color as BorderColor, Dir, Line, Side};
pub use stylesheet::font::{FontColor, FontColorKind, Scheme, Script, Underline};

/// Validate and own an authored SpreadsheetML number-format code.
///
/// The stylesheet model owns the format grammar so native source-preserving
/// edits and the semantic writer apply the same authoring policy.  The native
/// codec intentionally keeps the returned text byte-for-byte as authored; it
/// does not resolve friendly aliases or otherwise normalize the code.
pub(crate) fn validate_number_format_code(value: &str) -> Result<Arc<str>> {
    stylesheet::NumberFormat::validate_code(value).map_err(|error| invalid(error.to_string()))?;
    Ok(Arc::<str>::from(value))
}

/// Default native staged-memory ceiling used before copying authored font text.
pub(crate) const MAX_FONT_CHANGE_NAME_BYTES: usize = 64 * 1024 * 1024;
const MAX_FONT_FAMILY: u32 = 5;
const MAX_FONT_CHARSET: u32 = 255;

/// Checked finite Excel font size in points.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct FontSize(u64);

impl FontSize {
    fn new(value: f64) -> Result<Self> {
        if !(0.5..=409.0).contains(&value) {
            return Err(invalid(
                "font size must be finite, at least 0.5, and at most 409 points",
            ));
        }
        Ok(Self(value.to_bits()))
    }

    const fn get(self) -> f64 {
        f64::from_bits(self.0)
    }
}

/// Typed changes to the font properties currently supported by the immutable
/// worksheet edit API.
///
/// A builder which has not mentioned a property leaves it unchanged. Passing
/// `Some(value)` authors that value; passing `None` to a mentioned property
/// removes its authored element from the cloned font. This operation does not
/// resolve row, column or named-style inheritance.
/// The physical font and cell-format indexes allocated by a commit remain an
/// implementation detail.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct FontChange {
    bold: Option<Option<bool>>,
    italic: Option<Option<bool>>,
    strike: Option<Option<bool>>,
    condense: Option<Option<bool>>,
    extend: Option<Option<bool>>,
    outline: Option<Option<bool>>,
    shadow: Option<Option<bool>>,
    underline: Option<Option<Underline>>,
    color: Option<Option<FontColor>>,
    name: Option<Option<Arc<str>>>,
    size: Option<Option<FontSize>>,
    script: Option<Option<Script>>,
    family: Option<Option<u32>>,
    charset: Option<Option<u32>>,
    scheme: Option<Option<Scheme>>,
}

impl FontChange {
    /// Start an empty font change.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            bold: None,
            italic: None,
            strike: None,
            condense: None,
            extend: None,
            outline: None,
            shadow: None,
            underline: None,
            color: None,
            name: None,
            size: None,
            script: None,
            family: None,
            charset: None,
            scheme: None,
        }
    }

    /// Set or clear the authored bold value.
    #[must_use]
    pub const fn bold(mut self, value: Option<bool>) -> Self {
        self.bold = Some(value);
        self
    }

    /// Set or clear the authored italic value.
    #[must_use]
    pub const fn italic(mut self, value: Option<bool>) -> Self {
        self.italic = Some(value);
        self
    }

    /// Set or clear the authored strike-through value.
    #[must_use]
    pub const fn strike(mut self, value: Option<bool>) -> Self {
        self.strike = Some(value);
        self
    }

    /// Set or clear the authored condensed-spacing value.
    #[must_use]
    pub const fn condense(mut self, value: Option<bool>) -> Self {
        self.condense = Some(value);
        self
    }

    /// Set or clear the authored extended-spacing value.
    #[must_use]
    pub const fn extend(mut self, value: Option<bool>) -> Self {
        self.extend = Some(value);
        self
    }

    /// Set or clear the authored outlined-glyph value.
    #[must_use]
    pub const fn outline(mut self, value: Option<bool>) -> Self {
        self.outline = Some(value);
        self
    }

    /// Set or clear the authored shadowed-glyph value.
    #[must_use]
    pub const fn shadow(mut self, value: Option<bool>) -> Self {
        self.shadow = Some(value);
        self
    }

    /// Set or clear the authored underline style.
    #[must_use]
    pub const fn underline(mut self, value: Option<Underline>) -> Self {
        self.underline = Some(value);
        self
    }

    /// Set or clear the authored font color.
    ///
    /// Source-only lexical markup carried by a parsed [`FontColor`] is
    /// discarded when the value enters this request; its checked base and
    /// tint remain part of the semantic request.
    #[must_use]
    pub fn color(mut self, value: Option<FontColor>) -> Self {
        self.color = Some(value.map(FontColor::into_authoring));
        self
    }

    /// Set or clear the authored typeface name.
    ///
    /// The name must be non-empty and contain only XML 1.0 characters. The
    /// value is copied into shared owned storage so a staged change can outlive
    /// the caller's input buffer.
    pub fn name(mut self, value: Option<&str>) -> Result<Self> {
        self.name = Some(match value {
            Some(value) => Some(validate_font_name(value)?),
            None => None,
        });
        Ok(self)
    }

    /// Set or clear the authored font size in points.
    ///
    /// Excel's authoring domain is finite `[0.5, 409]` points.
    pub fn size(mut self, value: Option<f64>) -> Result<Self> {
        self.size = Some(match value {
            Some(value) => Some(FontSize::new(value)?),
            None => None,
        });
        Ok(self)
    }

    /// Set or clear the authored baseline placement.
    #[must_use]
    pub const fn script(mut self, value: Option<Script>) -> Self {
        self.script = Some(value);
        self
    }

    /// Set or clear the authored font family class (0..=5).
    pub fn family(mut self, value: Option<u32>) -> Result<Self> {
        self.family = Some(match value {
            Some(value) if value <= MAX_FONT_FAMILY => Some(value),
            Some(value) => {
                return Err(invalid(format!(
                    "font family must be between 0 and {MAX_FONT_FAMILY} (got {value})"
                )));
            },
            None => None,
        });
        Ok(self)
    }

    /// Set or clear the authored character-set identifier (0..=255).
    pub fn charset(mut self, value: Option<u32>) -> Result<Self> {
        self.charset = Some(match value {
            Some(value) if value <= MAX_FONT_CHARSET => Some(value),
            Some(value) => {
                return Err(invalid(format!(
                    "font charset must be between 0 and {MAX_FONT_CHARSET} (got {value})"
                )));
            },
            None => None,
        });
        Ok(self)
    }

    /// Set or clear the authored theme font scheme.
    #[must_use]
    pub const fn scheme(mut self, value: Option<Scheme>) -> Self {
        self.scheme = Some(value);
        self
    }

    /// Whether this request changes no font property.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.bold.is_none()
            && self.italic.is_none()
            && self.strike.is_none()
            && self.condense.is_none()
            && self.extend.is_none()
            && self.outline.is_none()
            && self.shadow.is_none()
            && self.underline.is_none()
            && self.color.is_none()
            && self.name.is_none()
            && self.size.is_none()
            && self.script.is_none()
            && self.family.is_none()
            && self.charset.is_none()
            && self.scheme.is_none()
    }

    pub(crate) const fn bold_change(&self) -> Option<Option<bool>> {
        self.bold
    }

    pub(crate) const fn italic_change(&self) -> Option<Option<bool>> {
        self.italic
    }

    pub(crate) const fn strike_change(&self) -> Option<Option<bool>> {
        self.strike
    }

    pub(crate) const fn condense_change(&self) -> Option<Option<bool>> {
        self.condense
    }

    pub(crate) const fn extend_change(&self) -> Option<Option<bool>> {
        self.extend
    }

    pub(crate) const fn outline_change(&self) -> Option<Option<bool>> {
        self.outline
    }

    pub(crate) const fn shadow_change(&self) -> Option<Option<bool>> {
        self.shadow
    }

    pub(crate) const fn underline_change(&self) -> Option<Option<Underline>> {
        self.underline
    }

    pub(crate) fn color_change(&self) -> Option<Option<&FontColor>> {
        self.color.as_ref().map(|value| value.as_ref())
    }

    pub(crate) fn name_change(&self) -> Option<Option<&str>> {
        self.name.as_ref().map(|value| value.as_deref())
    }

    pub(crate) const fn size_change(&self) -> Option<Option<f64>> {
        match self.size {
            Some(Some(value)) => Some(Some(value.get())),
            Some(None) => Some(None),
            None => None,
        }
    }

    pub(crate) const fn script_change(&self) -> Option<Option<Script>> {
        self.script
    }

    pub(crate) const fn family_change(&self) -> Option<Option<u32>> {
        self.family
    }

    pub(crate) const fn charset_change(&self) -> Option<Option<u32>> {
        self.charset
    }

    pub(crate) const fn scheme_change(&self) -> Option<Option<Scheme>> {
        self.scheme
    }

    /// Merge a later change into this staged request.
    pub(crate) fn merge(&mut self, later: &Self) {
        if later.bold.is_some() {
            self.bold = later.bold;
        }
        if later.italic.is_some() {
            self.italic = later.italic;
        }
        if later.strike.is_some() {
            self.strike = later.strike;
        }
        if later.condense.is_some() {
            self.condense = later.condense;
        }
        if later.extend.is_some() {
            self.extend = later.extend;
        }
        if later.outline.is_some() {
            self.outline = later.outline;
        }
        if later.shadow.is_some() {
            self.shadow = later.shadow;
        }
        if later.underline.is_some() {
            self.underline = later.underline;
        }
        if later.color.is_some() {
            self.color = later.color.clone();
        }
        if later.name.is_some() {
            self.name = later.name.clone();
        }
        if later.size.is_some() {
            self.size = later.size;
        }
        if later.script.is_some() {
            self.script = later.script;
        }
        if later.family.is_some() {
            self.family = later.family;
        }
        if later.charset.is_some() {
            self.charset = later.charset;
        }
        if later.scheme.is_some() {
            self.scheme = later.scheme;
        }
    }

    pub(crate) const fn overlaps(&self, other: &Self) -> bool {
        (self.bold.is_some() && other.bold.is_some())
            || (self.italic.is_some() && other.italic.is_some())
            || (self.strike.is_some() && other.strike.is_some())
            || (self.condense.is_some() && other.condense.is_some())
            || (self.extend.is_some() && other.extend.is_some())
            || (self.outline.is_some() && other.outline.is_some())
            || (self.shadow.is_some() && other.shadow.is_some())
            || (self.underline.is_some() && other.underline.is_some())
            || (self.color.is_some() && other.color.is_some())
            || (self.name.is_some() && other.name.is_some())
            || (self.size.is_some() && other.size.is_some())
            || (self.script.is_some() && other.script.is_some())
            || (self.family.is_some() && other.family.is_some())
            || (self.charset.is_some() && other.charset.is_some())
            || (self.scheme.is_some() && other.scheme.is_some())
    }

    /// Logical bytes retained by this authored font request while it waits
    /// for stylesheet lowering.
    pub(crate) fn retained_bytes(&self) -> usize {
        size_of::<Self>().saturating_add(
            self.name
                .as_ref()
                .and_then(|value| value.as_ref())
                .map_or(0, |value| value.len()),
        )
    }
}

/// Typed changes to an authored fill.
///
/// Each builder is tri-state: an untouched property remains unchanged, an
/// `Some` value authors that property, and `None` clears the authored value.
/// The request edits only the local fill record; it does not resolve theme or
/// style inheritance. Granular pattern fields retain their existing tri-state
/// behavior. [`Self::replace`] replaces the complete local fill while keeping
/// the native codec's source-preserving outer `fill` markup.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct FillChange {
    pattern: Option<Option<FillPattern>>,
    foreground: Option<Option<FillColor>>,
    background: Option<Option<FillColor>>,
    replacement: Option<Arc<cell_format::Fill>>,
}

impl FillChange {
    /// Start an empty fill change.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            pattern: None,
            foreground: None,
            background: None,
            replacement: None,
        }
    }

    /// Replace the complete authored fill record.
    #[must_use]
    pub fn replace(value: cell_format::Fill) -> Self {
        Self {
            pattern: None,
            foreground: None,
            background: None,
            replacement: Some(Arc::new(value)),
        }
    }

    /// Set or clear the authored pattern token.
    #[must_use]
    pub const fn pattern(mut self, value: Option<FillPattern>) -> Self {
        self.pattern = Some(value);
        self
    }

    /// Set or clear the authored foreground (`fgColor`).
    #[must_use]
    pub const fn foreground(mut self, value: Option<FillColor>) -> Self {
        self.foreground = Some(value);
        self
    }

    /// Set or clear the authored background (`bgColor`).
    #[must_use]
    pub const fn background(mut self, value: Option<FillColor>) -> Self {
        self.background = Some(value);
        self
    }

    /// Whether this request mentions no fill property.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.pattern.is_none()
            && self.foreground.is_none()
            && self.background.is_none()
            && self.replacement.is_none()
    }

    pub(crate) fn replacement(&self) -> Option<&cell_format::Fill> {
        self.replacement.as_deref()
    }

    pub(crate) const fn pattern_change(&self) -> Option<Option<FillPattern>> {
        self.pattern
    }

    pub(crate) const fn foreground_change(&self) -> Option<Option<FillColor>> {
        self.foreground
    }

    pub(crate) const fn background_change(&self) -> Option<Option<FillColor>> {
        self.background
    }

    /// Merge a later change into this staged request.
    pub(crate) fn merge(&mut self, later: &Self) {
        if later.replacement.is_some() {
            *self = later.clone();
            return;
        }
        if later.pattern.is_some() {
            self.pattern = later.pattern;
        }
        if later.foreground.is_some() {
            self.foreground = later.foreground;
        }
        if later.background.is_some() {
            self.background = later.background;
        }
    }

    pub(crate) fn overlaps(&self, other: &Self) -> bool {
        if self.replacement.is_some() || other.replacement.is_some() {
            return !self.is_empty() && !other.is_empty();
        }
        (self.pattern.is_some() && other.pattern.is_some())
            || (self.foreground.is_some() && other.foreground.is_some())
            || (self.background.is_some() && other.background.is_some())
    }

    /// Logical memory retained by a whole-fill replacement in a semantic
    /// transaction. Granular changes contain only scalar values and carry no
    /// owned payload. The replacement's `Arc` control block and gradient-stop
    /// backing storage are charged conservatively before native lowering.
    pub(crate) fn retained_bytes(&self) -> usize {
        let Some(fill) = self.replacement() else {
            return 0;
        };
        let stop_bytes = fill.gradient().map_or(0, |gradient| {
            gradient
                .stops()
                .len()
                .saturating_mul(size_of::<cell_format::GradientStop>())
        });
        size_of::<cell_format::Fill>()
            .saturating_add(stop_bytes)
            .saturating_add(size_of::<usize>().saturating_mul(2))
    }
}

/// The border edge addressed by a granular border change.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) enum BorderEdge {
    Start,
    End,
    Left,
    Right,
    Top,
    Bottom,
    Diagonal,
    Vertical,
    Horizontal,
}

/// A tri-state change to one authored border side.
///
/// The whole-side setter is useful when a caller already has a complete
/// typed side. The line and color setters intentionally remain independent so
/// changing a color can retain the existing line and vice versa.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub(crate) struct BorderSideChange {
    whole: Option<Option<Side>>,
    line: Option<Option<Line>>,
    color: Option<Option<BorderColor>>,
}

impl BorderSideChange {
    /// Start an empty side change.
    #[must_use]
    pub(crate) const fn new() -> Self {
        Self {
            whole: None,
            line: None,
            color: None,
        }
    }

    /// Replace the complete side, or remove it when `value` is `None`.
    #[must_use]
    pub(crate) fn side(mut self, value: Option<Side>) -> Self {
        self.whole = Some(value);
        self.line = None;
        self.color = None;
        self
    }

    /// Set or clear only the side's line style.
    #[must_use]
    pub(crate) const fn line(mut self, value: Option<Line>) -> Self {
        self.line = Some(value);
        self
    }

    /// Set or clear only the side's color.
    #[must_use]
    pub(crate) const fn color(mut self, value: Option<BorderColor>) -> Self {
        self.color = Some(value);
        self
    }

    /// Return whether this side change contains no request.
    #[must_use]
    pub(crate) const fn is_empty(&self) -> bool {
        self.whole.is_none() && self.line.is_none() && self.color.is_none()
    }

    pub(crate) fn whole_change(&self) -> Option<Option<&Side>> {
        match self.whole.as_ref() {
            Some(Some(value)) => Some(Some(value)),
            Some(None) => Some(None),
            None => None,
        }
    }

    pub(crate) const fn line_change(&self) -> Option<Option<Line>> {
        self.line
    }

    pub(crate) const fn color_change(&self) -> Option<Option<BorderColor>> {
        self.color
    }

    pub(crate) fn merge(&mut self, later: &Self) {
        if later.whole.is_some() {
            self.whole = later.whole.clone();
            self.line = None;
            self.color = None;
        }
        if later.line.is_some() {
            self.line = later.line;
        }
        if later.color.is_some() {
            self.color = later.color;
        }
    }

    pub(crate) const fn overlaps(&self, other: &Self) -> bool {
        (self.whole.is_some()
            && (other.line.is_some() || other.color.is_some() || other.whole.is_some()))
            || (other.whole.is_some()
                && (self.line.is_some() || self.color.is_some() || self.whole.is_some()))
            || (self.line.is_some() && other.line.is_some())
            || (self.color.is_some() && other.color.is_some())
    }
}

/// Typed tri-state changes to the local cell border.
///
/// Each side can be replaced as a complete typed side or changed one field at
/// a time. Diagonal direction flags and outline are independent authored
/// attributes. An unset field is retained; a mentioned None removes its
/// authored value.
///
/// ```
/// use litchi_xlsx::style::{BorderChange, BorderColor, Line, Side};
///
/// let change = BorderChange::new()
///     .left(Some(Side::new(Line::Thin)))
///     .left_color(Some(BorderColor::theme(4)))
///     .top_line(Some(Line::Double))
///     .diagonal_up(Some(true));
/// assert!(!change.is_empty());
/// ```
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct BorderChange {
    start: Option<BorderSideChange>,
    end: Option<BorderSideChange>,
    left: Option<BorderSideChange>,
    right: Option<BorderSideChange>,
    top: Option<BorderSideChange>,
    bottom: Option<BorderSideChange>,
    diagonal: Option<BorderSideChange>,
    vertical: Option<BorderSideChange>,
    horizontal: Option<BorderSideChange>,
    diagonal_up: Option<Option<bool>>,
    diagonal_down: Option<Option<bool>>,
    outline: Option<Option<bool>>,
}

impl BorderChange {
    /// Start an empty border change.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            start: None,
            end: None,
            left: None,
            right: None,
            top: None,
            bottom: None,
            diagonal: None,
            vertical: None,
            horizontal: None,
            diagonal_up: None,
            diagonal_down: None,
            outline: None,
        }
    }

    /// Replace the logical start side, or remove it when `value` is `None`.
    #[must_use]
    pub fn start(mut self, value: Option<Side>) -> Self {
        self.start = Some(BorderSideChange::new().side(value));
        self
    }

    /// Replace the logical end side, or remove it when `value` is `None`.
    #[must_use]
    pub fn end(mut self, value: Option<Side>) -> Self {
        self.end = Some(BorderSideChange::new().side(value));
        self
    }

    /// Replace the physical left side, or remove it when `value` is `None`.
    #[must_use]
    pub fn left(mut self, value: Option<Side>) -> Self {
        self.left = Some(BorderSideChange::new().side(value));
        self
    }

    /// Replace the physical right side, or remove it when `value` is `None`.
    #[must_use]
    pub fn right(mut self, value: Option<Side>) -> Self {
        self.right = Some(BorderSideChange::new().side(value));
        self
    }

    /// Replace the top side, or remove it when `value` is `None`.
    #[must_use]
    pub fn top(mut self, value: Option<Side>) -> Self {
        self.top = Some(BorderSideChange::new().side(value));
        self
    }

    /// Replace the bottom side, or remove it when `value` is `None`.
    #[must_use]
    pub fn bottom(mut self, value: Option<Side>) -> Self {
        self.bottom = Some(BorderSideChange::new().side(value));
        self
    }

    /// Replace the diagonal side, or remove it when `value` is `None`.
    #[must_use]
    pub fn diagonal(mut self, value: Option<Side>) -> Self {
        self.diagonal = Some(BorderSideChange::new().side(value));
        self
    }

    /// Replace the vertical interior side, or remove it when `value` is `None`.
    #[must_use]
    pub fn vertical(mut self, value: Option<Side>) -> Self {
        self.vertical = Some(BorderSideChange::new().side(value));
        self
    }

    /// Replace the horizontal interior side, or remove it when `value` is `None`.
    #[must_use]
    pub fn horizontal(mut self, value: Option<Side>) -> Self {
        self.horizontal = Some(BorderSideChange::new().side(value));
        self
    }

    /// Set or clear only the logical start line style.
    #[must_use]
    pub fn start_line(self, value: Option<Line>) -> Self {
        self.with_side(BorderEdge::Start, |side| side.line(value))
    }

    /// Set or clear only the logical end line style.
    #[must_use]
    pub fn end_line(self, value: Option<Line>) -> Self {
        self.with_side(BorderEdge::End, |side| side.line(value))
    }

    /// Set or clear only the physical left line style.
    #[must_use]
    pub fn left_line(self, value: Option<Line>) -> Self {
        self.with_side(BorderEdge::Left, |side| side.line(value))
    }

    /// Set or clear only the physical right line style.
    #[must_use]
    pub fn right_line(self, value: Option<Line>) -> Self {
        self.with_side(BorderEdge::Right, |side| side.line(value))
    }

    /// Set or clear only the top line style.
    #[must_use]
    pub fn top_line(self, value: Option<Line>) -> Self {
        self.with_side(BorderEdge::Top, |side| side.line(value))
    }

    /// Set or clear only the bottom line style.
    #[must_use]
    pub fn bottom_line(self, value: Option<Line>) -> Self {
        self.with_side(BorderEdge::Bottom, |side| side.line(value))
    }

    /// Set or clear only the diagonal line style.
    #[must_use]
    pub fn diagonal_line(self, value: Option<Line>) -> Self {
        self.with_side(BorderEdge::Diagonal, |side| side.line(value))
    }

    /// Set or clear only the vertical line style.
    #[must_use]
    pub fn vertical_line(self, value: Option<Line>) -> Self {
        self.with_side(BorderEdge::Vertical, |side| side.line(value))
    }

    /// Set or clear only the horizontal line style.
    #[must_use]
    pub fn horizontal_line(self, value: Option<Line>) -> Self {
        self.with_side(BorderEdge::Horizontal, |side| side.line(value))
    }

    /// Set or clear only the logical start color.
    #[must_use]
    pub fn start_color(self, value: Option<BorderColor>) -> Self {
        self.with_side(BorderEdge::Start, |side| side.color(value))
    }

    /// Set or clear only the logical end color.
    #[must_use]
    pub fn end_color(self, value: Option<BorderColor>) -> Self {
        self.with_side(BorderEdge::End, |side| side.color(value))
    }

    /// Set or clear only the physical left color.
    #[must_use]
    pub fn left_color(self, value: Option<BorderColor>) -> Self {
        self.with_side(BorderEdge::Left, |side| side.color(value))
    }

    /// Set or clear only the physical right color.
    #[must_use]
    pub fn right_color(self, value: Option<BorderColor>) -> Self {
        self.with_side(BorderEdge::Right, |side| side.color(value))
    }

    /// Set or clear only the top color.
    #[must_use]
    pub fn top_color(self, value: Option<BorderColor>) -> Self {
        self.with_side(BorderEdge::Top, |side| side.color(value))
    }

    /// Set or clear only the bottom color.
    #[must_use]
    pub fn bottom_color(self, value: Option<BorderColor>) -> Self {
        self.with_side(BorderEdge::Bottom, |side| side.color(value))
    }

    /// Set or clear only the diagonal color.
    #[must_use]
    pub fn diagonal_color(self, value: Option<BorderColor>) -> Self {
        self.with_side(BorderEdge::Diagonal, |side| side.color(value))
    }

    /// Set or clear only the vertical color.
    #[must_use]
    pub fn vertical_color(self, value: Option<BorderColor>) -> Self {
        self.with_side(BorderEdge::Vertical, |side| side.color(value))
    }

    /// Set or clear only the horizontal color.
    #[must_use]
    pub fn horizontal_color(self, value: Option<BorderColor>) -> Self {
        self.with_side(BorderEdge::Horizontal, |side| side.color(value))
    }

    /// Set or clear the authored diagonal-up flag.
    #[must_use]
    pub const fn diagonal_up(mut self, value: Option<bool>) -> Self {
        self.diagonal_up = Some(value);
        self
    }

    /// Set or clear the authored diagonal-down flag.
    #[must_use]
    pub const fn diagonal_down(mut self, value: Option<bool>) -> Self {
        self.diagonal_down = Some(value);
        self
    }

    /// Set or clear the authored outline flag.
    #[must_use]
    pub const fn outline(mut self, value: Option<bool>) -> Self {
        self.outline = Some(value);
        self
    }

    /// Return whether this change contains no authored border request.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.start.as_ref().is_none_or(BorderSideChange::is_empty)
            && self.end.as_ref().is_none_or(BorderSideChange::is_empty)
            && self.left.as_ref().is_none_or(BorderSideChange::is_empty)
            && self.right.as_ref().is_none_or(BorderSideChange::is_empty)
            && self.top.as_ref().is_none_or(BorderSideChange::is_empty)
            && self.bottom.as_ref().is_none_or(BorderSideChange::is_empty)
            && self
                .diagonal
                .as_ref()
                .is_none_or(BorderSideChange::is_empty)
            && self
                .vertical
                .as_ref()
                .is_none_or(BorderSideChange::is_empty)
            && self
                .horizontal
                .as_ref()
                .is_none_or(BorderSideChange::is_empty)
            && self.diagonal_up.is_none()
            && self.diagonal_down.is_none()
            && self.outline.is_none()
    }

    pub(crate) fn side_change(&self, edge: BorderEdge) -> Option<&BorderSideChange> {
        match edge {
            BorderEdge::Start => self.start.as_ref(),
            BorderEdge::End => self.end.as_ref(),
            BorderEdge::Left => self.left.as_ref(),
            BorderEdge::Right => self.right.as_ref(),
            BorderEdge::Top => self.top.as_ref(),
            BorderEdge::Bottom => self.bottom.as_ref(),
            BorderEdge::Diagonal => self.diagonal.as_ref(),
            BorderEdge::Vertical => self.vertical.as_ref(),
            BorderEdge::Horizontal => self.horizontal.as_ref(),
        }
    }

    pub(crate) const fn diagonal_up_change(&self) -> Option<Option<bool>> {
        self.diagonal_up
    }

    pub(crate) const fn diagonal_down_change(&self) -> Option<Option<bool>> {
        self.diagonal_down
    }

    pub(crate) const fn outline_change(&self) -> Option<Option<bool>> {
        self.outline
    }

    pub(crate) fn merge(&mut self, later: &Self) {
        for edge in [
            BorderEdge::Start,
            BorderEdge::End,
            BorderEdge::Left,
            BorderEdge::Right,
            BorderEdge::Top,
            BorderEdge::Bottom,
            BorderEdge::Diagonal,
            BorderEdge::Vertical,
            BorderEdge::Horizontal,
        ] {
            if let Some(change) = later.side_change(edge) {
                let slot = self.edge_mut(edge);
                if let Some(current) = slot {
                    current.merge(change);
                } else {
                    *slot = Some(change.clone());
                }
            }
        }
        if later.diagonal_up.is_some() {
            self.diagonal_up = later.diagonal_up;
        }
        if later.diagonal_down.is_some() {
            self.diagonal_down = later.diagonal_down;
        }
        if later.outline.is_some() {
            self.outline = later.outline;
        }
    }

    pub(crate) fn overlaps(&self, other: &Self) -> bool {
        [
            BorderEdge::Start,
            BorderEdge::End,
            BorderEdge::Left,
            BorderEdge::Right,
            BorderEdge::Top,
            BorderEdge::Bottom,
            BorderEdge::Diagonal,
            BorderEdge::Vertical,
            BorderEdge::Horizontal,
        ]
        .into_iter()
        .any(|edge| {
            self.side_change(edge)
                .zip(other.side_change(edge))
                .is_some_and(|(left, right)| left.overlaps(right))
        }) || self.diagonal_up.is_some() && other.diagonal_up.is_some()
            || self.diagonal_down.is_some() && other.diagonal_down.is_some()
            || self.outline.is_some() && other.outline.is_some()
    }

    /// Logical bytes retained by this typed border request.
    pub(crate) const fn retained_bytes(&self) -> usize {
        size_of::<Self>()
    }

    fn edge_mut(&mut self, edge: BorderEdge) -> &mut Option<BorderSideChange> {
        match edge {
            BorderEdge::Start => &mut self.start,
            BorderEdge::End => &mut self.end,
            BorderEdge::Left => &mut self.left,
            BorderEdge::Right => &mut self.right,
            BorderEdge::Top => &mut self.top,
            BorderEdge::Bottom => &mut self.bottom,
            BorderEdge::Diagonal => &mut self.diagonal,
            BorderEdge::Vertical => &mut self.vertical,
            BorderEdge::Horizontal => &mut self.horizontal,
        }
    }

    fn with_side(
        mut self,
        edge: BorderEdge,
        update: impl FnOnce(BorderSideChange) -> BorderSideChange,
    ) -> Self {
        let current = self.edge_mut(edge).take().unwrap_or_default();
        *self.edge_mut(edge) = Some(update(current));
        self
    }
}

/// Typed changes to the authored alignment properties of a cell format.
///
/// An unset builder leaves a property unchanged. Passing `Some(value)` writes
/// that value and passing `None` to a mentioned property removes its authored
/// attribute. The request only describes local `xf` markup; it does not
/// resolve row, column, or named-style inheritance.
///
/// ```
/// use litchi_xlsx::{Package, style::{AlignmentChange, Horizontal, Indent}};
///
/// let workbook = Package::new()?;
/// let mut edit = workbook.edit()?;
/// let change = AlignmentChange::new()
///     .horizontal(Some(Horizontal::Center))
///     .wrap_text(Some(true))
///     .indent(Some(Indent::new(2)));
/// edit.sheet("Sheet1")?.expect("initial worksheet")
///     .alignment("A1", &change)?;
/// let updated = edit.commit()?.into_package();
/// # let _ = updated;
/// # Ok::<(), litchi_xlsx::Error>(())
/// ```
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct AlignmentChange {
    horizontal: Option<Option<Horizontal>>,
    vertical: Option<Option<Vertical>>,
    text_rotation: Option<Option<Rotation>>,
    wrap_text: Option<Option<bool>>,
    indent: Option<Option<Indent>>,
    relative_indent: Option<Option<i32>>,
    justify_last_line: Option<Option<bool>>,
    shrink_to_fit: Option<Option<bool>>,
    reading_order: Option<Option<Reading>>,
}

impl AlignmentChange {
    /// Start an empty alignment change.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            horizontal: None,
            vertical: None,
            text_rotation: None,
            wrap_text: None,
            indent: None,
            relative_indent: None,
            justify_last_line: None,
            shrink_to_fit: None,
            reading_order: None,
        }
    }

    /// Set or clear the horizontal alignment.
    #[must_use]
    pub const fn horizontal(mut self, value: Option<Horizontal>) -> Self {
        self.horizontal = Some(value);
        self
    }

    /// Set or clear the vertical alignment.
    #[must_use]
    pub const fn vertical(mut self, value: Option<Vertical>) -> Self {
        self.vertical = Some(value);
        self
    }

    /// Set or clear text rotation.
    #[must_use]
    pub const fn text_rotation(mut self, value: Option<Rotation>) -> Self {
        self.text_rotation = Some(value);
        self
    }

    /// Set or clear the authored wrap-text flag.
    #[must_use]
    pub const fn wrap_text(mut self, value: Option<bool>) -> Self {
        self.wrap_text = Some(value);
        self
    }

    /// Set or clear the indentation level.
    #[must_use]
    pub const fn indent(mut self, value: Option<Indent>) -> Self {
        self.indent = Some(value);
        self
    }

    /// Set or clear the signed relative indentation.
    #[must_use]
    pub const fn relative_indent(mut self, value: Option<i32>) -> Self {
        self.relative_indent = Some(value);
        self
    }

    /// Set or clear the authored justify-last-line flag.
    #[must_use]
    pub const fn justify_last_line(mut self, value: Option<bool>) -> Self {
        self.justify_last_line = Some(value);
        self
    }

    /// Set or clear the authored shrink-to-fit flag.
    #[must_use]
    pub const fn shrink_to_fit(mut self, value: Option<bool>) -> Self {
        self.shrink_to_fit = Some(value);
        self
    }

    /// Set or clear the reading order.
    #[must_use]
    pub const fn reading_order(mut self, value: Option<Reading>) -> Self {
        self.reading_order = Some(value);
        self
    }

    /// Whether this request mentions no alignment property.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.horizontal.is_none()
            && self.vertical.is_none()
            && self.text_rotation.is_none()
            && self.wrap_text.is_none()
            && self.indent.is_none()
            && self.relative_indent.is_none()
            && self.justify_last_line.is_none()
            && self.shrink_to_fit.is_none()
            && self.reading_order.is_none()
    }

    pub(crate) const fn horizontal_change(&self) -> Option<Option<Horizontal>> {
        self.horizontal
    }

    pub(crate) const fn vertical_change(&self) -> Option<Option<Vertical>> {
        self.vertical
    }

    pub(crate) const fn text_rotation_change(&self) -> Option<Option<Rotation>> {
        self.text_rotation
    }

    pub(crate) const fn wrap_text_change(&self) -> Option<Option<bool>> {
        self.wrap_text
    }

    pub(crate) const fn indent_change(&self) -> Option<Option<Indent>> {
        self.indent
    }

    pub(crate) const fn relative_indent_change(&self) -> Option<Option<i32>> {
        self.relative_indent
    }

    pub(crate) const fn justify_last_line_change(&self) -> Option<Option<bool>> {
        self.justify_last_line
    }

    pub(crate) const fn shrink_to_fit_change(&self) -> Option<Option<bool>> {
        self.shrink_to_fit
    }

    pub(crate) const fn reading_order_change(&self) -> Option<Option<Reading>> {
        self.reading_order
    }

    /// Merge a later change into this staged request.
    pub(crate) const fn merge(&mut self, later: &Self) {
        if later.horizontal.is_some() {
            self.horizontal = later.horizontal;
        }
        if later.vertical.is_some() {
            self.vertical = later.vertical;
        }
        if later.text_rotation.is_some() {
            self.text_rotation = later.text_rotation;
        }
        if later.wrap_text.is_some() {
            self.wrap_text = later.wrap_text;
        }
        if later.indent.is_some() {
            self.indent = later.indent;
        }
        if later.relative_indent.is_some() {
            self.relative_indent = later.relative_indent;
        }
        if later.justify_last_line.is_some() {
            self.justify_last_line = later.justify_last_line;
        }
        if later.shrink_to_fit.is_some() {
            self.shrink_to_fit = later.shrink_to_fit;
        }
        if later.reading_order.is_some() {
            self.reading_order = later.reading_order;
        }
    }

    pub(crate) const fn overlaps(&self, other: &Self) -> bool {
        (self.horizontal.is_some() && other.horizontal.is_some())
            || (self.vertical.is_some() && other.vertical.is_some())
            || (self.text_rotation.is_some() && other.text_rotation.is_some())
            || (self.wrap_text.is_some() && other.wrap_text.is_some())
            || (self.indent.is_some() && other.indent.is_some())
            || (self.relative_indent.is_some() && other.relative_indent.is_some())
            || (self.justify_last_line.is_some() && other.justify_last_line.is_some())
            || (self.shrink_to_fit.is_some() && other.shrink_to_fit.is_some())
            || (self.reading_order.is_some() && other.reading_order.is_some())
    }

    /// Logical bytes retained by this typed alignment request.
    pub(crate) const fn retained_bytes(&self) -> usize {
        size_of::<Self>()
    }
}

/// Typed changes to authored cell protection properties.
///
/// An unset builder leaves a property unchanged. Passing `Some(value)` writes
/// that value and passing `None` to a mentioned property removes its authored
/// attribute. Protection activation remains an independent XF concern.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct ProtectionChange {
    locked: Option<Option<bool>>,
    hidden: Option<Option<bool>>,
}

impl ProtectionChange {
    /// Start an empty protection change.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            locked: None,
            hidden: None,
        }
    }

    /// Set or clear the authored locked flag.
    #[must_use]
    pub const fn locked(mut self, value: Option<bool>) -> Self {
        self.locked = Some(value);
        self
    }

    /// Set or clear the authored hidden-formula flag.
    #[must_use]
    pub const fn hidden(mut self, value: Option<bool>) -> Self {
        self.hidden = Some(value);
        self
    }

    /// Whether this request mentions no protection property.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.locked.is_none() && self.hidden.is_none()
    }

    pub(crate) const fn locked_change(&self) -> Option<Option<bool>> {
        self.locked
    }

    pub(crate) const fn hidden_change(&self) -> Option<Option<bool>> {
        self.hidden
    }

    /// Merge a later change into this staged request.
    pub(crate) const fn merge(&mut self, later: &Self) {
        if later.locked.is_some() {
            self.locked = later.locked;
        }
        if later.hidden.is_some() {
            self.hidden = later.hidden;
        }
    }

    pub(crate) const fn overlaps(&self, other: &Self) -> bool {
        (self.locked.is_some() && other.locked.is_some())
            || (self.hidden.is_some() && other.hidden.is_some())
    }

    /// Logical bytes retained by this typed protection request.
    pub(crate) const fn retained_bytes(&self) -> usize {
        size_of::<Self>()
    }
}

fn validate_font_name(value: &str) -> Result<Arc<str>> {
    if value.is_empty() {
        return Err(invalid("font name must not be empty"));
    }
    if value.len() > MAX_FONT_CHANGE_NAME_BYTES {
        return Err(crate::Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: litchi_core::Resource::Memory,
            observed: u64::try_from(value.len()).unwrap_or(u64::MAX),
            limit: u64::try_from(MAX_FONT_CHANGE_NAME_BYTES).unwrap_or(u64::MAX),
            scope: Arc::from("font name staged bytes"),
        }));
    }
    if let Some(character) = value
        .chars()
        .find(|character| !is_xml_10_character(*character))
    {
        return Err(invalid(format!(
            "font name contains an XML 1.0-forbidden character U+{:04X}",
            character as u32
        )));
    }
    Ok(Arc::<str>::from(value))
}

const fn is_xml_10_character(character: char) -> bool {
    matches!(character as u32, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

/// Opaque style identity used by semantic patch states.
///
/// The physical `SpreadsheetML` index remains private. Keys are meaningful only
/// within the patch lineage that produced them.
#[derive(Debug)]
pub(crate) struct StyleLineage;

#[derive(Clone)]
pub struct StyleKey {
    raw: u32,
    lineage: Arc<StyleLineage>,
}

impl StyleKey {
    pub(crate) fn new(raw: u32, lineage: Arc<StyleLineage>) -> Self {
        Self { raw, lineage }
    }

    pub(crate) fn belongs_to(&self, lineage: &Arc<StyleLineage>) -> bool {
        Arc::ptr_eq(&self.lineage, lineage)
    }

    pub(crate) const fn raw(&self) -> u32 {
        self.raw
    }
}

impl PartialEq for StyleKey {
    fn eq(&self, other: &Self) -> bool {
        self.raw == other.raw && Arc::ptr_eq(&self.lineage, &other.lineage)
    }
}

impl Eq for StyleKey {}

impl Hash for StyleKey {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.raw.hash(state);
        Arc::as_ptr(&self.lineage).hash(state);
    }
}

impl fmt::Debug for StyleKey {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("StyleKey(..)")
    }
}

/// Exact local style state for a stored cell.
///
/// `Default` means the cell has no local `s` attribute. It remains distinct
/// from an explicit reference to the workbook's base shared style.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum LocalStyle {
    Default,
    Shared(Style),
}

/// Exact local style identity recorded in a semantic patch state.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum StyleState {
    Default,
    Shared(StyleKey),
}

impl StyleState {
    pub(crate) fn rebind(&mut self, lineage: &Arc<StyleLineage>) {
        if let Self::Shared(key) = self {
            key.lineage = Arc::clone(lineage);
        }
    }
}

/// Lazy, immutable view of the workbook's shared cell formats.
#[derive(Clone)]
pub struct Styles {
    owner: Arc<Inner>,
    len: u32,
}

impl Styles {
    pub(crate) fn new(owner: Arc<Inner>, len: u32) -> Self {
        Self { owner, len }
    }

    /// Number of shared cell formats.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.len as usize
    }

    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.len == 0
    }

    /// Base shared cell format used by cells without an explicit local style.
    #[must_use]
    pub fn base(&self) -> Option<Style> {
        self.get(0)
    }

    /// Checked zero-based physical-position lookup for diagnostics and import.
    ///
    /// Copying a style obtained from [`crate::Worksheet::style`] is the preferred
    /// semantic entry point.
    #[must_use]
    pub fn get(&self, position: usize) -> Option<Style> {
        let key = u32::try_from(position).ok().filter(|key| *key < self.len)?;
        Some(Style {
            owner: Arc::clone(&self.owner),
            raw: key,
        })
    }

    /// Resolve an opaque key when it belongs to this shared-style table.
    #[must_use]
    pub fn find(&self, key: &StyleKey) -> Option<Style> {
        (key.raw < self.len && Arc::ptr_eq(&key.lineage, &self.owner.style_lineage)).then(|| {
            Style {
                owner: Arc::clone(&self.owner),
                raw: key.raw,
            }
        })
    }

    #[must_use]
    pub fn iter(&self) -> StylesIter {
        StylesIter {
            owner: Arc::clone(&self.owner),
            remaining: 0..self.len,
        }
    }
}

impl fmt::Debug for Styles {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Styles")
            .field("len", &self.len)
            .finish_non_exhaustive()
    }
}

impl IntoIterator for &Styles {
    type Item = Style;
    type IntoIter = StylesIter;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

/// Iterator over cheap shared-style handles.
#[derive(Debug, Clone)]
pub struct StylesIter {
    owner: Arc<Inner>,
    remaining: Range<u32>,
}

impl Iterator for StylesIter {
    type Item = Style;

    fn next(&mut self) -> Option<Self::Item> {
        self.remaining.next().map(|key| Style {
            owner: Arc::clone(&self.owner),
            raw: key,
        })
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        self.remaining.size_hint()
    }
}

impl DoubleEndedIterator for StylesIter {
    fn next_back(&mut self) -> Option<Self::Item> {
        self.remaining.next_back().map(|key| Style {
            owner: Arc::clone(&self.owner),
            raw: key,
        })
    }
}

impl ExactSizeIterator for StylesIter {}
impl std::iter::FusedIterator for StylesIter {}

/// Lineage-checked handle to one immutable shared cell format.
#[derive(Clone)]
pub struct Style {
    pub(crate) owner: Arc<Inner>,
    raw: u32,
}

impl Style {
    pub(crate) fn from_raw(owner: Arc<Inner>, key: u32) -> Self {
        Self { owner, raw: key }
    }

    pub(crate) const fn raw(&self) -> u32 {
        self.raw
    }

    /// Opaque, lineage-checked identity for maps and semantic patch states.
    #[must_use]
    pub fn key(&self) -> StyleKey {
        StyleKey::new(self.raw, Arc::clone(&self.owner.style_lineage))
    }

    /// Whether two handles name the same format in the same shared-style lineage.
    ///
    /// This remains true across descendant snapshots whose shared-style table
    /// is byte-for-byte unchanged. Use [`Self::same_workbook`] when exact
    /// snapshot identity matters.
    #[must_use]
    pub fn same(&self, other: &Self) -> bool {
        self.raw == other.raw && Arc::ptr_eq(&self.owner.style_lineage, &other.owner.style_lineage)
    }

    /// Whether two handles belong to the same immutable workbook snapshot.
    #[must_use]
    pub fn same_workbook(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.owner, &other.owner)
    }

    /// Count stored worksheet cells whose effective format is this resource.
    ///
    /// The base style includes stored cells with no explicit local style.
    /// Row/column defaults and unstored grid positions are intentionally not
    /// approximated.
    pub fn fan_out(&self) -> Result<usize> {
        self.owner.style_fan_out(self.raw)
    }

    /// Readable local formatting of this shared cell format: its number
    /// format, font, fill, border, alignment and protection.
    ///
    /// The workbook's `styles.xml` is parsed once per payload generation
    /// into a bounded, evictable cache shared by this workbook's descendant
    /// snapshots (see [`cell_format`]); later calls are table lookups.
    ///
    /// # Errors
    ///
    /// Returns an error when the styles part cannot be read or a property
    /// value lies outside its schema domain.
    pub fn format(&self) -> Result<Format> {
        self.owner.format(self.raw)
    }
}

impl PartialEq for Style {
    fn eq(&self, other: &Self) -> bool {
        self.same(other)
    }
}

impl Eq for Style {}

impl Hash for Style {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.raw.hash(state);
        Arc::as_ptr(&self.owner.style_lineage).hash(state);
    }
}

impl fmt::Debug for Style {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Style").finish_non_exhaustive()
    }
}
