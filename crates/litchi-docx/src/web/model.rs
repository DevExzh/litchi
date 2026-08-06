//! Typed WordprocessingML web-settings models and semantic mutation.

use super::codec::{
    div_position, parse_i64, validate_border_style, validate_divs, validate_encoding,
    validate_pixels_per_inch, validate_relationship_id, validate_text, validate_word_color, write,
};
use super::{
    STRICT_WEB_SETTINGS_RELATIONSHIP, STRICT_WORD_NAMESPACE, WORD_NAMESPACE, invalid, reserve_one,
};
use crate::color::Theme;
use crate::{Error, Result};
use std::num::NonZeroI64;

/// Scalar settings from a Word `webSettings.xml` part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Settings {
    pub(super) frameset: Option<Frameset>,
    pub(super) divs: Option<Vec<Div>>,
    pub(super) encoding: Option<String>,
    pub(super) optimize_for_browser: Option<bool>,
    pub(super) rely_on_vml: Option<bool>,
    pub(super) allow_png: Option<bool>,
    pub(super) do_not_rely_on_css: Option<bool>,
    pub(super) do_not_save_as_single_file: Option<bool>,
    pub(super) do_not_organize_in_folder: Option<bool>,
    pub(super) do_not_use_long_file_names: Option<bool>,
    pub(super) pixels_per_inch: Option<u16>,
    pub(super) target_screen_size: Option<Screen>,
    pub(super) save_smart_tags_as_xml: Option<bool>,
}

/// Namespace family used when serializing `word/webSettings.xml`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Conformance {
    /// ECMA-376 / ISO Transitional namespaces.
    #[default]
    Transitional,
    /// ISO/IEC 29500 Strict namespaces.
    Strict,
}

impl Conformance {
    /// Relationship type used by a document to own its web-settings part.
    pub const fn relationship(self) -> &'static str {
        match self {
            Self::Transitional => litchi_opc::constants::relationship_type::WEB_SETTINGS,
            Self::Strict => STRICT_WEB_SETTINGS_RELATIONSHIP,
        }
    }

    pub(super) const fn wordprocessingml(self) -> &'static str {
        match self {
            Self::Transitional => "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            Self::Strict => "http://purl.oclc.org/ooxml/wordprocessingml/main",
        }
    }

    pub(super) const fn relationships(self) -> &'static str {
        match self {
            Self::Transitional => {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            },
            Self::Strict => "http://purl.oclc.org/ooxml/officeDocument/relationships",
        }
    }

    pub(super) fn from_word_namespace(namespace: &[u8]) -> Option<Self> {
        match namespace {
            WORD_NAMESPACE => Some(Self::Transitional),
            STRICT_WORD_NAMESPACE => Some(Self::Strict),
            _ => None,
        }
    }

    pub(super) fn from_relationship(value: &str) -> Option<Self> {
        match value {
            litchi_opc::constants::relationship_type::WEB_SETTINGS => Some(Self::Transitional),
            STRICT_WEB_SETTINGS_RELATIONSHIP => Some(Self::Strict),
            _ => None,
        }
    }
}

/// Fidelity information for one HTML `div`, `body`, or `blockquote` element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Div {
    pub(super) id: Id,
    pub(super) block_quote: Option<bool>,
    pub(super) body_div: Option<bool>,
    pub(super) left: Twips,
    pub(super) right: Twips,
    pub(super) top: Twips,
    pub(super) bottom: Twips,
    pub(super) borders: Option<Borders>,
    pub(super) children: Vec<Div>,
}

/// A nonzero Word HTML-division identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct Id(pub(super) NonZeroI64);

/// A signed twip measure used by required HTML-division margins.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct Twips(pub(super) i64);

/// Stable semantic or positional selector for an HTML division.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Key {
    /// Select the unique division identifier.
    Id(Id),
    /// Select the raw document-order position.
    Index(usize),
}

impl From<Id> for Key {
    fn from(value: Id) -> Self {
        Self::Id(value)
    }
}

impl From<&Id> for Key {
    fn from(value: &Id) -> Self {
        Self::Id(*value)
    }
}

impl From<usize> for Key {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

impl Id {
    /// Create a checked nonzero identifier.
    pub fn new(value: i64) -> Result<Self> {
        NonZeroI64::new(value)
            .map(Self)
            .ok_or_else(|| invalid("Word HTML division ID must not be zero"))
    }

    /// Parse a checked decimal identifier.
    pub fn parse(value: &str) -> Result<Self> {
        let value = parse_i64(value, "HTML division ID")?;
        Self::new(value)
    }

    /// Return the numeric identifier.
    pub const fn get(self) -> i64 {
        self.0.get()
    }
}

impl std::fmt::Display for Id {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        self.get().fmt(formatter)
    }
}

impl TryFrom<i64> for Id {
    type Error = Error;

    fn try_from(value: i64) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<&str> for Id {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::parse(value)
    }
}

impl Twips {
    /// Create a signed twip measure.
    pub const fn new(value: i64) -> Self {
        Self(value)
    }

    /// Parse a signed decimal twip measure.
    pub fn parse(value: &str) -> Result<Self> {
        parse_i64(value, "HTML division margin").map(Self)
    }

    /// Return the signed twip count.
    pub const fn get(self) -> i64 {
        self.0
    }
}

impl From<i64> for Twips {
    fn from(value: i64) -> Self {
        Self::new(value)
    }
}

impl std::fmt::Display for Twips {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        self.get().fmt(formatter)
    }
}

/// Borders around an HTML division.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Borders {
    pub(super) top: Option<Border>,
    pub(super) left: Option<Border>,
    pub(super) bottom: Option<Border>,
    pub(super) right: Option<Border>,
}

/// One side of an HTML division border.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BorderSide {
    Top,
    Left,
    Bottom,
    Right,
}

impl BorderSide {
    /// Return the WordprocessingML element name for this side.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Left => "left",
            Self::Bottom => "bottom",
            Self::Right => "right",
        }
    }
}

/// One border around an HTML division.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Border {
    pub(super) style: String,
    pub(super) color: Option<String>,
    pub(super) theme_color: Option<Theme>,
    pub(super) theme_tint: Option<u8>,
    pub(super) theme_shade: Option<u8>,
    pub(super) size_eighth_points: Option<u64>,
    pub(super) space_points: Option<u64>,
    pub(super) shadow: Option<bool>,
    pub(super) frame: Option<bool>,
}

/// A recursive web frameset definition.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Frameset {
    pub(super) size: Option<String>,
    pub(super) split_bar: Option<SplitBar>,
    pub(super) layout: Option<Layout>,
    pub(super) children: Vec<Child>,
}

/// A child of a web frameset, retained in document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Child {
    /// A nested frameset.
    Frameset(Frameset),
    /// A leaf frame.
    Frame(Frame),
}

/// Properties for one frame in a web frameset.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Frame {
    pub(super) size: Option<String>,
    pub(super) name: Option<String>,
    pub(super) source_file_relationship_id: Option<String>,
    pub(super) margin_width: Option<u64>,
    pub(super) margin_height: Option<u64>,
    pub(super) scrollbar: Option<Scrollbar>,
    pub(super) no_resize_allowed: Option<bool>,
    pub(super) linked_to_file: Option<bool>,
}

/// Visual properties for the splitter bars of a frameset.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SplitBar {
    pub(super) width_twips: Option<u64>,
    pub(super) color: Option<Color>,
    pub(super) no_border: Option<bool>,
    pub(super) flat_borders: Option<bool>,
}

/// A frameset splitter color with optional theme modifiers.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Color {
    pub(super) value: String,
    pub(super) theme_color: Option<Theme>,
    pub(super) theme_tint: Option<u8>,
    pub(super) theme_shade: Option<u8>,
}

/// The direction in which a frameset stacks its children.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Layout {
    Rows,
    Columns,
    None,
}

/// The scrollbar policy for one frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Scrollbar {
    On,
    Off,
    Auto,
}

/// A spec-defined target size for generated web pages.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Screen {
    Pixels544x376,
    Pixels640x480,
    Pixels720x512,
    Pixels800x600,
    Pixels1024x768,
    Pixels1152x882,
    Pixels1152x900,
    Pixels1280x1024,
    Pixels1600x1200,
    Pixels1800x1440,
    Pixels1920x1200,
}

impl Screen {
    pub(super) fn from_xml(value: &str) -> Option<Self> {
        match value {
            "544x376" => Some(Self::Pixels544x376),
            "640x480" => Some(Self::Pixels640x480),
            "720x512" => Some(Self::Pixels720x512),
            "800x600" => Some(Self::Pixels800x600),
            "1024x768" => Some(Self::Pixels1024x768),
            "1152x882" => Some(Self::Pixels1152x882),
            "1152x900" => Some(Self::Pixels1152x900),
            "1280x1024" => Some(Self::Pixels1280x1024),
            "1600x1200" => Some(Self::Pixels1600x1200),
            "1800x1440" => Some(Self::Pixels1800x1440),
            "1920x1200" => Some(Self::Pixels1920x1200),
            _ => None,
        }
    }

    /// Return the OOXML lexical representation.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Pixels544x376 => "544x376",
            Self::Pixels640x480 => "640x480",
            Self::Pixels720x512 => "720x512",
            Self::Pixels800x600 => "800x600",
            Self::Pixels1024x768 => "1024x768",
            Self::Pixels1152x882 => "1152x882",
            Self::Pixels1152x900 => "1152x900",
            Self::Pixels1280x1024 => "1280x1024",
            Self::Pixels1600x1200 => "1600x1200",
            Self::Pixels1800x1440 => "1800x1440",
            Self::Pixels1920x1200 => "1920x1200",
        }
    }
}

impl Layout {
    pub(super) fn from_xml(value: &str) -> Option<Self> {
        match value {
            "rows" => Some(Self::Rows),
            "cols" => Some(Self::Columns),
            "none" => Some(Self::None),
            _ => None,
        }
    }

    /// Return the OOXML lexical representation.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Rows => "rows",
            Self::Columns => "cols",
            Self::None => "none",
        }
    }
}

impl Scrollbar {
    pub(super) fn from_xml(value: &str) -> Option<Self> {
        match value {
            "on" => Some(Self::On),
            "off" => Some(Self::Off),
            "auto" => Some(Self::Auto),
            _ => None,
        }
    }

    /// Return the OOXML lexical representation.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::On => "on",
            Self::Off => "off",
            Self::Auto => "auto",
        }
    }
}

impl Frameset {
    /// Return the size expression for this frameset, if explicitly present.
    pub fn size(&self) -> Option<&str> {
        self.size.as_deref()
    }

    /// Return splitter-bar properties, if explicitly present.
    pub fn split_bar(&self) -> Option<&SplitBar> {
        self.split_bar.as_ref()
    }

    /// Return the explicit child layout. Absence has the schema-defined row default.
    pub fn layout(&self) -> Option<Layout> {
        self.layout
    }

    /// Return nested frames and framesets in document order.
    pub fn children(&self) -> &[Child] {
        &self.children
    }

    /// Set the frameset size expression.
    pub fn set_size(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        let value = value.into();
        validate_text(&value, "frameset size", false)?;
        self.size = Some(value);
        Ok(self)
    }

    /// Remove the explicit frameset size expression.
    pub fn clear_size(&mut self) -> &mut Self {
        self.size = None;
        self
    }

    /// Set splitter-bar properties.
    pub fn set_split_bar(&mut self, value: SplitBar) -> &mut Self {
        self.split_bar = Some(value);
        self
    }

    /// Return mutable splitter-bar properties, if present.
    pub fn split_bar_mut(&mut self) -> Option<&mut SplitBar> {
        self.split_bar.as_mut()
    }

    /// Remove explicit splitter-bar properties.
    pub fn clear_split_bar(&mut self) -> &mut Self {
        self.split_bar = None;
        self
    }

    /// Set the child layout.
    pub fn set_layout(&mut self, value: Layout) -> &mut Self {
        self.layout = Some(value);
        self
    }

    /// Restore the schema-defined row layout.
    pub fn clear_layout(&mut self) -> &mut Self {
        self.layout = None;
        self
    }

    /// Append an empty frame and return it for configuration.
    pub fn add_frame(&mut self) -> Result<&mut Frame> {
        reserve_one(&mut self.children, "frameset child insertion")?;
        self.children.push(Child::Frame(Frame::default()));
        match self.children.last_mut() {
            Some(Child::Frame(frame)) => Ok(frame),
            _ => Err(invalid("new frame was not retained")),
        }
    }

    /// Append a configured frame.
    pub fn push_frame(&mut self, frame: Frame) -> Result<&mut Self> {
        reserve_one(&mut self.children, "frameset child insertion")?;
        self.children.push(Child::Frame(frame));
        Ok(self)
    }

    /// Append an empty nested frameset and return it for configuration.
    pub fn add_frameset(&mut self) -> Result<&mut Frameset> {
        reserve_one(&mut self.children, "frameset child insertion")?;
        self.children.push(Child::Frameset(Frameset::default()));
        match self.children.last_mut() {
            Some(Child::Frameset(frameset)) => Ok(frameset),
            _ => Err(invalid("new frameset was not retained")),
        }
    }

    /// Append a configured nested frameset.
    pub fn push_frameset(&mut self, frameset: Frameset) -> Result<&mut Self> {
        reserve_one(&mut self.children, "frameset child insertion")?;
        self.children.push(Child::Frameset(frameset));
        Ok(self)
    }

    /// Remove all nested frames and framesets.
    pub fn clear_children(&mut self) -> &mut Self {
        self.children.clear();
        self
    }
}

impl Frame {
    pub fn size(&self) -> Option<&str> {
        self.size.as_deref()
    }

    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Return the focused low-level frame relationship ID, if present.
    ///
    /// Prefer package-level [`load`] and [`put`], which validate this identifier
    /// against the part relationship graph.
    pub fn rel(&self) -> Option<&str> {
        self.source_file_relationship_id.as_deref()
    }

    pub fn margin_width(&self) -> Option<u64> {
        self.margin_width
    }

    pub fn margin_height(&self) -> Option<u64> {
        self.margin_height
    }

    pub fn scrollbar(&self) -> Option<Scrollbar> {
        self.scrollbar
    }

    pub fn no_resize_allowed(&self) -> Option<bool> {
        self.no_resize_allowed
    }

    pub fn linked_to_file(&self) -> Option<bool> {
        self.linked_to_file
    }

    pub fn set_size(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        let value = value.into();
        validate_text(&value, "frame size", false)?;
        self.size = Some(value);
        Ok(self)
    }

    pub fn clear_size(&mut self) -> &mut Self {
        self.size = None;
        self
    }

    pub fn set_name(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        let value = value.into();
        validate_text(&value, "frame name", false)?;
        self.name = Some(value);
        Ok(self)
    }

    pub fn clear_name(&mut self) -> &mut Self {
        self.name = None;
        self
    }

    /// Set a checked low-level frame relationship ID.
    pub fn set_rel(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        let value = value.into();
        validate_relationship_id(&value)?;
        self.source_file_relationship_id = Some(value);
        Ok(self)
    }

    /// Remove the low-level frame relationship ID.
    pub fn clear_rel(&mut self) -> &mut Self {
        self.source_file_relationship_id = None;
        self
    }

    pub fn set_margin_width(&mut self, value: u64) -> &mut Self {
        self.margin_width = Some(value);
        self
    }

    pub fn clear_margin_width(&mut self) -> &mut Self {
        self.margin_width = None;
        self
    }

    pub fn set_margin_height(&mut self, value: u64) -> &mut Self {
        self.margin_height = Some(value);
        self
    }

    pub fn clear_margin_height(&mut self) -> &mut Self {
        self.margin_height = None;
        self
    }

    pub fn set_scrollbar(&mut self, value: Scrollbar) -> &mut Self {
        self.scrollbar = Some(value);
        self
    }

    pub fn clear_scrollbar(&mut self) -> &mut Self {
        self.scrollbar = None;
        self
    }

    pub fn set_no_resize_allowed(&mut self, value: bool) -> &mut Self {
        self.no_resize_allowed = Some(value);
        self
    }

    pub fn clear_no_resize_allowed(&mut self) -> &mut Self {
        self.no_resize_allowed = None;
        self
    }

    pub fn set_linked_to_file(&mut self, value: bool) -> &mut Self {
        self.linked_to_file = Some(value);
        self
    }

    pub fn clear_linked_to_file(&mut self) -> &mut Self {
        self.linked_to_file = None;
        self
    }
}

impl SplitBar {
    pub fn width_twips(&self) -> Option<u64> {
        self.width_twips
    }

    pub fn color(&self) -> Option<&Color> {
        self.color.as_ref()
    }

    pub fn no_border(&self) -> Option<bool> {
        self.no_border
    }

    pub fn flat_borders(&self) -> Option<bool> {
        self.flat_borders
    }

    pub fn set_width_twips(&mut self, value: u64) -> &mut Self {
        self.width_twips = Some(value);
        self
    }

    pub fn clear_width_twips(&mut self) -> &mut Self {
        self.width_twips = None;
        self
    }

    pub fn set_color(&mut self, value: Color) -> &mut Self {
        self.color = Some(value);
        self
    }

    pub fn color_mut(&mut self) -> Option<&mut Color> {
        self.color.as_mut()
    }

    pub fn clear_color(&mut self) -> &mut Self {
        self.color = None;
        self
    }

    pub fn set_no_border(&mut self, value: bool) -> &mut Self {
        self.no_border = Some(value);
        self
    }

    pub fn clear_no_border(&mut self) -> &mut Self {
        self.no_border = None;
        self
    }

    pub fn set_flat_borders(&mut self, value: bool) -> &mut Self {
        self.flat_borders = Some(value);
        self
    }

    pub fn clear_flat_borders(&mut self) -> &mut Self {
        self.flat_borders = None;
        self
    }
}

impl Color {
    /// Create a validated automatic or six-digit RGB splitter color.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        Ok(Self {
            value: validate_word_color(value.into(), "frameset splitter color")?,
            theme_color: None,
            theme_tint: None,
            theme_shade: None,
        })
    }

    pub fn value(&self) -> &str {
        &self.value
    }

    pub fn theme_color(&self) -> Option<Theme> {
        self.theme_color
    }

    pub fn theme_tint(&self) -> Option<u8> {
        self.theme_tint
    }

    pub fn theme_shade(&self) -> Option<u8> {
        self.theme_shade
    }

    pub fn set_value(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        self.value = validate_word_color(value.into(), "frameset splitter color")?;
        Ok(self)
    }

    pub fn set_theme_color(&mut self, value: Theme) -> &mut Self {
        self.theme_color = Some(value);
        self
    }

    pub fn clear_theme_color(&mut self) -> &mut Self {
        self.theme_color = None;
        self
    }

    pub fn set_theme_tint(&mut self, value: u8) -> &mut Self {
        self.theme_tint = Some(value);
        self
    }

    pub fn clear_theme_tint(&mut self) -> &mut Self {
        self.theme_tint = None;
        self
    }

    pub fn set_theme_shade(&mut self, value: u8) -> &mut Self {
        self.theme_shade = Some(value);
        self
    }

    pub fn clear_theme_shade(&mut self) -> &mut Self {
        self.theme_shade = None;
        self
    }
}

impl Div {
    /// Create a schema-valid HTML division with zero margins.
    pub fn new(id: Id) -> Self {
        Self {
            id,
            block_quote: None,
            body_div: None,
            left: Twips::default(),
            right: Twips::default(),
            top: Twips::default(),
            bottom: Twips::default(),
            borders: None,
            children: Vec::new(),
        }
    }

    pub const fn id(&self) -> Id {
        self.id
    }

    pub fn is_block_quote(&self) -> Option<bool> {
        self.block_quote
    }

    pub fn is_body_div(&self) -> Option<bool> {
        self.body_div
    }

    /// Return the required left margin.
    pub const fn left(&self) -> Twips {
        self.left
    }

    /// Return the required right margin.
    pub const fn right(&self) -> Twips {
        self.right
    }

    /// Return the required top margin.
    pub const fn top(&self) -> Twips {
        self.top
    }

    /// Return the required bottom margin.
    pub const fn bottom(&self) -> Twips {
        self.bottom
    }

    pub fn borders(&self) -> Option<&Borders> {
        self.borders.as_ref()
    }

    pub fn children(&self) -> &[Div] {
        &self.children
    }

    pub fn set_id(&mut self, value: Id) -> &mut Self {
        self.id = value;
        self
    }

    pub fn set_block_quote(&mut self, value: bool) -> &mut Self {
        self.block_quote = Some(value);
        self
    }

    pub fn clear_block_quote(&mut self) -> &mut Self {
        self.block_quote = None;
        self
    }

    pub fn set_body_div(&mut self, value: bool) -> &mut Self {
        self.body_div = Some(value);
        self
    }

    pub fn clear_body_div(&mut self) -> &mut Self {
        self.body_div = None;
        self
    }

    /// Set the required left margin.
    pub fn set_left(&mut self, value: impl Into<Twips>) -> &mut Self {
        self.left = value.into();
        self
    }

    /// Set the required right margin.
    pub fn set_right(&mut self, value: impl Into<Twips>) -> &mut Self {
        self.right = value.into();
        self
    }

    /// Set the required top margin.
    pub fn set_top(&mut self, value: impl Into<Twips>) -> &mut Self {
        self.top = value.into();
        self
    }

    /// Set the required bottom margin.
    pub fn set_bottom(&mut self, value: impl Into<Twips>) -> &mut Self {
        self.bottom = value.into();
        self
    }

    pub fn set_borders(&mut self, value: Borders) -> &mut Self {
        self.borders = Some(value);
        self
    }

    pub fn borders_mut(&mut self) -> Option<&mut Borders> {
        self.borders.as_mut()
    }

    pub fn clear_borders(&mut self) -> &mut Self {
        self.borders = None;
        self
    }

    /// Select a child by its stable identifier or raw position.
    pub fn child(&self, key: impl Into<Key>) -> Result<Option<&Div>> {
        Ok(div_position(&self.children, key.into())?.and_then(|index| self.children.get(index)))
    }

    /// Append a uniquely identified child division.
    pub fn add_child(&mut self, child: Div) -> Result<&mut Self> {
        if div_position(&self.children, Key::Id(child.id()))?.is_some() {
            return Err(invalid(format!(
                "HTML division '{}' already exists",
                child.id()
            )));
        }
        reserve_one(&mut self.children, "HTML child division insertion")?;
        self.children.push(child);
        Ok(self)
    }

    /// Insert or replace a child by its identifier, retaining its position.
    pub fn put_child(&mut self, child: Div) -> Result<Option<Div>> {
        match div_position(&self.children, Key::Id(child.id()))? {
            Some(index) => {
                let slot = self
                    .children
                    .get_mut(index)
                    .ok_or_else(|| invalid("HTML child selector changed during replacement"))?;
                Ok(Some(std::mem::replace(slot, child)))
            },
            None => {
                reserve_one(&mut self.children, "HTML child division insertion")?;
                self.children.push(child);
                Ok(None)
            },
        }
    }

    /// Remove a child selected semantically or positionally.
    pub fn remove_child(&mut self, key: impl Into<Key>) -> Result<Option<Div>> {
        Ok(div_position(&self.children, key.into())?.map(|index| self.children.remove(index)))
    }

    /// Move a selected child to a checked final position.
    pub fn move_child(&mut self, key: impl Into<Key>, to: usize) -> Result<&mut Self> {
        let Some(from) = div_position(&self.children, key.into())? else {
            return Err(invalid("HTML child selector does not exist"));
        };
        if to >= self.children.len() {
            return Err(invalid(format!(
                "HTML child destination {to} is outside 0..{}",
                self.children.len()
            )));
        }
        if from != to {
            let child = self.children.remove(from);
            self.children.insert(to, child);
        }
        Ok(self)
    }

    pub fn clear_children(&mut self) -> &mut Self {
        self.children.clear();
        self
    }
}

impl Borders {
    pub fn top(&self) -> Option<&Border> {
        self.top.as_ref()
    }

    pub fn left(&self) -> Option<&Border> {
        self.left.as_ref()
    }

    pub fn bottom(&self) -> Option<&Border> {
        self.bottom.as_ref()
    }

    pub fn right(&self) -> Option<&Border> {
        self.right.as_ref()
    }

    pub fn set_top(&mut self, value: Border) -> &mut Self {
        self.top = Some(value);
        self
    }

    pub fn top_mut(&mut self) -> Option<&mut Border> {
        self.top.as_mut()
    }

    pub fn clear_top(&mut self) -> &mut Self {
        self.top = None;
        self
    }

    pub fn set_left(&mut self, value: Border) -> &mut Self {
        self.left = Some(value);
        self
    }

    pub fn left_mut(&mut self) -> Option<&mut Border> {
        self.left.as_mut()
    }

    pub fn clear_left(&mut self) -> &mut Self {
        self.left = None;
        self
    }

    pub fn set_bottom(&mut self, value: Border) -> &mut Self {
        self.bottom = Some(value);
        self
    }

    pub fn bottom_mut(&mut self) -> Option<&mut Border> {
        self.bottom.as_mut()
    }

    pub fn clear_bottom(&mut self) -> &mut Self {
        self.bottom = None;
        self
    }

    pub fn set_right(&mut self, value: Border) -> &mut Self {
        self.right = Some(value);
        self
    }

    pub fn right_mut(&mut self) -> Option<&mut Border> {
        self.right.as_mut()
    }

    pub fn clear_right(&mut self) -> &mut Self {
        self.right = None;
        self
    }
}

impl Border {
    /// Create a division border with its required style.
    pub fn new(style: impl Into<String>) -> Result<Self> {
        let style = style.into();
        validate_border_style(&style)?;
        Ok(Self {
            style,
            color: None,
            theme_color: None,
            theme_tint: None,
            theme_shade: None,
            size_eighth_points: None,
            space_points: None,
            shadow: None,
            frame: None,
        })
    }

    pub fn style(&self) -> &str {
        &self.style
    }

    pub fn color(&self) -> Option<&str> {
        self.color.as_deref()
    }

    pub fn theme_color(&self) -> Option<Theme> {
        self.theme_color
    }

    pub fn theme_tint(&self) -> Option<u8> {
        self.theme_tint
    }

    pub fn theme_shade(&self) -> Option<u8> {
        self.theme_shade
    }

    pub fn size_eighth_points(&self) -> Option<u64> {
        self.size_eighth_points
    }

    pub fn space_points(&self) -> Option<u64> {
        self.space_points
    }

    pub fn shadow(&self) -> Option<bool> {
        self.shadow
    }

    pub fn frame(&self) -> Option<bool> {
        self.frame
    }

    pub fn set_style(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        let value = value.into();
        validate_border_style(&value)?;
        self.style = value;
        Ok(self)
    }

    pub fn set_color(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        self.color = Some(validate_word_color(
            value.into(),
            "HTML division border color",
        )?);
        Ok(self)
    }

    pub fn clear_color(&mut self) -> &mut Self {
        self.color = None;
        self
    }

    pub fn set_theme_color(&mut self, value: Theme) -> &mut Self {
        self.theme_color = Some(value);
        self
    }

    pub fn clear_theme_color(&mut self) -> &mut Self {
        self.theme_color = None;
        self
    }

    pub fn set_theme_tint(&mut self, value: u8) -> &mut Self {
        self.theme_tint = Some(value);
        self
    }

    pub fn clear_theme_tint(&mut self) -> &mut Self {
        self.theme_tint = None;
        self
    }

    pub fn set_theme_shade(&mut self, value: u8) -> &mut Self {
        self.theme_shade = Some(value);
        self
    }

    pub fn clear_theme_shade(&mut self) -> &mut Self {
        self.theme_shade = None;
        self
    }

    pub fn set_size_eighth_points(&mut self, value: u64) -> &mut Self {
        self.size_eighth_points = Some(value);
        self
    }

    pub fn clear_size_eighth_points(&mut self) -> &mut Self {
        self.size_eighth_points = None;
        self
    }

    pub fn set_space_points(&mut self, value: u64) -> &mut Self {
        self.space_points = Some(value);
        self
    }

    pub fn clear_space_points(&mut self) -> &mut Self {
        self.space_points = None;
        self
    }

    pub fn set_shadow(&mut self, value: bool) -> &mut Self {
        self.shadow = Some(value);
        self
    }

    pub fn clear_shadow(&mut self) -> &mut Self {
        self.shadow = None;
        self
    }

    pub fn set_frame(&mut self, value: bool) -> &mut Self {
        self.frame = Some(value);
        self
    }

    pub fn clear_frame(&mut self) -> &mut Self {
        self.frame = None;
        self
    }
}

impl Settings {
    /// Return the root frameset definition, if present.
    pub fn frameset(&self) -> Option<&Frameset> {
        self.frameset.as_ref()
    }

    /// Set the root web frameset definition.
    pub fn set_frameset(&mut self, value: Frameset) -> &mut Self {
        self.frameset = Some(value);
        self
    }

    /// Return the mutable root web frameset definition, if present.
    pub fn frameset_mut(&mut self) -> Option<&mut Frameset> {
        self.frameset.as_mut()
    }

    /// Remove the root web frameset definition.
    pub fn clear_frameset(&mut self) -> &mut Self {
        self.frameset = None;
        self
    }

    /// Return the top-level HTML division definitions, preserving part absence.
    pub fn divs(&self) -> Option<&[Div]> {
        self.divs.as_deref()
    }

    /// Replace the top-level HTML divisions after validating unique identifiers.
    pub fn set_divs(&mut self, value: Vec<Div>) -> Result<&mut Self> {
        if value.is_empty() {
            return Err(invalid("Word HTML division container must not be empty"));
        }
        validate_divs(&value, 1)?;
        self.divs = Some(value);
        Ok(self)
    }

    /// Select a top-level division by stable identifier or raw position.
    pub fn get(&self, key: impl Into<Key>) -> Result<Option<&Div>> {
        let key = key.into();
        let Some(divs) = &self.divs else {
            let _ = div_position(&[], key)?;
            return Ok(None);
        };
        Ok(div_position(divs, key)?.and_then(|index| divs.get(index)))
    }

    /// Append a uniquely identified top-level division.
    pub fn add(&mut self, div: Div) -> Result<&mut Self> {
        match &mut self.divs {
            Some(divs) => {
                if div_position(divs, Key::Id(div.id()))?.is_some() {
                    return Err(invalid(format!(
                        "HTML division '{}' already exists",
                        div.id()
                    )));
                }
                reserve_one(divs, "HTML division insertion")?;
                divs.push(div);
            },
            None => {
                let mut divs = Vec::new();
                reserve_one(&mut divs, "HTML division insertion")?;
                divs.push(div);
                self.divs = Some(divs);
            },
        }
        Ok(self)
    }

    /// Insert or replace a top-level division by identifier.
    pub fn put(&mut self, div: Div) -> Result<Option<Div>> {
        match &mut self.divs {
            Some(divs) => match div_position(divs, Key::Id(div.id()))? {
                Some(index) => {
                    let slot = divs.get_mut(index).ok_or_else(|| {
                        invalid("HTML division selector changed during replacement")
                    })?;
                    Ok(Some(std::mem::replace(slot, div)))
                },
                None => {
                    reserve_one(divs, "HTML division insertion")?;
                    divs.push(div);
                    Ok(None)
                },
            },
            None => {
                let mut divs = Vec::new();
                reserve_one(&mut divs, "HTML division insertion")?;
                divs.push(div);
                self.divs = Some(divs);
                Ok(None)
            },
        }
    }

    /// Remove a top-level division selected semantically or positionally.
    pub fn remove(&mut self, key: impl Into<Key>) -> Result<Option<Div>> {
        let key = key.into();
        let (removed, empty) = {
            let Some(divs) = &mut self.divs else {
                let _ = div_position(&[], key)?;
                return Ok(None);
            };
            let Some(index) = div_position(divs, key)? else {
                return Ok(None);
            };
            (divs.remove(index), divs.is_empty())
        };
        if empty {
            self.divs = None;
        }
        Ok(Some(removed))
    }

    /// Move a selected division to a checked final position.
    pub fn move_to(&mut self, key: impl Into<Key>, to: usize) -> Result<&mut Self> {
        let Some(divs) = &mut self.divs else {
            return Err(invalid("HTML division container does not exist"));
        };
        let Some(from) = div_position(divs, key.into())? else {
            return Err(invalid("HTML division selector does not exist"));
        };
        if to >= divs.len() {
            return Err(invalid(format!(
                "HTML division destination {to} is outside 0..{}",
                divs.len()
            )));
        }
        if from != to {
            let div = divs.remove(from);
            divs.insert(to, div);
        }
        Ok(self)
    }

    /// Remove the complete top-level HTML division container.
    pub fn clear_divs(&mut self) -> &mut Self {
        self.divs = None;
        self
    }

    /// Return the requested output encoding, if declared.
    pub fn encoding(&self) -> Option<&str> {
        self.encoding.as_deref()
    }

    /// Return the browser-optimization setting, preserving absence.
    pub fn optimize_for_browser(&self) -> Option<bool> {
        self.optimize_for_browser
    }

    /// Return whether web output should rely on VML, preserving absence.
    pub fn rely_on_vml(&self) -> Option<bool> {
        self.rely_on_vml
    }

    /// Return whether PNG images are allowed, preserving absence.
    pub fn allow_png(&self) -> Option<bool> {
        self.allow_png
    }

    /// Return whether web output should avoid CSS, preserving absence.
    pub fn do_not_rely_on_css(&self) -> Option<bool> {
        self.do_not_rely_on_css
    }

    /// Return whether web output should use multiple files, preserving absence.
    pub fn do_not_save_as_single_file(&self) -> Option<bool> {
        self.do_not_save_as_single_file
    }

    /// Return whether supporting files should avoid a folder, preserving absence.
    pub fn do_not_organize_in_folder(&self) -> Option<bool> {
        self.do_not_organize_in_folder
    }

    /// Return whether web output should avoid long file names, preserving absence.
    pub fn do_not_use_long_file_names(&self) -> Option<bool> {
        self.do_not_use_long_file_names
    }

    /// Return Word's bounded web-output pixel density.
    pub fn pixels_per_inch(&self) -> Option<u16> {
        self.pixels_per_inch
    }

    /// Return the target screen size, if declared.
    pub fn target_screen_size(&self) -> Option<Screen> {
        self.target_screen_size
    }

    /// Return whether smart tags should be saved as XML, preserving absence.
    pub fn save_smart_tags_as_xml(&self) -> Option<bool> {
        self.save_smart_tags_as_xml
    }

    /// Set the requested output encoding.
    pub fn set_encoding(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        let value = value.into();
        validate_encoding(&value)?;
        self.encoding = Some(value);
        Ok(self)
    }

    /// Remove the requested output encoding.
    pub fn clear_encoding(&mut self) -> &mut Self {
        self.encoding = None;
        self
    }

    /// Set whether output should be optimized for browsers.
    pub fn set_optimize_for_browser(&mut self, value: bool) -> &mut Self {
        self.optimize_for_browser = Some(value);
        self
    }

    /// Restore schema-defined behavior for browser optimization.
    pub fn clear_optimize_for_browser(&mut self) -> &mut Self {
        self.optimize_for_browser = None;
        self
    }

    /// Set whether web output should rely on VML.
    pub fn set_rely_on_vml(&mut self, value: bool) -> &mut Self {
        self.rely_on_vml = Some(value);
        self
    }

    /// Restore schema-defined behavior for VML use.
    pub fn clear_rely_on_vml(&mut self) -> &mut Self {
        self.rely_on_vml = None;
        self
    }

    /// Set whether PNG images are allowed.
    pub fn set_allow_png(&mut self, value: bool) -> &mut Self {
        self.allow_png = Some(value);
        self
    }

    /// Restore schema-defined behavior for PNG images.
    pub fn clear_allow_png(&mut self) -> &mut Self {
        self.allow_png = None;
        self
    }

    /// Set whether web output should avoid CSS.
    pub fn set_do_not_rely_on_css(&mut self, value: bool) -> &mut Self {
        self.do_not_rely_on_css = Some(value);
        self
    }

    /// Restore schema-defined behavior for CSS output.
    pub fn clear_do_not_rely_on_css(&mut self) -> &mut Self {
        self.do_not_rely_on_css = None;
        self
    }

    /// Set whether web output should avoid a single-file representation.
    pub fn set_do_not_save_as_single_file(&mut self, value: bool) -> &mut Self {
        self.do_not_save_as_single_file = Some(value);
        self
    }

    /// Restore schema-defined behavior for single-file output.
    pub fn clear_do_not_save_as_single_file(&mut self) -> &mut Self {
        self.do_not_save_as_single_file = None;
        self
    }

    /// Set whether supporting files should avoid a dedicated folder.
    pub fn set_do_not_organize_in_folder(&mut self, value: bool) -> &mut Self {
        self.do_not_organize_in_folder = Some(value);
        self
    }

    /// Restore schema-defined behavior for supporting-file folders.
    pub fn clear_do_not_organize_in_folder(&mut self) -> &mut Self {
        self.do_not_organize_in_folder = None;
        self
    }

    /// Set whether web output should avoid long file names.
    pub fn set_do_not_use_long_file_names(&mut self, value: bool) -> &mut Self {
        self.do_not_use_long_file_names = Some(value);
        self
    }

    /// Restore schema-defined behavior for long file names.
    pub fn clear_do_not_use_long_file_names(&mut self) -> &mut Self {
        self.do_not_use_long_file_names = None;
        self
    }

    /// Set Word's web-output pixel density in the inclusive range `0..=1023`.
    pub fn set_pixels_per_inch(&mut self, value: u16) -> Result<&mut Self> {
        validate_pixels_per_inch(value)?;
        self.pixels_per_inch = Some(value);
        Ok(self)
    }

    /// Remove the explicit web-output pixel density.
    pub fn clear_pixels_per_inch(&mut self) -> &mut Self {
        self.pixels_per_inch = None;
        self
    }

    /// Set the target display size for generated web pages.
    pub fn set_target_screen_size(&mut self, value: Screen) -> &mut Self {
        self.target_screen_size = Some(value);
        self
    }

    /// Remove the explicit target display size.
    pub fn clear_target_screen_size(&mut self) -> &mut Self {
        self.target_screen_size = None;
        self
    }

    /// Set whether smart tags should be retained in generated XML.
    pub fn set_save_smart_tags_as_xml(&mut self, value: bool) -> &mut Self {
        self.save_smart_tags_as_xml = Some(value);
        self
    }

    /// Restore schema-defined smart-tag serialization behavior.
    pub fn clear_save_smart_tags_as_xml(&mut self) -> &mut Self {
        self.save_smart_tags_as_xml = None;
        self
    }

    /// Serialize deterministically using the selected namespace family.
    pub fn xml(&self, conformance: Conformance) -> Result<Vec<u8>> {
        write(self, conformance)
    }
}
