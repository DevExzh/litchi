//! Typed WordprocessingML web-settings semantics and OPC ownership.

use crate::color::Theme;
use crate::{Error, Result};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::escape::escape;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::fmt::Write as _;
use std::num::NonZeroI64;

const WORD_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml";
const STRICT_OFFICE_DOCUMENT_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
const STRICT_WEB_SETTINGS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/webSettings";
const MAX_XML_BYTES: usize = 8 * 1024 * 1024;
const MAX_XML_EVENTS: usize = 262_144;
const MAX_TEXT_BYTES: usize = 64 * 1024;

#[derive(Default)]
struct ParseBudget {
    events: usize,
}

impl ParseBudget {
    fn event(&mut self) -> Result<()> {
        self.events = self
            .events
            .checked_add(1)
            .ok_or_else(|| invalid("web-settings event count overflow"))?;
        if self.events > MAX_XML_EVENTS {
            return Err(invalid(format!(
                "web-settings XML exceeds {MAX_XML_EVENTS} events"
            )));
        }
        Ok(())
    }
}

fn is_web_settings_relationship(value: &str) -> bool {
    value == litchi_opc::constants::relationship_type::WEB_SETTINGS
        || value == STRICT_WEB_SETTINGS_RELATIONSHIP
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn reserve_one<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })
}

fn is_wordprocessing_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == WORD_NAMESPACE || *value == STRICT_WORD_NAMESPACE
    )
}

fn word_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_word_attribute = is_wordprocessing_namespace(&namespace)
            || matches!(namespace, ResolveResult::Unbound)
            || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"w");
        if !is_word_attribute {
            continue;
        }
        if value.is_some() {
            let name = std::str::from_utf8(name)
                .map_err(|_| invalid("Word attribute name is not UTF-8"))?;
            return Err(invalid(format!("duplicate Word attribute '{name}'")));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

/// Scalar settings from a Word `webSettings.xml` part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Settings {
    frameset: Option<Frameset>,
    divs: Option<Vec<Div>>,
    encoding: Option<String>,
    optimize_for_browser: Option<bool>,
    rely_on_vml: Option<bool>,
    allow_png: Option<bool>,
    do_not_rely_on_css: Option<bool>,
    do_not_save_as_single_file: Option<bool>,
    do_not_organize_in_folder: Option<bool>,
    do_not_use_long_file_names: Option<bool>,
    pixels_per_inch: Option<u16>,
    target_screen_size: Option<Screen>,
    save_smart_tags_as_xml: Option<bool>,
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

    const fn wordprocessingml(self) -> &'static str {
        match self {
            Self::Transitional => "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            Self::Strict => "http://purl.oclc.org/ooxml/wordprocessingml/main",
        }
    }

    const fn relationships(self) -> &'static str {
        match self {
            Self::Transitional => {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            },
            Self::Strict => "http://purl.oclc.org/ooxml/officeDocument/relationships",
        }
    }

    fn from_word_namespace(namespace: &[u8]) -> Option<Self> {
        match namespace {
            WORD_NAMESPACE => Some(Self::Transitional),
            STRICT_WORD_NAMESPACE => Some(Self::Strict),
            _ => None,
        }
    }

    fn from_relationship(value: &str) -> Option<Self> {
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
    id: Id,
    block_quote: Option<bool>,
    body_div: Option<bool>,
    left: Twips,
    right: Twips,
    top: Twips,
    bottom: Twips,
    borders: Option<Borders>,
    children: Vec<Div>,
}

/// A nonzero Word HTML-division identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct Id(NonZeroI64);

/// A signed twip measure used by required HTML-division margins.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct Twips(i64);

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
    top: Option<Border>,
    left: Option<Border>,
    bottom: Option<Border>,
    right: Option<Border>,
}

/// One border around an HTML division.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Border {
    style: String,
    color: Option<String>,
    theme_color: Option<Theme>,
    theme_tint: Option<u8>,
    theme_shade: Option<u8>,
    size_eighth_points: Option<u64>,
    space_points: Option<u64>,
    shadow: Option<bool>,
    frame: Option<bool>,
}

/// A recursive web frameset definition.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Frameset {
    size: Option<String>,
    split_bar: Option<SplitBar>,
    layout: Option<Layout>,
    children: Vec<Child>,
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
    size: Option<String>,
    name: Option<String>,
    source_file_relationship_id: Option<String>,
    margin_width: Option<u64>,
    margin_height: Option<u64>,
    scrollbar: Option<Scrollbar>,
    no_resize_allowed: Option<bool>,
    linked_to_file: Option<bool>,
}

/// Visual properties for the splitter bars of a frameset.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SplitBar {
    width_twips: Option<u64>,
    color: Option<Color>,
    no_border: Option<bool>,
    flat_borders: Option<bool>,
}

/// A frameset splitter color with optional theme modifiers.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Color {
    value: String,
    theme_color: Option<Theme>,
    theme_tint: Option<u8>,
    theme_shade: Option<u8>,
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
    fn from_xml(value: &str) -> Option<Self> {
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
    fn from_xml(value: &str) -> Option<Self> {
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
    fn from_xml(value: &str) -> Option<Self> {
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

    fn encode(&self, conformance: Conformance) -> Result<Vec<u8>> {
        if conformance == Conformance::Strict && self.rely_on_vml.is_some() {
            return Err(invalid("relyOnVML is not valid in Strict web settings"));
        }
        let capacity = validate_value(self)?;
        let mut xml = String::new();
        xml.try_reserve_exact(capacity)
            .map_err(|source| Error::Allocation {
                resource: "web-settings XML",
                source,
            })?;
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str("<w:webSettings xmlns:w=\"");
        xml.push_str(conformance.wordprocessingml());
        xml.push_str("\" xmlns:r=\"");
        xml.push_str(conformance.relationships());
        xml.push_str("\">");

        if let Some(frameset) = &self.frameset {
            write_frameset(&mut xml, frameset, 1)?;
        }
        if let Some(divs) = &self.divs {
            xml.push_str("<w:divs>");
            for div in divs {
                write_html_div(&mut xml, div, 1)?;
            }
            xml.push_str("</w:divs>");
        }
        if let Some(value) = &self.encoding {
            write_value_element(&mut xml, "encoding", value)?;
        }
        write_optional_on_off(&mut xml, "optimizeForBrowser", self.optimize_for_browser)?;
        if conformance == Conformance::Transitional {
            write_optional_on_off(&mut xml, "relyOnVML", self.rely_on_vml)?;
        }
        write_optional_on_off(&mut xml, "allowPNG", self.allow_png)?;
        write_optional_on_off(&mut xml, "doNotRelyOnCSS", self.do_not_rely_on_css)?;
        write_optional_on_off(
            &mut xml,
            "doNotSaveAsSingleFile",
            self.do_not_save_as_single_file,
        )?;
        write_optional_on_off(
            &mut xml,
            "doNotOrganizeInFolder",
            self.do_not_organize_in_folder,
        )?;
        write_optional_on_off(
            &mut xml,
            "doNotUseLongFileNames",
            self.do_not_use_long_file_names,
        )?;
        if let Some(value) = self.pixels_per_inch {
            write!(xml, "<w:pixelsPerInch w:val=\"{value}\"/>")?;
        }
        if let Some(value) = self.target_screen_size {
            write_value_element(&mut xml, "targetScreenSz", value.as_str())?;
        }
        write_optional_on_off(&mut xml, "saveSmartTagsAsXml", self.save_smart_tags_as_xml)?;

        xml.push_str("</w:webSettings>");
        if xml.len() > MAX_XML_BYTES {
            return Err(invalid(format!(
                "web-settings XML exceeds {MAX_XML_BYTES} bytes"
            )));
        }
        Ok(xml.into_bytes())
    }

    fn read_part(part: &dyn Part) -> Result<(Self, Conformance)> {
        if part.content_type() != CONTENT_TYPE {
            return Err(Error::ContentType {
                expected: CONTENT_TYPE.to_owned(),
                actual: part.content_type().to_owned(),
            });
        }
        if part.blob().len() > MAX_XML_BYTES {
            return Err(invalid(format!(
                "web-settings XML exceeds {MAX_XML_BYTES} bytes"
            )));
        }
        let xml = process_web_xml(part.blob())?;
        let (settings, conformance) = Self::parse_xml(xml.as_ref())?;
        validate_frame_relationships(part, &settings, conformance)?;
        Ok((settings, conformance))
    }

    fn parse_xml(xml: &[u8]) -> Result<(Self, Conformance)> {
        if xml.len() > MAX_XML_BYTES {
            return Err(invalid(format!(
                "web-settings XML exceeds {MAX_XML_BYTES} bytes"
            )));
        }
        let mut reader = NsReader::from_reader(xml);
        let mut settings = Self::default();
        let mut depth = 0usize;
        let mut saw_root = false;
        let mut conformance = None;
        let mut last_child_rank = None;
        let mut budget = ParseBudget::default();

        loop {
            budget.event()?;
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);

            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::Invalid("Word web-settings XML nesting is too deep".into())
                    })?;
                    if depth > MAX_FRAMESET_NESTING {
                        return Err(invalid(
                            "Word web-settings XML nesting exceeds the safety limit",
                        ));
                    }
                    if depth == 1 {
                        conformance = Some(validate_root(&namespace, &element, saw_root)?);
                        saw_root = true;
                    } else if depth == 2 && saw_root && is_wordprocessing_namespace(&namespace) {
                        let profile = conformance.ok_or_else(|| {
                            invalid("web-settings root conformance was not resolved")
                        })?;
                        validate_web_child(&namespace, &element, profile, &mut last_child_rank)?;
                        if element.local_name().as_ref() == b"frameset" {
                            let frameset = parse_frameset(&mut reader, 1, &mut budget)?;
                            set_once(&mut settings.frameset, frameset, "frameset")?;
                            depth = depth.checked_sub(1).ok_or_else(|| {
                                Error::Invalid("invalid Word web-settings XML nesting".into())
                            })?;
                        } else if element.local_name().as_ref() == b"divs" {
                            let divs = parse_div_container(&mut reader, b"divs", 1, &mut budget)?;
                            set_once(&mut settings.divs, divs, "divs")?;
                            depth = depth.checked_sub(1).ok_or_else(|| {
                                Error::Invalid("invalid Word web-settings XML nesting".into())
                            })?;
                        } else if is_scalar_setting(element.local_name().as_ref()) {
                            parse_setting(&element, decoder, &resolver, &mut settings)?;
                            finish_leaf(
                                &mut reader,
                                element.local_name().as_ref(),
                                "web setting",
                                &mut budget,
                            )?;
                            depth = depth.checked_sub(1).ok_or_else(|| {
                                Error::Invalid("invalid Word web-settings XML nesting".into())
                            })?;
                        }
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        Error::Invalid("Word web-settings XML nesting is too deep".into())
                    })?;
                    if child_depth > MAX_FRAMESET_NESTING {
                        return Err(invalid(
                            "Word web-settings XML nesting exceeds the safety limit",
                        ));
                    }
                    if child_depth == 1 {
                        conformance = Some(validate_root(&namespace, &element, saw_root)?);
                        saw_root = true;
                    } else if child_depth == 2
                        && saw_root
                        && is_wordprocessing_namespace(&namespace)
                    {
                        let profile = conformance.ok_or_else(|| {
                            invalid("web-settings root conformance was not resolved")
                        })?;
                        validate_web_child(&namespace, &element, profile, &mut last_child_rank)?;
                        if element.local_name().as_ref() == b"frameset" {
                            set_once(&mut settings.frameset, Frameset::default(), "frameset")?;
                        } else if element.local_name().as_ref() == b"divs" {
                            return Err(invalid("Word HTML division container must not be empty"));
                        } else if is_scalar_setting(element.local_name().as_ref()) {
                            parse_setting(&element, decoder, &resolver, &mut settings)?;
                        }
                    }
                },
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::Invalid("invalid Word web-settings XML nesting".into())
                    })?;
                },
                Event::Eof if depth != 0 => {
                    return Err(Error::Invalid("unterminated Word web-settings XML".into()));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        if !saw_root {
            return Err(invalid("web-settings part has no webSettings root"));
        }
        let conformance =
            conformance.ok_or_else(|| invalid("web-settings root conformance was not resolved"))?;
        validate_value(&settings)?;
        Ok((settings, conformance))
    }
}

/// Parse bounded web-settings XML without resolving frame relationships.
pub fn parse(xml: &[u8]) -> Result<(Settings, Conformance)> {
    let processed = process_web_xml(xml)?;
    Settings::parse_xml(processed.as_ref())
}

fn process_web_xml(xml: &[u8]) -> Result<std::borrow::Cow<'_, [u8]>> {
    let limits = litchi_ooxml_common::MceLimits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        ..litchi_ooxml_common::MceLimits::default()
    };
    litchi_ooxml_common::process_markup_compatibility(
        xml,
        &litchi_ooxml_common::MceCapabilities::default(),
        &limits,
    )
    .map(|output| output.xml)
    .map_err(Error::from)
}

/// Serialize a checked web-settings model.
pub fn write(value: &Settings, conformance: Conformance) -> Result<Vec<u8>> {
    value.encode(conformance)
}

/// Read one bounded web-settings part and validate its frame relationships.
pub fn read(part: &dyn Part) -> Result<(Settings, Conformance)> {
    Settings::read_part(part)
}

#[derive(Debug, Clone)]
struct Owner {
    main: PackURI,
    target: PackURI,
    relationship_id: String,
    conformance: Conformance,
}

/// Load the document-owned web-settings model, if present.
pub fn load(package: &OpcPackage) -> Result<Option<(Settings, Conformance)>> {
    let Some(owner) = locate(package)? else {
        return Ok(None);
    };
    let part = package.get_part(&owner.target)?;
    let (settings, conformance) = read(part)?;
    if conformance != owner.conformance {
        return Err(invalid(
            "web-settings relationship and XML use different conformance families",
        ));
    }
    Ok(Some((settings, conformance)))
}

/// Move a complete model into package ownership.
///
/// Serialization, graph validation, and a semantic round trip complete before
/// signatures or package members are changed. Semantic/conformance equality is
/// a no-op that retains the producer's exact original bytes and signatures.
/// A different requested conformance is rejected before mutation. A real
/// semantic edit writes canonical modeled XML; ignored or unknown extension
/// markup is not retained source-surgically.
pub fn put(package: &mut OpcPackage, value: Settings, conformance: Conformance) -> Result<bool> {
    let xml = write(&value, conformance)?;
    let package_conformance = package_conformance(package)?;
    if package_conformance != conformance {
        return Err(invalid(
            "web-settings conformance does not match the document package",
        ));
    }

    let existing = locate(package)?;
    if let Some(owner) = &existing {
        let part = package.get_part(&owner.target)?;
        let (current, parsed_conformance) = read(part)?;
        if parsed_conformance != owner.conformance {
            return Err(invalid(
                "web-settings relationship and XML use different conformance families",
            ));
        }
        if owner.conformance == conformance && current == value {
            return Ok(false);
        }
        if has_other_inbound(package, owner)? {
            return Err(invalid(format!(
                "shared web-settings part '{}' cannot be overwritten",
                owner.target
            )));
        }

        let mut staged = BlobPart::new(owner.target.clone(), CONTENT_TYPE.to_owned(), xml);
        copy_relationships(part, &mut staged);
        let (round_trip, staged_conformance) = read(&staged)?;
        if round_trip != value || staged_conformance != conformance {
            return Err(invalid("staged web-settings XML did not round-trip"));
        }
        validate_internal_targets(package)?;

        package.add_part(Box::new(staged));
        package.unsign();
        return Ok(true);
    }

    let main = package.main_document_part()?.partname().clone();
    let target = next_part_name(package)?;
    let staged = BlobPart::new(target.clone(), CONTENT_TYPE.to_owned(), xml);
    let (round_trip, staged_conformance) = read(&staged)?;
    if round_trip != value || staged_conformance != conformance {
        return Err(invalid("staged web-settings XML did not round-trip"));
    }
    validate_internal_targets(package)?;
    let target_ref = target.relative_ref(main.base_uri());
    let relationship_type = conformance.relationship();

    package
        .get_part_mut(&main)?
        .rels_mut()
        .get_or_add(relationship_type, &target_ref);
    package.add_part(Box::new(staged));
    package.unsign();
    Ok(true)
}

/// Remove the document-owned web-settings part.
///
/// A part shared by another relationship is rejected rather than silently
/// detached or deleted. Absence is an exact, signature-preserving no-op.
pub fn remove(package: &mut OpcPackage) -> Result<bool> {
    let Some(owner) = locate(package)? else {
        return Ok(false);
    };
    let part = package.get_part(&owner.target)?;
    let (_, conformance) = read(part)?;
    if conformance != owner.conformance {
        return Err(invalid(
            "web-settings relationship and XML use different conformance families",
        ));
    }
    if has_other_inbound(package, &owner)? {
        return Err(invalid(format!(
            "shared web-settings part '{}' cannot be removed",
            owner.target
        )));
    }
    validate_internal_targets(package)?;

    package
        .get_part_mut(&owner.main)?
        .rels_mut()
        .remove(&owner.relationship_id);
    package.remove_part(&owner.target);
    package.unsign();
    Ok(true)
}

fn locate(package: &OpcPackage) -> Result<Option<Owner>> {
    use litchi_opc::constants::content_type as ct;

    let main_part = package.main_document_part()?;
    if !matches!(
        main_part.content_type(),
        ct::WML_DOCUMENT_MAIN
            | ct::WML_TEMPLATE_MAIN
            | ct::WML_DOCUMENT_MACRO_MAIN
            | ct::WML_TEMPLATE_MACRO_MAIN
    ) {
        return Err(invalid("main document is not a WordprocessingML document"));
    }
    let main = main_part.partname().clone();
    let expected_conformance = package_conformance(package)?;

    if package
        .rels()
        .iter()
        .any(|relationship| is_web_settings_relationship(relationship.reltype()))
    {
        return Err(invalid(
            "package root cannot own a web-settings relationship",
        ));
    }
    for part in package.iter_parts() {
        if part.partname() != &main
            && part
                .rels()
                .iter()
                .any(|relationship| is_web_settings_relationship(relationship.reltype()))
        {
            return Err(invalid(format!(
                "web-settings relationship has invalid source '{}'",
                part.partname()
            )));
        }
    }

    let mut relationships = main_part
        .rels()
        .iter()
        .filter(|relationship| is_web_settings_relationship(relationship.reltype()));
    let relationship = relationships.next();
    if relationships.next().is_some() {
        return Err(invalid("document has multiple web-settings relationships"));
    }

    let mut parts = package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE);
    let part_name = parts.next().map(|part| part.partname());
    if parts.next().is_some() {
        return Err(invalid("package has multiple web-settings parts"));
    }

    let Some(relationship) = relationship else {
        if part_name.is_some() {
            return Err(invalid(
                "package contains a web-settings part without document ownership",
            ));
        }
        return Ok(None);
    };
    if relationship.is_external() {
        return Err(invalid("web-settings relationship cannot be external"));
    }
    let conformance = Conformance::from_relationship(relationship.reltype())
        .ok_or_else(|| invalid("web-settings relationship type is unsupported"))?;
    if conformance != expected_conformance {
        return Err(invalid(
            "web-settings relationship conformance does not match the package",
        ));
    }
    let target = relationship.target_partname()?;
    let Some(part_name) = part_name else {
        let part = package.get_part(&target)?;
        return Err(Error::ContentType {
            expected: CONTENT_TYPE.to_owned(),
            actual: part.content_type().to_owned(),
        });
    };
    if part_name != &target {
        return Err(invalid(
            "web-settings relationship does not target the web-settings part",
        ));
    }

    Ok(Some(Owner {
        main,
        target,
        relationship_id: relationship.r_id().to_owned(),
        conformance,
    }))
}

fn package_conformance(package: &OpcPackage) -> Result<Conformance> {
    let mut relationships = package.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::OFFICE_DOCUMENT
                | STRICT_OFFICE_DOCUMENT_RELATIONSHIP
        )
    });
    let relationship = relationships
        .next()
        .ok_or_else(|| invalid("main-document relationship is missing"))?;
    if relationships.next().is_some() {
        return Err(invalid("package has multiple main-document relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("main-document relationship cannot be external"));
    }
    Ok(
        if relationship.reltype() == STRICT_OFFICE_DOCUMENT_RELATIONSHIP {
            Conformance::Strict
        } else {
            Conformance::Transitional
        },
    )
}

fn copy_relationships(source: &dyn Part, target: &mut dyn Part) {
    for relationship in source.rels().iter() {
        target.rels_mut().add_relationship(
            relationship.reltype().to_owned(),
            relationship.target_ref().to_owned(),
            relationship.r_id().to_owned(),
            relationship.is_external(),
        );
    }
}

fn has_other_inbound(package: &OpcPackage, owner: &Owner) -> Result<bool> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        if relationship.target_partname()? == owner.target {
            return Ok(true);
        }
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            if relationship.target_partname()? == owner.target
                && (part.partname() != &owner.main || relationship.r_id() != owner.relationship_id)
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn validate_internal_targets(package: &OpcPackage) -> Result<()> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        relationship.target_partname()?;
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            relationship.target_partname()?;
        }
    }
    Ok(())
}

fn next_part_name(package: &OpcPackage) -> Result<PackURI> {
    for index in 1..=4096 {
        let name = if index == 1 {
            "/word/webSettings.xml".to_owned()
        } else {
            format!("/word/webSettings{index}.xml")
        };
        let candidate = PackURI::new(name).map_err(Error::Uri)?;
        if package.validate_new_part_name(&candidate).is_ok() {
            return Ok(candidate);
        }
    }
    Err(invalid("no bounded web-settings part name is available"))
}

#[derive(Default)]
struct Budget {
    nodes: usize,
    encoded_bytes: usize,
}

impl Budget {
    fn node(&mut self, description: &str) -> Result<()> {
        self.nodes = self
            .nodes
            .checked_add(1)
            .ok_or_else(|| invalid("web-settings node count overflow"))?;
        if self.nodes > MAX_XML_EVENTS / 2 {
            return Err(invalid(format!(
                "web-settings {description} exceeds the node limit"
            )));
        }
        self.bytes(256, description)
    }

    fn text(&mut self, value: &str, description: &str) -> Result<()> {
        validate_text(value, description, true)?;
        let escaped = value
            .len()
            .checked_mul(6)
            .ok_or_else(|| invalid("web-settings escaped text size overflow"))?;
        self.bytes(escaped, description)
    }

    fn bytes(&mut self, bytes: usize, description: &str) -> Result<()> {
        self.encoded_bytes = self
            .encoded_bytes
            .checked_add(bytes)
            .ok_or_else(|| invalid("web-settings output size overflow"))?;
        if self.encoded_bytes > MAX_XML_BYTES {
            return Err(invalid(format!(
                "web-settings {description} exceeds {MAX_XML_BYTES} output bytes"
            )));
        }
        Ok(())
    }
}

fn validate_value(value: &Settings) -> Result<usize> {
    let mut budget = Budget {
        nodes: 1,
        encoded_bytes: 512,
    };
    if let Some(encoding) = &value.encoding {
        validate_encoding(encoding)?;
        budget.node("encoding")?;
        budget.text(encoding, "encoding")?;
    }
    if let Some(pixels) = value.pixels_per_inch {
        validate_pixels_per_inch(pixels)?;
        budget.node("pixels-per-inch")?;
    }
    if let Some(frameset) = &value.frameset {
        validate_frameset(frameset, 1, &mut budget)?;
    }
    if let Some(divs) = &value.divs {
        if divs.is_empty() {
            return Err(invalid("Word HTML division container must not be empty"));
        }
        budget.node("division container")?;
        validate_div_slice(divs, 1, &mut budget)?;
    }
    for present in [
        value.optimize_for_browser,
        value.rely_on_vml,
        value.allow_png,
        value.do_not_rely_on_css,
        value.do_not_save_as_single_file,
        value.do_not_organize_in_folder,
        value.do_not_use_long_file_names,
        value.target_screen_size.map(|_| true),
        value.save_smart_tags_as_xml,
    ] {
        if present.is_some() {
            budget.node("scalar setting")?;
        }
    }
    Ok(budget.encoded_bytes)
}

fn validate_frameset(value: &Frameset, depth: usize, budget: &mut Budget) -> Result<()> {
    if depth > MAX_FRAMESET_NESTING {
        return Err(invalid("web frameset nesting exceeds the safety limit"));
    }
    budget.node("frameset")?;
    if let Some(size) = &value.size {
        budget.node("frameset size")?;
        budget.text(size, "frameset size")?;
    }
    if let Some(split) = &value.split_bar {
        budget.node("frameset split bar")?;
        if let Some(color) = &split.color {
            validate_color(&color.value, "frameset splitter color")?;
            budget.node("frameset splitter color")?;
            budget.text(&color.value, "frameset splitter color")?;
        }
    }
    if value.layout.is_some() {
        budget.node("frameset layout")?;
    }
    for child in &value.children {
        match child {
            Child::Frameset(nested) => validate_frameset(nested, depth + 1, budget)?,
            Child::Frame(frame) => {
                budget.node("frame")?;
                for (text, description) in [
                    (frame.size.as_deref(), "frame size"),
                    (frame.name.as_deref(), "frame name"),
                    (
                        frame.source_file_relationship_id.as_deref(),
                        "frame relationship ID",
                    ),
                ] {
                    if let Some(text) = text {
                        if description == "frame relationship ID" {
                            validate_relationship_id(text)?;
                        }
                        budget.node(description)?;
                        budget.text(text, description)?;
                    }
                }
            },
        }
    }
    Ok(())
}

fn validate_divs(divs: &[Div], depth: usize) -> Result<()> {
    if divs.is_empty() {
        return Err(invalid("Word HTML division container must not be empty"));
    }
    let mut budget = Budget::default();
    validate_div_slice(divs, depth, &mut budget)
}

fn validate_div_slice(divs: &[Div], depth: usize, budget: &mut Budget) -> Result<()> {
    if depth > MAX_FRAMESET_NESTING {
        return Err(invalid("HTML division nesting exceeds the safety limit"));
    }
    let mut ids = std::collections::HashSet::new();
    ids.try_reserve(divs.len())
        .map_err(|source| Error::Allocation {
            resource: "HTML division identifier index",
            source,
        })?;
    for div in divs {
        if !ids.insert(div.id) {
            return Err(invalid(format!("HTML division '{}' is ambiguous", div.id)));
        }
        budget.node("HTML division")?;
        budget.bytes(decimal_len(div.id.get()), "HTML division ID")?;
        for margin in [div.left, div.right, div.top, div.bottom] {
            budget.node("HTML division margin")?;
            budget.bytes(decimal_len(margin.get()), "HTML division margin")?;
        }
        if let Some(borders) = &div.borders {
            budget.node("HTML division borders")?;
            for border in [&borders.top, &borders.left, &borders.bottom, &borders.right]
                .into_iter()
                .flatten()
            {
                validate_border_style(&border.style)?;
                budget.node("HTML division border")?;
                budget.text(&border.style, "HTML division border style")?;
                if let Some(color) = &border.color {
                    validate_color(color, "HTML division border color")?;
                    budget.text(color, "HTML division border color")?;
                }
            }
        }
        if !div.children.is_empty() {
            validate_div_slice(&div.children, depth + 1, budget)?;
        }
    }
    Ok(())
}

fn div_position(divs: &[Div], key: Key) -> Result<Option<usize>> {
    match key {
        Key::Index(index) => {
            if index >= divs.len() {
                Err(invalid(format!(
                    "HTML division position {index} is outside 0..{}",
                    divs.len()
                )))
            } else {
                Ok(Some(index))
            }
        },
        Key::Id(id) => {
            let mut matches = divs
                .iter()
                .enumerate()
                .filter(|(_, div)| div.id == id)
                .map(|(index, _)| index);
            let first = matches.next();
            if first.is_some() && matches.next().is_some() {
                Err(invalid(format!("HTML division ID '{id}' is ambiguous")))
            } else {
                Ok(first)
            }
        },
    }
}

fn validate_text(value: &str, description: &str, allow_empty: bool) -> Result<()> {
    if (!allow_empty && value.is_empty()) || value.len() > MAX_TEXT_BYTES {
        return Err(invalid(format!("invalid {description} length")));
    }
    if value
        .chars()
        .any(|character| character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
    {
        return Err(invalid(format!(
            "{description} contains a control character"
        )));
    }
    Ok(())
}

fn validate_encoding(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 255
        || !value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'_' | b'.'))
    {
        return Err(invalid("web encoding is not a bounded character-set name"));
    }
    Ok(())
}

fn validate_relationship_id(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 255
        || !value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.' | b':'))
    {
        return Err(invalid(
            "frame relationship ID is not a safe XML identifier",
        ));
    }
    Ok(())
}

fn validate_pixels_per_inch(value: u16) -> Result<()> {
    if value <= 1023 {
        Ok(())
    } else {
        Err(invalid("pixels-per-inch must be in the range 0..=1023"))
    }
}

fn parse_i64(value: &str, description: &str) -> Result<i64> {
    let value = value.trim();
    if value.is_empty() || value.len() > MAX_TEXT_BYTES {
        return Err(invalid(format!("invalid {description} length")));
    }
    value
        .parse::<i64>()
        .map_err(|_| invalid(format!("invalid {description} value '{value}'")))
}

fn decimal_len(value: i64) -> usize {
    let magnitude = value.unsigned_abs();
    let digits = if magnitude == 0 {
        1
    } else {
        magnitude.ilog10() as usize + 1
    };
    digits + usize::from(value.is_negative())
}

fn validate_border_style(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 64
        || !value.bytes().all(|byte| byte.is_ascii_alphanumeric())
    {
        return Err(invalid("HTML division border style is not a schema token"));
    }
    Ok(())
}

fn validate_color(value: &str, description: &str) -> Result<()> {
    if value.eq_ignore_ascii_case("auto")
        || (value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
    {
        Ok(())
    } else {
        Err(invalid(format!(
            "invalid {description} '{value}'; expected auto or six hexadecimal digits"
        )))
    }
}

fn write_value_element(xml: &mut String, name: &str, value: &str) -> Result<()> {
    write!(xml, "<w:{name} w:val=\"{}\"/>", escape(value))
        .map_err(|error| Error::Xml(error.to_string()))
}

fn write_optional_on_off(xml: &mut String, name: &str, value: Option<bool>) -> Result<()> {
    match value {
        Some(true) => {
            write!(xml, "<w:{name}/>")?;
        },
        Some(false) => {
            write!(xml, "<w:{name} w:val=\"false\"/>")?;
        },
        None => {},
    }
    Ok(())
}

/// Write the explicit numeric form required by desktop Word for `CT_Div`
/// role markers. Word rejects the otherwise schema-valid empty true form for
/// both `blockQuote` and `bodyDiv`.
fn write_explicit_on_off(xml: &mut String, name: &str, value: Option<bool>) -> Result<()> {
    if let Some(value) = value {
        write!(xml, "<w:{name} w:val=\"{}\"/>", u8::from(value))?;
    }
    Ok(())
}

fn write_frameset(xml: &mut String, frameset: &Frameset, nesting: usize) -> Result<()> {
    if nesting > MAX_FRAMESET_NESTING {
        return Err(Error::Invalid(
            "Word web frameset nesting exceeds the supported safety limit".into(),
        ));
    }
    xml.push_str("<w:frameset>");
    if let Some(value) = &frameset.size {
        write_value_element(xml, "sz", value)?;
    }
    if let Some(split_bar) = &frameset.split_bar {
        write_frameset_split_bar(xml, split_bar)?;
    }
    if let Some(layout) = frameset.layout {
        write_value_element(xml, "frameLayout", layout.as_str())?;
    }
    for child in &frameset.children {
        match child {
            Child::Frameset(nested) => write_frameset(xml, nested, nesting + 1)?,
            Child::Frame(frame) => write_frame(xml, frame)?,
        }
    }
    xml.push_str("</w:frameset>");
    Ok(())
}

fn write_frame(xml: &mut String, frame: &Frame) -> Result<()> {
    xml.push_str("<w:frame>");
    if let Some(value) = &frame.size {
        write_value_element(xml, "sz", value)?;
    }
    if let Some(value) = &frame.name {
        write_value_element(xml, "name", value)?;
    }
    if let Some(value) = &frame.source_file_relationship_id {
        write!(xml, "<w:sourceFileName r:id=\"{}\"/>", escape(value))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(value) = frame.margin_width {
        write_value_element(xml, "marW", &value.to_string())?;
    }
    if let Some(value) = frame.margin_height {
        write_value_element(xml, "marH", &value.to_string())?;
    }
    if let Some(value) = frame.scrollbar {
        write_value_element(xml, "scrollbar", value.as_str())?;
    }
    write_optional_on_off(xml, "noResizeAllowed", frame.no_resize_allowed)?;
    write_optional_on_off(xml, "linkedToFile", frame.linked_to_file)?;
    xml.push_str("</w:frame>");
    Ok(())
}

fn write_frameset_split_bar(xml: &mut String, split_bar: &SplitBar) -> Result<()> {
    xml.push_str("<w:framesetSplitbar>");
    if let Some(value) = split_bar.width_twips {
        write_value_element(xml, "w", &value.to_string())?;
    }
    if let Some(color) = &split_bar.color {
        xml.push_str("<w:color");
        write_color_attributes(
            xml,
            &color.value,
            color.theme_color,
            color.theme_tint,
            color.theme_shade,
        )?;
        xml.push_str("/>");
    }
    write_optional_on_off(xml, "noBorder", split_bar.no_border)?;
    write_optional_on_off(xml, "flatBorders", split_bar.flat_borders)?;
    xml.push_str("</w:framesetSplitbar>");
    Ok(())
}

fn write_html_div(xml: &mut String, div: &Div, nesting: usize) -> Result<()> {
    if nesting > MAX_FRAMESET_NESTING {
        return Err(Error::Invalid(
            "Word HTML division nesting exceeds the supported safety limit".into(),
        ));
    }
    write!(xml, "<w:div w:id=\"{}\">", div.id)?;
    write_explicit_on_off(xml, "blockQuote", div.block_quote)?;
    write_explicit_on_off(xml, "bodyDiv", div.body_div)?;
    for (name, value) in [
        ("marLeft", div.left),
        ("marRight", div.right),
        ("marTop", div.top),
        ("marBottom", div.bottom),
    ] {
        write!(xml, "<w:{name} w:val=\"{value}\"/>")?;
    }
    if let Some(borders) = &div.borders {
        write_html_div_borders(xml, borders)?;
    }
    if !div.children.is_empty() {
        xml.push_str("<w:divsChild>");
        for child in &div.children {
            write_html_div(xml, child, nesting + 1)?;
        }
        xml.push_str("</w:divsChild>");
    }
    xml.push_str("</w:div>");
    Ok(())
}

fn write_html_div_borders(xml: &mut String, borders: &Borders) -> Result<()> {
    xml.push_str("<w:divBdr>");
    for (name, border) in [
        ("top", &borders.top),
        ("left", &borders.left),
        ("bottom", &borders.bottom),
        ("right", &borders.right),
    ] {
        let Some(border) = border else {
            continue;
        };
        write!(xml, "<w:{name} w:val=\"{}\"", escape(&border.style))
            .map_err(|error| Error::Xml(error.to_string()))?;
        if let Some(color) = &border.color {
            write!(xml, " w:color=\"{}\"", escape(color))
                .map_err(|error| Error::Xml(error.to_string()))?;
        }
        write_theme_attributes(
            xml,
            border.theme_color,
            border.theme_tint,
            border.theme_shade,
        )?;
        if let Some(value) = border.size_eighth_points {
            write!(xml, " w:sz=\"{value}\"").map_err(|error| Error::Xml(error.to_string()))?;
        }
        if let Some(value) = border.space_points {
            write!(xml, " w:space=\"{value}\"").map_err(|error| Error::Xml(error.to_string()))?;
        }
        write_optional_on_off_attribute(xml, "shadow", border.shadow)?;
        write_optional_on_off_attribute(xml, "frame", border.frame)?;
        xml.push_str("/>");
    }
    xml.push_str("</w:divBdr>");
    Ok(())
}

fn write_color_attributes(
    xml: &mut String,
    value: &str,
    theme_color: Option<Theme>,
    theme_tint: Option<u8>,
    theme_shade: Option<u8>,
) -> Result<()> {
    write!(xml, " w:val=\"{}\"", escape(value)).map_err(|error| Error::Xml(error.to_string()))?;
    write_theme_attributes(xml, theme_color, theme_tint, theme_shade)
}

fn write_theme_attributes(
    xml: &mut String,
    theme_color: Option<Theme>,
    theme_tint: Option<u8>,
    theme_shade: Option<u8>,
) -> Result<()> {
    if let Some(value) = theme_color {
        write!(xml, " w:themeColor=\"{}\"", value.as_str())
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(value) = theme_tint {
        write!(xml, " w:themeTint=\"{value:02X}\"")
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(value) = theme_shade {
        write!(xml, " w:themeShade=\"{value:02X}\"")
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    Ok(())
}

fn write_optional_on_off_attribute(
    xml: &mut String,
    name: &str,
    value: Option<bool>,
) -> Result<()> {
    if let Some(value) = value {
        write!(
            xml,
            " w:{name}=\"{}\"",
            if value { "true" } else { "false" }
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    Ok(())
}

fn validate_frame_relationships(
    part: &dyn Part,
    settings: &Settings,
    conformance: Conformance,
) -> Result<()> {
    const FRAME_RELATIONSHIP: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame";
    const STRICT_FRAME_RELATIONSHIP: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/frame";

    fn validate(part: &dyn Part, frameset: &Frameset, expected_type: &str) -> Result<()> {
        for child in &frameset.children {
            match child {
                Child::Frameset(nested) => {
                    validate(part, nested, expected_type)?;
                },
                Child::Frame(frame) => {
                    let Some(id) = &frame.source_file_relationship_id else {
                        continue;
                    };
                    let relationship = part.rels().get(id).ok_or_else(|| {
                        Error::Invalid(format!("frame source relationship '{id}' does not exist"))
                    })?;
                    if relationship.reltype() != expected_type {
                        return Err(Error::Invalid(format!(
                            "frame source relationship '{id}' has an invalid type"
                        )));
                    }
                },
            }
        }
        Ok(())
    }

    if let Some(frameset) = &settings.frameset {
        let expected = match conformance {
            Conformance::Transitional => FRAME_RELATIONSHIP,
            Conformance::Strict => STRICT_FRAME_RELATIONSHIP,
        };
        validate(part, frameset, expected)?;
    }
    Ok(())
}

fn validate_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    saw_root: bool,
) -> Result<Conformance> {
    if saw_root || element.local_name().as_ref() != b"webSettings" {
        return Err(invalid(
            "web-settings part has an invalid or trailing root element",
        ));
    }
    let ResolveResult::Bound(Namespace(namespace)) = namespace else {
        return Err(invalid(
            "web-settings root has no WordprocessingML namespace",
        ));
    };
    Conformance::from_word_namespace(namespace)
        .ok_or_else(|| invalid("web-settings root uses an unsupported namespace"))
}

fn validate_web_child(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    conformance: Conformance,
    last_rank: &mut Option<u8>,
) -> Result<()> {
    let ResolveResult::Bound(Namespace(namespace)) = namespace else {
        return Err(invalid(
            "web-settings child has no WordprocessingML namespace",
        ));
    };
    if *namespace != conformance.wordprocessingml().as_bytes() {
        return Err(invalid(
            "web-settings child namespace does not match the root conformance",
        ));
    }

    let name = element.local_name();
    let name = name.as_ref();
    if conformance == Conformance::Strict && name == b"relyOnVML" {
        return Err(invalid("relyOnVML is not valid in Strict web settings"));
    }
    let rank = match name {
        b"frameset" => Some(0),
        b"divs" => Some(1),
        b"encoding" => Some(2),
        b"optimizeForBrowser" => Some(3),
        b"relyOnVML" => Some(4),
        b"allowPNG" => Some(5),
        b"doNotRelyOnCSS" => Some(6),
        b"doNotSaveAsSingleFile" => Some(7),
        b"doNotOrganizeInFolder" => Some(8),
        b"doNotUseLongFileNames" => Some(9),
        b"pixelsPerInch" => Some(10),
        b"targetScreenSz" => Some(11),
        b"saveSmartTagsAsXml" => Some(12),
        _ => None,
    };
    if let Some(rank) = rank {
        if last_rank.is_some_and(|last| rank < last) {
            return Err(invalid("web-settings children are out of schema order"));
        }
        *last_rank = Some(rank);
    }
    Ok(())
}

fn is_scalar_setting(name: &[u8]) -> bool {
    matches!(
        name,
        b"encoding"
            | b"optimizeForBrowser"
            | b"relyOnVML"
            | b"allowPNG"
            | b"doNotRelyOnCSS"
            | b"doNotSaveAsSingleFile"
            | b"doNotOrganizeInFolder"
            | b"doNotUseLongFileNames"
            | b"pixelsPerInch"
            | b"targetScreenSz"
            | b"saveSmartTagsAsXml"
    )
}

fn parse_setting(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    settings: &mut Settings,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"encoding" => set_once(
            &mut settings.encoding,
            required_value(element, decoder, resolver, "web encoding")?,
            "encoding",
        ),
        b"optimizeForBrowser" => set_on_off(
            &mut settings.optimize_for_browser,
            element,
            decoder,
            resolver,
            "optimizeForBrowser",
        ),
        b"relyOnVML" => set_on_off(
            &mut settings.rely_on_vml,
            element,
            decoder,
            resolver,
            "relyOnVML",
        ),
        b"allowPNG" => set_on_off(
            &mut settings.allow_png,
            element,
            decoder,
            resolver,
            "allowPNG",
        ),
        b"doNotRelyOnCSS" => set_on_off(
            &mut settings.do_not_rely_on_css,
            element,
            decoder,
            resolver,
            "doNotRelyOnCSS",
        ),
        b"doNotSaveAsSingleFile" => set_on_off(
            &mut settings.do_not_save_as_single_file,
            element,
            decoder,
            resolver,
            "doNotSaveAsSingleFile",
        ),
        b"doNotOrganizeInFolder" => set_on_off(
            &mut settings.do_not_organize_in_folder,
            element,
            decoder,
            resolver,
            "doNotOrganizeInFolder",
        ),
        b"doNotUseLongFileNames" => set_on_off(
            &mut settings.do_not_use_long_file_names,
            element,
            decoder,
            resolver,
            "doNotUseLongFileNames",
        ),
        b"pixelsPerInch" => {
            let value = required_value(element, decoder, resolver, "pixels per inch")?;
            let value = value
                .trim()
                .parse::<u16>()
                .map_err(|_| invalid(format!("invalid pixels-per-inch value '{value}'")))?;
            validate_pixels_per_inch(value)?;
            set_once(&mut settings.pixels_per_inch, value, "pixelsPerInch")
        },
        b"targetScreenSz" => {
            let value = required_value(element, decoder, resolver, "target screen size")?;
            let value = Screen::from_xml(&value).ok_or_else(|| {
                Error::Invalid(format!("invalid target-screen-size value '{value}'"))
            })?;
            set_once(&mut settings.target_screen_size, value, "targetScreenSz")
        },
        b"saveSmartTagsAsXml" => set_on_off(
            &mut settings.save_smart_tags_as_xml,
            element,
            decoder,
            resolver,
            "saveSmartTagsAsXml",
        ),
        _ => Ok(()),
    }
}

const MAX_FRAMESET_NESTING: usize = 128;

fn parse_frameset(
    reader: &mut NsReader<&[u8]>,
    nesting: usize,
    budget: &mut ParseBudget,
) -> Result<Frameset> {
    if nesting > MAX_FRAMESET_NESTING {
        return Err(Error::Invalid(
            "Word web frameset nesting exceeds the supported safety limit".into(),
        ));
    }
    let mut frameset = Frameset::default();
    loop {
        budget.event()?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if is_wordprocessing_namespace(&namespace) => {
                match element.local_name().as_ref() {
                    b"sz" => {
                        let value = required_value(&element, decoder, &resolver, "frame size")?;
                        set_once(&mut frameset.size, value, "frameset size")?;
                        finish_leaf(
                            reader,
                            element.local_name().as_ref(),
                            "frameset size",
                            budget,
                        )?;
                    },
                    b"framesetSplitbar" => {
                        let split_bar = parse_frameset_split_bar(reader, budget)?;
                        set_once(&mut frameset.split_bar, split_bar, "frameset split bar")?;
                    },
                    b"frameLayout" => {
                        let layout = parse_frame_layout(&element, decoder, &resolver)?;
                        set_once(&mut frameset.layout, layout, "frame layout")?;
                        finish_leaf(
                            reader,
                            element.local_name().as_ref(),
                            "frame layout",
                            budget,
                        )?;
                    },
                    b"frameset" => {
                        reserve_one(&mut frameset.children, "parsed frameset child")?;
                        let child = parse_frameset(reader, nesting + 1, budget)?;
                        frameset.children.push(Child::Frameset(child));
                    },
                    b"frame" => {
                        reserve_one(&mut frameset.children, "parsed frameset child")?;
                        let child = parse_frame(reader, budget)?;
                        frameset.children.push(Child::Frame(child));
                    },
                    _ => skip_element(reader, budget)?,
                }
            },
            Event::Empty(element) if is_wordprocessing_namespace(&namespace) => {
                match element.local_name().as_ref() {
                    b"sz" => set_once(
                        &mut frameset.size,
                        required_value(&element, decoder, &resolver, "frame size")?,
                        "frameset size",
                    )?,
                    b"framesetSplitbar" => set_once(
                        &mut frameset.split_bar,
                        SplitBar::default(),
                        "frameset split bar",
                    )?,
                    b"frameLayout" => set_once(
                        &mut frameset.layout,
                        parse_frame_layout(&element, decoder, &resolver)?,
                        "frame layout",
                    )?,
                    b"frameset" => {
                        reserve_one(&mut frameset.children, "parsed frameset child")?;
                        frameset.children.push(Child::Frameset(Frameset::default()));
                    },
                    b"frame" => {
                        reserve_one(&mut frameset.children, "parsed frameset child")?;
                        frameset.children.push(Child::Frame(Frame::default()));
                    },
                    _ => {},
                }
            },
            Event::Start(_) => skip_element(reader, budget)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"frameset" =>
            {
                return Ok(frameset);
            },
            Event::Eof => {
                return Err(Error::Invalid("unterminated Word web frameset".into()));
            },
            _ => {},
        }
    }
}

fn parse_div_container(
    reader: &mut NsReader<&[u8]>,
    end_name: &[u8],
    nesting: usize,
    budget: &mut ParseBudget,
) -> Result<Vec<Div>> {
    if nesting > MAX_FRAMESET_NESTING {
        return Err(Error::Invalid(
            "Word HTML division nesting exceeds the supported safety limit".into(),
        ));
    }
    let mut divs = Vec::new();
    loop {
        budget.event()?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"div" =>
            {
                reserve_one(&mut divs, "parsed HTML division")?;
                let div = parse_html_div(reader, &element, decoder, &resolver, nesting, budget)?;
                divs.push(div);
            },
            Event::Empty(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"div" =>
            {
                let _ = parse_div_id(&element, decoder, &resolver)?;
                return Err(invalid(
                    "Word HTML division is missing its four required margins",
                ));
            },
            Event::Start(_) => skip_element(reader, budget)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == end_name =>
            {
                if divs.is_empty() {
                    return Err(invalid("Word HTML division container must not be empty"));
                }
                return Ok(divs);
            },
            Event::Eof => {
                return Err(Error::Invalid(
                    "unterminated Word HTML division container".into(),
                ));
            },
            _ => {},
        }
    }
}

fn parse_div_id(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Id> {
    let id = word_attribute_value(element, b"id", decoder, resolver)?
        .ok_or_else(|| Error::Invalid("Word HTML division ID is required".into()))?;
    Id::parse(&id)
}

struct DivParse {
    value: Div,
    margins: u8,
    last_rank: Option<u8>,
}

impl DivParse {
    const LEFT: u8 = 1;
    const RIGHT: u8 = 2;
    const TOP: u8 = 4;
    const BOTTOM: u8 = 8;
    const ALL_MARGINS: u8 = Self::LEFT | Self::RIGHT | Self::TOP | Self::BOTTOM;

    fn new(id: Id) -> Self {
        Self {
            value: Div::new(id),
            margins: 0,
            last_rank: None,
        }
    }

    fn advance(&mut self, rank: u8) -> Result<()> {
        if self.last_rank.is_some_and(|last| rank < last) {
            return Err(invalid("HTML division children are out of schema order"));
        }
        self.last_rank = Some(rank);
        Ok(())
    }

    fn set_margin(&mut self, bit: u8, value: Twips, description: &'static str) -> Result<()> {
        if self.margins & bit != 0 {
            return Err(invalid(format!("duplicate Word {description}")));
        }
        self.margins |= bit;
        match bit {
            Self::LEFT => self.value.left = value,
            Self::RIGHT => self.value.right = value,
            Self::TOP => self.value.top = value,
            Self::BOTTOM => self.value.bottom = value,
            _ => return Err(invalid("invalid HTML division margin selector")),
        }
        Ok(())
    }

    fn append_children(&mut self, mut children: Vec<Div>) -> Result<()> {
        self.value
            .children
            .try_reserve(children.len())
            .map_err(|source| Error::Allocation {
                resource: "parsed HTML child divisions",
                source,
            })?;
        self.value.children.append(&mut children);
        Ok(())
    }

    fn finish(self) -> Result<Div> {
        if self.margins != Self::ALL_MARGINS {
            return Err(invalid(
                "Word HTML division is missing one or more required margins",
            ));
        }
        Ok(self.value)
    }
}

fn parse_html_div(
    reader: &mut NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    nesting: usize,
    budget: &mut ParseBudget,
) -> Result<Div> {
    let mut div = DivParse::new(parse_div_id(element, decoder, resolver)?);
    loop {
        budget.event()?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if is_wordprocessing_namespace(&namespace) => {
                match element.local_name().as_ref() {
                    name if is_html_div_leaf(name) => {
                        parse_html_div_leaf(&element, decoder, &resolver, &mut div)?;
                        finish_leaf(
                            reader,
                            element.local_name().as_ref(),
                            "HTML division property",
                            budget,
                        )?;
                    },
                    b"divBdr" => {
                        div.advance(6)?;
                        let borders = parse_html_div_borders(reader, budget)?;
                        set_once(&mut div.value.borders, borders, "HTML division borders")?;
                    },
                    b"divsChild" => {
                        div.advance(7)?;
                        let children =
                            parse_div_container(reader, b"divsChild", nesting + 1, budget)?;
                        div.append_children(children)?;
                    },
                    _ => skip_element(reader, budget)?,
                }
            },
            Event::Empty(element) if is_wordprocessing_namespace(&namespace) => {
                match element.local_name().as_ref() {
                    name if is_html_div_leaf(name) => {
                        parse_html_div_leaf(&element, decoder, &resolver, &mut div)?;
                    },
                    b"divBdr" => {
                        div.advance(6)?;
                        set_once(
                            &mut div.value.borders,
                            Borders::default(),
                            "HTML division borders",
                        )?;
                    },
                    b"divsChild" => {
                        return Err(invalid(
                            "Word HTML child division container must not be empty",
                        ));
                    },
                    _ => {},
                }
            },
            Event::Start(_) => skip_element(reader, budget)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"div" =>
            {
                return div.finish();
            },
            Event::Eof => {
                return Err(Error::Invalid("unterminated Word HTML division".into()));
            },
            _ => {},
        }
    }
}

fn is_html_div_leaf(name: &[u8]) -> bool {
    matches!(
        name,
        b"blockQuote" | b"bodyDiv" | b"marLeft" | b"marRight" | b"marTop" | b"marBottom"
    )
}

fn parse_html_div_leaf(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    div: &mut DivParse,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"blockQuote" => {
            div.advance(0)?;
            set_on_off(
                &mut div.value.block_quote,
                element,
                decoder,
                resolver,
                "HTML blockQuote",
            )
        },
        b"bodyDiv" => {
            div.advance(1)?;
            set_on_off(
                &mut div.value.body_div,
                element,
                decoder,
                resolver,
                "HTML bodyDiv",
            )
        },
        b"marLeft" => set_signed_twips(
            div,
            DivParse::LEFT,
            2,
            element,
            decoder,
            resolver,
            "HTML division left margin",
        ),
        b"marRight" => set_signed_twips(
            div,
            DivParse::RIGHT,
            3,
            element,
            decoder,
            resolver,
            "HTML division right margin",
        ),
        b"marTop" => set_signed_twips(
            div,
            DivParse::TOP,
            4,
            element,
            decoder,
            resolver,
            "HTML division top margin",
        ),
        b"marBottom" => set_signed_twips(
            div,
            DivParse::BOTTOM,
            5,
            element,
            decoder,
            resolver,
            "HTML division bottom margin",
        ),
        _ => Ok(()),
    }
}

fn set_signed_twips(
    div: &mut DivParse,
    bit: u8,
    rank: u8,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &'static str,
) -> Result<()> {
    div.advance(rank)?;
    let value = required_value(element, decoder, resolver, description)?;
    let value = Twips::parse(&value)?;
    div.set_margin(bit, value, description)
}

fn parse_html_div_borders(
    reader: &mut NsReader<&[u8]>,
    budget: &mut ParseBudget,
) -> Result<Borders> {
    let mut borders = Borders::default();
    loop {
        budget.event()?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element)
                if is_wordprocessing_namespace(&namespace)
                    && is_html_div_border_side(element.local_name().as_ref()) =>
            {
                set_html_div_border(&mut borders, &element, decoder, &resolver)?;
                finish_leaf(
                    reader,
                    element.local_name().as_ref(),
                    "HTML division border",
                    budget,
                )?;
            },
            Event::Empty(element)
                if is_wordprocessing_namespace(&namespace)
                    && is_html_div_border_side(element.local_name().as_ref()) =>
            {
                set_html_div_border(&mut borders, &element, decoder, &resolver)?;
            },
            Event::Start(_) => skip_element(reader, budget)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"divBdr" =>
            {
                return Ok(borders);
            },
            Event::Eof => {
                return Err(Error::Invalid(
                    "unterminated Word HTML division borders".into(),
                ));
            },
            _ => {},
        }
    }
}

fn is_html_div_border_side(name: &[u8]) -> bool {
    matches!(name, b"top" | b"left" | b"bottom" | b"right")
}

fn set_html_div_border(
    borders: &mut Borders,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    let border = parse_html_div_border(element, decoder, resolver)?;
    let (slot, description) = match element.local_name().as_ref() {
        b"top" => (&mut borders.top, "top HTML division border"),
        b"left" => (&mut borders.left, "left HTML division border"),
        b"bottom" => (&mut borders.bottom, "bottom HTML division border"),
        b"right" => (&mut borders.right, "right HTML division border"),
        _ => return Ok(()),
    };
    set_once(slot, border, description)
}

fn parse_html_div_border(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Border> {
    let style = required_value(element, decoder, resolver, "HTML division border style")?;
    let color = word_attribute_value(element, b"color", decoder, resolver)?
        .map(|value| validate_word_color(value, "HTML division border color"))
        .transpose()?;
    let theme_color = word_attribute_value(element, b"themeColor", decoder, resolver)?
        .map(|value| {
            Theme::parse(&value)
                .ok_or_else(|| Error::Invalid(format!("invalid theme color '{value}'")))
        })
        .transpose()?;
    Ok(Border {
        style,
        color,
        theme_color,
        theme_tint: optional_hex_byte(element, b"themeTint", decoder, resolver)?,
        theme_shade: optional_hex_byte(element, b"themeShade", decoder, resolver)?,
        size_eighth_points: optional_unsigned_long_attribute(element, b"sz", decoder, resolver)?,
        space_points: optional_unsigned_long_attribute(element, b"space", decoder, resolver)?,
        shadow: optional_on_off_attribute(element, b"shadow", decoder, resolver)?,
        frame: optional_on_off_attribute(element, b"frame", decoder, resolver)?,
    })
}

fn parse_frame(reader: &mut NsReader<&[u8]>, budget: &mut ParseBudget) -> Result<Frame> {
    let mut frame = Frame::default();
    loop {
        budget.event()?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if is_wordprocessing_namespace(&namespace) => {
                if is_frame_property(element.local_name().as_ref()) {
                    parse_frame_property(&element, decoder, &resolver, &mut frame)?;
                    finish_leaf(
                        reader,
                        element.local_name().as_ref(),
                        "frame property",
                        budget,
                    )?;
                } else {
                    skip_element(reader, budget)?;
                }
            },
            Event::Empty(element)
                if is_wordprocessing_namespace(&namespace)
                    && is_frame_property(element.local_name().as_ref()) =>
            {
                parse_frame_property(&element, decoder, &resolver, &mut frame)?;
            },
            Event::Start(_) => skip_element(reader, budget)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"frame" =>
            {
                return Ok(frame);
            },
            Event::Eof => {
                return Err(Error::Invalid("unterminated Word web frame".into()));
            },
            _ => {},
        }
    }
}

fn is_frame_property(name: &[u8]) -> bool {
    matches!(
        name,
        b"sz"
            | b"name"
            | b"sourceFileName"
            | b"marW"
            | b"marH"
            | b"scrollbar"
            | b"noResizeAllowed"
            | b"linkedToFile"
    )
}

fn parse_frame_property(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    frame: &mut Frame,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"sz" => set_once(
            &mut frame.size,
            required_value(element, decoder, resolver, "frame size")?,
            "frame size",
        ),
        b"name" => set_once(
            &mut frame.name,
            required_value(element, decoder, resolver, "frame name")?,
            "frame name",
        ),
        b"sourceFileName" => set_once(
            &mut frame.source_file_relationship_id,
            required_relationship_id(element, decoder, resolver)?,
            "frame source file",
        ),
        b"marW" => set_once(
            &mut frame.margin_width,
            required_unsigned_long(element, decoder, resolver, "frame margin width")?,
            "frame margin width",
        ),
        b"marH" => set_once(
            &mut frame.margin_height,
            required_unsigned_long(element, decoder, resolver, "frame margin height")?,
            "frame margin height",
        ),
        b"scrollbar" => {
            let value = required_value(element, decoder, resolver, "frame scrollbar")?;
            let value = Scrollbar::from_xml(&value).ok_or_else(|| {
                Error::Invalid(format!("invalid frame scrollbar value '{value}'"))
            })?;
            set_once(&mut frame.scrollbar, value, "frame scrollbar")
        },
        b"noResizeAllowed" => set_on_off(
            &mut frame.no_resize_allowed,
            element,
            decoder,
            resolver,
            "frame noResizeAllowed",
        ),
        b"linkedToFile" => set_on_off(
            &mut frame.linked_to_file,
            element,
            decoder,
            resolver,
            "frame linkedToFile",
        ),
        _ => Ok(()),
    }
}

fn parse_frameset_split_bar(
    reader: &mut NsReader<&[u8]>,
    budget: &mut ParseBudget,
) -> Result<SplitBar> {
    let mut split_bar = SplitBar::default();
    loop {
        budget.event()?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if is_wordprocessing_namespace(&namespace) => {
                if is_split_bar_property(element.local_name().as_ref()) {
                    parse_split_bar_property(&element, decoder, &resolver, &mut split_bar)?;
                    finish_leaf(
                        reader,
                        element.local_name().as_ref(),
                        "frameset split-bar property",
                        budget,
                    )?;
                } else {
                    skip_element(reader, budget)?;
                }
            },
            Event::Empty(element)
                if is_wordprocessing_namespace(&namespace)
                    && is_split_bar_property(element.local_name().as_ref()) =>
            {
                parse_split_bar_property(&element, decoder, &resolver, &mut split_bar)?;
            },
            Event::Start(_) => skip_element(reader, budget)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"framesetSplitbar" =>
            {
                return Ok(split_bar);
            },
            Event::Eof => {
                return Err(Error::Invalid(
                    "unterminated Word frameset split bar".into(),
                ));
            },
            _ => {},
        }
    }
}

fn is_split_bar_property(name: &[u8]) -> bool {
    matches!(name, b"w" | b"color" | b"noBorder" | b"flatBorders")
}

fn parse_split_bar_property(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    split_bar: &mut SplitBar,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"w" => set_once(
            &mut split_bar.width_twips,
            required_unsigned_long(element, decoder, resolver, "split-bar width")?,
            "split-bar width",
        ),
        b"color" => set_once(
            &mut split_bar.color,
            parse_frameset_color(element, decoder, resolver)?,
            "split-bar color",
        ),
        b"noBorder" => set_on_off(
            &mut split_bar.no_border,
            element,
            decoder,
            resolver,
            "split-bar noBorder",
        ),
        b"flatBorders" => set_on_off(
            &mut split_bar.flat_borders,
            element,
            decoder,
            resolver,
            "split-bar flatBorders",
        ),
        _ => Ok(()),
    }
}

fn parse_frame_layout(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Layout> {
    let value = required_value(element, decoder, resolver, "frame layout")?;
    Layout::from_xml(&value)
        .ok_or_else(|| Error::Invalid(format!("invalid frame-layout value '{value}'")))
}

fn parse_frameset_color(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Color> {
    let value = validate_word_color(
        required_value(element, decoder, resolver, "frameset splitter color")?,
        "frameset splitter color",
    )?;
    let theme_color = word_attribute_value(element, b"themeColor", decoder, resolver)?
        .map(|value| {
            Theme::parse(&value)
                .ok_or_else(|| Error::Invalid(format!("invalid theme color '{value}'")))
        })
        .transpose()?;
    let theme_tint = optional_hex_byte(element, b"themeTint", decoder, resolver)?;
    let theme_shade = optional_hex_byte(element, b"themeShade", decoder, resolver)?;
    Ok(Color {
        value,
        theme_color,
        theme_tint,
        theme_shade,
    })
}

fn validate_word_color(value: String, description: &str) -> Result<String> {
    if value != "auto" && (value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()))
    {
        return Err(Error::Invalid(format!("invalid {description} '{value}'")));
    }
    Ok(value)
}

fn required_unsigned_long(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<u64> {
    let value = required_value(element, decoder, resolver, description)?;
    value
        .trim()
        .parse::<u64>()
        .map_err(|_| Error::Invalid(format!("invalid unsigned {description} value '{value}'")))
}

fn optional_unsigned_long_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<u64>> {
    word_attribute_value(element, name, decoder, resolver)?
        .map(|value| {
            value.trim().parse::<u64>().map_err(|_| {
                Error::Invalid(format!("invalid unsigned Word attribute value '{value}'"))
            })
        })
        .transpose()
}

fn optional_on_off_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<bool>> {
    word_attribute_value(element, name, decoder, resolver)?
        .map(|value| match value.as_str() {
            "true" | "1" | "on" => Ok(true),
            "false" | "0" | "off" => Ok(false),
            _ => Err(Error::Invalid(format!(
                "invalid Word on/off value '{value}'"
            ))),
        })
        .transpose()
}

fn optional_hex_byte(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<u8>> {
    word_attribute_value(element, name, decoder, resolver)?
        .map(|value| {
            if value.len() != 2 {
                return Err(Error::Invalid(format!(
                    "invalid hexadecimal byte '{value}'"
                )));
            }
            u8::from_str_radix(&value, 16)
                .map_err(|_| Error::Invalid(format!("invalid hexadecimal byte '{value}'")))
        })
        .transpose()
}

fn required_relationship_id(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<String> {
    const RELATIONSHIPS: &[u8] =
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_RELATIONSHIPS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";

    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_relationship = matches!(
            namespace,
            ResolveResult::Bound(Namespace(namespace))
                if namespace == RELATIONSHIPS || namespace == STRICT_RELATIONSHIPS
        );
        if !is_relationship {
            continue;
        }
        if value.is_some() {
            return Err(Error::Invalid(
                "duplicate frame source relationship ID".into(),
            ));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    value.ok_or_else(|| Error::Invalid("frame source relationship ID is required".into()))
}

fn finish_leaf(
    reader: &mut NsReader<&[u8]>,
    expected_name: &[u8],
    description: &str,
    budget: &mut ParseBudget,
) -> Result<()> {
    loop {
        budget.event()?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == expected_name =>
            {
                return Ok(());
            },
            Event::End(_) => {
                return Err(invalid(format!(
                    "Word {description} has a mismatched closing element"
                )));
            },
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => {
                return Err(Error::Invalid(format!(
                    "unterminated Word {description} element"
                )));
            },
            _ => {
                return Err(Error::Invalid(format!(
                    "Word {description} must not contain child content"
                )));
            },
        }
    }
}

fn skip_element(reader: &mut NsReader<&[u8]>, budget: &mut ParseBudget) -> Result<()> {
    let mut depth = 1usize;
    loop {
        budget.event()?;
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("Word web XML nesting is too deep".into()))?;
                if depth > MAX_FRAMESET_NESTING {
                    return Err(invalid("Word web XML nesting exceeds the safety limit"));
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Invalid("invalid Word web XML nesting".into()))?;
                if depth == 0 {
                    return Ok(());
                }
            },
            Event::Eof => {
                return Err(Error::Invalid("unterminated Word web XML element".into()));
            },
            _ => {},
        }
    }
}

fn required_value(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<String> {
    word_attribute_value(element, b"val", decoder, resolver)?
        .ok_or_else(|| Error::Invalid(format!("Word {description} value is required")))
}

fn set_on_off(
    slot: &mut Option<bool>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<()> {
    let value = match word_attribute_value(element, b"val", decoder, resolver)? {
        Some(value) => match value.as_str() {
            "true" | "1" | "on" => true,
            "false" | "0" | "off" => false,
            _ => {
                return Err(Error::Invalid(format!(
                    "invalid Word on/off value '{value}'"
                )));
            },
        },
        None => true,
    };
    set_once(slot, value, description)
}

fn set_once<T>(slot: &mut Option<T>, value: T, description: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(Error::Invalid(format!(
            "duplicate Word web setting '{description}'"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::packuri::PackURI;
    use litchi_opc::part::BlobPart;

    fn parse_settings(xml: &[u8]) -> Result<Settings> {
        parse(xml).map(|(settings, _)| settings)
    }

    fn read_settings_part(part: &dyn Part) -> Result<Settings> {
        read(part).map(|(settings, _)| settings)
    }

    fn contains_xml(xml: &[u8], needle: &str) -> bool {
        xml.windows(needle.len())
            .any(|window| window == needle.as_bytes())
    }

    fn id(value: i64) -> Id {
        Id::new(value).unwrap()
    }

    fn make_div(value: i64) -> Div {
        Div::new(id(value))
    }

    fn package(conformance: Conformance) -> OpcPackage {
        use litchi_opc::constants::{content_type as ct, relationship_type as rt};

        let mut package = OpcPackage::new();
        let document = PackURI::new("/word/document.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            document,
            ct::WML_DOCUMENT_MAIN.to_owned(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#.to_vec(),
        )));
        let relationship = match conformance {
            Conformance::Transitional => rt::OFFICE_DOCUMENT,
            Conformance::Strict => STRICT_OFFICE_DOCUMENT_RELATIONSHIP,
        };
        package.relate_to("word/document.xml", relationship);
        package
    }

    fn add_raw_web(package: &mut OpcPackage, xml: &[u8], conformance: Conformance) -> PackURI {
        let name = PackURI::new("/word/webSettings.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            name.clone(),
            CONTENT_TYPE.to_owned(),
            xml.to_vec(),
        )));
        package
            .get_part_mut(&PackURI::new("/word/document.xml").unwrap())
            .unwrap()
            .relate_to("webSettings.xml", conformance.relationship());
        name
    }

    #[test]
    fn semantic_and_positional_div_crud_is_checked_and_atomic() {
        let mut settings = Settings::default();
        let mut first = make_div(1);
        first.set_body_div(true);
        settings.add(first).unwrap();
        assert_eq!(settings.get(id(1)).unwrap().unwrap().id(), id(1));
        assert_eq!(settings.get(0).unwrap().unwrap().id(), id(1));

        let before = settings.clone();
        assert!(settings.add(make_div(1)).is_err());
        assert_eq!(settings, before);
        assert!(settings.move_to(id(1), 1).is_err());
        assert_eq!(settings, before);

        let mut replacement = make_div(1);
        replacement.set_block_quote(true);
        let old = settings.put(replacement).unwrap().unwrap();
        assert_eq!(old.is_body_div(), Some(true));
        assert_eq!(
            settings.get(id(1)).unwrap().unwrap().is_block_quote(),
            Some(true)
        );

        settings.add(make_div(2)).unwrap();
        settings.move_to(id(2), 0).unwrap();
        assert_eq!(settings.get(0).unwrap().unwrap().id(), id(2));
        assert_eq!(settings.remove(id(1)).unwrap().unwrap().id(), id(1));
        assert!(settings.remove(id(3)).unwrap().is_none());
        assert_eq!(settings.remove(0).unwrap().unwrap().id(), id(2));
        assert!(settings.divs().is_none());
    }

    #[test]
    fn numeric_div_selectors_reject_missing_positions_without_mutation() {
        let mut settings = Settings::default();
        assert!(settings.get(id(99)).unwrap().is_none());
        assert!(settings.remove(id(99)).unwrap().is_none());
        assert!(settings.get(0usize).is_err());
        assert!(settings.remove(0usize).is_err());

        settings.add(make_div(1)).unwrap();
        let before = settings.clone();
        assert!(settings.get(1usize).is_err());
        assert!(settings.remove(1usize).is_err());
        assert_eq!(settings, before);

        let mut parent = make_div(10);
        assert!(parent.child(id(99)).unwrap().is_none());
        assert!(parent.remove_child(id(99)).unwrap().is_none());
        assert!(parent.child(0usize).is_err());
        assert!(parent.remove_child(0usize).is_err());

        parent.add_child(make_div(11)).unwrap();
        assert_eq!(parent.child(0usize).unwrap().unwrap().id(), id(11));
        let before = parent.clone();
        assert!(parent.child(1usize).is_err());
        assert!(parent.remove_child(1usize).is_err());
        assert_eq!(parent, before);
    }

    #[test]
    fn package_graph_crud_round_trips_and_is_idempotent() {
        let mut package = package(Conformance::Transitional);
        assert!(load(&package).unwrap().is_none());
        assert!(!remove(&mut package).unwrap());

        let mut settings = Settings::default();
        settings.set_encoding("utf-8").unwrap().set_allow_png(true);
        assert!(put(&mut package, settings.clone(), Conformance::Transitional).unwrap());
        assert_eq!(
            load(&package).unwrap(),
            Some((settings.clone(), Conformance::Transitional))
        );
        assert!(!put(&mut package, settings.clone(), Conformance::Transitional).unwrap());

        settings.set_allow_png(false);
        assert!(put(&mut package, settings, Conformance::Transitional).unwrap());
        assert!(remove(&mut package).unwrap());
        assert!(load(&package).unwrap().is_none());
        assert!(!remove(&mut package).unwrap());
    }

    #[test]
    fn semantic_noop_preserves_noncanonical_bytes_and_signatures() {
        use litchi_opc::constants::relationship_type as rt;

        let mut package = package(Conformance::Transitional);
        let source = br#"<?xml version="1.0"?>
          <w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:encoding w:val="utf-8"></w:encoding>
            <w:allowPNG w:val="1" />
          </w:webSettings>"#;
        let name = add_raw_web(&mut package, source, Conformance::Transitional);
        package.rels_mut().add_relationship(
            rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rSignature".to_owned(),
            false,
        );
        let (settings, conformance) = load(&package).unwrap().unwrap();
        assert!(package.is_signed());

        assert!(put(&mut package, settings.clone(), Conformance::Strict).is_err());
        assert_eq!(package.get_part(&name).unwrap().blob(), source);
        assert!(package.is_signed());
        assert!(!put(&mut package, settings, conformance).unwrap());
        assert_eq!(package.get_part(&name).unwrap().blob(), source);
        assert!(package.is_signed());
    }

    #[test]
    fn strict_graph_and_mce_round_trip() {
        let mut package = package(Conformance::Strict);
        let mut settings = Settings::default();
        settings.set_target_screen_size(Screen::Pixels1920x1200);
        assert!(put(&mut package, settings.clone(), Conformance::Strict).unwrap());
        assert_eq!(
            load(&package).unwrap(),
            Some((settings, Conformance::Strict))
        );

        let xml = br#"<w:webSettings xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:x="urn:unsupported" mc:Ignorable="x">
            <x:ignored/><w:allowPNG/>
        </w:webSettings>"#;
        let (parsed, conformance) = parse(xml).unwrap();
        assert_eq!(conformance, Conformance::Strict);
        assert_eq!(parsed.allow_png(), Some(true));
    }

    #[test]
    fn graph_failures_are_atomic_and_shared_parts_are_rejected() {
        let mut duplicate = package(Conformance::Transitional);
        let document = PackURI::new("/word/document.xml").unwrap();
        {
            let part = duplicate.get_part_mut(&document).unwrap();
            part.rels_mut().add_relationship(
                Conformance::Transitional.relationship().to_owned(),
                "webSettings.xml".to_owned(),
                "rWeb1".to_owned(),
                false,
            );
            part.rels_mut().add_relationship(
                Conformance::Transitional.relationship().to_owned(),
                "webSettings2.xml".to_owned(),
                "rWeb2".to_owned(),
                false,
            );
        }
        let parts = duplicate.part_count();
        let relationships = duplicate.get_part(&document).unwrap().rels().iter().count();
        assert!(
            put(
                &mut duplicate,
                Settings::default(),
                Conformance::Transitional
            )
            .is_err()
        );
        assert_eq!(duplicate.part_count(), parts);
        assert_eq!(
            duplicate.get_part(&document).unwrap().rels().iter().count(),
            relationships
        );

        let mut shared = package(Conformance::Transitional);
        assert!(put(&mut shared, Settings::default(), Conformance::Transitional).unwrap());
        let name = PackURI::new("/word/webSettings.xml").unwrap();
        let bytes = shared.get_part(&name).unwrap().blob().to_vec();
        let mut other = BlobPart::new(
            PackURI::new("/word/other.xml").unwrap(),
            "application/xml".to_owned(),
            b"<other/>".to_vec(),
        );
        other.rels_mut().add_relationship(
            "urn:shared".to_owned(),
            "webSettings.xml".to_owned(),
            "rShared".to_owned(),
            false,
        );
        shared.add_part(Box::new(other));
        assert!(remove(&mut shared).is_err());
        assert_eq!(shared.get_part(&name).unwrap().blob(), bytes);
        assert!(load(&shared).unwrap().is_some());
    }

    #[test]
    fn adversarial_xml_never_unwinds() {
        for xml in [
            b"".as_slice(),
            b"<w:webSettings".as_slice(),
            b"<webSettings/>".as_slice(),
            b"\xFF\xFE".as_slice(),
            br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:divs><w:div w:id="same"/><w:div w:id="same"/></w:divs></w:webSettings>"#.as_slice(),
        ] {
            let result = std::panic::catch_unwind(|| parse(xml));
            assert!(result.is_ok());
            assert!(result.unwrap().is_err());
        }
        let oversized = vec![b' '; MAX_XML_BYTES + 1];
        assert!(parse(&oversized).is_err());
    }

    #[test]
    fn parser_budget_covers_nested_readers() {
        let mut xml = String::from(
            r#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset>"#,
        );
        xml.try_reserve(MAX_XML_EVENTS * 8).unwrap();
        for _ in 0..=MAX_XML_EVENTS {
            xml.push_str("<!--x-->");
        }
        xml.push_str("</w:frameset></w:webSettings>");

        let result = std::panic::catch_unwind(|| Settings::parse_xml(xml.as_bytes()));
        assert!(result.is_ok());
        assert!(result.unwrap().is_err());
    }

    #[test]
    fn mce_preprocessing_respects_web_settings_xml_bound() {
        let prefix = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">"#;
        let suffix = b"</w:webSettings>";
        let mut xml = Vec::with_capacity(MAX_XML_BYTES + suffix.len() + 1);
        xml.extend_from_slice(prefix);
        xml.resize(MAX_XML_BYTES + 1, b' ');
        xml.extend_from_slice(suffix);

        assert!(matches!(
            parse(&xml),
            Err(Error::Mce(litchi_ooxml_common::MceError::LimitExceeded(_)))
        ));
    }

    #[test]
    fn rejects_mismatched_leaf_end_tags() {
        const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        for xml in [
            format!(r#"<w:webSettings xmlns:w="{W}"><w:allowPNG></w:encoding></w:webSettings>"#),
            format!(
                r#"<w:webSettings xmlns:w="{W}"><w:frameset><w:sz w:val="1*"></w:name></w:frameset></w:webSettings>"#
            ),
            format!(
                r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1"><w:marLeft w:val="0"></w:marRight><w:marRight w:val="0"/><w:marTop w:val="0"/><w:marBottom w:val="0"/></w:div></w:divs></w:webSettings>"#
            ),
        ] {
            assert!(parse_settings(xml.as_bytes()).is_err(), "accepted {xml}");
        }
    }

    #[test]
    fn parses_all_scalar_web_settings_with_strict_namespaces() {
        let xml = br#"<s:webSettings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"
            xmlns:false="urn:not-wordprocessingml">
            <s:encoding s:val="utf-8"/>
            <s:optimizeForBrowser s:val="on"/>
            <s:allowPNG/>
            <s:doNotRelyOnCSS s:val="0"/>
            <s:doNotSaveAsSingleFile s:val="1"/>
            <s:doNotOrganizeInFolder s:val="false"/>
            <s:doNotUseLongFileNames s:val="true"/>
            <s:pixelsPerInch s:val=" 1023 "/>
            <s:targetScreenSz s:val="1920x1200"/>
            <s:saveSmartTagsAsXml s:val="on"/>
            <false:saveSmartTagsAsXml false:val="off"/>
        </s:webSettings>"#;

        let settings = parse_settings(xml).unwrap();
        assert_eq!(settings.encoding(), Some("utf-8"));
        assert_eq!(settings.optimize_for_browser(), Some(true));
        assert_eq!(settings.rely_on_vml(), None);
        assert_eq!(settings.allow_png(), Some(true));
        assert_eq!(settings.do_not_rely_on_css(), Some(false));
        assert_eq!(settings.do_not_save_as_single_file(), Some(true));
        assert_eq!(settings.do_not_organize_in_folder(), Some(false));
        assert_eq!(settings.do_not_use_long_file_names(), Some(true));
        assert_eq!(settings.pixels_per_inch(), Some(1023));
        assert_eq!(settings.target_screen_size(), Some(Screen::Pixels1920x1200));
        assert_eq!(settings.save_smart_tags_as_xml(), Some(true));
    }

    #[test]
    fn rejects_invalid_or_duplicate_scalar_web_settings() {
        let missing_value = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pixelsPerInch/></w:webSettings>"#;
        assert!(parse_settings(missing_value).is_err());

        let invalid_on_off = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:saveSmartTagsAsXml w:val="maybe"/></w:webSettings>"#;
        assert!(parse_settings(invalid_on_off).is_err());

        let duplicate = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:allowPNG/><w:allowPNG/></w:webSettings>"#;
        assert!(parse_settings(duplicate).is_err());

        let invalid_screen = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:targetScreenSz w:val="1366x768"/></w:webSettings>"#;
        assert!(parse_settings(invalid_screen).is_err());

        let excessive_pixels = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pixelsPerInch w:val="1024"/></w:webSettings>"#;
        assert!(parse_settings(excessive_pixels).is_err());

        let strict_rely = br#"<w:webSettings xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:relyOnVML/></w:webSettings>"#;
        assert!(parse_settings(strict_rely).is_err());

        let out_of_order = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:allowPNG/><w:encoding w:val="utf-8"/></w:webSettings>"#;
        assert!(parse_settings(out_of_order).is_err());

        let nested_scalar = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:allowPNG><w:doNotRelyOnCSS/></w:allowPNG></w:webSettings>"#;
        assert!(parse_settings(nested_scalar).is_err());

        let scalar_text = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:encoding w:val="utf-8">unexpected</w:encoding></w:webSettings>"#;
        assert!(parse_settings(scalar_text).is_err());
    }

    #[test]
    fn parses_recursive_framesets_and_all_frame_properties() {
        let xml = br#"<s:webSettings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"
            xmlns:rel="http://purl.oclc.org/ooxml/officeDocument/relationships"
            xmlns:false="urn:not-wordprocessingml">
          <s:frameset>
            <s:sz s:val="2*"/>
            <s:framesetSplitbar>
              <s:w s:val="90"/>
              <s:color s:val="auto" s:themeColor="accent2" s:themeTint="7f" s:themeShade="00"/>
              <s:noBorder s:val="off"/>
              <s:flatBorders/>
            </s:framesetSplitbar>
            <s:frameLayout s:val="cols"/>
            <s:frame>
              <s:sz s:val="50%"/>
              <s:name s:val="navigation"/>
              <s:sourceFileName rel:id="rId7"/>
              <s:marW s:val="18446744073709551615"/>
              <s:marH s:val="24"/>
              <s:scrollbar s:val="auto"/>
              <s:noResizeAllowed/>
              <s:linkedToFile s:val="false"/>
              <s:futureExtension><s:nested/></s:futureExtension>
            </s:frame>
            <s:frameset>
              <s:frameLayout s:val="none"/>
              <s:frame><s:name s:val="content"/></s:frame>
            </s:frameset>
            <false:frame><false:name false:val="ignored"/></false:frame>
          </s:frameset>
        </s:webSettings>"#;

        let settings = parse_settings(xml).unwrap();
        let frameset = settings.frameset().unwrap();
        assert_eq!(frameset.size(), Some("2*"));
        assert_eq!(frameset.layout(), Some(Layout::Columns));
        let split_bar = frameset.split_bar().unwrap();
        assert_eq!(split_bar.width_twips(), Some(90));
        assert_eq!(split_bar.no_border(), Some(false));
        assert_eq!(split_bar.flat_borders(), Some(true));
        let color = split_bar.color().unwrap();
        assert_eq!(color.value(), "auto");
        assert_eq!(color.theme_color(), Some(Theme::Accent2));
        assert_eq!(color.theme_tint(), Some(0x7f));
        assert_eq!(color.theme_shade(), Some(0));
        assert_eq!(frameset.children().len(), 2);

        let Child::Frame(frame) = &frameset.children()[0] else {
            panic!("first frameset child must be a frame");
        };
        assert_eq!(frame.size(), Some("50%"));
        assert_eq!(frame.name(), Some("navigation"));
        assert_eq!(frame.rel(), Some("rId7"));
        assert_eq!(frame.margin_width(), Some(u64::MAX));
        assert_eq!(frame.margin_height(), Some(24));
        assert_eq!(frame.scrollbar(), Some(Scrollbar::Auto));
        assert_eq!(frame.no_resize_allowed(), Some(true));
        assert_eq!(frame.linked_to_file(), Some(false));

        let Child::Frameset(nested) = &frameset.children()[1] else {
            panic!("second frameset child must be nested");
        };
        assert_eq!(nested.layout(), Some(Layout::None));
        let Child::Frame(frame) = &nested.children()[0] else {
            panic!("nested child must be a frame");
        };
        assert_eq!(frame.name(), Some("content"));
    }

    #[test]
    fn validates_frame_values_and_source_relationships() {
        let invalid_layout = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset><w:frameLayout w:val="diagonal"/></w:frameset></w:webSettings>"#;
        assert!(parse_settings(invalid_layout).is_err());

        let overflowing_pixels = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset><w:frame><w:marW w:val="18446744073709551616"/></w:frame></w:frameset></w:webSettings>"#;
        assert!(parse_settings(overflowing_pixels).is_err());

        let child_in_leaf = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset><w:frame><w:name w:val="bad"><w:frame/></w:name></w:frame></w:frameset></w:webSettings>"#;
        assert!(parse_settings(child_in_leaf).is_err());

        let duplicate = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset><w:frame><w:name w:val="one"/><w:name w:val="two"/></w:frame></w:frameset></w:webSettings>"#;
        assert!(parse_settings(duplicate).is_err());

        let xml = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:frameset><w:frame><w:sourceFileName r:id="rId1"/></w:frame></w:frameset></w:webSettings>"#;
        let mut part = BlobPart::new(
            PackURI::new("/word/webSettings.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml".into(),
            xml.to_vec(),
        );
        assert!(read_settings_part(&part).is_err());
        part.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame".into(),
            "https://example.test/frame.html".into(),
            "rId1".into(),
            true,
        );
        assert!(read_settings_part(&part).is_ok());
    }

    #[test]
    fn parses_recursive_html_divisions_and_border_properties() {
        let xml = br#"<s:webSettings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:false="urn:not-wordprocessingml">
          <s:divs>
            <s:div s:id="1785730240">
              <s:blockQuote/>
              <s:bodyDiv s:val="off"/>
              <s:marLeft s:val=" -1234567890 "/>
              <s:marRight s:val="+42"/>
              <s:marTop s:val="0"/>
              <s:marBottom s:val="700"/>
              <s:divBdr>
                <s:top s:val="single" s:color="A0b1C2" s:themeColor="text2" s:themeTint="10" s:themeShade="ff" s:sz="18446744073709551615" s:space="6" s:shadow="on" s:frame="0"/>
                <s:left s:val="zigZagStitch"/>
              </s:divBdr>
              <s:divsChild>
                <s:div s:id="1785730241"><s:bodyDiv/><s:marLeft s:val="0"/><s:marRight s:val="0"/><s:marTop s:val="0"/><s:marBottom s:val="0"/></s:div>
              </s:divsChild>
              <s:divsChild><s:div s:id="1785730242"><s:marLeft s:val="1"/><s:marRight s:val="2"/><s:marTop s:val="3"/><s:marBottom s:val="4"/></s:div></s:divsChild>
            </s:div>
            <s:div s:id="1785730243"><s:marLeft s:val="0"/><s:marRight s:val="0"/><s:marTop s:val="0"/><s:marBottom s:val="0"/></s:div>
            <false:div false:id="ignored"/>
          </s:divs>
        </s:webSettings>"#;

        let settings = parse_settings(xml).unwrap();
        let divs = settings.divs().unwrap();
        assert_eq!(divs.len(), 2);
        let div = &divs[0];
        assert_eq!(div.id(), id(1785730240));
        assert_eq!(div.is_block_quote(), Some(true));
        assert_eq!(div.is_body_div(), Some(false));
        assert_eq!(div.left(), Twips::new(-1234567890));
        assert_eq!(div.right(), Twips::new(42));
        assert_eq!(div.top(), Twips::new(0));
        assert_eq!(div.bottom(), Twips::new(700));
        assert_eq!(div.children().len(), 2);
        assert_eq!(div.children()[0].id(), id(1785730241));
        assert_eq!(div.children()[0].is_body_div(), Some(true));
        assert_eq!(div.children()[1].id(), id(1785730242));

        let borders = div.borders().unwrap();
        let top = borders.top().unwrap();
        assert_eq!(top.style(), "single");
        assert_eq!(top.color(), Some("A0b1C2"));
        assert_eq!(top.theme_color(), Some(Theme::Text2));
        assert_eq!(top.theme_tint(), Some(0x10));
        assert_eq!(top.theme_shade(), Some(0xff));
        assert_eq!(top.size_eighth_points(), Some(u64::MAX));
        assert_eq!(top.space_points(), Some(6));
        assert_eq!(top.shadow(), Some(true));
        assert_eq!(top.frame(), Some(false));
        assert_eq!(borders.left().unwrap().style(), "zigZagStitch");
        assert!(borders.bottom().is_none());
        assert!(borders.right().is_none());
    }

    #[test]
    fn validates_html_division_structure_and_values() {
        const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        let margins = r#"<w:marLeft w:val="0"/><w:marRight w:val="0"/><w:marTop w:val="0"/><w:marBottom w:val="0"/>"#;
        let missing_id = format!(
            r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div>{margins}</w:div></w:divs></w:webSettings>"#
        );
        assert!(parse_settings(missing_id.as_bytes()).is_err());

        for invalid_id in ["0", "-0", "not-a-number", "9223372036854775808"] {
            let xml = format!(
                r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="{invalid_id}">{margins}</w:div></w:divs></w:webSettings>"#
            );
            assert!(parse_settings(xml.as_bytes()).is_err());
        }

        let missing_margin = format!(
            r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1"><w:marLeft w:val="0"/><w:marRight w:val="0"/><w:marTop w:val="0"/></w:div></w:divs></w:webSettings>"#
        );
        assert!(parse_settings(missing_margin.as_bytes()).is_err());

        let invalid_margin = format!(
            r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1"><w:marLeft w:val="1.5"/><w:marRight w:val="0"/><w:marTop w:val="0"/><w:marBottom w:val="0"/></w:div></w:divs></w:webSettings>"#
        );
        assert!(parse_settings(invalid_margin.as_bytes()).is_err());

        let invalid_color = format!(
            r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1">{margins}<w:divBdr><w:left w:val="single" w:color="xyz"/></w:divBdr></w:div></w:divs></w:webSettings>"#
        );
        assert!(parse_settings(invalid_color.as_bytes()).is_err());

        let empty_child_container = format!(
            r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1">{margins}<w:divsChild/></w:div></w:divs></w:webSettings>"#
        );
        assert!(parse_settings(empty_child_container.as_bytes()).is_err());

        let empty_divs = format!(r#"<w:webSettings xmlns:w="{W}"><w:divs/></w:webSettings>"#);
        assert!(parse_settings(empty_divs.as_bytes()).is_err());

        let out_of_order = format!(
            r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1"><w:marRight w:val="0"/><w:marLeft w:val="0"/><w:marTop w:val="0"/><w:marBottom w:val="0"/></w:div></w:divs></w:webSettings>"#
        );
        assert!(parse_settings(out_of_order.as_bytes()).is_err());
    }

    #[test]
    fn serializes_every_modeled_web_setting_for_round_trip() {
        let xml = br#"<w:webSettings
          xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:frameset>
            <w:sz w:val="2* &amp; 1*"/>
            <w:framesetSplitbar>
              <w:w w:val="18446744073709551615"/>
              <w:color w:val="A0b1C2" w:themeColor="accent4" w:themeTint="0a" w:themeShade="FF"/>
              <w:noBorder/>
              <w:flatBorders w:val="false"/>
            </w:framesetSplitbar>
            <w:frameLayout w:val="cols"/>
            <w:frame>
              <w:sz w:val="50%"/>
              <w:name w:val="main &amp; detail"/>
              <w:sourceFileName r:id="rId7"/>
              <w:marW w:val="42"/>
              <w:marH w:val="24"/>
              <w:scrollbar w:val="auto"/>
              <w:noResizeAllowed w:val="off"/>
              <w:linkedToFile/>
            </w:frame>
            <w:frameset><w:frameLayout w:val="none"/></w:frameset>
          </w:frameset>
          <w:divs>
            <w:div w:id="1">
              <w:blockQuote/>
              <w:bodyDiv w:val="0"/>
              <w:marLeft w:val="-1234567890"/>
              <w:marRight w:val="+42"/>
              <w:marTop w:val="0"/>
              <w:marBottom w:val="700"/>
              <w:divBdr>
                <w:top w:val="single" w:color="auto" w:themeColor="text2" w:themeTint="10" w:themeShade="ff" w:sz="18446744073709551615" w:space="6" w:shadow="on" w:frame="0"/>
                <w:left w:val="zigZagStitch"/>
              </w:divBdr>
              <w:divsChild><w:div w:id="2"><w:bodyDiv/><w:marLeft w:val="0"/><w:marRight w:val="0"/><w:marTop w:val="0"/><w:marBottom w:val="0"/></w:div></w:divsChild>
            </w:div>
          </w:divs>
          <w:encoding w:val="utf-8"/>
          <w:optimizeForBrowser/>
          <w:relyOnVML w:val="false"/>
          <w:allowPNG/>
          <w:doNotRelyOnCSS w:val="off"/>
          <w:doNotSaveAsSingleFile/>
          <w:doNotOrganizeInFolder w:val="0"/>
          <w:doNotUseLongFileNames/>
          <w:pixelsPerInch w:val="1023"/>
          <w:targetScreenSz w:val="1920x1200"/>
          <w:saveSmartTagsAsXml w:val="false"/>
        </w:webSettings>"#;

        let settings = parse_settings(xml).unwrap();
        let serialized = settings.xml(Conformance::Transitional).unwrap();
        let reparsed = parse_settings(&serialized).unwrap();

        assert_eq!(reparsed, settings);
        assert!(contains_xml(&serialized, "main &amp; detail"));
        assert!(contains_xml(&serialized, "w:themeTint=\"0A\""));
        assert!(contains_xml(&serialized, "w:themeShade=\"FF\""));
        assert!(contains_xml(&serialized, "<w:blockQuote w:val=\"1\"/>"));
        assert!(contains_xml(&serialized, "<w:bodyDiv w:val=\"0\"/>"));
        assert!(contains_xml(&serialized, "<w:bodyDiv w:val=\"1\"/>"));

        let mut strict_settings = settings.clone();
        strict_settings.clear_rely_on_vml();
        let strict = strict_settings.xml(Conformance::Strict).unwrap();
        assert!(contains_xml(&strict, "<w:blockQuote w:val=\"1\"/>"));
        assert!(contains_xml(&strict, "<w:bodyDiv w:val=\"0\"/>"));
        assert!(contains_xml(&strict, "<w:bodyDiv w:val=\"1\"/>"));
    }

    #[test]
    fn edits_and_clears_every_scalar_web_setting() {
        let mut settings = Settings::default();
        settings
            .set_encoding("utf-8")
            .unwrap()
            .set_optimize_for_browser(true)
            .set_rely_on_vml(false)
            .set_allow_png(true)
            .set_do_not_rely_on_css(false)
            .set_do_not_save_as_single_file(true)
            .set_do_not_organize_in_folder(false)
            .set_do_not_use_long_file_names(true);
        settings
            .set_pixels_per_inch(96)
            .unwrap()
            .set_target_screen_size(Screen::Pixels1800x1440)
            .set_save_smart_tags_as_xml(false);

        let serialized = settings.xml(Conformance::Transitional).unwrap();
        let reparsed = parse_settings(&serialized).unwrap();
        assert_eq!(reparsed, settings);
        assert_eq!(reparsed.encoding(), Some("utf-8"));
        assert_eq!(reparsed.pixels_per_inch(), Some(96));
        assert_eq!(reparsed.target_screen_size(), Some(Screen::Pixels1800x1440));

        let previous_pixels = settings.pixels_per_inch().unwrap();
        assert!(settings.set_pixels_per_inch(1024).is_err());
        assert_eq!(settings.pixels_per_inch(), Some(previous_pixels));

        assert!(settings.xml(Conformance::Strict).is_err());

        settings
            .clear_encoding()
            .clear_optimize_for_browser()
            .clear_rely_on_vml()
            .clear_allow_png()
            .clear_do_not_rely_on_css()
            .clear_do_not_save_as_single_file()
            .clear_do_not_organize_in_folder()
            .clear_do_not_use_long_file_names()
            .clear_pixels_per_inch()
            .clear_target_screen_size()
            .clear_save_smart_tags_as_xml();
        assert_eq!(settings, Settings::default());
        assert_eq!(
            parse_settings(&settings.xml(Conformance::Transitional).unwrap()).unwrap(),
            Settings::default()
        );
    }

    #[test]
    fn builds_and_edits_recursive_framesets_for_round_trip() {
        let mut color = Color::new("A0b1C2").unwrap();
        color
            .set_theme_color(Theme::Accent4)
            .set_theme_tint(0x0a)
            .set_theme_shade(0xff);

        let mut split_bar = SplitBar::default();
        split_bar
            .set_width_twips(u64::MAX)
            .set_color(color)
            .set_no_border(true)
            .set_flat_borders(false);

        let mut frameset = Frameset::default();
        frameset
            .set_size("2* & 1*")
            .unwrap()
            .set_split_bar(split_bar)
            .set_layout(Layout::Columns);
        let frame = frameset.add_frame().unwrap();
        frame.set_size("50%").unwrap();
        frame.set_name("main & detail").unwrap();
        frame
            .set_rel("rId7")
            .unwrap()
            .set_margin_width(42)
            .set_margin_height(24)
            .set_scrollbar(Scrollbar::Auto)
            .set_no_resize_allowed(false)
            .set_linked_to_file(true);
        let nested = frameset.add_frameset().unwrap();
        nested.set_size("1*").unwrap().set_layout(Layout::None);
        nested.add_frame().unwrap().set_name("nested").unwrap();

        let mut settings = Settings::default();
        settings.set_frameset(frameset);
        let serialized = settings.xml(Conformance::Transitional).unwrap();
        let reparsed = parse_settings(&serialized).unwrap();
        assert_eq!(reparsed, settings);
        assert!(contains_xml(&serialized, "main &amp; detail"));

        let frameset = settings.frameset_mut().unwrap();
        assert_eq!(frameset.children().len(), 2);
        assert!(matches!(frameset.children()[0], Child::Frame(_)));
        assert!(matches!(frameset.children()[1], Child::Frameset(_)));
        frameset
            .clear_size()
            .clear_split_bar()
            .clear_layout()
            .clear_children();
        assert_eq!(frameset, &Frameset::default());
        settings.clear_frameset();
        assert!(settings.frameset().is_none());
    }

    #[test]
    fn validates_mutable_frameset_colors_without_losing_prior_value() {
        assert!(Color::new("12345").is_err());
        let mut color = Color::new("auto").unwrap();
        assert!(color.set_value("GG0000").is_err());
        assert_eq!(color.value(), "auto");
        color.set_value("00ffAA").unwrap();
        assert_eq!(color.value(), "00ffAA");
        color
            .set_theme_color(Theme::Text1)
            .set_theme_tint(1)
            .set_theme_shade(2)
            .clear_theme_color()
            .clear_theme_tint()
            .clear_theme_shade();
        assert_eq!(color.theme_color(), None);
        assert_eq!(color.theme_tint(), None);
        assert_eq!(color.theme_shade(), None);
    }

    #[test]
    fn builds_and_edits_recursive_html_divisions_for_round_trip() {
        let mut top = Border::new("single").unwrap();
        top.set_color("A0b1C2")
            .unwrap()
            .set_theme_color(Theme::Text2)
            .set_theme_tint(0x10)
            .set_theme_shade(0xff)
            .set_size_eighth_points(u64::MAX)
            .set_space_points(6)
            .set_shadow(true)
            .set_frame(false);
        let mut borders = Borders::default();
        borders
            .set_top(top)
            .set_left(Border::new("zigZagStitch").unwrap())
            .set_bottom(Border::new("double").unwrap())
            .set_right(Border::new("nil").unwrap());

        let mut div = make_div(1);
        div.set_block_quote(true)
            .set_body_div(false)
            .set_left(-1_234_567_890)
            .set_right(42)
            .set_top(0)
            .set_bottom(700)
            .set_borders(borders);
        let mut grandchild = make_div(3);
        grandchild.set_block_quote(false);
        let mut child = make_div(2);
        child.set_body_div(true);
        child.add_child(grandchild).unwrap();
        div.add_child(child).unwrap();

        let mut settings = Settings::default();
        settings.set_divs(vec![div]).unwrap();
        let serialized = settings.xml(Conformance::Transitional).unwrap();
        let reparsed = parse_settings(&serialized).unwrap();
        assert_eq!(reparsed, settings);
        assert!(contains_xml(&serialized, "w:id=\"1\""));
        assert_eq!(
            reparsed.divs().unwrap()[0].left(),
            Twips::new(-1_234_567_890)
        );
        let left = serialized
            .windows(b"<w:marLeft".len())
            .position(|window| window == b"<w:marLeft")
            .unwrap();
        let right = serialized
            .windows(b"<w:marRight".len())
            .position(|window| window == b"<w:marRight")
            .unwrap();
        let top = serialized
            .windows(b"<w:marTop".len())
            .position(|window| window == b"<w:marTop")
            .unwrap();
        let bottom = serialized
            .windows(b"<w:marBottom".len())
            .position(|window| window == b"<w:marBottom")
            .unwrap();
        assert!(left < right && right < top && top < bottom);

        settings.add(make_div(4)).unwrap();
        assert_eq!(settings.divs().unwrap().len(), 2);
        assert_eq!(settings.get(id(4)).unwrap().unwrap().id(), id(4));
        settings.move_to(id(4), 0).unwrap();
        let mut first = settings.remove(id(1)).unwrap().unwrap();
        first
            .clear_block_quote()
            .clear_body_div()
            .set_left(0)
            .set_right(0)
            .set_top(0)
            .set_bottom(0)
            .clear_children();
        let borders = first.borders_mut().unwrap();
        borders
            .clear_top()
            .clear_left()
            .clear_bottom()
            .clear_right();
        first.clear_borders();
        settings.add(first).unwrap();
        settings.clear_divs();
        assert!(settings.divs().is_none());
    }

    #[test]
    fn validates_mutable_html_division_values_atomically() {
        assert!(Id::new(0).is_err());
        assert!(Id::parse("word-id").is_err());
        assert!(Twips::parse("1.5").is_err());
        let mut div = make_div(1);
        div.set_left(-42);
        assert_eq!(div.left(), Twips::new(-42));

        let mut border = Border::new("single").unwrap();
        border.set_color("auto").unwrap();
        assert!(border.set_color("xyz").is_err());
        assert_eq!(border.color(), Some("auto"));
        border
            .set_style("double")
            .unwrap()
            .clear_color()
            .set_theme_color(Theme::Accent1)
            .set_theme_tint(1)
            .set_theme_shade(2)
            .set_size_eighth_points(8)
            .set_space_points(1)
            .set_shadow(false)
            .set_frame(true)
            .clear_theme_color()
            .clear_theme_tint()
            .clear_theme_shade()
            .clear_size_eighth_points()
            .clear_space_points()
            .clear_shadow()
            .clear_frame();
        assert_eq!(border.style(), "double");
        assert_eq!(border.color(), None);
        assert_eq!(border.theme_color(), None);
        assert_eq!(border.theme_tint(), None);
        assert_eq!(border.theme_shade(), None);
        assert_eq!(border.size_eighth_points(), None);
        assert_eq!(border.space_points(), None);
        assert_eq!(border.shadow(), None);
        assert_eq!(border.frame(), None);
    }

    #[test]
    fn serialization_rejects_empty_division_containers() {
        let xml = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset/><w:divs/></w:webSettings>"#;
        assert!(parse_settings(xml).is_err());
        assert!(Settings::default().set_divs(Vec::new()).is_err());
    }

    #[test]
    fn serialization_rejects_excessive_recursive_nesting() {
        let mut frameset = Frameset::default();
        for _ in 0..=MAX_FRAMESET_NESTING {
            frameset = Frameset {
                children: vec![Child::Frameset(frameset)],
                ..Frameset::default()
            };
        }
        let settings = Settings {
            frameset: Some(frameset),
            ..Settings::default()
        };
        assert!(settings.xml(Conformance::Transitional).is_err());

        let mut div = make_div(1);
        for value in 2..=(MAX_FRAMESET_NESTING as i64 + 3) {
            let mut parent = make_div(value);
            parent.children.push(div);
            div = parent;
        }
        let settings = Settings {
            divs: Some(vec![div]),
            ..Settings::default()
        };
        assert!(settings.xml(Conformance::Transitional).is_err());
    }
}
