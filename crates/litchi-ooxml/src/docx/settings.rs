//! Document settings and protection support.
//!
//! This module provides types and methods for accessing document settings
//! and protection status.

use crate::docx::mail_merge::{
    MailMergeSettings, parse_settings_mail_merge, validate_mail_merge_relationships,
};
use crate::docx::namespace::{is_wordprocessing_namespace, word_attribute_value};
use crate::docx::numbering::NumberFormat;
use crate::docx::variables::DocumentVariables;
use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::ops::Range;

const TRANSITIONAL_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";

/// Relationship type required by Word for an attached document template.
pub const ATTACHED_TEMPLATE_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate";
/// Strict OOXML alias accepted when reading existing packages.
pub const STRICT_ATTACHED_TEMPLATE_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/attachedTemplate";
const MAX_ATTACHED_TEMPLATE_TARGET_LEN: usize = 32 * 1024;

/// An inert reference to the external template associated with a document.
///
/// The target is never opened, fetched, normalized, or executed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AttachedTemplate {
    relationship_id: String,
    target_uri: String,
}

impl AttachedTemplate {
    /// Relationship ID used by `w:attachedTemplate` in `settings.xml`.
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// External relationship target exactly as stored in the package.
    pub fn target_uri(&self) -> &str {
        &self.target_uri
    }
}

/// Document settings including protection status.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// if let Some(settings) = doc.settings()? {
///     if settings.is_protected() {
///         println!("Document is protected");
///         println!("Protection type: {:?}", settings.protection_type());
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct DocumentSettings {
    /// Whether document is protected
    protected: bool,
    /// Type of protection
    protection_type: Option<ProtectionType>,
    /// Whether to track revisions
    track_revisions: bool,
    /// Zoom percentage
    zoom_percent: Option<u32>,
    /// Smart-tag type declarations.
    smart_tag_types: Vec<SmartTagType>,
    /// Whether applications should omit embedded smart-tag data when saving.
    do_not_embed_smart_tags: bool,
    /// Inert mail-merge connection and display metadata.
    mail_merge: Option<MailMergeSettings>,
    /// Inert external attached-template reference.
    attached_template: Option<AttachedTemplate>,
    /// On/off compatibility option flags from `w:compat`.
    compatibility_options: Vec<CompatibilityOption>,
    /// `w:compatSetting` triples from `w:compat`.
    compatibility_settings: Vec<CompatibilitySetting>,
    /// Document-level footnote properties from `w:footnotePr`.
    footnote_properties: Option<NoteNumberingProperties>,
    /// Document-level endnote properties from `w:endnotePr`.
    endnote_properties: Option<NoteNumberingProperties>,
    /// Whether applications should recommend write protection (`w:writeProtection`).
    write_protection: bool,
    /// Document view mode from `w:view`.
    view: Option<DocumentView>,
    /// Proofing completion markers from `w:proofState`.
    proofing_state: Option<ProofingState>,
    /// Default tab stop interval in twips from `w:defaultTabStop`.
    default_tab_stop_twips: Option<u32>,
    /// Theme font language defaults from `w:themeFontLang`.
    theme_font_languages: Option<ThemeFontLanguages>,
    /// Theme color slot remapping from `w:clrSchemeMapping`.
    color_scheme_mapping: Option<ColorSchemeMapping>,
}

/// A smart-tag vocabulary declaration from `settings.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SmartTagType {
    namespace_uri: String,
    name: String,
    url: String,
}

impl SmartTagType {
    /// Return the smart-tag vocabulary namespace URI.
    #[inline]
    pub fn namespace_uri(&self) -> &str {
        &self.namespace_uri
    }

    /// Return the smart-tag type name.
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the vocabulary download URL.
    #[inline]
    pub fn url(&self) -> &str {
        &self.url
    }
}

/// `w:compatSetting` name identifying the targeted Word compatibility mode.
pub const COMPATIBILITY_MODE_SETTING_NAME: &str = "compatibilityMode";
/// `w:compatSetting` URI under which Word stores its compatibility settings.
pub const COMPATIBILITY_SETTING_URI: &str = "http://schemas.microsoft.com/office/word";

/// An on/off compatibility option flag from `w:compat` (for example
/// `w:useFELayout`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CompatibilityOption {
    name: String,
    enabled: bool,
}

impl CompatibilityOption {
    /// Local element name of the option (for example `useFELayout`).
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Whether the option is enabled.
    #[inline]
    pub fn is_enabled(&self) -> bool {
        self.enabled
    }
}

/// A `w:compatSetting` name/URI/value triple from `w:compat`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CompatibilitySetting {
    name: String,
    uri: String,
    value: String,
}

impl CompatibilitySetting {
    /// Return the setting name (for example `compatibilityMode`).
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the URI scoping the setting name.
    #[inline]
    pub fn uri(&self) -> &str {
        &self.uri
    }

    /// Return the raw setting value.
    #[inline]
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// Placement of footnote or endnote text (`ST_FtnPos`/`ST_EdnPos`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum NotePosition {
    /// At the bottom of the page.
    PageBottom,
    /// Immediately beneath the page's text.
    BeneathText,
    /// At the end of the section.
    SectionEnd,
    /// At the end of the document.
    DocumentEnd,
    /// A position token outside the standardized value set.
    Other(String),
}

impl NotePosition {
    fn from_xml(value: String) -> Self {
        match value.as_str() {
            "pageBottom" => Self::PageBottom,
            "beneathText" => Self::BeneathText,
            "sectEnd" => Self::SectionEnd,
            "docEnd" => Self::DocumentEnd,
            _ => Self::Other(value),
        }
    }

    /// Get the XML value for this position.
    pub fn as_str(&self) -> &str {
        match self {
            Self::PageBottom => "pageBottom",
            Self::BeneathText => "beneathText",
            Self::SectionEnd => "sectEnd",
            Self::DocumentEnd => "docEnd",
            Self::Other(value) => value,
        }
    }
}

/// Numbering restart behavior for footnotes or endnotes (`w:numRestart`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NoteNumberingRestart {
    /// Numbering continues throughout the document.
    Continuous,
    /// Numbering restarts at each section.
    EachSection,
    /// Numbering restarts at each page.
    EachPage,
}

impl NoteNumberingRestart {
    fn from_xml(value: &str) -> Result<Self> {
        match value {
            "continuous" => Ok(Self::Continuous),
            "eachSect" => Ok(Self::EachSection),
            "eachPage" => Ok(Self::EachPage),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid note numbering restart value '{value}'"
            ))),
        }
    }

    /// Get the XML value for this restart behavior.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Continuous => "continuous",
            Self::EachSection => "eachSect",
            Self::EachPage => "eachPage",
        }
    }
}

/// Document-level footnote or endnote properties from `w:footnotePr` or
/// `w:endnotePr` in `settings.xml`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct NoteNumberingProperties {
    position: Option<NotePosition>,
    format: Option<NumberFormat>,
    start: Option<u32>,
    restart: Option<NoteNumberingRestart>,
}

impl NoteNumberingProperties {
    /// Return the note placement, when specified.
    #[inline]
    pub fn position(&self) -> Option<&NotePosition> {
        self.position.as_ref()
    }

    /// Return the numbering format, when specified (defaults to decimal).
    #[inline]
    pub fn format(&self) -> Option<&NumberFormat> {
        self.format.as_ref()
    }

    /// Return the first note number, when specified.
    #[inline]
    pub fn start(&self) -> Option<u32> {
        self.start
    }

    /// Return the numbering restart behavior, when specified.
    #[inline]
    pub fn restart(&self) -> Option<NoteNumberingRestart> {
        self.restart
    }
}

/// Type of document protection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProtectionType {
    /// No editing allowed
    ReadOnly,
    /// Only comments allowed
    Comments,
    /// Only tracked changes allowed
    TrackedChanges,
    /// Only form fields allowed
    Forms,
}

impl ProtectionType {
    /// Parse protection type from XML value.
    fn from_xml(s: &str) -> Option<Self> {
        match s {
            "readOnly" => Some(Self::ReadOnly),
            "comments" => Some(Self::Comments),
            "trackedChanges" => Some(Self::TrackedChanges),
            "forms" => Some(Self::Forms),
            _ => None,
        }
    }

    /// Get XML value for this protection type.
    pub const fn to_xml(self) -> &'static str {
        match self {
            Self::ReadOnly => "readOnly",
            Self::Comments => "comments",
            Self::TrackedChanges => "trackedChanges",
            Self::Forms => "forms",
        }
    }
}

/// Document view mode from `w:view` (`ST_View`, ECMA-376 section 17.18.102).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentView {
    /// No explicit view is specified.
    None,
    /// Print layout view.
    Print,
    /// Outline view.
    Outline,
    /// Master pages view.
    MasterPages,
    /// Normal (draft) view.
    Normal,
    /// Web layout view.
    Web,
}

impl DocumentView {
    fn from_xml(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "print" => Ok(Self::Print),
            "outline" => Ok(Self::Outline),
            "masterPages" => Ok(Self::MasterPages),
            "normal" => Ok(Self::Normal),
            "web" => Ok(Self::Web),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid document view value '{value}'"
            ))),
        }
    }

    /// Get the XML value for this view mode.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Print => "print",
            Self::Outline => "outline",
            Self::MasterPages => "masterPages",
            Self::Normal => "normal",
            Self::Web => "web",
        }
    }
}

/// Proofing completion marker (`ST_ProofState`, ECMA-376 section 17.18.67).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProofState {
    /// Proofing for the region completed without errors.
    Clean,
    /// The region changed since proofing last ran.
    Dirty,
}

impl ProofState {
    fn from_xml(value: &str) -> Result<Self> {
        match value {
            "clean" => Ok(Self::Clean),
            "dirty" => Ok(Self::Dirty),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid proof state value '{value}'"
            ))),
        }
    }

    /// Get the XML value for this proof state.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Clean => "clean",
            Self::Dirty => "dirty",
        }
    }
}

/// Proofing completion markers from `w:proofState` (ECMA-376 section
/// 17.15.1.66).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ProofingState {
    spelling: Option<ProofState>,
    grammar: Option<ProofState>,
}

impl ProofingState {
    /// Create a proofing state with no markers.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the spelling proofing marker.
    pub fn set_spelling(&mut self, value: Option<ProofState>) -> &mut Self {
        self.spelling = value;
        self
    }

    /// Set the grammar proofing marker.
    pub fn set_grammar(&mut self, value: Option<ProofState>) -> &mut Self {
        self.grammar = value;
        self
    }

    /// Return the spelling proofing marker, when specified.
    #[inline]
    pub fn spelling(&self) -> Option<ProofState> {
        self.spelling
    }

    /// Return the grammar proofing marker, when specified.
    #[inline]
    pub fn grammar(&self) -> Option<ProofState> {
        self.grammar
    }

    /// Serialize a standalone `w:proofState` fragment.
    pub fn to_xml(&self, prefix: &str) -> String {
        let mut xml = format!("<{prefix}:proofState");
        if let Some(spelling) = self.spelling {
            xml.push_str(&format!(" {prefix}:spelling=\"{}\"", spelling.as_str()));
        }
        if let Some(grammar) = self.grammar {
            xml.push_str(&format!(" {prefix}:grammar=\"{}\"", grammar.as_str()));
        }
        xml.push_str("/>");
        xml
    }
}

/// Maximum accepted length of a `w:themeFontLang` language tag. `ST_Lang`
/// (ECMA-376 section 17.18.51) is an unbounded string; this bound keeps
/// hostile documents from forcing unbounded allocations.
pub const MAX_LANGUAGE_TAG_LENGTH: usize = 255;

fn validate_language_tag(value: &str, description: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > MAX_LANGUAGE_TAG_LENGTH
        || value.chars().any(char::is_control)
    {
        return Err(OoxmlError::InvalidFormat(format!(
            "invalid Word {description} language tag '{value}'"
        )));
    }
    Ok(())
}

/// Theme font language defaults from `w:themeFontLang` (ECMA-376 section
/// 17.15.1.91). Each value is an `ST_Lang` BCP 47 tag or LCID string.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ThemeFontLanguages {
    latin: Option<String>,
    east_asia: Option<String>,
    bidi: Option<String>,
}

impl ThemeFontLanguages {
    /// Create theme font language defaults with no languages set.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the Latin (`w:val`) theme language.
    pub fn set_latin(&mut self, value: Option<String>) -> Result<&mut Self> {
        if let Some(tag) = value.as_deref() {
            validate_language_tag(tag, "Latin theme font")?;
        }
        self.latin = value;
        Ok(self)
    }

    /// Set the East Asian (`w:eastAsia`) theme language.
    pub fn set_east_asia(&mut self, value: Option<String>) -> Result<&mut Self> {
        if let Some(tag) = value.as_deref() {
            validate_language_tag(tag, "East Asian theme font")?;
        }
        self.east_asia = value;
        Ok(self)
    }

    /// Set the complex-script (`w:bidi`) theme language.
    pub fn set_bidi(&mut self, value: Option<String>) -> Result<&mut Self> {
        if let Some(tag) = value.as_deref() {
            validate_language_tag(tag, "complex-script theme font")?;
        }
        self.bidi = value;
        Ok(self)
    }

    /// Return the Latin theme language, when specified.
    #[inline]
    pub fn latin(&self) -> Option<&str> {
        self.latin.as_deref()
    }

    /// Return the East Asian theme language, when specified.
    #[inline]
    pub fn east_asia(&self) -> Option<&str> {
        self.east_asia.as_deref()
    }

    /// Return the complex-script theme language, when specified.
    #[inline]
    pub fn bidi(&self) -> Option<&str> {
        self.bidi.as_deref()
    }

    /// Serialize a standalone `w:themeFontLang` fragment.
    pub fn to_xml(&self, prefix: &str) -> String {
        let mut xml = format!("<{prefix}:themeFontLang");
        for (name, value) in [
            ("val", &self.latin),
            ("eastAsia", &self.east_asia),
            ("bidi", &self.bidi),
        ] {
            if let Some(tag) = value {
                xml.push_str(&format!(" {prefix}:{name}=\""));
                escape_attribute(&mut xml, tag);
                xml.push('"');
            }
        }
        xml.push_str("/>");
        xml
    }
}

/// Theme color slot produced by a `w:clrSchemeMapping` value
/// (`ST_ColorSchemeIndex`, ECMA-376 section 17.18.11).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColorSchemeIndex {
    /// First dark theme color.
    Dark1,
    /// First light theme color.
    Light1,
    /// Second dark theme color.
    Dark2,
    /// Second light theme color.
    Light2,
    /// First accent theme color.
    Accent1,
    /// Second accent theme color.
    Accent2,
    /// Third accent theme color.
    Accent3,
    /// Fourth accent theme color.
    Accent4,
    /// Fifth accent theme color.
    Accent5,
    /// Sixth accent theme color.
    Accent6,
    /// Hyperlink theme color.
    Hyperlink,
    /// Followed hyperlink theme color.
    FollowedHyperlink,
}

impl ColorSchemeIndex {
    fn from_xml(value: &str) -> Result<Self> {
        match value {
            "dark1" => Ok(Self::Dark1),
            "light1" => Ok(Self::Light1),
            "dark2" => Ok(Self::Dark2),
            "light2" => Ok(Self::Light2),
            "accent1" => Ok(Self::Accent1),
            "accent2" => Ok(Self::Accent2),
            "accent3" => Ok(Self::Accent3),
            "accent4" => Ok(Self::Accent4),
            "accent5" => Ok(Self::Accent5),
            "accent6" => Ok(Self::Accent6),
            "hyperlink" => Ok(Self::Hyperlink),
            "followedHyperlink" => Ok(Self::FollowedHyperlink),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid color scheme index value '{value}'"
            ))),
        }
    }

    /// Get the XML value for this theme color slot.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Dark1 => "dark1",
            Self::Light1 => "light1",
            Self::Dark2 => "dark2",
            Self::Light2 => "light2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hyperlink",
            Self::FollowedHyperlink => "followedHyperlink",
        }
    }
}

/// A remappable theme color slot on `w:clrSchemeMapping` (ECMA-376 section
/// 17.15.1.21).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColorSchemeSlot {
    /// First background color slot (`w:bg1`).
    Background1,
    /// First text color slot (`w:t1`).
    Text1,
    /// Second background color slot (`w:bg2`).
    Background2,
    /// Second text color slot (`w:t2`).
    Text2,
    /// First accent color slot (`w:accent1`).
    Accent1,
    /// Second accent color slot (`w:accent2`).
    Accent2,
    /// Third accent color slot (`w:accent3`).
    Accent3,
    /// Fourth accent color slot (`w:accent4`).
    Accent4,
    /// Fifth accent color slot (`w:accent5`).
    Accent5,
    /// Sixth accent color slot (`w:accent6`).
    Accent6,
    /// Hyperlink color slot (`w:hyperlink`).
    Hyperlink,
    /// Followed hyperlink color slot (`w:followedHyperlink`).
    FollowedHyperlink,
}

impl ColorSchemeSlot {
    /// Number of remappable color slots.
    pub const COUNT: usize = 12;

    /// Every slot in schema attribute order.
    pub const ALL: [Self; Self::COUNT] = [
        Self::Background1,
        Self::Text1,
        Self::Background2,
        Self::Text2,
        Self::Accent1,
        Self::Accent2,
        Self::Accent3,
        Self::Accent4,
        Self::Accent5,
        Self::Accent6,
        Self::Hyperlink,
        Self::FollowedHyperlink,
    ];

    const fn index(self) -> usize {
        self as usize
    }

    /// Get the attribute name carrying this slot on `w:clrSchemeMapping`.
    pub const fn attribute_name(self) -> &'static str {
        match self {
            Self::Background1 => "bg1",
            Self::Text1 => "t1",
            Self::Background2 => "bg2",
            Self::Text2 => "t2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hyperlink",
            Self::FollowedHyperlink => "followedHyperlink",
        }
    }
}

/// Theme color slot remapping from `w:clrSchemeMapping` (ECMA-376 section
/// 17.15.1.21).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ColorSchemeMapping {
    slots: [Option<ColorSchemeIndex>; ColorSchemeSlot::COUNT],
}

impl ColorSchemeMapping {
    /// Create a mapping with every slot left at its default.
    pub fn new() -> Self {
        Self::default()
    }

    /// Remap a slot to a theme color index.
    pub fn set(&mut self, slot: ColorSchemeSlot, index: ColorSchemeIndex) -> &mut Self {
        self.slots[slot.index()] = Some(index);
        self
    }

    /// Restore a slot to its default mapping.
    pub fn clear(&mut self, slot: ColorSchemeSlot) -> &mut Self {
        self.slots[slot.index()] = None;
        self
    }

    /// Return the theme color index a slot maps to, when remapped.
    #[inline]
    pub fn get(&self, slot: ColorSchemeSlot) -> Option<ColorSchemeIndex> {
        self.slots[slot.index()]
    }

    /// Whether no slot is remapped.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.slots.iter().all(Option::is_none)
    }

    /// Iterate the remapped slots in schema attribute order.
    pub fn iter(&self) -> impl Iterator<Item = (ColorSchemeSlot, ColorSchemeIndex)> + '_ {
        ColorSchemeSlot::ALL
            .into_iter()
            .filter_map(|slot| self.get(slot).map(|index| (slot, index)))
    }

    /// Serialize a standalone `w:clrSchemeMapping` fragment.
    pub fn to_xml(&self, prefix: &str) -> String {
        let mut xml = format!("<{prefix}:clrSchemeMapping");
        for (slot, index) in self.iter() {
            xml.push_str(&format!(
                " {prefix}:{}=\"{}\"",
                slot.attribute_name(),
                index.as_str()
            ));
        }
        xml.push_str("/>");
        xml
    }
}

impl DocumentSettings {
    /// Create a new DocumentSettings with default values.
    pub fn new() -> Self {
        Self {
            protected: false,
            protection_type: None,
            track_revisions: false,
            zoom_percent: None,
            smart_tag_types: Vec::new(),
            do_not_embed_smart_tags: false,
            mail_merge: None,
            attached_template: None,
            compatibility_options: Vec::new(),
            compatibility_settings: Vec::new(),
            footnote_properties: None,
            endnote_properties: None,
            write_protection: false,
            view: None,
            proofing_state: None,
            default_tab_stop_twips: None,
            theme_font_languages: None,
            color_scheme_mapping: None,
        }
    }

    /// Check if the document is protected.
    #[inline]
    pub fn is_protected(&self) -> bool {
        self.protected
    }

    /// Get the type of protection applied.
    #[inline]
    pub fn protection_type(&self) -> Option<ProtectionType> {
        self.protection_type
    }

    /// Check if track revisions is enabled.
    #[inline]
    pub fn track_revisions(&self) -> bool {
        self.track_revisions
    }

    /// Get the zoom percentage.
    #[inline]
    pub fn zoom_percent(&self) -> Option<u32> {
        self.zoom_percent
    }

    /// Return the declared smart-tag vocabularies in document order.
    #[inline]
    pub fn smart_tag_types(&self) -> &[SmartTagType] {
        &self.smart_tag_types
    }

    /// Whether embedded smart-tag data should be omitted when saving.
    #[inline]
    pub fn do_not_embed_smart_tags(&self) -> bool {
        self.do_not_embed_smart_tags
    }

    /// Return the document's inert mail-merge metadata, if present.
    #[inline]
    pub fn mail_merge(&self) -> Option<&MailMergeSettings> {
        self.mail_merge.as_ref()
    }

    /// Return the inert attached-template reference, if present.
    #[inline]
    pub fn attached_template(&self) -> Option<&AttachedTemplate> {
        self.attached_template.as_ref()
    }

    /// Return the on/off compatibility option flags in document order.
    #[inline]
    pub fn compatibility_options(&self) -> &[CompatibilityOption] {
        &self.compatibility_options
    }

    /// Return the `w:compatSetting` triples in document order.
    #[inline]
    pub fn compatibility_settings(&self) -> &[CompatibilitySetting] {
        &self.compatibility_settings
    }

    /// Look up a `w:compatSetting` triple by name and URI.
    pub fn compatibility_setting(&self, name: &str, uri: &str) -> Option<&CompatibilitySetting> {
        self.compatibility_settings
            .iter()
            .find(|setting| setting.name == name && setting.uri == uri)
    }

    /// Return the Word compatibility mode (`compatibilityMode` value), when
    /// declared — for example `15` targets Word 2013 behavior.
    pub fn compatibility_mode(&self) -> Option<u32> {
        self.compatibility_setting(COMPATIBILITY_MODE_SETTING_NAME, COMPATIBILITY_SETTING_URI)
            .and_then(|setting| setting.value.parse().ok())
    }

    /// Return the document-level footnote properties, if present.
    #[inline]
    pub fn footnote_properties(&self) -> Option<&NoteNumberingProperties> {
        self.footnote_properties.as_ref()
    }

    /// Return the document-level endnote properties, if present.
    #[inline]
    pub fn endnote_properties(&self) -> Option<&NoteNumberingProperties> {
        self.endnote_properties.as_ref()
    }

    /// Whether applications should recommend write protection for the
    /// document (`w:writeProtection`).
    #[inline]
    pub fn is_write_protected(&self) -> bool {
        self.write_protection
    }

    /// Return the document view mode (`w:view`), when specified.
    #[inline]
    pub fn view(&self) -> Option<DocumentView> {
        self.view
    }

    /// Return the proofing completion markers (`w:proofState`), if present.
    #[inline]
    pub fn proofing_state(&self) -> Option<&ProofingState> {
        self.proofing_state.as_ref()
    }

    /// Return the default tab stop interval in twips (`w:defaultTabStop`),
    /// when specified.
    #[inline]
    pub fn default_tab_stop_twips(&self) -> Option<u32> {
        self.default_tab_stop_twips
    }

    /// Return the theme font language defaults (`w:themeFontLang`), if
    /// present.
    #[inline]
    pub fn theme_font_languages(&self) -> Option<&ThemeFontLanguages> {
        self.theme_font_languages.as_ref()
    }

    /// Return the theme color slot remapping (`w:clrSchemeMapping`), if
    /// present.
    #[inline]
    pub fn color_scheme_mapping(&self) -> Option<&ColorSchemeMapping> {
        self.color_scheme_mapping.as_ref()
    }

    /// Serialize the editing view, proofing, and theme default elements
    /// (`w:writeProtection`, `w:view`, `w:proofState`, `w:defaultTabStop`,
    /// `w:themeFontLang`, `w:clrSchemeMapping`) in ECMA-376 schema order.
    pub fn to_editing_settings_xml(&self, prefix: &str) -> String {
        let mut xml = String::new();
        if self.write_protection {
            xml.push_str(&format!("<{prefix}:writeProtection/>"));
        }
        if let Some(view) = self.view {
            xml.push_str(&format!(
                "<{prefix}:view {prefix}:val=\"{}\"/>",
                view.as_str()
            ));
        }
        if let Some(state) = &self.proofing_state {
            xml.push_str(&state.to_xml(prefix));
        }
        if let Some(twips) = self.default_tab_stop_twips {
            xml.push_str(&format!(
                "<{prefix}:defaultTabStop {prefix}:val=\"{twips}\"/>"
            ));
        }
        if let Some(languages) = &self.theme_font_languages {
            xml.push_str(&languages.to_xml(prefix));
        }
        if let Some(mapping) = &self.color_scheme_mapping {
            xml.push_str(&mapping.to_xml(prefix));
        }
        xml
    }

    /// Extract settings from a settings.xml part.
    ///
    /// # Arguments
    ///
    /// * `part` - The settings part
    ///
    /// # Returns
    ///
    /// A DocumentSettings object
    pub(crate) fn extract_from_part(part: &dyn Part) -> Result<Self> {
        let xml = litchi_ooxml_common::mce::process_part(part)?;
        let mut settings = Self::extract_from_xml(xml.as_ref())?;
        validate_mail_merge_relationships(part, settings.mail_merge.as_ref())?;
        validate_attached_template_relationship(part, &mut settings)?;
        Ok(settings)
    }

    fn extract_from_xml(xml_bytes: &[u8]) -> Result<Self> {
        let mut reader = NsReader::from_reader(xml_bytes);

        let mut settings = Self::new();
        settings.mail_merge = parse_settings_mail_merge(xml_bytes)?;
        let mut depth = 0usize;
        let mut saw_root = false;
        let mut seen = SeenSettings::default();
        let mut saw_compat = false;
        let mut pending_group: Option<PendingGroup> = None;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);

            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("Word settings XML nesting is too deep".into())
                    })?;
                    if depth == 1 {
                        validate_settings_root(&namespace, &element, saw_root)?;
                        saw_root = true;
                    } else if saw_root && is_wordprocessing_namespace(&namespace) {
                        if depth == 2 {
                            if let Some(group) =
                                begin_settings_group(&element, &settings, &mut saw_compat)?
                            {
                                pending_group = Some(group);
                            } else {
                                parse_setting(
                                    &element,
                                    decoder,
                                    &resolver,
                                    &mut settings,
                                    &mut seen,
                                )?;
                            }
                        } else if depth == 3
                            && let Some(group) = pending_group.as_mut()
                        {
                            parse_group_child(group, &element, decoder, &resolver)?;
                        }
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("Word settings XML nesting is too deep".into())
                    })?;
                    if child_depth == 1 {
                        validate_settings_root(&namespace, &element, saw_root)?;
                        saw_root = true;
                    } else if saw_root && is_wordprocessing_namespace(&namespace) {
                        if child_depth == 2 {
                            if let Some(group) =
                                begin_settings_group(&element, &settings, &mut saw_compat)?
                            {
                                finish_settings_group(&mut settings, group)?;
                            } else {
                                parse_setting(
                                    &element,
                                    decoder,
                                    &resolver,
                                    &mut settings,
                                    &mut seen,
                                )?;
                            }
                        } else if child_depth == 3
                            && let Some(group) = pending_group.as_mut()
                        {
                            parse_group_child(group, &element, decoder, &resolver)?;
                        }
                    }
                },
                Event::End(_) => {
                    if depth == 2
                        && let Some(group) = pending_group.take()
                    {
                        finish_settings_group(&mut settings, group)?;
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid Word settings XML nesting".into())
                    })?;
                },
                Event::Eof if depth != 0 => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated Word settings XML".into(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        if !saw_root {
            return Err(OoxmlError::InvalidFormat(
                "settings part has no settings root".into(),
            ));
        }
        Ok(settings)
    }
}

fn validate_settings_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    saw_root: bool,
) -> Result<()> {
    if saw_root
        || !is_wordprocessing_namespace(namespace)
        || element.local_name().as_ref() != b"settings"
    {
        return Err(OoxmlError::InvalidFormat(
            "settings part has an invalid or trailing root element".into(),
        ));
    }
    Ok(())
}

/// A grouped settings element (`w:compat`, `w:footnotePr`, `w:endnotePr`)
/// currently being collected from the stream.
enum PendingGroup {
    Compatibility {
        options: Vec<CompatibilityOption>,
        settings: Vec<CompatibilitySetting>,
    },
    FootnoteProperties(NoteNumberingProperties),
    EndnoteProperties(NoteNumberingProperties),
}

fn begin_settings_group(
    element: &BytesStart<'_>,
    settings: &DocumentSettings,
    saw_compat: &mut bool,
) -> Result<Option<PendingGroup>> {
    match element.local_name().as_ref() {
        b"compat" => {
            if std::mem::replace(saw_compat, true) {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate compat settings group".into(),
                ));
            }
            Ok(Some(PendingGroup::Compatibility {
                options: Vec::new(),
                settings: Vec::new(),
            }))
        },
        b"footnotePr" => {
            if settings.footnote_properties.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate footnotePr settings group".into(),
                ));
            }
            Ok(Some(PendingGroup::FootnoteProperties(
                NoteNumberingProperties::default(),
            )))
        },
        b"endnotePr" => {
            if settings.endnote_properties.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate endnotePr settings group".into(),
                ));
            }
            Ok(Some(PendingGroup::EndnoteProperties(
                NoteNumberingProperties::default(),
            )))
        },
        _ => Ok(None),
    }
}

fn finish_settings_group(settings: &mut DocumentSettings, group: PendingGroup) -> Result<()> {
    match group {
        PendingGroup::Compatibility {
            options,
            settings: triples,
        } => {
            settings.compatibility_options = options;
            settings.compatibility_settings = triples;
        },
        PendingGroup::FootnoteProperties(properties) => {
            if settings.footnote_properties.replace(properties).is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate footnotePr settings group".into(),
                ));
            }
        },
        PendingGroup::EndnoteProperties(properties) => {
            if settings.endnote_properties.replace(properties).is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate endnotePr settings group".into(),
                ));
            }
        },
    }
    Ok(())
}

fn parse_group_child(
    group: &mut PendingGroup,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    match group {
        PendingGroup::Compatibility { options, settings } => {
            if element.local_name().as_ref() == b"compatSetting" {
                settings.push(CompatibilitySetting {
                    name: required_attribute(
                        element,
                        b"name",
                        decoder,
                        resolver,
                        "compatSetting name",
                    )?,
                    uri: required_attribute(
                        element,
                        b"uri",
                        decoder,
                        resolver,
                        "compatSetting URI",
                    )?,
                    value: required_attribute(
                        element,
                        b"val",
                        decoder,
                        resolver,
                        "compatSetting value",
                    )?,
                });
            } else {
                options.push(CompatibilityOption {
                    name: String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
                    enabled: parse_on_off(element, decoder, resolver)?,
                });
            }
        },
        PendingGroup::FootnoteProperties(properties)
        | PendingGroup::EndnoteProperties(properties) => {
            parse_note_property_child(properties, element, decoder, resolver)?;
        },
    }
    Ok(())
}

fn parse_note_property_child(
    properties: &mut NoteNumberingProperties,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"pos" => {
            if properties.position.is_some() {
                return Err(OoxmlError::InvalidFormat("duplicate note position".into()));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note position")?;
            properties.position = Some(NotePosition::from_xml(value));
        },
        b"numFmt" => {
            if properties.format.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate note numbering format".into(),
                ));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note numFmt")?;
            properties.format = Some(NumberFormat::parse(&value));
        },
        b"numStart" => {
            if properties.start.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate note numbering start".into(),
                ));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note numStart")?;
            properties.start = Some(value.parse().map_err(|_| {
                OoxmlError::InvalidFormat(format!("invalid note numbering start '{value}'"))
            })?);
        },
        b"numRestart" => {
            if properties.restart.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate note numbering restart".into(),
                ));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note numRestart")?;
            properties.restart = Some(NoteNumberingRestart::from_xml(&value)?);
        },
        // `w:footnote`/`w:endnote` separator references carry no properties.
        _ => {},
    }
    Ok(())
}

/// Cardinality flags for on/off settings whose "not seen" state cannot be
/// told apart from an explicit `false` value.
#[derive(Debug, Default)]
struct SeenSettings {
    do_not_embed_smart_tags: bool,
    attached_template: bool,
    write_protection: bool,
}

fn parse_setting(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    settings: &mut DocumentSettings,
    seen: &mut SeenSettings,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"documentProtection" => {
            settings.protected = true;
            if let Some(value) = word_attribute_value(element, b"edit", decoder, resolver)? {
                settings.protection_type = ProtectionType::from_xml(&value);
            }
            if let Some(value) = word_attribute_value(element, b"enforcement", decoder, resolver)? {
                settings.protected = parse_on_off_value(&value)?;
            }
        },
        b"trackRevisions" => {
            settings.track_revisions = parse_on_off(element, decoder, resolver)?;
        },
        b"zoom" => {
            if let Some(value) = word_attribute_value(element, b"percent", decoder, resolver)? {
                settings.zoom_percent =
                    atoi_simd::parse::<u32, false, false>(value.as_bytes()).ok();
            }
        },
        b"smartTagType" => {
            settings.smart_tag_types.push(SmartTagType {
                namespace_uri: required_attribute(
                    element,
                    b"namespaceuri",
                    decoder,
                    resolver,
                    "smart-tag namespace URI",
                )?,
                name: required_attribute(
                    element,
                    b"name",
                    decoder,
                    resolver,
                    "smart-tag type name",
                )?,
                url: required_attribute(
                    element,
                    b"url",
                    decoder,
                    resolver,
                    "smart-tag vocabulary URL",
                )?,
            });
        },
        b"doNotEmbedSmartTags" => {
            if std::mem::replace(&mut seen.do_not_embed_smart_tags, true) {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate doNotEmbedSmartTags setting".into(),
                ));
            }
            settings.do_not_embed_smart_tags = parse_on_off(element, decoder, resolver)?;
        },
        b"attachedTemplate" => {
            if std::mem::replace(&mut seen.attached_template, true) {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate attachedTemplate setting".into(),
                ));
            }
            let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?
                .ok_or_else(|| {
                    OoxmlError::InvalidFormat("attachedTemplate relationship ID is required".into())
                })?;
            if relationship_id.is_empty() {
                return Err(OoxmlError::InvalidFormat(
                    "attachedTemplate relationship ID cannot be empty".into(),
                ));
            }
            settings.attached_template = Some(AttachedTemplate {
                relationship_id,
                target_uri: String::new(),
            });
        },
        b"writeProtection" => {
            if std::mem::replace(&mut seen.write_protection, true) {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate writeProtection setting".into(),
                ));
            }
            settings.write_protection = parse_on_off(element, decoder, resolver)?;
        },
        b"view" => {
            if settings.view.is_some() {
                return Err(OoxmlError::InvalidFormat("duplicate view setting".into()));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "view mode")?;
            settings.view = Some(DocumentView::from_xml(&value)?);
        },
        b"proofState" => {
            if settings.proofing_state.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate proofState setting".into(),
                ));
            }
            let mut state = ProofingState::new();
            if let Some(value) = word_attribute_value(element, b"spelling", decoder, resolver)? {
                state.set_spelling(Some(ProofState::from_xml(&value)?));
            }
            if let Some(value) = word_attribute_value(element, b"grammar", decoder, resolver)? {
                state.set_grammar(Some(ProofState::from_xml(&value)?));
            }
            settings.proofing_state = Some(state);
        },
        b"defaultTabStop" => {
            if settings.default_tab_stop_twips.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate defaultTabStop setting".into(),
                ));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "default tab stop")?;
            settings.default_tab_stop_twips = Some(value.parse().map_err(|_| {
                OoxmlError::InvalidFormat(format!("invalid default tab stop '{value}'"))
            })?);
        },
        b"themeFontLang" => {
            if settings.theme_font_languages.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate themeFontLang setting".into(),
                ));
            }
            let mut languages = ThemeFontLanguages::new();
            if let Some(value) = word_attribute_value(element, b"val", decoder, resolver)? {
                languages.set_latin(Some(value))?;
            }
            if let Some(value) = word_attribute_value(element, b"eastAsia", decoder, resolver)? {
                languages.set_east_asia(Some(value))?;
            }
            if let Some(value) = word_attribute_value(element, b"bidi", decoder, resolver)? {
                languages.set_bidi(Some(value))?;
            }
            settings.theme_font_languages = Some(languages);
        },
        b"clrSchemeMapping" => {
            if settings.color_scheme_mapping.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate clrSchemeMapping setting".into(),
                ));
            }
            let mut mapping = ColorSchemeMapping::new();
            for slot in ColorSchemeSlot::ALL {
                if let Some(value) = word_attribute_value(
                    element,
                    slot.attribute_name().as_bytes(),
                    decoder,
                    resolver,
                )? {
                    mapping.set(slot, ColorSchemeIndex::from_xml(&value)?);
                }
            }
            settings.color_scheme_mapping = Some(mapping);
        },
        _ => {},
    }
    Ok(())
}

fn relationship_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_relationship_attribute = matches!(
            namespace,
            ResolveResult::Bound(quick_xml::name::Namespace(uri))
                if uri == TRANSITIONAL_RELATIONSHIPS_NAMESPACE
                    || uri == STRICT_RELATIONSHIPS_NAMESPACE
        );
        if !is_relationship_attribute {
            continue;
        }
        if value.is_some() {
            return Err(OoxmlError::InvalidFormat(
                "duplicate attachedTemplate relationship ID attribute".into(),
            ));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

pub(crate) fn is_attached_template_relationship(value: &str) -> bool {
    matches!(
        value,
        ATTACHED_TEMPLATE_RELATIONSHIP | STRICT_ATTACHED_TEMPLATE_RELATIONSHIP
    )
}

pub(crate) fn validate_attached_template_target(target: &str) -> Result<()> {
    if target.is_empty() || target.len() > MAX_ATTACHED_TEMPLATE_TARGET_LEN {
        return Err(OoxmlError::InvalidFormat(
            "attached-template target must contain 1 to 32768 bytes".into(),
        ));
    }
    if target
        .chars()
        .any(|character| character.is_control() || character.is_whitespace())
    {
        return Err(OoxmlError::InvalidFormat(
            "attached-template target contains an invalid control or whitespace character".into(),
        ));
    }
    Ok(())
}

fn validate_attached_template_relationship(
    part: &dyn Part,
    settings: &mut DocumentSettings,
) -> Result<()> {
    let matching = part
        .rels()
        .iter()
        .filter(|relationship| is_attached_template_relationship(relationship.reltype()))
        .collect::<Vec<_>>();
    if matching.len() > 1 {
        return Err(OoxmlError::InvalidFormat(
            "settings part has multiple attached-template relationships".into(),
        ));
    }

    let Some(attached_template) = settings.attached_template.as_mut() else {
        if matching.is_empty() {
            return Ok(());
        }
        return Err(OoxmlError::InvalidFormat(
            "settings part has an attached-template relationship without an attachedTemplate element"
                .into(),
        ));
    };
    let relationship = part
        .rels()
        .get(&attached_template.relationship_id)
        .ok_or_else(|| {
            OoxmlError::InvalidFormat(format!(
                "attachedTemplate references missing relationship {:?}",
                attached_template.relationship_id
            ))
        })?;
    if !is_attached_template_relationship(relationship.reltype()) {
        return Err(OoxmlError::InvalidFormat(
            "attachedTemplate relationship has the wrong type".into(),
        ));
    }
    if !relationship.is_external() {
        return Err(OoxmlError::InvalidFormat(
            "attachedTemplate relationship must use external target mode".into(),
        ));
    }
    validate_attached_template_target(relationship.target_ref())?;
    attached_template.target_uri = relationship.target_ref().to_owned();
    Ok(())
}

struct SettingsXmlLayout {
    attached_template_range: Option<Range<usize>>,
    doc_vars_range: Option<Range<usize>>,
    doc_vars_insert_at: Option<usize>,
    mail_merge_range: Option<Range<usize>>,
    mail_merge_insert_at: Option<usize>,
    root_empty_range: Option<Range<usize>>,
    root_end: Option<usize>,
    root_qname: Vec<u8>,
    word_prefix: Option<Vec<u8>>,
    relationship_prefix: Option<Vec<u8>>,
    strict: bool,
}

pub(crate) fn patch_mail_merge(
    xml: &[u8],
    mail_merge: Option<&MailMergeSettings>,
    conformance: crate::docx::mail_merge::MailMergeConformance,
) -> Result<Vec<u8>> {
    DocumentSettings::extract_from_xml(xml)?;
    let layout = scan_settings_xml_layout(xml)?;
    let replacement = mail_merge
        .map(|value| value.to_xml(conformance))
        .transpose()?
        .unwrap_or_default();
    if let Some(range) = layout.mail_merge_range {
        let mut output = Vec::with_capacity(xml.len() - range.len() + replacement.len());
        output.extend_from_slice(&xml[..range.start]);
        output.extend_from_slice(replacement.as_bytes());
        output.extend_from_slice(&xml[range.end..]);
        return Ok(output);
    }
    if replacement.is_empty() {
        return Ok(xml.to_vec());
    }
    if let Some(range) = layout.root_empty_range {
        let root = &xml[range.clone()];
        let slash = root
            .windows(2)
            .rposition(|window| window == b"/>")
            .ok_or_else(|| OoxmlError::InvalidFormat("invalid empty settings root".into()))?;
        let mut output =
            Vec::with_capacity(xml.len() + replacement.len() + layout.root_qname.len() + 4);
        output.extend_from_slice(&xml[..range.start]);
        output.extend_from_slice(&root[..slash]);
        output.push(b'>');
        output.extend_from_slice(replacement.as_bytes());
        output.extend_from_slice(b"</");
        output.extend_from_slice(&layout.root_qname);
        output.push(b'>');
        output.extend_from_slice(&xml[range.end..]);
        return Ok(output);
    }
    let insert_at = layout
        .mail_merge_insert_at
        .or(layout.root_end)
        .ok_or_else(|| {
            OoxmlError::InvalidFormat("settings root has no mailMerge insertion point".into())
        })?;
    let mut output = Vec::with_capacity(xml.len() + replacement.len());
    output.extend_from_slice(&xml[..insert_at]);
    output.extend_from_slice(replacement.as_bytes());
    output.extend_from_slice(&xml[insert_at..]);
    Ok(output)
}

pub(crate) fn patch_document_variables(
    xml: &[u8],
    variables: &DocumentVariables,
) -> Result<Vec<u8>> {
    variables.validate()?;
    DocumentVariables::extract_from_xml(xml)?;
    let layout = scan_settings_xml_layout(xml)?;
    let replacement = if variables.is_empty() {
        String::new()
    } else {
        document_variables_element(&layout, variables)
    };

    if let Some(range) = layout.doc_vars_range {
        let mut output = Vec::with_capacity(xml.len() - range.len() + replacement.len());
        output.extend_from_slice(&xml[..range.start]);
        output.extend_from_slice(replacement.as_bytes());
        output.extend_from_slice(&xml[range.end..]);
        return Ok(output);
    }
    if replacement.is_empty() {
        return Ok(xml.to_vec());
    }
    if let Some(range) = layout.root_empty_range {
        let root = &xml[range.clone()];
        let slash = root
            .windows(2)
            .rposition(|window| window == b"/>")
            .ok_or_else(|| OoxmlError::InvalidFormat("invalid empty settings root".into()))?;
        let mut output =
            Vec::with_capacity(xml.len() + replacement.len() + layout.root_qname.len() + 4);
        output.extend_from_slice(&xml[..range.start]);
        output.extend_from_slice(&root[..slash]);
        output.push(b'>');
        output.extend_from_slice(replacement.as_bytes());
        output.extend_from_slice(b"</");
        output.extend_from_slice(&layout.root_qname);
        output.push(b'>');
        output.extend_from_slice(&xml[range.end..]);
        return Ok(output);
    }
    let insert_at = layout
        .doc_vars_insert_at
        .or(layout.root_end)
        .ok_or_else(|| OoxmlError::InvalidFormat("settings root has no insertion point".into()))?;
    let mut output = Vec::with_capacity(xml.len() + replacement.len());
    output.extend_from_slice(&xml[..insert_at]);
    output.extend_from_slice(replacement.as_bytes());
    output.extend_from_slice(&xml[insert_at..]);
    Ok(output)
}

pub(crate) fn patch_attached_template(
    xml: &[u8],
    relationship_id: Option<&str>,
) -> Result<Vec<u8>> {
    // Validate the original tree and its direct-child cardinality before using offsets.
    DocumentSettings::extract_from_xml(xml)?;
    let layout = scan_settings_xml_layout(xml)?;
    if relationship_id.is_none() && layout.attached_template_range.is_none() {
        return Ok(xml.to_vec());
    }

    let replacement = relationship_id
        .map(|id| attached_template_element(&layout, id))
        .unwrap_or_default();
    if let Some(range) = layout.attached_template_range {
        let mut output = Vec::with_capacity(xml.len() - range.len() + replacement.len());
        output.extend_from_slice(&xml[..range.start]);
        output.extend_from_slice(replacement.as_bytes());
        output.extend_from_slice(&xml[range.end..]);
        return Ok(output);
    }

    if let Some(range) = layout.root_empty_range {
        let root = &xml[range.clone()];
        let slash = root
            .windows(2)
            .rposition(|window| window == b"/>")
            .ok_or_else(|| OoxmlError::InvalidFormat("invalid empty settings root".into()))?;
        let mut output =
            Vec::with_capacity(xml.len() + replacement.len() + layout.root_qname.len() + 4);
        output.extend_from_slice(&xml[..range.start]);
        output.extend_from_slice(&root[..slash]);
        output.push(b'>');
        output.extend_from_slice(replacement.as_bytes());
        output.extend_from_slice(b"</");
        output.extend_from_slice(&layout.root_qname);
        output.push(b'>');
        output.extend_from_slice(&xml[range.end..]);
        return Ok(output);
    }

    let insert_at = layout
        .root_end
        .ok_or_else(|| OoxmlError::InvalidFormat("settings root has no closing element".into()))?;
    let mut output = Vec::with_capacity(xml.len() + replacement.len());
    output.extend_from_slice(&xml[..insert_at]);
    output.extend_from_slice(replacement.as_bytes());
    output.extend_from_slice(&xml[insert_at..]);
    Ok(output)
}

fn scan_settings_xml_layout(xml: &[u8]) -> Result<SettingsXmlLayout> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut root_qname = None;
    let mut word_prefix = None;
    let mut relationship_prefix = None;
    let mut strict = false;
    let mut root_empty_range = None;
    let mut root_end = None;
    let mut attached_template_range = None;
    let mut attached_start = None;
    let mut doc_vars_range = None;
    let mut doc_vars_start = None;
    let mut doc_vars_insert_at = None;
    let mut mail_merge_range = None;
    let mut mail_merge_start = None;
    let mut mail_merge_insert_at = None;

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_| OoxmlError::InvalidFormat("settings XML offset is too large".into()))?;
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_| OoxmlError::InvalidFormat("settings XML offset is too large".into()))?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("settings XML nesting is too deep".into())
                })?;
                if depth == 1 {
                    capture_settings_root(
                        &namespace,
                        &element,
                        reader.decoder(),
                        &mut root_qname,
                        &mut word_prefix,
                        &mut relationship_prefix,
                        &mut strict,
                    )?;
                } else if depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"docVars"
                {
                    doc_vars_start = Some(event_start);
                } else if depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"attachedTemplate"
                {
                    attached_start = Some(event_start);
                } else if depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"mailMerge"
                {
                    mail_merge_start = Some(event_start);
                }
                if depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && mail_merge_insert_at.is_none()
                    && is_after_mail_merge(element.local_name().as_ref())
                {
                    mail_merge_insert_at = Some(event_start);
                }
                if depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && doc_vars_insert_at.is_none()
                    && is_after_doc_vars(element.local_name().as_ref())
                {
                    doc_vars_insert_at = Some(event_start);
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("settings XML nesting is too deep".into())
                })?;
                if child_depth == 1 {
                    capture_settings_root(
                        &namespace,
                        &element,
                        reader.decoder(),
                        &mut root_qname,
                        &mut word_prefix,
                        &mut relationship_prefix,
                        &mut strict,
                    )?;
                    root_empty_range = Some(event_start..event_end);
                } else if child_depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"docVars"
                {
                    doc_vars_range = Some(event_start..event_end);
                } else if child_depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"attachedTemplate"
                {
                    attached_template_range = Some(event_start..event_end);
                } else if child_depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"mailMerge"
                {
                    mail_merge_range = Some(event_start..event_end);
                }
                if child_depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && mail_merge_insert_at.is_none()
                    && is_after_mail_merge(element.local_name().as_ref())
                {
                    mail_merge_insert_at = Some(event_start);
                }
                if child_depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && doc_vars_insert_at.is_none()
                    && is_after_doc_vars(element.local_name().as_ref())
                {
                    doc_vars_insert_at = Some(event_start);
                }
            },
            Event::End(_) => {
                if depth == 2
                    && let Some(start) = attached_start.take()
                {
                    attached_template_range = Some(start..event_end);
                }
                if depth == 2
                    && let Some(start) = doc_vars_start.take()
                {
                    doc_vars_range = Some(start..event_end);
                }
                if depth == 2
                    && let Some(start) = mail_merge_start.take()
                {
                    mail_merge_range = Some(start..event_end);
                }
                if depth == 1 {
                    root_end = Some(event_start);
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid settings XML nesting".into())
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(SettingsXmlLayout {
        attached_template_range,
        doc_vars_range,
        doc_vars_insert_at,
        mail_merge_range,
        mail_merge_insert_at,
        root_empty_range,
        root_end,
        root_qname: root_qname
            .ok_or_else(|| OoxmlError::InvalidFormat("settings root is missing".into()))?,
        word_prefix,
        relationship_prefix,
        strict,
    })
}

fn is_after_mail_merge(local_name: &[u8]) -> bool {
    matches!(
        local_name,
        b"revisionView"
            | b"trackRevisions"
            | b"doNotTrackMoves"
            | b"doNotTrackFormatting"
            | b"documentProtection"
            | b"autoFormatOverride"
            | b"styleLockTheme"
            | b"styleLockQFSet"
            | b"defaultTabStop"
            | b"hyphenationZone"
            | b"consecutiveHyphenLimit"
            | b"doNotHyphenateCaps"
            | b"showEnvelope"
            | b"summaryLength"
            | b"clickAndTypeStyle"
            | b"defaultTableStyle"
            | b"evenAndOddHeaders"
            | b"bookFoldRevPrinting"
            | b"bookFoldPrinting"
            | b"bookFoldPrintingSheets"
            | b"drawingGridHorizontalSpacing"
            | b"drawingGridVerticalSpacing"
            | b"displayHorizontalDrawingGridEvery"
            | b"displayVerticalDrawingGridEvery"
    )
}

fn is_after_doc_vars(local_name: &[u8]) -> bool {
    matches!(
        local_name,
        b"rsids"
            | b"uiCompat97To2003"
            | b"attachedSchema"
            | b"themeFontLang"
            | b"clrSchemeMapping"
            | b"doNotIncludeSubdocsInStats"
            | b"doNotAutoCompressPictures"
            | b"forceUpgrade"
            | b"captions"
            | b"readModeInkLockDown"
            | b"smartTagType"
            | b"schemaLibrary"
            | b"shapeDefaults"
            | b"doNotEmbedSmartTags"
            | b"decimalSymbol"
            | b"listSeparator"
    )
}

fn document_variables_element(layout: &SettingsXmlLayout, variables: &DocumentVariables) -> String {
    let prefix = layout
        .word_prefix
        .as_deref()
        .map(String::from_utf8_lossy)
        .unwrap_or_else(|| "w".into());
    let mut output = format!("<{prefix}:docVars");
    if layout.word_prefix.is_none() {
        let namespace = if layout.strict {
            crate::docx::namespace::STRICT_WORDPROCESSINGML_NAMESPACE
        } else {
            crate::docx::namespace::WORDPROCESSINGML_NAMESPACE
        };
        output.push_str(&format!(
            " xmlns:{prefix}=\"{}\"",
            String::from_utf8_lossy(namespace)
        ));
    }
    output.push('>');
    variables.write_entries(&mut output, &prefix);
    output.push_str(&format!("</{prefix}:docVars>"));
    output
}

#[allow(clippy::too_many_arguments)]
fn capture_settings_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    root_qname: &mut Option<Vec<u8>>,
    word_prefix: &mut Option<Vec<u8>>,
    relationship_prefix: &mut Option<Vec<u8>>,
    strict: &mut bool,
) -> Result<()> {
    *root_qname = Some(element.name().as_ref().to_vec());
    *word_prefix = element
        .name()
        .prefix()
        .map(|prefix| prefix.into_inner().to_vec());
    *strict = matches!(
        namespace,
        ResolveResult::Bound(quick_xml::name::Namespace(uri))
            if *uri == crate::docx::namespace::STRICT_WORDPROCESSINGML_NAMESPACE
    );
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        let Some(prefix) = key.strip_prefix(b"xmlns:") else {
            continue;
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if value.as_bytes() == TRANSITIONAL_RELATIONSHIPS_NAMESPACE
            || value.as_bytes() == STRICT_RELATIONSHIPS_NAMESPACE
        {
            *relationship_prefix = Some(prefix.to_vec());
            break;
        }
    }
    Ok(())
}

fn attached_template_element(layout: &SettingsXmlLayout, relationship_id: &str) -> String {
    let word_name = layout.word_prefix.as_ref().map_or_else(
        || "attachedTemplate".to_owned(),
        |prefix| format!("{}:attachedTemplate", String::from_utf8_lossy(prefix)),
    );
    let relationship_prefix = layout
        .relationship_prefix
        .as_deref()
        .map(String::from_utf8_lossy)
        .unwrap_or_else(|| "r".into());
    let mut output = format!("<{word_name} {relationship_prefix}:id=\"");
    escape_attribute(&mut output, relationship_id);
    output.push('"');
    if layout.relationship_prefix.is_none() {
        let namespace = if layout.strict {
            String::from_utf8_lossy(STRICT_RELATIONSHIPS_NAMESPACE)
        } else {
            String::from_utf8_lossy(TRANSITIONAL_RELATIONSHIPS_NAMESPACE)
        };
        output.push_str(&format!(" xmlns:{relationship_prefix}=\"{namespace}\""));
    }
    output.push_str("/>");
    output
}

fn escape_attribute(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            _ => output.push(character),
        }
    }
}

fn required_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<String> {
    word_attribute_value(element, name, decoder, resolver)?.ok_or_else(|| {
        OoxmlError::InvalidFormat(format!("Word {description} attribute is required"))
    })
}

fn parse_on_off(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<bool> {
    word_attribute_value(element, b"val", decoder, resolver)?
        .as_deref()
        .map_or(Ok(true), parse_on_off_value)
}

fn parse_on_off_value(value: &str) -> Result<bool> {
    match value {
        "true" | "1" | "on" => Ok(true),
        "false" | "0" | "off" => Ok(false),
        _ => Err(OoxmlError::InvalidFormat(format!(
            "invalid Word on/off value '{value}'"
        ))),
    }
}

impl Default for DocumentSettings {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackURI;
    use litchi_opc::part::{BlobPart, Part};

    #[test]
    fn test_settings_creation() {
        let settings = DocumentSettings::new();
        assert!(!settings.is_protected());
        assert!(settings.protection_type().is_none());
        assert!(!settings.track_revisions());
    }

    #[test]
    fn test_protection_type() {
        assert_eq!(
            ProtectionType::from_xml("readOnly"),
            Some(ProtectionType::ReadOnly)
        );
        assert_eq!(
            ProtectionType::from_xml("comments"),
            Some(ProtectionType::Comments)
        );
        assert_eq!(
            ProtectionType::from_xml("trackedChanges"),
            Some(ProtectionType::TrackedChanges)
        );
        assert_eq!(
            ProtectionType::from_xml("forms"),
            Some(ProtectionType::Forms)
        );
        assert_eq!(ProtectionType::from_xml("invalid"), None);
    }

    #[test]
    fn parses_smart_tag_settings_with_strict_namespaces() {
        let xml = br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"
            xmlns:false="urn:not-wordprocessingml">
            <s:trackRevisions s:val="on"/>
            <s:smartTagType s:namespaceuri="urn:contacts" s:name="person"
                s:url="https://example.test/schema?a=1&amp;b=2"/>
            <false:smartTagType false:namespaceuri="urn:false" false:name="ignored" false:url="ignored"/>
            <s:doNotEmbedSmartTags s:val="off"/>
        </s:settings>"#;

        let settings = DocumentSettings::extract_from_xml(xml).unwrap();
        assert!(settings.track_revisions());
        assert!(!settings.do_not_embed_smart_tags());
        assert_eq!(settings.smart_tag_types().len(), 1);
        assert_eq!(
            settings.smart_tag_types()[0].namespace_uri(),
            "urn:contacts"
        );
        assert_eq!(settings.smart_tag_types()[0].name(), "person");
        assert_eq!(
            settings.smart_tag_types()[0].url(),
            "https://example.test/schema?a=1&b=2"
        );
    }

    #[test]
    fn validates_smart_tag_settings() {
        let enabled = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:doNotEmbedSmartTags/></w:settings>"#;
        assert!(
            DocumentSettings::extract_from_xml(enabled)
                .unwrap()
                .do_not_embed_smart_tags()
        );

        let missing_url = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:smartTagType w:namespaceuri="urn:test" w:name="test"/></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(missing_url).is_err());

        let invalid_on_off = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:doNotEmbedSmartTags w:val="maybe"/></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(invalid_on_off).is_err());

        let duplicate = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:doNotEmbedSmartTags/><w:doNotEmbedSmartTags/></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(duplicate).is_err());
    }

    #[test]
    fn parses_compat_options_and_settings() {
        let xml = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:compat><w:useFELayout/><w:doNotExpandShiftReturn w:val="off"/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat><w:rsids/></w:settings>"#;
        let settings = DocumentSettings::extract_from_xml(xml).unwrap();
        assert_eq!(settings.zoom_percent(), Some(100));
        assert_eq!(settings.compatibility_options().len(), 2);
        assert_eq!(settings.compatibility_options()[0].name(), "useFELayout");
        assert!(settings.compatibility_options()[0].is_enabled());
        assert_eq!(
            settings.compatibility_options()[1].name(),
            "doNotExpandShiftReturn"
        );
        assert!(!settings.compatibility_options()[1].is_enabled());
        assert_eq!(settings.compatibility_settings().len(), 2);
        let mode = settings
            .compatibility_setting(COMPATIBILITY_MODE_SETTING_NAME, COMPATIBILITY_SETTING_URI)
            .unwrap();
        assert_eq!(mode.value(), "14");
        assert_eq!(settings.compatibility_mode(), Some(14));
        assert!(
            settings
                .compatibility_setting("missing", COMPATIBILITY_SETTING_URI)
                .is_none()
        );
    }

    #[test]
    fn parses_empty_and_strict_compat_groups() {
        let empty = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat/></w:settings>"#;
        let settings = DocumentSettings::extract_from_xml(empty).unwrap();
        assert!(settings.compatibility_options().is_empty());
        assert!(settings.compatibility_settings().is_empty());
        assert_eq!(settings.compatibility_mode(), None);

        let strict = br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:compat><s:compatSetting s:name="compatibilityMode" s:uri="http://schemas.microsoft.com/office/word" s:val="15"/></s:compat></s:settings>"#;
        let settings = DocumentSettings::extract_from_xml(strict).unwrap();
        assert_eq!(settings.compatibility_mode(), Some(15));
    }

    #[test]
    fn rejects_invalid_compat_groups() {
        let duplicate = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat/><w:compat/></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(duplicate).is_err());

        let missing_value = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word"/></w:compat></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(missing_value).is_err());

        let unterminated = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat>"#;
        assert!(DocumentSettings::extract_from_xml(unterminated).is_err());
    }

    #[test]
    fn parses_document_level_note_properties() {
        let xml = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnotePr><w:footnote w:type="separator" w:id="-1"/><w:pos w:val="pageBottom"/><w:numFmt w:val="lowerRoman"/><w:numStart w:val="2"/><w:numRestart w:val="eachPage"/></w:footnotePr><w:endnotePr><w:pos w:val="docEnd"/><w:numFmt w:val="upperLetter"/></w:endnotePr></w:settings>"#;
        let settings = DocumentSettings::extract_from_xml(xml).unwrap();

        let footnotes = settings.footnote_properties().unwrap();
        assert_eq!(footnotes.position(), Some(&NotePosition::PageBottom));
        assert_eq!(footnotes.format(), Some(&NumberFormat::LowerRoman));
        assert_eq!(footnotes.start(), Some(2));
        assert_eq!(footnotes.restart(), Some(NoteNumberingRestart::EachPage));

        let endnotes = settings.endnote_properties().unwrap();
        assert_eq!(endnotes.position(), Some(&NotePosition::DocumentEnd));
        assert_eq!(endnotes.format(), Some(&NumberFormat::UpperLetter));
        assert_eq!(endnotes.start(), None);
        assert_eq!(endnotes.restart(), None);
    }

    #[test]
    fn note_property_enums_round_trip() {
        for (raw, expected) in [
            ("pageBottom", NotePosition::PageBottom),
            ("beneathText", NotePosition::BeneathText),
            ("sectEnd", NotePosition::SectionEnd),
            ("docEnd", NotePosition::DocumentEnd),
        ] {
            assert_eq!(NotePosition::from_xml(raw.to_owned()), expected);
            assert_eq!(expected.as_str(), raw);
        }
        let extension = NotePosition::from_xml("vendorPosition".to_owned());
        assert_eq!(extension.as_str(), "vendorPosition");

        for (raw, expected) in [
            ("continuous", NoteNumberingRestart::Continuous),
            ("eachSect", NoteNumberingRestart::EachSection),
            ("eachPage", NoteNumberingRestart::EachPage),
        ] {
            assert_eq!(NoteNumberingRestart::from_xml(raw).unwrap(), expected);
            assert_eq!(expected.as_str(), raw);
        }
        assert!(NoteNumberingRestart::from_xml("sometimes").is_err());
    }

    #[test]
    fn rejects_invalid_note_property_groups() {
        let duplicate = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnotePr/><w:footnotePr/></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(duplicate).is_err());

        let bad_start = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnotePr><w:numStart w:val="soon"/></w:endnotePr></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(bad_start).is_err());

        let bad_restart = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnotePr><w:numRestart w:val="sometimes"/></w:footnotePr></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(bad_restart).is_err());

        let duplicate_child = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnotePr><w:numFmt w:val="decimal"/><w:numFmt w:val="bullet"/></w:footnotePr></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(duplicate_child).is_err());
    }

    #[test]
    fn parses_view_proofing_and_theme_defaults() {
        let xml = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:writeProtection/><w:view w:val="print"/><w:proofState w:spelling="clean" w:grammar="dirty"/><w:defaultTabStop w:val="720"/><w:themeFontLang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:hyperlink="hyperlink"/></w:settings>"#;
        let settings = DocumentSettings::extract_from_xml(xml).unwrap();
        assert!(settings.is_write_protected());
        assert_eq!(settings.view(), Some(DocumentView::Print));
        let proofing = settings.proofing_state().unwrap();
        assert_eq!(proofing.spelling(), Some(ProofState::Clean));
        assert_eq!(proofing.grammar(), Some(ProofState::Dirty));
        assert_eq!(settings.default_tab_stop_twips(), Some(720));
        let languages = settings.theme_font_languages().unwrap();
        assert_eq!(languages.latin(), Some("en-US"));
        assert_eq!(languages.east_asia(), Some("ja-JP"));
        assert_eq!(languages.bidi(), Some("ar-SA"));
        let mapping = settings.color_scheme_mapping().unwrap();
        assert!(!mapping.is_empty());
        assert_eq!(
            mapping.get(ColorSchemeSlot::Background1),
            Some(ColorSchemeIndex::Light1)
        );
        assert_eq!(
            mapping.get(ColorSchemeSlot::Text1),
            Some(ColorSchemeIndex::Dark1)
        );
        assert_eq!(
            mapping.get(ColorSchemeSlot::Hyperlink),
            Some(ColorSchemeIndex::Hyperlink)
        );
        assert_eq!(mapping.get(ColorSchemeSlot::Accent1), None);

        let strict = br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:writeProtection s:val="off"/><s:view s:val="web"/><s:proofState/></s:settings>"#;
        let settings = DocumentSettings::extract_from_xml(strict).unwrap();
        assert!(!settings.is_write_protected());
        assert_eq!(settings.view(), Some(DocumentView::Web));
        let proofing = settings.proofing_state().unwrap();
        assert_eq!(proofing.spelling(), None);
        assert_eq!(proofing.grammar(), None);
    }

    #[test]
    fn editing_settings_enums_round_trip() {
        for (raw, expected) in [
            ("none", DocumentView::None),
            ("print", DocumentView::Print),
            ("outline", DocumentView::Outline),
            ("masterPages", DocumentView::MasterPages),
            ("normal", DocumentView::Normal),
            ("web", DocumentView::Web),
        ] {
            assert_eq!(DocumentView::from_xml(raw).unwrap(), expected);
            assert_eq!(expected.as_str(), raw);
        }
        assert!(DocumentView::from_xml("immersive").is_err());

        for (raw, expected) in [("clean", ProofState::Clean), ("dirty", ProofState::Dirty)] {
            assert_eq!(ProofState::from_xml(raw).unwrap(), expected);
            assert_eq!(expected.as_str(), raw);
        }
        assert!(ProofState::from_xml("pending").is_err());

        assert_eq!(ColorSchemeSlot::COUNT, 12);
        for (raw, expected) in [
            ("dark1", ColorSchemeIndex::Dark1),
            ("light2", ColorSchemeIndex::Light2),
            ("accent6", ColorSchemeIndex::Accent6),
            ("followedHyperlink", ColorSchemeIndex::FollowedHyperlink),
        ] {
            assert_eq!(ColorSchemeIndex::from_xml(raw).unwrap(), expected);
            assert_eq!(expected.as_str(), raw);
        }
        assert!(ColorSchemeIndex::from_xml("accent7").is_err());
    }

    #[test]
    fn editing_settings_serialize_and_reparse() {
        let xml = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:writeProtection/><w:view w:val="outline"/><w:proofState w:spelling="dirty"/><w:defaultTabStop w:val="1440"/><w:themeFontLang w:val="en-US" w:bidi="he-IL"/><w:clrSchemeMapping w:bg1="dark2" w:accent3="accent1" w:followedHyperlink="hyperlink"/></w:settings>"#;
        let settings = DocumentSettings::extract_from_xml(xml).unwrap();

        let fragment = settings.to_editing_settings_xml("w");
        assert_eq!(
            fragment,
            r#"<w:writeProtection/><w:view w:val="outline"/><w:proofState w:spelling="dirty"/><w:defaultTabStop w:val="1440"/><w:themeFontLang w:val="en-US" w:bidi="he-IL"/><w:clrSchemeMapping w:bg1="dark2" w:accent3="accent1" w:followedHyperlink="hyperlink"/>"#
        );

        let reparsed_xml = format!(
            r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{fragment}</w:settings>"#
        );
        let reparsed = DocumentSettings::extract_from_xml(reparsed_xml.as_bytes()).unwrap();
        assert_eq!(reparsed.is_write_protected(), settings.is_write_protected());
        assert_eq!(reparsed.view(), settings.view());
        assert_eq!(reparsed.proofing_state(), settings.proofing_state());
        assert_eq!(
            reparsed.default_tab_stop_twips(),
            settings.default_tab_stop_twips()
        );
        assert_eq!(
            reparsed.theme_font_languages(),
            settings.theme_font_languages()
        );
        assert_eq!(
            reparsed.color_scheme_mapping(),
            settings.color_scheme_mapping()
        );
        // Serializing the reparsed settings is stable.
        assert_eq!(reparsed.to_editing_settings_xml("w"), fragment);
    }

    #[test]
    fn editing_settings_builders_write_fragments() {
        let mut proofing = ProofingState::new();
        proofing
            .set_spelling(Some(ProofState::Clean))
            .set_grammar(Some(ProofState::Dirty));
        assert_eq!(
            proofing.to_xml("w"),
            r#"<w:proofState w:spelling="clean" w:grammar="dirty"/>"#
        );

        let mut languages = ThemeFontLanguages::new();
        languages
            .set_latin(Some("fr-FR".to_owned()))
            .unwrap()
            .set_bidi(None)
            .unwrap();
        assert!(languages.set_latin(Some(String::new())).is_err());
        assert_eq!(languages.to_xml("w"), r#"<w:themeFontLang w:val="fr-FR"/>"#);

        let mut mapping = ColorSchemeMapping::new();
        assert!(mapping.is_empty());
        mapping
            .set(ColorSchemeSlot::Background1, ColorSchemeIndex::Light1)
            .set(ColorSchemeSlot::Text1, ColorSchemeIndex::Dark1);
        mapping.clear(ColorSchemeSlot::Text1);
        assert_eq!(
            mapping.iter().collect::<Vec<_>>(),
            [(ColorSchemeSlot::Background1, ColorSchemeIndex::Light1)]
        );
        assert_eq!(
            mapping.to_xml("w"),
            r#"<w:clrSchemeMapping w:bg1="light1"/>"#
        );
    }

    #[test]
    fn rejects_invalid_editing_settings() {
        let prefix = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#;
        let suffix = br#"</w:settings>"#;
        let reject = |body: &[u8]| {
            let mut xml = prefix.to_vec();
            xml.extend_from_slice(body);
            xml.extend_from_slice(suffix);
            DocumentSettings::extract_from_xml(&xml)
        };

        assert!(reject(br#"<w:view/>"#).is_err());
        assert!(reject(br#"<w:view w:val="immersive"/>"#).is_err());
        assert!(reject(br#"<w:view w:val="print"/><w:view w:val="web"/>"#).is_err());
        assert!(reject(br#"<w:proofState w:spelling="pending"/>"#).is_err());
        assert!(reject(br#"<w:proofState w:grammar="maybe"/>"#).is_err());
        assert!(reject(br#"<w:proofState/><w:proofState/>"#).is_err());
        assert!(reject(br#"<w:defaultTabStop/>"#).is_err());
        assert!(reject(br#"<w:defaultTabStop w:val="wide"/>"#).is_err());
        assert!(reject(br#"<w:defaultTabStop w:val="-720"/>"#).is_err());
        assert!(reject(br#"<w:defaultTabStop w:val="99999999999999999999"/>"#).is_err());
        assert!(
            reject(br#"<w:defaultTabStop w:val="720"/><w:defaultTabStop w:val="720"/>"#).is_err()
        );
        assert!(reject(br#"<w:themeFontLang w:val=""/>"#).is_err());
        assert!(reject(br#"<w:themeFontLang w:eastAsia="ja&#0;-JP"/>"#).is_err());
        assert!(reject(br#"<w:themeFontLang w:val="en-US"/><w:themeFontLang/>"#).is_err());
        assert!(reject(br#"<w:clrSchemeMapping w:bg1="light7"/>"#).is_err());
        assert!(reject(br#"<w:clrSchemeMapping/><w:clrSchemeMapping/>"#).is_err());
        assert!(reject(br#"<w:writeProtection/><w:writeProtection/>"#).is_err());
        assert!(reject(br#"<w:writeProtection w:val="maybe"/>"#).is_err());
    }

    #[test]
    fn parses_bundled_settings_resource() {
        let settings =
            DocumentSettings::extract_from_xml(include_bytes!("resources/settings.xml")).unwrap();
        assert_eq!(settings.compatibility_mode(), Some(14));
        assert_eq!(settings.compatibility_settings().len(), 4);
        assert!(
            settings
                .compatibility_options()
                .iter()
                .any(|option| option.name() == "useFELayout" && option.is_enabled())
        );
        let proofing = settings.proofing_state().unwrap();
        assert_eq!(proofing.spelling(), Some(ProofState::Clean));
        assert_eq!(proofing.grammar(), Some(ProofState::Clean));
        assert_eq!(settings.default_tab_stop_twips(), Some(720));
        let languages = settings.theme_font_languages().unwrap();
        assert_eq!(languages.latin(), Some("en-US"));
        assert_eq!(languages.east_asia(), Some("ja-JP"));
        let mapping = settings.color_scheme_mapping().unwrap();
        for (slot, expected) in [
            (ColorSchemeSlot::Background1, ColorSchemeIndex::Light1),
            (ColorSchemeSlot::Text1, ColorSchemeIndex::Dark1),
            (ColorSchemeSlot::Background2, ColorSchemeIndex::Light2),
            (ColorSchemeSlot::Text2, ColorSchemeIndex::Dark2),
            (ColorSchemeSlot::Accent1, ColorSchemeIndex::Accent1),
            (ColorSchemeSlot::Accent6, ColorSchemeIndex::Accent6),
            (ColorSchemeSlot::Hyperlink, ColorSchemeIndex::Hyperlink),
            (
                ColorSchemeSlot::FollowedHyperlink,
                ColorSchemeIndex::FollowedHyperlink,
            ),
        ] {
            assert_eq!(mapping.get(slot), Some(expected));
        }
    }

    fn attached_template_part(xml: &[u8], reltype: &str, target: &str, external: bool) -> BlobPart {
        let mut part = BlobPart::new(
            PackURI::new("/word/settings.xml").unwrap(),
            litchi_opc::constants::content_type::WML_SETTINGS.to_owned(),
            xml.to_vec(),
        );
        part.rels_mut().add_relationship(
            reltype.to_owned(),
            target.to_owned(),
            "customRel".to_owned(),
            external,
        );
        part
    }

    #[test]
    fn parses_transitional_and_strict_attached_templates() {
        for (word, relationships, reltype) in [
            (
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                ATTACHED_TEMPLATE_RELATIONSHIP,
            ),
            (
                "http://purl.oclc.org/ooxml/wordprocessingml/main",
                "http://purl.oclc.org/ooxml/officeDocument/relationships",
                STRICT_ATTACHED_TEMPLATE_RELATIONSHIP,
            ),
        ] {
            let xml = format!(
                r#"<q:settings xmlns:q="{word}" xmlns:rel="{relationships}"><q:attachedTemplate rel:id="customRel"/></q:settings>"#
            );
            let part = attached_template_part(
                xml.as_bytes(),
                reltype,
                "file:///templates/Corporate.dotx",
                true,
            );
            let settings = DocumentSettings::extract_from_part(&part).unwrap();
            let attached = settings.attached_template().unwrap();
            assert_eq!(attached.relationship_id(), "customRel");
            assert_eq!(attached.target_uri(), "file:///templates/Corporate.dotx");
        }
    }

    #[test]
    fn rejects_invalid_attached_template_graphs() {
        let xml = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:attachedTemplate r:id="customRel"/></w:settings>"#;
        let missing = BlobPart::new(
            PackURI::new("/word/settings.xml").unwrap(),
            litchi_opc::constants::content_type::WML_SETTINGS.to_owned(),
            xml.to_vec(),
        );
        assert!(DocumentSettings::extract_from_part(&missing).is_err());

        let wrong_type = attached_template_part(xml, "urn:wrong", "file:///a.dotx", true);
        assert!(DocumentSettings::extract_from_part(&wrong_type).is_err());
        let internal =
            attached_template_part(xml, ATTACHED_TEMPLATE_RELATIONSHIP, "template.dotx", false);
        assert!(DocumentSettings::extract_from_part(&internal).is_err());
        let whitespace = attached_template_part(
            xml,
            ATTACHED_TEMPLATE_RELATIONSHIP,
            "file:///bad path.dotx",
            true,
        );
        assert!(DocumentSettings::extract_from_part(&whitespace).is_err());

        let mut duplicate =
            attached_template_part(xml, ATTACHED_TEMPLATE_RELATIONSHIP, "file:///a.dotx", true);
        duplicate.rels_mut().add_relationship(
            STRICT_ATTACHED_TEMPLATE_RELATIONSHIP.to_owned(),
            "file:///b.dotx".to_owned(),
            "duplicate".to_owned(),
            true,
        );
        assert!(DocumentSettings::extract_from_part(&duplicate).is_err());
    }

    #[test]
    fn patches_only_the_attached_template_element() {
        let xml = br#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:opaque"><!--keep--><q:zoom q:percent="125"/><q:attachedTemplate rel:id="old"><x:ignored/></q:attachedTemplate><x:opaque><![CDATA[a < b]]></x:opaque></q:settings>"#;
        let replaced = patch_attached_template(xml, Some("new-id")).unwrap();
        assert_eq!(
            String::from_utf8(replaced).unwrap(),
            r#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:opaque"><!--keep--><q:zoom q:percent="125"/><q:attachedTemplate rel:id="new-id"/><x:opaque><![CDATA[a < b]]></x:opaque></q:settings>"#
        );

        let empty = br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"/>"#;
        let inserted =
            String::from_utf8(patch_attached_template(empty, Some("rId7")).unwrap()).unwrap();
        assert_eq!(
            inserted,
            r#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:attachedTemplate r:id="rId7" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"/></s:settings>"#
        );
    }

    #[test]
    fn patches_document_variables_without_touching_unrelated_settings() {
        let xml = br#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><!--before--><q:compat/><q:docVars><q:docVar q:name="old" q:val="old-value"/></q:docVars><x:opaque><![CDATA[a < b]]></x:opaque><q:rsids/></q:settings>"#;
        let mut variables = DocumentVariables::new();
        variables.insert("Company & Team", "A < B").unwrap();
        variables.insert("empty", "").unwrap();
        let patched =
            String::from_utf8(patch_document_variables(xml, &variables).unwrap()).unwrap();
        assert_eq!(
            patched,
            r#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><!--before--><q:compat/><q:docVars><q:docVar q:name="Company &amp; Team" q:val="A &lt; B"/><q:docVar q:name="empty" q:val=""/></q:docVars><x:opaque><![CDATA[a < b]]></x:opaque><q:rsids/></q:settings>"#
        );

        variables.clear();
        let removed =
            String::from_utf8(patch_document_variables(patched.as_bytes(), &variables).unwrap())
                .unwrap();
        assert_eq!(
            removed,
            r#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><!--before--><q:compat/><x:opaque><![CDATA[a < b]]></x:opaque><q:rsids/></q:settings>"#
        );
    }

    #[test]
    fn inserts_document_variables_into_empty_strict_root_in_schema_order() {
        let mut variables = DocumentVariables::new();
        variables.insert("strict", "value").unwrap();
        let empty = br#"<settings xmlns="http://purl.oclc.org/ooxml/wordprocessingml/main"/>"#;
        assert_eq!(
            String::from_utf8(patch_document_variables(empty, &variables).unwrap()).unwrap(),
            r#"<settings xmlns="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:docVars xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:docVar w:name="strict" w:val="value"/></w:docVars></settings>"#
        );

        let ordered = br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:compat/><s:rsids/></s:settings>"#;
        assert_eq!(
            String::from_utf8(patch_document_variables(ordered, &variables).unwrap()).unwrap(),
            r#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:compat/><s:docVars><s:docVar s:name="strict" s:val="value"/></s:docVars><s:rsids/></s:settings>"#
        );
    }
}
