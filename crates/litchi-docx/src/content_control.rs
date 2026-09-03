#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
#![expect(
    clippy::similar_names,
    reason = "domain names mirror distinct OOXML roles"
)]
#![expect(
    clippy::struct_excessive_bools,
    reason = "the public model preserves independent OOXML flags"
)]
#![expect(
    clippy::trivially_copy_pass_by_ref,
    reason = "the public API shape is retained for compatibility"
)]
#![expect(
    clippy::unnecessary_wraps,
    reason = "the Result signature preserves a uniform fallible codec API"
)]
use crate::error::{Error, Result};
/// Content control support for Word documents.
///
/// Content controls are structured regions in a document that can contain
/// specific types of content (text, dates, lists, etc.).
use crate::namespace::{is_wordprocessing_namespace, word_attribute_value};
use litchi_ooxml_common::custom_xml::valid_guid;
use litchi_ooxml_common::mce::{self, Capabilities};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;
use std::fmt;
use std::str::FromStr;

mod authoring;
mod model;
mod package;
mod patch;
mod snapshot;
mod transaction;

pub use authoring::{AuthoredProperties, AuthoringView, NamespaceRequirements, write_sdt_pr};
pub use model::{
    BindingFlavor, Checksum, ChecksumStatus, ChecksumValue, DataBinding,
    FORMATTING_ALLOWED_NAMESPACE, FormattingAllowed, Inventory, Limits, Lock, Occurrence,
    STORE_ITEM_CHECKSUM_NAMESPACE,
};
pub use package::{
    ChecksumEntry, PackageChecksumStatus, PackageCommit, PackageLimits, PackagePatch,
    PackageSnapshot, PackageTransaction, Story,
};
pub use patch::Patch;
pub use snapshot::{AttributeSpan, BindingSpan, LockSpan, Snapshot, SourceOccurrence, Span};
pub use transaction::{Commit, Edit, Transaction};

const WORD_2010_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/word/2010/wordml";
const WORD_2012_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/word/2012/wordml";
const WORD_2010_NAMESPACE_TEXT: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const WORD_2012_NAMESPACE_TEXT: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_SDT_ID: u32 = i32::MAX as u32;

/// Semantic kind of a Word content control.
///
/// The enum covers the complete set of core, Word 2010, and Word 2012 kind
/// markers recognized by this reader. A control with no marker is rich text.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Kind {
    /// Unrestricted rich text (`w:richText`), also the schema default.
    #[default]
    RichText,
    /// Plain text (`w:text`).
    Text,
    /// Picture (`w:picture`).
    Picture,
    /// Date (`w:date`).
    Date,
    /// Editable combo box (`w:comboBox`).
    ComboBox,
    /// Drop-down list (`w:dropDownList`).
    Dropdown,
    /// Citation (`w:citation`).
    Citation,
    /// Equation (`w:equation`).
    Equation,
    /// Group (`w:group`).
    Group,
    /// Document-part list (`w:docPartList`).
    DocPartList,
    /// Document-part object (`w:docPartObj`).
    DocPart,
    /// Bibliography (`w:bibliography`).
    Bibliography,
    /// Word 2010 checkbox (`w14:checkbox`).
    Checkbox,
    /// Word 2010 entity picker (`w14:entityPicker`).
    EntityPicker,
    /// Word 2012 repeating section (`w15:repeatingSection`).
    RepeatingSection,
    /// Word 2012 repeating-section item (`w15:repeatingSectionItem`).
    RepeatingItem,
}

/// Calendar used by a date content control.
///
/// `Umalqura` is the Word 2010 extension defined by `[MS-DOCX]` §2.2.7.
/// Its canonical writer form uses an MCE `AlternateContent` wrapper with a
/// `hijri` fallback. Unknown future tokens remain readable and writable.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum Calendar {
    Gregorian,
    Hijri,
    Umalqura,
    Hebrew,
    Taiwan,
    Japan,
    Thai,
    Korea,
    Saka,
    GregorianXlitEnglish,
    GregorianXlitFrench,
    GregorianUs,
    GregorianMeFrench,
    GregorianArabic,
    None,
    /// A calendar token not defined by the current schema vocabulary.
    Unknown(Box<str>),
}

impl Calendar {
    /// Return the exact schema token.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Gregorian => "gregorian",
            Self::Hijri => "hijri",
            Self::Umalqura => "umalqura",
            Self::Hebrew => "hebrew",
            Self::Taiwan => "taiwan",
            Self::Japan => "japan",
            Self::Thai => "thai",
            Self::Korea => "korea",
            Self::Saka => "saka",
            Self::GregorianXlitEnglish => "gregorianXlitEnglish",
            Self::GregorianXlitFrench => "gregorianXlitFrench",
            Self::GregorianUs => "gregorianUs",
            Self::GregorianMeFrench => "gregorianMeFrench",
            Self::GregorianArabic => "gregorianArabic",
            Self::None => "none",
            Self::Unknown(value) => value,
        }
    }

    fn from_xml(value: String) -> Result<Self> {
        let calendar = match value.as_str() {
            "gregorian" => Self::Gregorian,
            "hijri" => Self::Hijri,
            "umalqura" => Self::Umalqura,
            "hebrew" => Self::Hebrew,
            "taiwan" => Self::Taiwan,
            "japan" => Self::Japan,
            "thai" => Self::Thai,
            "korea" => Self::Korea,
            "saka" => Self::Saka,
            "gregorianXlitEnglish" => Self::GregorianXlitEnglish,
            "gregorianXlitFrench" => Self::GregorianXlitFrench,
            "gregorianUs" => Self::GregorianUs,
            "gregorianMeFrench" => Self::GregorianMeFrench,
            "gregorianArabic" => Self::GregorianArabic,
            "none" => Self::None,
            _ if value.is_empty() => {
                return Err(Error::InvalidFormat(
                    "content-control calendar token is empty".to_string(),
                ));
            },
            _ => Self::Unknown(value.into_boxed_str()),
        };
        Ok(calendar)
    }
}

/// Visual treatment requested for a structured document tag.
///
/// This is presentation metadata only.  Litchi neither renders controls nor
/// activates any associated Office Web Extension.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Appearance {
    /// Outline or shade the complete control region when needed.
    BoundingBox,
    /// Show the physical start and end tag characters.
    Tags,
    /// Do not show a visual indication for the control.
    Hidden,
}

impl Appearance {
    /// Return the schema token used by `w15:appearance`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::BoundingBox => "boundingBox",
            Self::Tags => "tags",
            Self::Hidden => "hidden",
        }
    }

    fn from_xml(value: &str) -> Result<Self> {
        match value {
            "boundingBox" => Ok(Self::BoundingBox),
            "tags" => Ok(Self::Tags),
            "hidden" => Ok(Self::Hidden),
            _ => Err(Error::InvalidFormat(format!(
                "invalid content-control appearance '{value}'"
            ))),
        }
    }
}

/// RGB color used as the visual basis of a structured document tag.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SdtColor {
    /// Let the consuming application choose the automatic color.
    Auto,
    /// A six-digit sRGB color.
    Rgb([u8; 3]),
}

impl SdtColor {
    /// Return the sRGB value when this is not the automatic color.
    #[must_use]
    pub const fn rgb(self) -> Option<[u8; 3]> {
        match self {
            Self::Auto => None,
            Self::Rgb(value) => Some(value),
        }
    }

    fn from_xml(value: &str) -> Result<Self> {
        if value == "auto" {
            return Ok(Self::Auto);
        }
        let bytes = value.as_bytes();
        if bytes.len() != 6 {
            return Err(Error::InvalidFormat(
                "content-control color must be 'auto' or six hexadecimal digits".to_string(),
            ));
        }
        let parse = |pair: &[u8]| {
            std::str::from_utf8(pair)
                .ok()
                .and_then(|value| u8::from_str_radix(value, 16).ok())
        };
        match (parse(&bytes[..2]), parse(&bytes[2..4]), parse(&bytes[4..])) {
            (Some(red), Some(green), Some(blue)) => Ok(Self::Rgb([red, green, blue])),
            _ => Err(Error::InvalidFormat(
                "content-control color must be 'auto' or six hexadecimal digits".to_string(),
            )),
        }
    }
}

/// Effective inert Office Web Extension provenance of a content control.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum WebExtensionBinding {
    /// The control was created by an Office Web Extension. This takes
    /// precedence over `webExtensionLinked` when both markers are present.
    Created(bool),
    /// The control is linked to an Office Web Extension.
    Linked(bool),
}

impl Kind {
    /// Return the exact OOXML marker name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::RichText => "richText",
            Self::Text => "text",
            Self::Picture => "picture",
            Self::Date => "date",
            Self::ComboBox => "comboBox",
            Self::Dropdown => "dropDownList",
            Self::Citation => "citation",
            Self::Equation => "equation",
            Self::Group => "group",
            Self::DocPartList => "docPartList",
            Self::DocPart => "docPartObj",
            Self::Bibliography => "bibliography",
            Self::Checkbox => "checkbox",
            Self::EntityPicker => "entityPicker",
            Self::RepeatingSection => "repeatingSection",
            Self::RepeatingItem => "repeatingSectionItem",
        }
    }
}

impl fmt::Display for Kind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl FromStr for Kind {
    type Err = ParseKindError;

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        match value {
            "richText" => Ok(Self::RichText),
            "text" => Ok(Self::Text),
            "picture" => Ok(Self::Picture),
            "date" => Ok(Self::Date),
            "comboBox" => Ok(Self::ComboBox),
            "dropDownList" => Ok(Self::Dropdown),
            "citation" => Ok(Self::Citation),
            "equation" => Ok(Self::Equation),
            "group" => Ok(Self::Group),
            "docPartList" => Ok(Self::DocPartList),
            "docPartObj" => Ok(Self::DocPart),
            "bibliography" => Ok(Self::Bibliography),
            "checkbox" => Ok(Self::Checkbox),
            "entityPicker" => Ok(Self::EntityPicker),
            "repeatingSection" => Ok(Self::RepeatingSection),
            "repeatingSectionItem" => Ok(Self::RepeatingItem),
            _ => Err(ParseKindError),
        }
    }
}

/// An unknown content-control kind token.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParseKindError;

impl fmt::Display for ParseKindError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid content-control kind")
    }
}

impl std::error::Error for ParseKindError {}

/// A content control in a Word document.
///
/// Content controls provide structured content regions that can be
/// bound to data or restricted to specific content types.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// for control in doc.content_controls()? {
///     if let Some(tag) = control.tag() {
///         println!("Control {}: {}", control.id(), tag);
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct ContentControl {
    /// Control ID
    id: Option<u32>,
    /// Control tag (optional identifier)
    tag: Option<String>,
    /// Control title
    title: Option<String>,
    /// Semantic control kind.
    kind: Kind,
    /// Whether the control can be deleted
    lock_delete: bool,
    /// Whether the content can be edited
    lock_content: bool,
    /// Whether Word should remove the control after its contents are edited.
    temporary: bool,
    /// Whether the control is currently displaying placeholder content.
    showing_placeholder: bool,
    /// Building-block name used for the placeholder.
    placeholder: Option<String>,
    /// Keyboard tab order.
    tab_index: Option<u32>,
    /// `XPath` of the custom XML data binding.
    data_bindings: Vec<DataBinding>,
    /// Word 2024 formatting exception attached to a content lock.
    formatting_allowed: Option<FormattingAllowed>,
    /// Display text and values declared by combo-box or drop-down controls.
    list_items: Vec<(String, String)>,
    /// Checked state for a Word 2010 checkbox control.
    checked: Option<bool>,
    /// Display format for a date control.
    date_format: Option<String>,
    /// ISO date value stored on a date control.
    date_value: Option<String>,
    /// Calendar used by a date control.
    date_calendar: Option<Calendar>,
    /// Title of a Word 2012 repeating section.
    repeating_section_title: Option<String>,
    /// Word 2012 visual treatment.
    appearance: Option<Appearance>,
    /// Word 2012 visual base color.
    color: Option<SdtColor>,
    /// Inert Word 2012 web-extension linked marker.
    web_extension_linked: Option<bool>,
    /// Inert Word 2012 web-extension-created marker.
    web_extension_created: Option<bool>,
}

impl ContentControl {
    /// Create a new `ContentControl`.
    #[must_use]
    pub fn new(
        id: u32,
        tag: Option<String>,
        title: Option<String>,
        kind: Option<Kind>,
        lock_delete: bool,
        lock_content: bool,
    ) -> Self {
        Self {
            id: Some(id),
            tag,
            title,
            kind: kind.unwrap_or_default(),
            lock_delete,
            lock_content,
            temporary: false,
            showing_placeholder: false,
            placeholder: None,
            tab_index: None,
            data_bindings: Vec::new(),
            formatting_allowed: None,
            list_items: Vec::new(),
            checked: None,
            date_format: None,
            date_value: None,
            date_calendar: None,
            repeating_section_title: None,
            appearance: None,
            color: None,
            web_extension_linked: None,
            web_extension_created: None,
        }
    }

    /// Get the control ID.
    #[inline]
    #[must_use]
    pub fn id(&self) -> u32 {
        self.id.unwrap_or_default()
    }

    /// Get the optional source ID without conflating a missing ID with zero.
    #[inline]
    #[must_use]
    pub const fn id_opt(&self) -> Option<u32> {
        self.id
    }

    /// Get the control tag.
    #[inline]
    #[must_use]
    pub fn tag(&self) -> Option<&str> {
        self.tag.as_deref()
    }

    /// Get the control title.
    #[inline]
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Get the semantic control kind.
    #[inline]
    #[must_use]
    pub fn kind(&self) -> Kind {
        self.kind
    }

    /// Check if the control is locked for deletion.
    #[inline]
    #[must_use]
    pub fn is_lock_delete(&self) -> bool {
        self.lock_delete
    }

    /// Check if the content is locked for editing.
    #[inline]
    #[must_use]
    pub fn is_lock_content(&self) -> bool {
        self.lock_content
    }

    /// Get the typed lock state.
    #[must_use]
    pub const fn lock(&self) -> Lock {
        match (self.lock_delete, self.lock_content) {
            (false, false) => Lock::Unlocked,
            (true, false) => Lock::SdtLocked,
            (false, true) => Lock::ContentLocked,
            (true, true) => Lock::SdtContentLocked,
        }
    }

    /// Check whether Word should remove the control after it is edited.
    #[inline]
    #[must_use]
    pub fn is_temporary(&self) -> bool {
        self.temporary
    }

    /// Check whether the control is displaying placeholder content.
    #[inline]
    #[must_use]
    pub fn is_showing_placeholder(&self) -> bool {
        self.showing_placeholder
    }

    /// Get the building-block name used for placeholder content.
    #[inline]
    #[must_use]
    pub fn placeholder(&self) -> Option<&str> {
        self.placeholder.as_deref()
    }

    /// Get the keyboard tab order.
    #[inline]
    #[must_use]
    pub fn tab_index(&self) -> Option<u32> {
        self.tab_index
    }

    /// Get the `XPath` of the custom XML data binding.
    #[inline]
    pub fn data_binding_xpath(&self) -> Option<&str> {
        self.data_binding().map(DataBinding::xpath)
    }

    /// Get the custom XML data-store item identifier.
    #[inline]
    pub fn data_binding_store_item_id(&self) -> Option<&str> {
        self.data_binding().map(DataBinding::store_item_id)
    }

    /// Get namespace prefix mappings used by the data-binding `XPath`.
    #[inline]
    pub fn data_binding_prefix_mappings(&self) -> Option<&str> {
        self.data_binding().and_then(DataBinding::prefix_mappings)
    }

    /// Get the complete typed inert data binding.
    #[inline]
    #[must_use]
    pub fn data_binding(&self) -> Option<&DataBinding> {
        self.data_bindings
            .iter()
            .find(|binding| binding.flavor() == BindingFlavor::Core)
            .or_else(|| self.data_bindings.first())
    }

    /// Get every exact binding occurrence in source order.
    #[inline]
    #[must_use]
    pub fn data_bindings(&self) -> &[DataBinding] {
        &self.data_bindings
    }

    /// Get the Word 2024 formatting exception, preserving absence.
    #[inline]
    #[must_use]
    pub const fn formatting_allowed(&self) -> Option<FormattingAllowed> {
        self.formatting_allowed
    }

    /// Validate binding metadata without evaluating the `XPath` or resolving URIs.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn validate_data_binding(&self) -> Result<()> {
        for binding in &self.data_bindings {
            validate_data_binding_values(
                binding.xpath(),
                binding.store_item_id(),
                binding.prefix_mappings(),
            )?;
        }
        Ok(())
    }

    /// Get the display text and values declared by a list control.
    #[inline]
    #[must_use]
    pub fn list_items(&self) -> &[(String, String)] {
        &self.list_items
    }

    /// Get the checked state of a checkbox control.
    #[inline]
    #[must_use]
    pub fn checked(&self) -> Option<bool> {
        self.checked
    }

    /// Get the display format of a date control.
    #[inline]
    #[must_use]
    pub fn date_format(&self) -> Option<&str> {
        self.date_format.as_deref()
    }

    /// Get the ISO date value stored on a date control.
    #[inline]
    #[must_use]
    pub fn date_value(&self) -> Option<&str> {
        self.date_value.as_deref()
    }

    /// Get the calendar selected by a date control, preserving absence.
    #[inline]
    #[must_use]
    pub fn date_calendar(&self) -> Option<&Calendar> {
        self.date_calendar.as_ref()
    }

    /// Get the title of a repeating-section control.
    #[inline]
    #[must_use]
    pub fn repeating_section_title(&self) -> Option<&str> {
        self.repeating_section_title.as_deref()
    }

    /// Return the requested visual treatment of this control.
    #[inline]
    #[must_use]
    pub const fn appearance(&self) -> Option<Appearance> {
        self.appearance
    }

    /// Return the requested visual base color of this control.
    #[inline]
    #[must_use]
    pub const fn color(&self) -> Option<SdtColor> {
        self.color
    }

    /// Return the exact inert `webExtensionLinked` marker, preserving absence.
    #[inline]
    #[must_use]
    pub const fn web_extension_linked(&self) -> Option<bool> {
        self.web_extension_linked
    }

    /// Return the exact inert `webExtensionCreated` marker, preserving absence.
    #[inline]
    #[must_use]
    pub const fn web_extension_created(&self) -> Option<bool> {
        self.web_extension_created
    }

    /// Return the effective inert Web Extension provenance marker.
    ///
    /// Per [MS-DOCX] 2.5.1.12-13, `webExtensionCreated` takes precedence
    /// when both markers are present. No relationship is resolved or run.
    #[inline]
    #[must_use]
    pub const fn web_extension_binding(&self) -> Option<WebExtensionBinding> {
        if let Some(value) = self.web_extension_created {
            Some(WebExtensionBinding::Created(value))
        } else if let Some(value) = self.web_extension_linked {
            Some(WebExtensionBinding::Linked(value))
        } else {
            None
        }
    }

    fn metadata_bytes(&self) -> usize {
        let mut total = 0usize;
        for value in [
            self.tag.as_deref(),
            self.title.as_deref(),
            self.placeholder.as_deref(),
            self.date_format.as_deref(),
            self.date_value.as_deref(),
            self.date_calendar.as_ref().map(Calendar::as_str),
            self.repeating_section_title.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            total = total.saturating_add(value.len());
        }
        for binding in &self.data_bindings {
            total = total
                .saturating_add(binding.xpath().len())
                .saturating_add(binding.store_item_id().len())
                .saturating_add(binding.prefix_mappings().map_or(0, str::len))
                .saturating_add(
                    binding
                        .checksum_value()
                        .map_or(0, |value| value.lexical().len()),
                );
        }
        for (display, value) in &self.list_items {
            total = total
                .saturating_add(display.len())
                .saturating_add(value.len());
        }
        total
    }

    /// Extract content controls from document XML bytes.
    pub(crate) fn extract_from_document(doc_xml: &[u8]) -> Result<Vec<ContentControl>> {
        Ok(Inventory::parse(doc_xml)?
            .into_controls()
            .into_iter()
            .filter(|control| control.id_opt().is_some())
            .collect())
    }
}

pub(crate) fn parse_inventory(doc_xml: &[u8], limits: &Limits) -> Result<Inventory> {
    limits.validate()?;
    if doc_xml.len() > limits.max_input_bytes {
        return Err(limit("input bytes"));
    }
    let capabilities = content_control_capabilities();
    validate_extension_ignorable(doc_xml, limits, &capabilities)?;
    let mce_limits = mce::Limits {
        max_input_bytes: limits.max_input_bytes,
        max_output_bytes: limits.max_mce_output_bytes,
        max_depth: limits.max_depth,
        max_namespace_bindings: limits.max_bindings,
        max_directive_tokens: limits.max_bindings,
        max_choices_per_alternate: limits.max_content_controls,
    };
    let selected = mce::process_markup_compatibility(doc_xml, &capabilities, &mce_limits)?.xml;
    if selected.len() > limits.max_mce_output_bytes {
        return Err(limit("MCE output bytes"));
    }
    parse_selected_inventory(selected.as_ref(), limits)
}

pub(crate) fn content_control_capabilities() -> Capabilities {
    let mut capabilities = Capabilities::default();
    for namespace in [
        WORD_2010_NAMESPACE_TEXT,
        WORD_2012_NAMESPACE_TEXT,
        STORE_ITEM_CHECKSUM_NAMESPACE,
        FORMATTING_ALLOWED_NAMESPACE,
    ] {
        capabilities.understand_namespace(namespace);
    }
    capabilities
}

fn parse_selected_inventory(doc_xml: &[u8], limits: &Limits) -> Result<Inventory> {
    let mut reader = NsReader::from_reader(doc_xml);
    let mut occurrences = Vec::new();
    let mut pending = None;
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut bindings = 0usize;
    let mut list_items = 0usize;
    let mut metadata_bytes = 0usize;

    loop {
        events = events.checked_add(1).ok_or_else(|| limit("event count"))?;
        if events > limits.max_events {
            return Err(limit("event count"));
        }
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
                    Error::InvalidFormat("content-control XML nesting is too deep".to_string())
                })?;
                if depth > limits.max_depth {
                    return Err(limit("XML depth"));
                }
                if is_word_element(&namespace, &element, b"sdtPr") {
                    if pending.is_some() {
                        return Err(Error::InvalidFormat(
                            "nested content-control properties are invalid".to_string(),
                        ));
                    }
                    if occurrences.len() >= limits.max_content_controls {
                        return Err(limit("content-control count"));
                    }
                    pending = Some(PendingContentControl::new(
                        depth,
                        limits.max_list_items_per_control,
                    ));
                } else if let Some(control) = pending.as_mut() {
                    parse_property(
                        &namespace, &element, decoder, &resolver, depth, false, control,
                    )?;
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("content-control XML nesting is too deep".to_string())
                })?;
                if is_word_element(&namespace, &element, b"sdtPr") {
                    if pending.is_some() {
                        return Err(Error::InvalidFormat(
                            "nested content-control properties are invalid".to_string(),
                        ));
                    }
                    push_occurrence(
                        &mut occurrences,
                        &mut bindings,
                        &mut list_items,
                        &mut metadata_bytes,
                        PendingContentControl::new(child_depth, limits.max_list_items_per_control)
                            .finish()?,
                        limits,
                    )?;
                } else if let Some(control) = pending.as_mut() {
                    parse_property(
                        &namespace,
                        &element,
                        decoder,
                        &resolver,
                        child_depth,
                        true,
                        control,
                    )?;
                }
            },
            Event::End(element) => {
                if pending.as_ref().is_some_and(|control| {
                    control.depth == depth
                        && is_wordprocessing_namespace(&namespace)
                        && element.local_name().as_ref() == b"sdtPr"
                }) {
                    let control = pending.take().ok_or_else(|| {
                        Error::InvalidFormat("missing content-control properties".to_string())
                    })?;
                    let control = control.finish()?;
                    push_occurrence(
                        &mut occurrences,
                        &mut bindings,
                        &mut list_items,
                        &mut metadata_bytes,
                        control,
                        limits,
                    )?;
                } else if let Some(control) = pending.as_mut()
                    && control
                        .context
                        .is_some_and(|(_, context_depth)| context_depth == depth)
                {
                    control.context = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid content-control XML nesting".to_string())
                })?;
            },
            Event::Eof if pending.is_some() => {
                return Err(Error::InvalidFormat(
                    "unterminated content-control properties".to_string(),
                ));
            },
            Event::Eof if depth != 0 => {
                return Err(Error::InvalidFormat(
                    "unterminated document XML".to_string(),
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }

    Ok(Inventory { occurrences })
}

fn push_occurrence(
    occurrences: &mut Vec<Occurrence>,
    bindings: &mut usize,
    list_items: &mut usize,
    metadata_bytes: &mut usize,
    control: ContentControl,
    limits: &Limits,
) -> Result<()> {
    if occurrences.len() >= limits.max_content_controls {
        return Err(limit("content-control count"));
    }
    if !control.data_bindings().is_empty() {
        *bindings = bindings
            .checked_add(control.data_bindings().len())
            .ok_or_else(|| limit("binding count"))?;
        if *bindings > limits.max_bindings {
            return Err(limit("binding count"));
        }
    }
    *list_items = list_items
        .checked_add(control.list_items().len())
        .ok_or_else(|| limit("list-item count"))?;
    if *list_items > limits.max_list_items {
        return Err(limit("list-item count"));
    }
    *metadata_bytes = metadata_bytes
        .checked_add(control.metadata_bytes())
        .ok_or_else(|| limit("metadata bytes"))?;
    if *metadata_bytes > limits.max_metadata_bytes {
        return Err(limit("metadata bytes"));
    }
    let ordinal = occurrences.len();
    occurrences
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "content-control inventory",
            source,
        })?;
    occurrences.push(Occurrence { ordinal, control });
    Ok(())
}

/// Validate the lexical form of an SDT data binding without executing `XPath`.
///
/// # Errors
///
/// Returns an error if the operation cannot be completed.
pub fn validate_data_binding_values(
    xpath: &str,
    store_item_id: &str,
    prefix_mappings: Option<&str>,
) -> Result<()> {
    const MAX_BINDING_BYTES: usize = 64 * 1024;
    if xpath.is_empty()
        || xpath.len() > MAX_BINDING_BYTES
        || xpath.chars().any(|character| {
            character == '\0' || character.is_control() && !character.is_whitespace()
        })
    {
        return Err(Error::InvalidFormat(
            "content-control XPath is empty or exceeds lexical limits".to_string(),
        ));
    }
    if !valid_guid(store_item_id) {
        return Err(Error::InvalidFormat(format!(
            "content-control storeItemID '{store_item_id}' is not ST_Guid"
        )));
    }
    let Some(mut remaining) = prefix_mappings else {
        return Ok(());
    };
    if remaining.len() > MAX_BINDING_BYTES {
        return Err(Error::InvalidFormat(
            "content-control prefixMappings exceeds lexical limits".to_string(),
        ));
    }
    let mut prefixes = HashSet::new();
    while !remaining.trim_start().is_empty() {
        remaining = remaining.trim_start();
        let after_xmlns = remaining.strip_prefix("xmlns").ok_or_else(|| {
            Error::InvalidFormat("prefixMappings requires xmlns declarations".to_string())
        })?;
        let (prefix, after_prefix) = if let Some(after_colon) = after_xmlns.strip_prefix(':') {
            let end = after_colon.find('=').ok_or_else(|| {
                Error::InvalidFormat("prefixMappings declaration has no '='".to_string())
            })?;
            let prefix = &after_colon[..end];
            if prefix.is_empty()
                || !prefix.bytes().enumerate().all(|(index, byte)| {
                    byte.is_ascii_alphabetic()
                        || byte == b'_'
                        || index > 0 && (byte.is_ascii_digit() || matches!(byte, b'.' | b'-'))
                })
            {
                return Err(Error::InvalidFormat(
                    "prefixMappings contains an invalid namespace prefix".to_string(),
                ));
            }
            (prefix, &after_colon[end..])
        } else {
            ("", after_xmlns)
        };
        if !prefixes.insert(prefix.to_string()) {
            return Err(Error::InvalidFormat(
                "prefixMappings contains a duplicate namespace prefix".to_string(),
            ));
        }
        let after_equals = after_prefix.strip_prefix('=').ok_or_else(|| {
            Error::InvalidFormat("prefixMappings declaration has no '='".to_string())
        })?;
        let quote = after_equals.as_bytes().first().copied().ok_or_else(|| {
            Error::InvalidFormat("prefixMappings declaration has no URI".to_string())
        })?;
        if !matches!(quote, b'\'' | b'"') {
            return Err(Error::InvalidFormat(
                "prefixMappings URI must be quoted".to_string(),
            ));
        }
        let quoted = &after_equals[1..];
        let end = quoted.find(quote as char).ok_or_else(|| {
            Error::InvalidFormat("prefixMappings URI quote is not closed".to_string())
        })?;
        if quoted[..end].is_empty() {
            return Err(Error::InvalidFormat(
                "prefixMappings namespace URI must not be empty".to_string(),
            ));
        }
        remaining = &quoted[end + 1..];
        if !remaining.is_empty()
            && !remaining
                .as_bytes()
                .first()
                .is_some_and(u8::is_ascii_whitespace)
        {
            return Err(Error::InvalidFormat(
                "prefixMappings declarations must be whitespace separated".to_string(),
            ));
        }
    }
    Ok(())
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum PropertyContext {
    Placeholder,
    Date,
    List,
    Checkbox,
    RepeatingSection,
}

struct PendingContentControl {
    depth: usize,
    id: Option<u32>,
    tag: Option<String>,
    title: Option<String>,
    kind: Option<Kind>,
    lock: Option<(bool, bool)>,
    temporary: Option<bool>,
    showing_placeholder: Option<bool>,
    placeholder: Option<String>,
    placeholder_seen: bool,
    tab_index: Option<u32>,
    data_bindings: Vec<DataBinding>,
    formatting_allowed: Option<FormattingAllowed>,
    list_items: Vec<(String, String)>,
    max_list_items: usize,
    checked: Option<bool>,
    date_format: Option<String>,
    date_value: Option<String>,
    date_calendar: Option<Calendar>,
    repeating_section_title: Option<String>,
    appearance: Option<Appearance>,
    color: Option<SdtColor>,
    web_extension_linked: Option<bool>,
    web_extension_created: Option<bool>,
    context: Option<(PropertyContext, usize)>,
}

impl PendingContentControl {
    fn new(depth: usize, max_list_items: usize) -> Self {
        Self {
            depth,
            id: None,
            tag: None,
            title: None,
            kind: None,
            lock: None,
            temporary: None,
            showing_placeholder: None,
            placeholder: None,
            placeholder_seen: false,
            tab_index: None,
            data_bindings: Vec::new(),
            formatting_allowed: None,
            list_items: Vec::new(),
            max_list_items,
            checked: None,
            date_format: None,
            date_value: None,
            date_calendar: None,
            repeating_section_title: None,
            appearance: None,
            color: None,
            web_extension_linked: None,
            web_extension_created: None,
            context: None,
        }
    }

    fn finish(self) -> Result<ContentControl> {
        let (lock_delete, lock_content) = self.lock.unwrap_or((false, false));
        Ok(ContentControl {
            id: self.id,
            tag: self.tag,
            title: self.title,
            kind: self.kind.unwrap_or_default(),
            lock_delete,
            lock_content,
            temporary: self.temporary.unwrap_or(false),
            showing_placeholder: self.showing_placeholder.unwrap_or(false),
            placeholder: self.placeholder,
            tab_index: self.tab_index,
            data_bindings: self.data_bindings,
            formatting_allowed: self.formatting_allowed,
            list_items: self.list_items,
            checked: self.checked,
            date_format: self.date_format,
            date_value: self.date_value,
            date_calendar: self.date_calendar,
            repeating_section_title: self.repeating_section_title,
            appearance: self.appearance,
            color: self.color,
            web_extension_linked: self.web_extension_linked,
            web_extension_created: self.web_extension_created,
        })
    }
}

fn parse_property(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    depth: usize,
    empty: bool,
    control: &mut PendingContentControl,
) -> Result<()> {
    if depth == control.depth + 1 {
        parse_direct_property(namespace, element, decoder, resolver, depth, empty, control)
    } else if depth == control.depth + 2 {
        parse_nested_property(namespace, element, decoder, resolver, control)
    } else {
        Ok(())
    }
}

fn parse_direct_property(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    depth: usize,
    empty: bool,
    control: &mut PendingContentControl,
) -> Result<()> {
    let name = element.local_name();
    if is_wordprocessing_namespace(namespace) {
        match name.as_ref() {
            b"id" => {
                let id = required_u32(element, b"val", decoder, resolver, "content-control ID")?;
                if id > MAX_SDT_ID {
                    return Err(Error::InvalidFormat(
                        "content-control ID exceeds the Int32Value maximum".to_string(),
                    ));
                }
                set_once(&mut control.id, id, "content-control ID")?;
            },
            b"tag" => {
                let value = required_word_attribute(element, b"val", decoder, resolver, "tag")?;
                set_once(&mut control.tag, value, "content-control tag")?;
            },
            b"alias" => {
                let value = required_word_attribute(element, b"val", decoder, resolver, "alias")?;
                set_once(&mut control.title, value, "content-control alias")?;
            },
            b"lock" => parse_lock(element, decoder, resolver, control)?,
            b"temporary" => {
                let value = parse_on_off(element, decoder, resolver)?;
                set_once(
                    &mut control.temporary,
                    value,
                    "content-control temporary property",
                )?;
            },
            b"showingPlcHdr" => {
                let value = parse_on_off(element, decoder, resolver)?;
                set_once(
                    &mut control.showing_placeholder,
                    value,
                    "content-control placeholder-display property",
                )?;
            },
            b"tabIndex" => {
                let value = required_u32(element, b"val", decoder, resolver, "tab index")?;
                set_once(&mut control.tab_index, value, "content-control tab index")?;
            },
            b"dataBinding" => {
                parse_data_binding(element, decoder, resolver, BindingFlavor::Core, control)?;
            },
            b"placeholder" => {
                if control.placeholder_seen {
                    return Err(Error::InvalidFormat(
                        "duplicate content-control placeholder".to_string(),
                    ));
                }
                control.placeholder_seen = true;
                set_context(control, PropertyContext::Placeholder, depth, empty);
            },
            b"date" => {
                set_kind(control, Kind::Date)?;
                control.date_value = word_attribute_value(element, b"fullDate", decoder, resolver)?;
                set_context(control, PropertyContext::Date, depth, empty);
            },
            b"comboBox" => {
                set_kind(control, Kind::ComboBox)?;
                set_context(control, PropertyContext::List, depth, empty);
            },
            b"dropDownList" => {
                set_kind(control, Kind::Dropdown)?;
                set_context(control, PropertyContext::List, depth, empty);
            },
            b"text" => set_kind(control, Kind::Text)?,
            b"picture" => set_kind(control, Kind::Picture)?,
            b"citation" => set_kind(control, Kind::Citation)?,
            b"equation" => set_kind(control, Kind::Equation)?,
            b"group" => set_kind(control, Kind::Group)?,
            b"docPartList" => set_kind(control, Kind::DocPartList)?,
            b"docPartObj" => set_kind(control, Kind::DocPart)?,
            b"bibliography" => set_kind(control, Kind::Bibliography)?,
            b"richText" => set_kind(control, Kind::RichText)?,
            _ => {},
        }
    } else if is_extension_namespace(namespace, WORD_2010_NAMESPACE) {
        match name.as_ref() {
            b"checkbox" => {
                set_kind(control, Kind::Checkbox)?;
                set_context(control, PropertyContext::Checkbox, depth, empty);
            },
            b"entityPicker" => set_kind(control, Kind::EntityPicker)?,
            _ => {},
        }
    } else if is_extension_namespace(namespace, WORD_2012_NAMESPACE) {
        match name.as_ref() {
            b"appearance" => {
                let value = extension_attribute_value(
                    element,
                    b"val",
                    WORD_2012_NAMESPACE,
                    b"w15",
                    decoder,
                    resolver,
                )?
                .ok_or_else(|| {
                    Error::InvalidFormat("content-control appearance has no value".to_string())
                })?;
                set_once(
                    &mut control.appearance,
                    Appearance::from_xml(&value)?,
                    "content-control appearance",
                )?;
            },
            b"color" => {
                let value = required_word_attribute(
                    element,
                    b"val",
                    decoder,
                    resolver,
                    "content-control color",
                )?;
                set_once(
                    &mut control.color,
                    SdtColor::from_xml(&value)?,
                    "content-control color",
                )?;
            },
            b"dataBinding" => {
                parse_data_binding(element, decoder, resolver, BindingFlavor::Word2012, control)?;
            },
            b"repeatingSection" => {
                set_kind(control, Kind::RepeatingSection)?;
                set_context(control, PropertyContext::RepeatingSection, depth, empty);
            },
            b"repeatingSectionItem" => set_kind(control, Kind::RepeatingItem)?,
            b"webExtensionLinked" => {
                set_once(
                    &mut control.web_extension_linked,
                    parse_on_off(element, decoder, resolver)?,
                    "content-control web-extension linked marker",
                )?;
            },
            b"webExtensionCreated" => {
                set_once(
                    &mut control.web_extension_created,
                    parse_on_off(element, decoder, resolver)?,
                    "content-control web-extension created marker",
                )?;
            },
            _ => {},
        }
    }
    Ok(())
}

fn parse_nested_property(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    control: &mut PendingContentControl,
) -> Result<()> {
    let Some((context, _)) = control.context else {
        return Ok(());
    };
    let name = element.local_name();
    match context {
        PropertyContext::Placeholder if is_word_element(namespace, element, b"docPart") => {
            let value = required_word_attribute(element, b"val", decoder, resolver, "placeholder")?;
            set_once(
                &mut control.placeholder,
                value,
                "content-control placeholder",
            )?;
        },
        PropertyContext::Date if is_word_element(namespace, element, b"dateFormat") => {
            let value = required_word_attribute(element, b"val", decoder, resolver, "date format")?;
            set_once(
                &mut control.date_format,
                value,
                "content-control date format",
            )?;
        },
        PropertyContext::Date if is_word_element(namespace, element, b"calendar") => {
            let value = required_word_attribute(element, b"val", decoder, resolver, "calendar")?;
            set_once(
                &mut control.date_calendar,
                Calendar::from_xml(value)?,
                "content-control calendar",
            )?;
        },
        PropertyContext::List if is_word_element(namespace, element, b"listItem") => {
            if control.list_items.len() >= control.max_list_items {
                return Err(limit("per-control list-item count"));
            }
            let value = required_word_attribute(element, b"value", decoder, resolver, "list item")?;
            let display = word_attribute_value(element, b"displayText", decoder, resolver)?
                .unwrap_or_else(|| value.clone());
            control
                .list_items
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "content-control list items",
                    source,
                })?;
            control.list_items.push((display, value));
        },
        PropertyContext::Checkbox
            if is_extension_namespace(namespace, WORD_2010_NAMESPACE)
                && name.as_ref() == b"checked" =>
        {
            let value = extension_attribute_value(
                element,
                b"val",
                WORD_2010_NAMESPACE,
                b"w14",
                decoder,
                resolver,
            )?;
            set_once(
                &mut control.checked,
                value.as_deref().map_or(Ok(true), parse_on_off_value)?,
                "checkbox state",
            )?;
        },
        PropertyContext::RepeatingSection
            if is_extension_namespace(namespace, WORD_2012_NAMESPACE)
                && name.as_ref() == b"sectionTitle" =>
        {
            let value =
                required_word_attribute(element, b"val", decoder, resolver, "section title")?;
            set_once(
                &mut control.repeating_section_title,
                value,
                "repeating-section title",
            )?;
        },
        PropertyContext::Placeholder
        | PropertyContext::Date
        | PropertyContext::List
        | PropertyContext::Checkbox
        | PropertyContext::RepeatingSection => {},
    }
    Ok(())
}

fn parse_lock(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    control: &mut PendingContentControl,
) -> Result<()> {
    if control.lock.is_some() {
        return Err(Error::InvalidFormat(
            "duplicate content-control lock".to_string(),
        ));
    }
    let value = required_word_attribute(element, b"val", decoder, resolver, "lock")?;
    let lock = match value.as_str() {
        "unlocked" => Lock::Unlocked,
        "sdtLocked" => Lock::SdtLocked,
        "contentLocked" => Lock::ContentLocked,
        "sdtContentLocked" => Lock::SdtContentLocked,
        _ => {
            return Err(Error::InvalidFormat(format!(
                "invalid content-control lock value '{value}'"
            )));
        },
    };
    let formatting = exact_extension_attribute_value(
        element,
        b"formattingAllowed",
        FORMATTING_ALLOWED_NAMESPACE.as_bytes(),
        decoder,
        resolver,
    )?;
    if let Some(value) = formatting {
        if !lock.locks_content() {
            return Err(Error::InvalidFormat(
                "formattingAllowed requires contentLocked or sdtContentLocked".to_string(),
            ));
        }
        control.formatting_allowed =
            Some(FormattingAllowed::from_bool(parse_on_off_value(&value)?));
    }
    control.lock = Some((lock.locks_control(), lock.locks_content()));
    Ok(())
}

fn parse_data_binding(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    flavor: BindingFlavor,
    control: &mut PendingContentControl,
) -> Result<()> {
    let xpath =
        required_word_attribute(element, b"xpath", decoder, resolver, "data-binding XPath")?;
    let store_item_id = required_word_attribute(
        element,
        b"storeItemID",
        decoder,
        resolver,
        "data-binding store item ID",
    )?;
    let prefix_mappings = word_attribute_value(element, b"prefixMappings", decoder, resolver)?;
    let checksum = exact_extension_attribute_value(
        element,
        b"storeItemChecksum",
        STORE_ITEM_CHECKSUM_NAMESPACE.as_bytes(),
        decoder,
        resolver,
    )?
    .map(ChecksumValue::from_source);
    control
        .data_bindings
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "content-control data bindings",
            source,
        })?;
    control.data_bindings.push(DataBinding::from_parsed(
        flavor,
        xpath,
        store_item_id,
        prefix_mappings,
        checksum,
    ));
    Ok(())
}

fn set_context(
    control: &mut PendingContentControl,
    context: PropertyContext,
    depth: usize,
    empty: bool,
) {
    if !empty {
        control.context = Some((context, depth));
    }
}

fn set_kind(control: &mut PendingContentControl, value: Kind) -> Result<()> {
    if let Some(existing) = control.kind {
        return Err(Error::InvalidFormat(format!(
            "multiple content-control types '{existing}' and '{value}'"
        )));
    }
    control.kind = Some(value);
    Ok(())
}

fn set_once<T>(target: &mut Option<T>, value: T, description: &str) -> Result<()> {
    if target.is_some() {
        return Err(Error::InvalidFormat(format!("duplicate {description}")));
    }
    *target = Some(value);
    Ok(())
}

fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<u32> {
    let value = required_word_attribute(element, name, decoder, resolver, description)?;
    value.parse::<u32>().map_err(|_source_error| {
        Error::InvalidFormat(format!("invalid {description} value '{value}'"))
    })
}

fn required_word_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<String> {
    word_attribute_value(element, name, decoder, resolver)?
        .ok_or_else(|| Error::InvalidFormat(format!("missing {description} attribute")))
}

fn parse_on_off(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<bool> {
    match word_attribute_value(element, b"val", decoder, resolver)? {
        Some(value) => parse_on_off_value(&value),
        None => Ok(true),
    }
}

fn parse_on_off_value(value: &str) -> Result<bool> {
    match value {
        "true" | "1" | "on" => Ok(true),
        "false" | "0" | "off" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid on/off value '{value}'"
        ))),
    }
}

fn extension_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    namespace: &[u8],
    conventional_prefix: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (resolved, _) = resolver.resolve_attribute(attribute.key);
        let matches = matches!(resolved, ResolveResult::Bound(Namespace(value)) if value == namespace)
            || matches!(resolved, ResolveResult::Unknown(prefix) if prefix.as_slice() == conventional_prefix);
        if !matches {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(format!(
                "duplicate extension attribute '{}'",
                String::from_utf8_lossy(name)
            )));
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

fn exact_extension_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    namespace: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (resolved, _) = resolver.resolve_attribute(attribute.key);
        if !matches!(resolved, ResolveResult::Bound(Namespace(value)) if value == namespace) {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(format!(
                "duplicate extension attribute '{}'",
                String::from_utf8_lossy(name)
            )));
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

#[derive(Clone, Copy, Default)]
struct IgnorableState {
    word_2010: bool,
    checksum: bool,
    formatting: bool,
    word_2012: bool,
}

impl IgnorableState {
    fn include(&mut self, namespace: &[u8]) {
        self.word_2010 |= namespace == WORD_2010_NAMESPACE;
        self.checksum |= namespace == STORE_ITEM_CHECKSUM_NAMESPACE.as_bytes();
        self.formatting |= namespace == FORMATTING_ALLOWED_NAMESPACE.as_bytes();
        self.word_2012 |= namespace == WORD_2012_NAMESPACE;
    }

    fn includes(self, namespace: &[u8]) -> bool {
        match namespace {
            value if value == WORD_2010_NAMESPACE => self.word_2010,
            value if value == STORE_ITEM_CHECKSUM_NAMESPACE.as_bytes() => self.checksum,
            value if value == FORMATTING_ALLOWED_NAMESPACE.as_bytes() => self.formatting,
            value if value == WORD_2012_NAMESPACE => self.word_2012,
            _ => false,
        }
    }
}

#[derive(Default)]
struct AlternateState {
    selected: bool,
}

struct ActivityFrame {
    active: bool,
    alternate: Option<AlternateState>,
}

fn validate_extension_ignorable(
    xml: &[u8],
    limits: &Limits,
    capabilities: &Capabilities,
) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    let mut stack: Vec<IgnorableState> = Vec::new();
    let mut activity: Vec<ActivityFrame> = Vec::new();
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut metadata = 0usize;
    let mut namespace_bindings = 0usize;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| limit("MCE validation event count"))?;
        if events > limits.max_events {
            return Err(limit("MCE validation event count"));
        }
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| limit("XML depth"))?;
                if depth > limits.max_depth {
                    return Err(limit("XML depth"));
                }
                let active = element_is_active(
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    &mut activity,
                    capabilities,
                )?;
                let mut effective = stack.last().copied().unwrap_or_default();
                extend_ignorable(
                    &element,
                    decoder,
                    &resolver,
                    &mut effective,
                    limits,
                    &mut metadata,
                    &mut namespace_bindings,
                )?;
                if active {
                    require_ignorable_extensions(&namespace, &element, &resolver, &effective)?;
                }
                stack.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "content-control MCE scope",
                    source,
                })?;
                stack.push(effective);
                activity
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "content-control MCE activity",
                        source,
                    })?;
                activity.push(ActivityFrame {
                    active,
                    alternate: is_mce_element(&namespace, &element, b"AlternateContent")
                        .then(AlternateState::default),
                });
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| limit("XML depth"))?;
                if child_depth > limits.max_depth {
                    return Err(limit("XML depth"));
                }
                let active = element_is_active(
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    &mut activity,
                    capabilities,
                )?;
                let mut effective = stack.last().copied().unwrap_or_default();
                extend_ignorable(
                    &element,
                    decoder,
                    &resolver,
                    &mut effective,
                    limits,
                    &mut metadata,
                    &mut namespace_bindings,
                )?;
                if active {
                    require_ignorable_extensions(&namespace, &element, &resolver, &effective)?;
                }
            },
            Event::End(_) => {
                stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("invalid content-control XML nesting".to_string())
                })?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid content-control XML nesting".to_string())
                })?;
                activity.pop().ok_or_else(|| {
                    Error::InvalidFormat("invalid content-control XML nesting".to_string())
                })?;
            },
            Event::Eof if depth != 0 => {
                return Err(Error::InvalidFormat(
                    "unterminated content-control XML".to_string(),
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(())
}

fn element_is_active(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    activity: &mut [ActivityFrame],
    capabilities: &Capabilities,
) -> Result<bool> {
    let Some(parent) = activity.last_mut() else {
        return Ok(true);
    };
    let Some(alternate) = parent.alternate.as_mut() else {
        return Ok(parent.active);
    };
    if is_mce_element(namespace, element, b"Choice") {
        let supported = choice_is_supported(element, decoder, resolver, capabilities)?;
        let active = parent.active && !alternate.selected && supported;
        alternate.selected |= active;
        Ok(active)
    } else if is_mce_element(namespace, element, b"Fallback") {
        let active = parent.active && !alternate.selected;
        alternate.selected |= active;
        Ok(active)
    } else {
        Ok(parent.active)
    }
}

fn choice_is_supported(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    capabilities: &Capabilities,
) -> Result<bool> {
    let mut requires = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() == b"Requires" {
            requires = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                    .map_err(|error| Error::Xml(error.to_string()))?,
            );
        }
    }
    let Some(requires) = requires else {
        return Ok(false);
    };
    for prefix in requires.split_whitespace() {
        let mut qualified = Vec::new();
        qualified
            .try_reserve_exact(prefix.len() + 2)
            .map_err(|source| Error::Allocation {
                resource: "content-control MCE Requires prefix",
                source,
            })?;
        qualified.extend_from_slice(prefix.as_bytes());
        qualified.extend_from_slice(b":_");
        let (namespace, _) = resolver.resolve(QName(&qualified), false);
        let ResolveResult::Bound(Namespace(namespace)) = namespace else {
            return Ok(false);
        };
        let Ok(namespace) = std::str::from_utf8(namespace) else {
            return Ok(false);
        };
        if !capabilities.understands(namespace) {
            return Ok(false);
        }
    }
    Ok(!requires.trim().is_empty())
}

fn is_mce_element(namespace: &ResolveResult<'_>, element: &BytesStart<'_>, local: &[u8]) -> bool {
    is_extension_namespace(namespace, MCE_NAMESPACE) && element.local_name().as_ref() == local
}

fn extend_ignorable(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    effective: &mut IgnorableState,
    limits: &Limits,
    metadata: &mut usize,
    namespace_bindings: &mut usize,
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"Ignorable" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == MCE_NAMESPACE) {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        *metadata = metadata
            .checked_add(value.len())
            .ok_or_else(|| limit("MCE metadata bytes"))?;
        if *metadata > limits.max_metadata_bytes {
            return Err(limit("MCE metadata bytes"));
        }
        for prefix in value.split_whitespace() {
            let mut qualified = Vec::new();
            qualified
                .try_reserve_exact(prefix.len() + 2)
                .map_err(|source| Error::Allocation {
                    resource: "content-control MCE prefix",
                    source,
                })?;
            qualified.extend_from_slice(prefix.as_bytes());
            qualified.extend_from_slice(b":_");
            let (namespace, _) = resolver.resolve(QName(&qualified), false);
            let ResolveResult::Bound(Namespace(namespace)) = namespace else {
                return Err(Error::InvalidFormat(format!(
                    "mc:Ignorable prefix '{prefix}' is not bound"
                )));
            };
            *namespace_bindings = namespace_bindings
                .checked_add(1)
                .ok_or_else(|| limit("MCE namespace bindings"))?;
            if *namespace_bindings > limits.max_bindings {
                return Err(limit("MCE namespace bindings"));
            }
            effective.include(namespace);
        }
    }
    Ok(())
}

fn require_ignorable_extensions(
    element_namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    effective: &IgnorableState,
) -> Result<()> {
    let local_name = element.local_name();
    let required_owner_namespace = if is_extension_namespace(element_namespace, WORD_2010_NAMESPACE)
        && matches!(local_name.as_ref(), b"checkbox" | b"entityPicker")
    {
        Some((WORD_2010_NAMESPACE, "Word 2010 content-control property"))
    } else if is_extension_namespace(element_namespace, WORD_2012_NAMESPACE)
        && matches!(
            local_name.as_ref(),
            b"dataBinding" | b"repeatingSection" | b"repeatingSectionItem"
        )
    {
        Some((WORD_2012_NAMESPACE, "Word 2012 content-control property"))
    } else {
        None
    };
    if let Some((namespace, description)) = required_owner_namespace
        && !effective.includes(namespace)
    {
        return Err(Error::InvalidFormat(format!(
            "{description} namespace is not declared in effective mc:Ignorable"
        )));
    }
    let mut checksum_seen = false;
    let mut formatting_seen = false;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let local = attribute.key.local_name();
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let (expected_namespace, expected_owner, seen, description) =
            match (&namespace, local.as_ref()) {
                (ResolveResult::Bound(Namespace(value)), b"storeItemChecksum")
                    if *value == STORE_ITEM_CHECKSUM_NAMESPACE.as_bytes() =>
                {
                    (
                        STORE_ITEM_CHECKSUM_NAMESPACE.as_bytes(),
                        b"dataBinding".as_slice(),
                        &mut checksum_seen,
                        "storeItemChecksum",
                    )
                },
                (ResolveResult::Bound(Namespace(value)), b"formattingAllowed")
                    if *value == FORMATTING_ALLOWED_NAMESPACE.as_bytes() =>
                {
                    (
                        FORMATTING_ALLOWED_NAMESPACE.as_bytes(),
                        b"lock".as_slice(),
                        &mut formatting_seen,
                        "formattingAllowed",
                    )
                },
                _ => continue,
            };
        if *seen {
            return Err(Error::InvalidFormat(format!(
                "duplicate expanded-name {description} attribute"
            )));
        }
        *seen = true;
        let owner_namespace_valid = is_wordprocessing_namespace(element_namespace)
            || description == "storeItemChecksum"
                && is_extension_namespace(element_namespace, WORD_2012_NAMESPACE);
        if !owner_namespace_valid || element.local_name().as_ref() != expected_owner {
            return Err(Error::InvalidFormat(format!(
                "{description} is attached to an invalid element"
            )));
        }
        if !effective.includes(expected_namespace) {
            return Err(Error::InvalidFormat(format!(
                "{description} namespace is not declared in effective mc:Ignorable"
            )));
        }
    }
    Ok(())
}

fn limit(resource: &str) -> Error {
    Error::InvalidFormat(format!("content-control {resource} limit exceeded"))
}

fn is_word_element(namespace: &ResolveResult<'_>, element: &BytesStart<'_>, name: &[u8]) -> bool {
    is_wordprocessing_namespace(namespace) && element.local_name().as_ref() == name
}

fn is_extension_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::mem::size_of;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    #[test]
    fn test_content_control_creation() {
        let control = ContentControl::new(
            1,
            Some("field1".to_string()),
            Some("My Field".to_string()),
            Some(Kind::Text),
            false,
            false,
        );

        assert_eq!(control.id(), 1);
        assert_eq!(control.tag(), Some("field1"));
        assert_eq!(control.title(), Some("My Field"));
        assert_eq!(control.kind(), Kind::Text);
        assert!(!control.is_lock_delete());
        assert!(!control.is_lock_content());

        let defaulted = ContentControl::new(2, None, None, None, false, false);
        assert_eq!(defaulted.kind(), Kind::RichText);
    }

    #[test]
    fn kind_tokens_are_complete_compact_and_round_trip() {
        let kinds = [
            Kind::RichText,
            Kind::Text,
            Kind::Picture,
            Kind::Date,
            Kind::ComboBox,
            Kind::Dropdown,
            Kind::Citation,
            Kind::Equation,
            Kind::Group,
            Kind::DocPartList,
            Kind::DocPart,
            Kind::Bibliography,
            Kind::Checkbox,
            Kind::EntityPicker,
            Kind::RepeatingSection,
            Kind::RepeatingItem,
        ];
        for kind in kinds {
            assert_eq!(kind.as_str().parse::<Kind>().unwrap(), kind);
            assert_eq!(kind.to_string(), kind.as_str());
        }
        assert_eq!(size_of::<Kind>(), 1);
        assert_eq!(Kind::default(), Kind::RichText);
        assert!("vendorControl".parse::<Kind>().is_err());
    }

    #[test]
    fn extracts_namespaced_properties_and_decodes_values() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:sdt><w:sdtPr><w:id w:val="42"/><w:tag w:val="customer&amp;id"/><w:alias w:val="Customer &amp; address"/><w:tabIndex w:val="7"/><w:temporary/><w:showingPlcHdr w:val="on"/><w:placeholder><w:docPart w:val="DefaultPlaceholder_1"/></w:placeholder><w:dataBinding w:prefixMappings="xmlns:x='urn:test&amp;more'" w:xpath="/x:root/x:name" w:storeItemID="{{ABC}}"/><w:dropDownList><w:listItem w:displayText="A &amp; B" w:value="ab"/><w:listItem w:value="fallback"/></w:dropDownList><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent/></w:sdt></w:body></w:document>"#
        );
        let controls = ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        let control = &controls[0];
        assert_eq!(control.id(), 42);
        assert_eq!(control.tag(), Some("customer&id"));
        assert_eq!(control.title(), Some("Customer & address"));
        assert_eq!(control.kind(), Kind::Dropdown);
        assert_eq!(control.tab_index(), Some(7));
        assert!(control.is_temporary());
        assert!(control.is_showing_placeholder());
        assert_eq!(control.placeholder(), Some("DefaultPlaceholder_1"));
        assert_eq!(control.data_binding_xpath(), Some("/x:root/x:name"));
        assert_eq!(control.data_binding_store_item_id(), Some("{ABC}"));
        assert_eq!(
            control.data_binding_prefix_mappings(),
            Some("xmlns:x='urn:test&more'")
        );
        assert_eq!(
            control.list_items(),
            &[
                ("A & B".into(), "ab".into()),
                ("fallback".into(), "fallback".into())
            ]
        );
        assert!(control.is_lock_delete());
        assert!(control.is_lock_content());
    }

    #[test]
    fn accepts_strict_and_aliased_word_namespaces() {
        let xml = r#"<x:document xmlns:x="http://purl.oclc.org/ooxml/wordprocessingml/main"><x:sdtPr><x:id x:val="1"/><x:text/></x:sdtPr><x:sdtPr><x:id x:val="2"/></x:sdtPr></x:document>"#;
        let controls = ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        assert_eq!(controls.len(), 2);
        assert_eq!(controls[0].kind(), Kind::Text);
        assert_eq!(controls[1].kind(), Kind::RichText);
    }

    #[test]
    fn extracts_checkbox_date_and_repeating_section_metadata() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:c="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="c r"><w:sdtPr><w:id w:val="1"/><c:checkbox><c:checked/></c:checkbox></w:sdtPr><w:sdtPr><w:id w:val="2"/><w:date w:fullDate="2026-07-14T00:00:00Z"><w:dateFormat w:val="yyyy-MM-dd"/></w:date></w:sdtPr><w:sdtPr><w:id w:val="3"/><r:repeatingSection><r:sectionTitle w:val="People"/></r:repeatingSection></w:sdtPr><w:sdtPr><w:id w:val="4"/><c:entityPicker/></w:sdtPr><w:sdtPr><w:id w:val="5"/><r:repeatingSectionItem/></w:sdtPr></w:document>"#
        );
        let controls = ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        assert_eq!(controls[0].kind(), Kind::Checkbox);
        assert_eq!(controls[0].checked(), Some(true));
        assert_eq!(controls[1].kind(), Kind::Date);
        assert_eq!(controls[1].date_value(), Some("2026-07-14T00:00:00Z"));
        assert_eq!(controls[1].date_format(), Some("yyyy-MM-dd"));
        assert_eq!(controls[2].kind(), Kind::RepeatingSection);
        assert_eq!(controls[2].repeating_section_title(), Some("People"));
        assert_eq!(controls[3].kind(), Kind::EntityPicker);
        assert_eq!(controls[4].kind(), Kind::RepeatingItem);
    }

    #[test]
    fn umalqura_calendar_uses_the_active_mce_choice() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"><w:sdtPr><w:id w:val="7"/><w:date><mc:AlternateContent><mc:Choice Requires="w14"><w:calendar w:val="umalqura"/></mc:Choice><mc:Fallback><w:calendar w:val="hijri"/></mc:Fallback></mc:AlternateContent></w:date></w:sdtPr></w:document>"#
        );
        let controls = ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        assert_eq!(controls[0].date_calendar(), Some(&Calendar::Umalqura));
    }

    #[test]
    fn unknown_date_calendar_tokens_remain_typed_and_lossless() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}"><w:sdtPr><w:id w:val="8"/><w:date><w:calendar w:val="futureCalendar"/></w:date></w:sdtPr></w:document>"#
        );
        let controls = ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        assert_eq!(
            controls[0].date_calendar().map(Calendar::as_str),
            Some("futureCalendar")
        );
    }

    #[test]
    fn ignores_foreign_lookalikes_and_idless_controls() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:f="urn:foreign"><f:sdtPr><f:id f:val="8"/><f:text/></f:sdtPr><w:sdtPr><f:id f:val="9"/><f:tag f:val="spoof"/></w:sdtPr><w:sdtPr/><w:sdtPr><w:id w:val="10"/><f:text/><w:vendorKind/></w:sdtPr></w:document>"#
        );
        let controls = ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        assert_eq!(controls.len(), 1);
        assert_eq!(controls[0].id(), 10);
        assert_eq!(controls[0].kind(), Kind::RichText);
    }

    #[test]
    fn recognizes_all_standard_type_markers() {
        let types = [
            "text",
            "picture",
            "citation",
            "equation",
            "group",
            "docPartList",
            "docPartObj",
            "bibliography",
            "richText",
            "comboBox",
        ];
        for (index, kind) in types.iter().enumerate() {
            let xml = format!(
                r#"<w:sdtPr xmlns:w="{W}"><w:id w:val="{}"/><w:{kind}/></w:sdtPr>"#,
                index + 1
            );
            let controls = ContentControl::extract_from_document(xml.as_bytes()).unwrap();
            assert_eq!(controls[0].kind().as_str(), *kind);
        }
    }

    #[test]
    fn rejects_invalid_or_duplicate_properties() {
        let invalid = [
            r#"<w:sdtPr xmlns:w="W"><w:id/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="x"/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:id w:val="2"/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:lock w:val="invalid"/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:text/><w:picture/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:text/><w:text/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:temporary w:val="maybe"/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:temporary/><w:temporary w:val="0"/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:lock w:val="unlocked"/><w:lock w:val="unlocked"/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:placeholder/><w:placeholder/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:dataBinding w:xpath="/x"/></w:sdtPr>"#,
        ];
        for xml in invalid {
            let xml = xml.replace("xmlns:w=\"W\"", &format!("xmlns:w=\"{W}\""));
            assert!(ContentControl::extract_from_document(xml.as_bytes()).is_err());
        }

        let duplicate = format!(
            r#"<w:document xmlns:w="{W}"><w:sdtPr><w:id w:val="1"/></w:sdtPr>
                <w:sdtPr><w:id w:val="1"/></w:sdtPr></w:document>"#
        );
        assert_eq!(
            ContentControl::extract_from_document(duplicate.as_bytes())
                .unwrap()
                .len(),
            2
        );
    }

    #[test]
    fn rejects_truncated_properties() {
        let xml = format!(r#"<w:document xmlns:w="{W}"><w:sdtPr><w:id w:val="1"/>"#);
        assert!(ContentControl::extract_from_document(xml.as_bytes()).is_err());
    }

    #[test]
    fn checksum_is_strict_canonical_base64_over_little_endian_word_value() {
        let checksum = Checksum::from_word_value(0xBD0B_E338);
        assert_eq!(checksum.as_bytes(), &[0x38, 0xE3, 0x0B, 0xBD]);
        assert_eq!(checksum.to_base64(), "OOMLvQ==");
        assert_eq!(
            Checksum::parse("OOMLvQ==").unwrap().word_value(),
            0xBD0B_E338
        );
        assert_eq!(
            Checksum::compute(b"123456789", &Limits::default()).unwrap(),
            checksum
        );
        for invalid in ["", "OOMLvQ=", "OOMLvQ===", "OOMLvQ==\n", "OOMLvQ--"] {
            assert!(Checksum::parse(invalid).is_err(), "accepted {invalid:?}");
        }
    }

    #[test]
    fn inventories_missing_and_duplicate_ids_as_distinct_occurrences() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}"><w:sdtPr/><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtPr><w:id w:val="7"/></w:sdtPr></w:document>"#
        );
        let inventory = Inventory::parse(xml.as_bytes()).unwrap();
        assert_eq!(
            inventory
                .occurrences()
                .iter()
                .map(Occurrence::id)
                .collect::<Vec<_>>(),
            [None, Some(7), Some(7)]
        );
        assert_eq!(
            ContentControl::extract_from_document(xml.as_bytes())
                .unwrap()
                .len(),
            2
        );
    }

    #[test]
    fn parses_checksum_by_expanded_name_and_retains_malformed_lexical_state() {
        let xml = format!(
            r#"<w:sdtPr xmlns:w="{W}" xmlns:mc="{mce}" xmlns:h="{hash}" mc:Ignorable="h"><w:id w:val="1"/><w:dataBinding w:xpath="/x" w:storeItemID="{{ABC}}" h:storeItemChecksum="bad"/></w:sdtPr>"#,
            mce = std::str::from_utf8(MCE_NAMESPACE).unwrap(),
            hash = STORE_ITEM_CHECKSUM_NAMESPACE,
        );
        let inventory = Inventory::parse(xml.as_bytes()).unwrap();
        let binding = inventory.occurrences()[0].control().data_binding().unwrap();
        assert!(binding.checksum().is_none());
        assert!(
            matches!(binding.checksum_status(), ChecksumStatus::Malformed(value) if &*value == "bad")
        );

        let missing_ignorable = xml.replace(" mc:Ignorable=\"h\"", "");
        assert!(Inventory::parse(missing_ignorable.as_bytes()).is_err());
        let fake_namespace = xml.replace(STORE_ITEM_CHECKSUM_NAMESPACE, "urn:fake");
        assert!(Inventory::parse(fake_namespace.as_bytes()).is_ok());
        assert!(
            Inventory::parse(fake_namespace.as_bytes())
                .unwrap()
                .occurrences()[0]
                .control()
                .data_binding()
                .unwrap()
                .checksum_value()
                .is_none()
        );
    }

    #[test]
    fn formatting_allowed_accepts_all_on_off_tokens_only_on_content_locks() {
        for (token, expected) in [
            ("true", true),
            ("1", true),
            ("on", true),
            ("false", false),
            ("0", false),
            ("off", false),
        ] {
            let xml = format!(
                r#"<w:sdtPr xmlns:w="{W}" xmlns:mc="{mce}" xmlns:f="{format}" mc:Ignorable="f"><w:id w:val="1"/><w:lock w:val="contentLocked" f:formattingAllowed="{token}"/></w:sdtPr>"#,
                mce = std::str::from_utf8(MCE_NAMESPACE).unwrap(),
                format = FORMATTING_ALLOWED_NAMESPACE,
            );
            let control = Inventory::parse(xml.as_bytes())
                .unwrap()
                .into_controls()
                .remove(0);
            assert_eq!(control.formatting_allowed().unwrap().as_bool(), expected);
        }
        let invalid = format!(
            r#"<w:sdtPr xmlns:w="{W}" xmlns:mc="{mce}" xmlns:f="{format}" mc:Ignorable="f"><w:lock w:val="sdtLocked" f:formattingAllowed="true"/></w:sdtPr>"#,
            mce = std::str::from_utf8(MCE_NAMESPACE).unwrap(),
            format = FORMATTING_ALLOWED_NAMESPACE,
        );
        assert!(Inventory::parse(invalid.as_bytes()).is_err());
    }

    #[test]
    fn selects_only_the_active_mce_branch() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{mce}" xmlns:u="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="u"><w:sdtPr><w:id w:val="1"/></w:sdtPr></mc:Choice><mc:Fallback><w:sdtPr><w:id w:val="2"/></w:sdtPr></mc:Fallback></mc:AlternateContent></w:document>"#,
            mce = std::str::from_utf8(MCE_NAMESPACE).unwrap(),
        );
        let inventory = Inventory::parse(xml.as_bytes()).unwrap();
        assert_eq!(inventory.occurrences().len(), 1);
        assert_eq!(inventory.occurrences()[0].id(), Some(2));
    }

    #[test]
    fn applies_configurable_event_depth_control_and_metadata_limits() {
        let xml = format!(r#"<w:sdtPr xmlns:w="{W}"><w:id w:val="1"/></w:sdtPr>"#);
        let mut limits = Limits {
            max_content_controls: 1,
            ..Limits::default()
        };
        assert!(Inventory::parse_with_limits(xml.as_bytes(), &limits).is_ok());
        limits.max_events = 1;
        assert!(Inventory::parse_with_limits(xml.as_bytes(), &limits).is_err());
        limits = Limits::default();
        limits.max_depth = 1;
        assert!(Inventory::parse_with_limits(xml.as_bytes(), &limits).is_err());
        limits = Limits::default();
        limits.max_metadata_bytes = 1;
        let tagged = format!(r#"<w:sdtPr xmlns:w="{W}"><w:tag w:val="ab"/></w:sdtPr>"#);
        assert!(Inventory::parse_with_limits(tagged.as_bytes(), &limits).is_err());
    }

    #[test]
    fn mce_output_budget_is_independent_and_accepts_the_exact_boundary() {
        let mce = std::str::from_utf8(MCE_NAMESPACE).unwrap();
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{mce}" xmlns:u="urn:u"><mc:AlternateContent><mc:Choice Requires="u"><w:sdtPr><w:id w:val="1"/></w:sdtPr></mc:Choice><mc:Fallback><w:sdtPr><w:id w:val="2"/></w:sdtPr></mc:Fallback></mc:AlternateContent></w:document>"#
        );
        let exact_mce_output = mce::process_markup_compatibility(
            xml.as_bytes(),
            &Capabilities::default(),
            &mce::Limits::default(),
        )
        .unwrap()
        .xml
        .len();
        let mut limits = Limits {
            max_output_bytes: 1,
            ..Limits::default()
        };
        limits.max_mce_output_bytes = exact_mce_output;
        assert_eq!(
            Inventory::parse_with_limits(xml.as_bytes(), &limits)
                .unwrap()
                .occurrences()[0]
                .id(),
            Some(2)
        );

        limits.max_mce_output_bytes = exact_mce_output - 1;
        assert!(Inventory::parse_with_limits(xml.as_bytes(), &limits).is_err());
        limits.max_mce_output_bytes = 1;
        assert!(Inventory::parse_with_limits(xml.as_bytes(), &limits).is_err());
    }

    #[test]
    fn checksum_equality_hash_and_lexical_access_ignore_source_provenance() {
        use std::collections::hash_map::DefaultHasher;
        use std::hash::{Hash, Hasher};

        let authored = Checksum::from_word_value(0xBD0B_E338);
        let parsed = Checksum::parse("OOMLvQ==").unwrap();
        assert_eq!(authored, parsed);
        assert_eq!(authored.original_lexical(), None);
        assert_eq!(parsed.original_lexical(), Some("OOMLvQ=="));
        assert_eq!(ChecksumValue::Valid(authored.clone()).lexical(), "OOMLvQ==");
        let mut authored_hash = DefaultHasher::new();
        let mut parsed_hash = DefaultHasher::new();
        authored.hash(&mut authored_hash);
        parsed.hash(&mut parsed_hash);
        assert_eq!(authored_hash.finish(), parsed_hash.finish());
    }

    #[test]
    fn deep_many_prefix_ignorable_scopes_use_constant_semantic_state() {
        use std::fmt::Write as _;

        assert!(std::mem::size_of::<IgnorableState>() <= 4);
        const PREFIXES: usize = 64;
        const SCOPES: usize = 32;
        let tokens = (0..PREFIXES)
            .map(|index| format!("x{index}"))
            .collect::<Vec<_>>()
            .join(" ");
        let mut xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{}""#,
            std::str::from_utf8(MCE_NAMESPACE).unwrap()
        );
        for index in 0..PREFIXES {
            write!(&mut xml, r#" xmlns:x{index}="urn:{index}""#).unwrap();
        }
        write!(&mut xml, r#" mc:Ignorable="{tokens}">"#).unwrap();
        for _ in 0..SCOPES {
            write!(&mut xml, r#"<w:body mc:Ignorable="{tokens}">"#).unwrap();
        }
        xml.push_str("<w:sdtPr/>");
        for _ in 0..SCOPES {
            xml.push_str("</w:body>");
        }
        xml.push_str("</w:document>");

        let mut limits = Limits {
            max_depth: SCOPES + 2,
            ..Limits::default()
        };
        limits.max_bindings = PREFIXES * (SCOPES + 1);
        assert!(Inventory::parse_with_limits(xml.as_bytes(), &limits).is_ok());
        limits.max_bindings -= 1;
        assert!(Inventory::parse_with_limits(xml.as_bytes(), &limits).is_err());
    }

    #[test]
    fn list_item_limits_accept_exact_and_reject_next_allocation() {
        let one = r#"<w:listItem w:displayText="A" w:value="a"/>"#;
        let two = r#"<w:listItem w:displayText="B" w:value="b"/>"#;
        let single_control = format!(
            r#"<w:sdtPr xmlns:w="{W}"><w:dropDownList>{one}{two}</w:dropDownList></w:sdtPr>"#
        );
        let mut limits = Limits {
            max_list_items_per_control: 2,
            max_list_items: 2,
            ..Limits::default()
        };
        assert_eq!(
            Inventory::parse_with_limits(single_control.as_bytes(), &limits)
                .unwrap()
                .occurrences()[0]
                .control()
                .list_items()
                .len(),
            2
        );
        limits.max_list_items_per_control = 1;
        assert!(Inventory::parse_with_limits(single_control.as_bytes(), &limits).is_err());

        let two_controls = format!(
            r#"<w:document xmlns:w="{W}"><w:sdtPr><w:dropDownList>{one}</w:dropDownList></w:sdtPr><w:sdtPr><w:dropDownList>{two}</w:dropDownList></w:sdtPr></w:document>"#
        );
        limits.max_list_items_per_control = 1;
        limits.max_list_items = 2;
        assert!(Inventory::parse_with_limits(two_controls.as_bytes(), &limits).is_ok());
        limits.max_list_items = 1;
        assert!(Inventory::parse_with_limits(two_controls.as_bytes(), &limits).is_err());
    }

    #[test]
    fn known_extension_owners_require_expanded_inherited_ignorable_namespaces() {
        let mce = std::str::from_utf8(MCE_NAMESPACE).unwrap();
        let aliased_inherited = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{mce}" xmlns:a="{}" xmlns:b="{}" mc:Ignorable="a b"><w:body><w:sdtPr><a:checkbox><a:checked a:val="1"/></a:checkbox></w:sdtPr><w:sdtPr><b:repeatingSectionItem/></w:sdtPr></w:body></w:document>"#,
            std::str::from_utf8(WORD_2010_NAMESPACE).unwrap(),
            std::str::from_utf8(WORD_2012_NAMESPACE).unwrap(),
        );
        let inventory = Inventory::parse(aliased_inherited.as_bytes()).unwrap();
        assert_eq!(inventory.occurrences()[0].control().kind(), Kind::Checkbox);
        assert_eq!(
            inventory.occurrences()[1].control().kind(),
            Kind::RepeatingItem
        );

        let missing = format!(
            r#"<w:sdtPr xmlns:w="{W}" xmlns:a="{}"><a:entityPicker/></w:sdtPr>"#,
            std::str::from_utf8(WORD_2010_NAMESPACE).unwrap(),
        );
        assert!(Inventory::parse(missing.as_bytes()).is_err());

        let rebound = format!(
            r#"<w:sdtPr xmlns:w="{W}" xmlns:mc="{mce}" xmlns:a="urn:foreign" mc:Ignorable="a"><a:wrapper xmlns:a="{}"><a:repeatingSection/></a:wrapper></w:sdtPr>"#,
            std::str::from_utf8(WORD_2012_NAMESPACE).unwrap(),
        );
        assert!(Inventory::parse(rebound.as_bytes()).is_err());
    }
}
