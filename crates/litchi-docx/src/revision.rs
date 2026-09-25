#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::items_after_statements,
    reason = "the local helper remains adjacent to its sole use"
)]
#![expect(
    clippy::module_name_repetitions,
    reason = "public names retain established OOXML facade terminology"
)]
#![expect(
    clippy::struct_field_names,
    reason = "the public model retains established field names"
)]
/// Track changes (revisions) support for DOCX documents.
///
/// This module provides structures and functions for reading tracked changes
/// (revisions) from Word documents. Tracked changes record insertions, deletions,
/// moves, and formatting changes made by document editors.
///
/// # Architecture
///
/// - `Revision`: A single tracked change
/// - `RevisionType`: Type of change (insert, delete, move, format)
/// - `RevisionInfo`: Metadata about who made the change and when
///
/// # Example
///
/// ```rust,no_run
/// use litchi_docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// // Get all revisions from the document
/// for para in doc.paragraphs()? {
///     for revision in para.revisions()? {
///         println!("Revision by {}: {} - {}",
///             revision.author().unwrap_or("Unattributed"),
///             revision.revision_type(),
///             revision.text()
///         );
///         if let Some(date) = revision.date() {
///             println!("  Made on: {}", date);
///         }
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
use crate::error::{Error, Result};
use crate::namespace::is_wordprocessing_namespace;
use litchi_ooxml_common::properties::time::DateTime;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, NamespaceResolver, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;
use smallvec::SmallVec;
use std::fmt;

/// Word 2023 namespace for the UTC timestamp on tracked changes.
pub const WORD_2023_DATE_UTC_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/word/2023/wordml/word16du";

const WORD_2023_DATE_UTC_NAMESPACE_BYTES: &[u8] =
    b"http://schemas.microsoft.com/office/word/2023/wordml/word16du";

const MAX_REVISION_DATE_UTC_BYTES: usize = 128;

pub mod authoring;
pub mod conflict;
mod limits;
#[cfg(test)]
mod limits_tests;
pub use limits::Limits;
use limits::{charge, check};

/// Type of tracked change.
///
/// Represents the different types of revisions that can be tracked
/// in a Word document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RevisionType {
    /// Text insertion
    Insert,
    /// Text deletion
    Delete,
    /// Move from (cut)
    MoveFrom,
    /// Move to (paste)
    MoveTo,
    /// Formatting change
    FormatChange,
    /// Insertion of the paragraph mark.
    ParagraphMarkInsert,
    /// Deletion of the paragraph mark.
    ParagraphMarkDelete,
    /// Movement of the paragraph mark from its original location.
    ParagraphMarkMoveFrom,
    /// Movement of the paragraph mark to its new location.
    ParagraphMarkMoveTo,
    /// Insertion of paragraph numbering properties.
    NumberingInsert,
    /// Revision information for section properties.
    SectionPropertiesChange,
    /// Revision information for table grid columns (identifier only).
    TableGridChange,
    /// Revision information for table-level property exceptions on a row.
    TablePropertyExceptionsChange,
    /// Previous paragraph or field numbering properties (Transitional only).
    NumberingChange,
    /// Table property change
    TablePropertiesChange,
    /// Table row insertion
    RowInsert,
    /// Table row deletion
    RowDelete,
    /// Table row property change
    RowPropertiesChange,
    /// Table cell insertion
    CellInsert,
    /// Table cell deletion
    CellDelete,
    /// Table cell merge change
    CellMerge,
    /// Table cell property change
    CellPropertiesChange,
    /// Custom or unknown revision type
    Unknown,
}

impl fmt::Display for RevisionType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Insert => write!(f, "Insert"),
            Self::Delete => write!(f, "Delete"),
            Self::MoveFrom => write!(f, "Move From"),
            Self::MoveTo => write!(f, "Move To"),
            Self::FormatChange => write!(f, "Format Change"),
            Self::ParagraphMarkInsert => write!(f, "Paragraph Mark Insert"),
            Self::ParagraphMarkDelete => write!(f, "Paragraph Mark Delete"),
            Self::ParagraphMarkMoveFrom => write!(f, "Paragraph Mark Move From"),
            Self::ParagraphMarkMoveTo => write!(f, "Paragraph Mark Move To"),
            Self::NumberingInsert => write!(f, "Numbering Insert"),
            Self::SectionPropertiesChange => write!(f, "Section Properties Change"),
            Self::TableGridChange => write!(f, "Table Grid Change"),
            Self::TablePropertyExceptionsChange => write!(f, "Table Property Exceptions Change"),
            Self::NumberingChange => write!(f, "Numbering Change"),
            Self::TablePropertiesChange => write!(f, "Table Properties Change"),
            Self::RowInsert => write!(f, "Row Insert"),
            Self::RowDelete => write!(f, "Row Delete"),
            Self::RowPropertiesChange => write!(f, "Row Properties Change"),
            Self::CellInsert => write!(f, "Cell Insert"),
            Self::CellDelete => write!(f, "Cell Delete"),
            Self::CellMerge => write!(f, "Cell Merge"),
            Self::CellPropertiesChange => write!(f, "Cell Properties Change"),
            Self::Unknown => write!(f, "Unknown"),
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum RevisionScope {
    Other,
    ParagraphProperties,
    ParagraphMark,
    ParagraphMarkHistory,
    NumberingProperties { phase: u8 },
    FieldCharacter { has_content: bool },
    SectionProperties { changed: bool },
    TableGrid { changed: bool },
    TableExceptions { changed: bool },
    OriginalProperties,
    OriginalPropertyContent,
    History { kind: RevisionType, seen: bool },
    RowProperties,
    Properties,
    Marker,
}

fn is_inline_revision(kind: RevisionType) -> bool {
    matches!(
        kind,
        RevisionType::Insert | RevisionType::Delete | RevisionType::MoveFrom | RevisionType::MoveTo
    )
}

fn revision_special_character(local_name: &[u8], is_word: bool) -> Option<char> {
    if !is_word {
        return None;
    }
    match local_name {
        b"tab" => Some('\t'),
        b"br" | b"cr" => Some('\n'),
        b"noBreakHyphen" => Some('\u{2011}'),
        b"softHyphen" => Some('\u{ad}'),
        _ => None,
    }
}

fn revision_scope(
    local_name: &[u8],
    parent: RevisionScope,
    is_word: bool,
    kind: Option<RevisionType>,
) -> RevisionScope {
    if matches!(
        parent,
        RevisionScope::OriginalProperties | RevisionScope::OriginalPropertyContent
    ) {
        return RevisionScope::OriginalPropertyContent;
    }
    if !is_word {
        return RevisionScope::Other;
    }
    if matches!(
        kind,
        Some(
            RevisionType::RowInsert
                | RevisionType::RowDelete
                | RevisionType::CellInsert
                | RevisionType::CellDelete
                | RevisionType::CellMerge
                | RevisionType::NumberingInsert
                | RevisionType::NumberingChange
                | RevisionType::ParagraphMarkInsert
                | RevisionType::ParagraphMarkDelete
                | RevisionType::ParagraphMarkMoveFrom
                | RevisionType::ParagraphMarkMoveTo
        )
    ) {
        return RevisionScope::Marker;
    }
    if let Some(
        kind @ (RevisionType::SectionPropertiesChange
        | RevisionType::TableGridChange
        | RevisionType::TablePropertyExceptionsChange),
    ) = kind
    {
        return RevisionScope::History { kind, seen: false };
    }
    if matches!(parent, RevisionScope::History { .. }) {
        return RevisionScope::OriginalProperties;
    }
    match local_name {
        b"sectPr" => RevisionScope::SectionProperties { changed: false },
        b"tblGrid" => RevisionScope::TableGrid { changed: false },
        b"tblPrEx" => RevisionScope::TableExceptions { changed: false },
        b"fldChar" => RevisionScope::FieldCharacter { has_content: false },
        b"pPr" => RevisionScope::ParagraphProperties,
        b"numPr" => RevisionScope::NumberingProperties { phase: 0 },
        b"trPr" => RevisionScope::RowProperties,
        b"rPr"
            if matches!(
                parent,
                RevisionScope::ParagraphProperties | RevisionScope::ParagraphMarkHistory
            ) =>
        {
            RevisionScope::ParagraphMark
        },
        b"rPrChange" if parent == RevisionScope::ParagraphMark => {
            RevisionScope::ParagraphMarkHistory
        },
        b"rPr" | b"rPrChange" | b"pPrChange" | b"tblPr" | b"tblPrChange" | b"trPrChange"
        | b"tcPr" | b"tcPrChange" => RevisionScope::Properties,
        _ => RevisionScope::Other,
    }
}

/// A tracked change (revision) in a Word document.
///
/// Represents a single change tracked by Word's revision system.
/// Contains information about what changed, who made the change, and when.
///
/// # Field Ordering
///
/// Fields are ordered to maximize CPU cache line utilization:
/// - Strings (24 bytes each on 64-bit systems)
/// - Enums and smaller types
#[derive(Debug, Clone)]
pub struct Revision {
    /// Author who made the change
    author: Option<String>,

    /// Previous numbering representation, retained without evaluation.
    original_numbering: Option<String>,

    /// Date/time of the change (ISO 8601 format)
    date: Option<String>,

    /// UTC date/time of the change from the Word 2023 `w16du` extension.
    date_utc: Option<String>,

    /// Text content affected by this revision
    text: String,

    /// Revision ID
    id: String,

    /// Type of revision
    revision_type: RevisionType,
}

impl Revision {
    /// Create a new Revision.
    ///
    /// # Arguments
    ///
    /// * `revision_type` - Type of revision
    /// * `author` - Author who made the change
    /// * `date` - Date/time of the change
    /// * `id` - Revision ID
    #[inline]
    #[must_use]
    pub fn new(
        revision_type: RevisionType,
        author: String,
        date: Option<String>,
        id: String,
    ) -> Self {
        Self {
            author: Some(author),
            original_numbering: None,
            date,
            date_utc: None,
            text: String::new(),
            id,
            revision_type,
        }
    }

    /// Get the revision type.
    #[inline]
    #[must_use]
    pub fn revision_type(&self) -> RevisionType {
        self.revision_type
    }

    /// Get the author who made the change.
    ///
    /// Table-grid revisions carry an identifier only and return `None`.
    /// An explicitly empty tracked-change author remains `Some("")`.
    #[inline]
    #[must_use]
    pub fn author(&self) -> Option<&str> {
        self.author.as_deref()
    }

    /// Previous paragraph/field numbering from Transitional `numberingChange`.
    /// The value is inert metadata; it is never evaluated or reconstructed.
    #[must_use]
    pub fn original_numbering(&self) -> Option<&str> {
        self.original_numbering.as_deref()
    }

    /// Get the date/time of the change.
    #[inline]
    #[must_use]
    pub fn date(&self) -> Option<&str> {
        self.date.as_deref()
    }

    /// Get the Word 2023 UTC date/time of the change.
    #[inline]
    #[must_use]
    pub fn date_utc(&self) -> Option<&str> {
        self.date_utc.as_deref()
    }

    /// Get the revision ID.
    #[inline]
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Get the text content affected by this revision.
    #[inline]
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Set the text content.
    #[inline]
    pub fn set_text(&mut self, text: String) {
        self.text = text;
    }

    /// Append text content.
    #[inline]
    pub fn append_text(&mut self, text: &str) {
        self.text.push_str(text);
    }
}

/// Parse revisions from paragraph XML.
///
/// Extracts all tracked changes (w:ins, w:del, w:moveFrom, w:moveTo) from
/// the paragraph XML.
///
/// # Arguments
///
/// * `xml_bytes` - The raw XML bytes of the paragraph
///
/// # Performance
///
/// Uses streaming XML parsing with pre-allocated `SmallVec` for efficient
/// storage of typically small revision collections.
///
/// # Example XML Structure
///
/// ```xml
/// <w:p>
///   <w:r>
///     <w:t>Normal text</w:t>
///   </w:r>
///   <w:ins w:id="0" w:author="John Doe" w:date="2024-11-05T10:30:00Z">
///     <w:r>
///       <w:t>inserted text</w:t>
///     </w:r>
///   </w:ins>
///   <w:del w:id="1" w:author="Jane Smith" w:date="2024-11-05T11:00:00Z">
///     <w:r>
///       <w:delText>deleted text</w:delText>
///     </w:r>
///   </w:del>
/// </w:p>
/// ```
#[cfg_attr(
    not(test),
    expect(
        dead_code,
        reason = "the namespace-free parser facade remains useful to in-module fragment tests"
    )
)]
pub(crate) fn parse_revisions(xml_bytes: &[u8]) -> Result<SmallVec<[Revision; 4]>> {
    parse_revisions_with_context(xml_bytes, &[])
}

pub(crate) fn parse_revisions_with_context(
    xml_bytes: &[u8],
    inherited_namespaces: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<SmallVec<[Revision; 4]>> {
    parse_revisions_with_limits(xml_bytes, inherited_namespaces, Limits::default())
}

pub(crate) fn parse_revisions_with_limits(
    xml_bytes: &[u8],
    inherited_namespaces: &[(Option<Vec<u8>>, Vec<u8>)],
    limits: Limits,
) -> Result<SmallVec<[Revision; 4]>> {
    let limits = limits.validate()?;
    check("source bytes", xml_bytes.len(), limits.max_source_bytes)?;
    check(
        "inherited namespaces",
        inherited_namespaces.len(),
        limits.max_inherited_namespaces,
    )?;
    let mut namespace_bytes = 0;
    for (prefix, namespace) in inherited_namespaces {
        charge(
            "inherited namespace bytes",
            &mut namespace_bytes,
            prefix.as_ref().map_or(0, Vec::len),
            limits.max_inherited_namespace_bytes,
        )?;
        charge(
            "inherited namespace bytes",
            &mut namespace_bytes,
            namespace.len(),
            limits.max_inherited_namespace_bytes,
        )?;
    }
    let mut reader = NsReader::from_reader(xml_bytes);
    for (prefix, namespace) in inherited_namespaces {
        let prefix = prefix
            .as_deref()
            .map_or(PrefixDeclaration::Default, |prefix| {
                PrefixDeclaration::Named(prefix)
            });
        reader
            .resolver_mut()
            .add(prefix, Namespace(namespace))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    // Keep text exactly as authored.  Whitespace outside a tracked text
    // element is ignored below, while spaces inside w:t/w:delText are part of
    // the document content and must survive xml:space="preserve".
    reader.config_mut().trim_text(false);

    // Use SmallVec for efficient storage of typically small revision collections
    let mut revisions = SmallVec::new();

    // State tracking for parsing
    // Each active annotation references its pre-order result slot. A nested
    // annotation retains its own metadata and contributes text to its parents.
    let mut active: SmallVec<[(usize, usize); 4]> = SmallVec::new();
    let mut text_element_depth = None;
    let mut scopes: SmallVec<[RevisionScope; 16]> = SmallVec::new();
    let mut element_depth = 0usize;
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut events = 0;
    let mut metadata_bytes = 0;
    let mut text_bytes = 0;

    fn revision_type(
        local_name: &[u8],
        parent: RevisionScope,
        is_word_element: bool,
    ) -> Result<Option<RevisionType>> {
        if !is_word_element {
            return Ok(None);
        }
        let kind = match (local_name, parent) {
            (b"ins", RevisionScope::RowProperties) => RevisionType::RowInsert,
            (b"del", RevisionScope::RowProperties) => RevisionType::RowDelete,
            (b"ins", RevisionScope::ParagraphMark) => RevisionType::ParagraphMarkInsert,
            (b"del", RevisionScope::ParagraphMark) => RevisionType::ParagraphMarkDelete,
            (b"moveFrom", RevisionScope::ParagraphMark) => RevisionType::ParagraphMarkMoveFrom,
            (b"moveTo", RevisionScope::ParagraphMark) => RevisionType::ParagraphMarkMoveTo,
            (b"ins", RevisionScope::NumberingProperties { .. }) => RevisionType::NumberingInsert,
            (b"ins" | b"del" | b"moveFrom" | b"moveTo", scope) if scope != RevisionScope::Other => {
                return Err(Error::InvalidFormat(
                    "tracked revision has an invalid property owner".into(),
                ));
            },
            (b"ins", _) => RevisionType::Insert,
            (b"del", _) => RevisionType::Delete,
            (b"moveFrom", _) => RevisionType::MoveFrom,
            (b"moveTo", _) => RevisionType::MoveTo,
            (b"rPrChange" | b"pPrChange", _) => RevisionType::FormatChange,
            (b"sectPrChange", RevisionScope::SectionProperties { .. }) => {
                RevisionType::SectionPropertiesChange
            },
            (b"tblGridChange", RevisionScope::TableGrid { .. }) => RevisionType::TableGridChange,
            (b"tblPrExChange", RevisionScope::TableExceptions { .. }) => {
                RevisionType::TablePropertyExceptionsChange
            },
            (
                b"numberingChange",
                RevisionScope::NumberingProperties { .. } | RevisionScope::FieldCharacter { .. },
            ) => RevisionType::NumberingChange,
            (b"sectPrChange" | b"tblGridChange" | b"tblPrExChange" | b"numberingChange", _) => {
                return Err(Error::InvalidFormat(
                    "revision has an invalid property owner".into(),
                ));
            },
            (b"tblPrChange", _) => RevisionType::TablePropertiesChange,
            (b"trPrChange", _) => RevisionType::RowPropertiesChange,
            (b"cellIns", _) => RevisionType::CellInsert,
            (b"cellDel", _) => RevisionType::CellDelete,
            (b"cellMerge", _) => RevisionType::CellMerge,
            (b"tcPrChange", _) => RevisionType::CellPropertiesChange,
            _ => return Ok(None),
        };
        Ok(Some(kind))
    }

    fn is_fragment_word_namespace(
        namespace: &ResolveResult<'_>,
        fragment_prefix: &Option<Option<Vec<u8>>>,
    ) -> bool {
        if is_wordprocessing_namespace(namespace) {
            return true;
        }
        match namespace {
            ResolveResult::Unknown(prefix) => {
                fragment_prefix.as_ref().and_then(Option::as_deref) == Some(prefix.as_slice())
            },
            ResolveResult::Unbound => fragment_prefix == &Some(None),
            ResolveResult::Bound(_) => false,
        }
    }

    fn is_date_utc_namespace(namespace: &ResolveResult<'_>) -> bool {
        matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if *value == WORD_2023_DATE_UTC_NAMESPACE_BYTES
        )
    }

    fn decode_attribute(
        attribute: &quick_xml::events::attributes::Attribute<'_>,
        decoder: quick_xml::encoding::Decoder,
        limits: Limits,
        metadata_bytes: &mut usize,
    ) -> Result<String> {
        check("value bytes", attribute.value.len(), limits.max_value_bytes)?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        validate_xml_characters(&value)?;
        charge(
            "metadata bytes",
            metadata_bytes,
            value.len(),
            limits.max_metadata_bytes,
        )?;
        let mut owned = String::new();
        owned
            .try_reserve_exact(value.len())
            .map_err(|source| Error::Allocation {
                resource: "revision metadata",
                source,
            })?;
        owned.push_str(&value);
        Ok(owned)
    }

    fn is_word_attribute(
        namespace: &ResolveResult<'_>,
        fragment_prefix: &Option<Option<Vec<u8>>>,
    ) -> bool {
        is_wordprocessing_namespace(namespace)
            || (matches!(namespace, ResolveResult::Unbound) && fragment_prefix.is_some())
            || matches!(namespace, ResolveResult::Unknown(prefix)
                if prefix.as_slice() == b"w"
                    && fragment_prefix.as_ref().and_then(Option::as_deref) == Some(b"w"))
    }

    fn validate_date_utc(value: &str) -> Result<()> {
        if value.len() > MAX_REVISION_DATE_UTC_BYTES || !value.is_ascii() {
            return Err(Error::InvalidFormat(
                "revision dateUtc must be a bounded ASCII UTC dateTime".into(),
            ));
        }
        DateTime::new(value.to_owned()).map_err(|_source_error| {
            Error::InvalidFormat("revision dateUtc must be a valid xsd:dateTime".into())
        })?;
        let normalized = collapse_xml_whitespace(value);
        if normalized.ends_with('Z')
            || normalized.ends_with("+00:00")
            || normalized.ends_with("-00:00")
        {
            Ok(())
        } else {
            Err(Error::InvalidFormat(
                "revision dateUtc must use a UTC timezone".into(),
            ))
        }
    }

    fn collapse_xml_whitespace(value: &str) -> String {
        let mut output = String::with_capacity(value.len());
        let mut pending_space = false;
        for character in value.chars() {
            if matches!(character, '\u{9}' | '\u{A}' | '\u{D}' | ' ') {
                pending_space = true;
            } else {
                if pending_space && !output.is_empty() {
                    output.push(' ');
                }
                output.push(character);
                pending_space = false;
            }
        }
        output
    }

    fn validate_xml_integer(value: &str) -> Result<()> {
        let normalized = collapse_xml_whitespace(value);
        let digits = normalized
            .strip_prefix('+')
            .or_else(|| normalized.strip_prefix('-'))
            .unwrap_or(&normalized);
        if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
            return Err(Error::InvalidFormat(
                "revision ID must be a valid xsd:integer".into(),
            ));
        }
        Ok(())
    }

    fn validate_date(value: &str) -> Result<()> {
        DateTime::new(value.to_owned()).map_err(|_source_error| {
            Error::InvalidFormat("revision date must be a valid xsd:dateTime".into())
        })?;
        Ok(())
    }

    fn revision_from_element(
        element: &quick_xml::events::BytesStart<'_>,
        revision_type: RevisionType,
        resolver: &NamespaceResolver,
        decoder: quick_xml::encoding::Decoder,
        fragment_prefix: &Option<Option<Vec<u8>>>,
        limits: Limits,
        metadata_bytes: &mut usize,
    ) -> Result<Revision> {
        let mut author = None;
        let mut date = None;
        let mut date_utc = None;
        let mut id = None;
        let mut original_numbering = None;

        for (index, attr) in element.checked_attributes().enumerate() {
            check("attributes", index.saturating_add(1), limits.max_attributes)?;
            let attr = attr.map_err(|error| Error::Xml(error.to_string()))?;
            let local_name = attr.key.local_name();
            let (namespace, _) = resolver.resolve_attribute(attr.key);
            if matches!(namespace, ResolveResult::Unknown(_)) && fragment_prefix.is_none() {
                return Err(Error::InvalidFormat(
                    "revision attribute has an undeclared namespace prefix".into(),
                ));
            }
            if is_word_attribute(&namespace, fragment_prefix) {
                if revision_type == RevisionType::TableGridChange && local_name.as_ref() != b"id" {
                    return Err(Error::InvalidFormat(
                        "table-grid revision permits only a Word ID attribute".into(),
                    ));
                }
                match local_name.as_ref() {
                    b"original" if revision_type == RevisionType::NumberingChange => {
                        if original_numbering.is_some() {
                            return Err(Error::InvalidFormat(
                                "duplicate original numbering attribute".into(),
                            ));
                        }
                        original_numbering =
                            Some(decode_attribute(&attr, decoder, limits, metadata_bytes)?);
                    },
                    b"author" => {
                        if author.is_some() {
                            return Err(Error::InvalidFormat(
                                "duplicate Word revision author attribute".into(),
                            ));
                        }
                        author = Some(decode_attribute(&attr, decoder, limits, metadata_bytes)?);
                    },
                    b"date" => {
                        if date.is_some() {
                            return Err(Error::InvalidFormat(
                                "duplicate Word revision date attribute".into(),
                            ));
                        }
                        let value = decode_attribute(&attr, decoder, limits, metadata_bytes)?;
                        validate_date(&value)?;
                        date = Some(value);
                    },
                    b"id" => {
                        if id.is_some() {
                            return Err(Error::InvalidFormat(
                                "duplicate Word revision ID attribute".into(),
                            ));
                        }
                        let value = decode_attribute(&attr, decoder, limits, metadata_bytes)?;
                        validate_xml_integer(&value)?;
                        id = Some(value);
                    },
                    _ => {},
                }
            } else if local_name.as_ref() == b"dateUtc" && is_date_utc_namespace(&namespace) {
                if date_utc.is_some() {
                    return Err(Error::InvalidFormat(
                        "duplicate Word revision dateUtc attribute".into(),
                    ));
                }
                let value = decode_attribute(&attr, decoder, limits, metadata_bytes)?;
                validate_date_utc(&value)?;
                date_utc = Some(value);
            }
        }

        if revision_type != RevisionType::TableGridChange && author.is_none() {
            return Err(Error::InvalidFormat(
                "tracked revision requires a Word author attribute".into(),
            ));
        }
        let id = id.ok_or_else(|| {
            Error::InvalidFormat("tracked revision requires a Word ID attribute".into())
        })?;
        Ok(Revision {
            revision_type,
            author,
            original_numbering,
            date,
            date_utc,
            id,
            text: String::new(),
        })
    }

    let mut first_element = true;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        if !matches!(event, Event::Eof) {
            charge("events", &mut events, 1, limits.max_events)?;
        }
        if first_element && matches!(&event, Event::Start(_) | Event::Empty(_)) {
            first_element = false;
            if inherited_namespaces.is_empty() && !matches!(&namespace, ResolveResult::Bound(_)) {
                let prefix = match &event {
                    Event::Start(element) | Event::Empty(element) => element
                        .name()
                        .prefix()
                        .map(|prefix| prefix.into_inner().to_vec()),
                    _ => None,
                };
                fragment_prefix = Some(prefix);
            }
        }
        let is_word_element = is_fragment_word_namespace(&namespace, &fragment_prefix);
        let legacy_attribute_prefix = if is_wordprocessing_namespace(&namespace) {
            &None
        } else {
            &fragment_prefix
        };
        match event {
            Event::Start(e) => {
                if root_closed {
                    return Err(Error::InvalidFormat(
                        "revision XML has multiple root elements".into(),
                    ));
                }
                if !root_seen {
                    root_seen = true;
                }
                if text_element_depth.is_some() {
                    return Err(Error::InvalidFormat(
                        "revision text cannot contain child elements".into(),
                    ));
                }
                charge("depth", &mut element_depth, 1, limits.max_depth)?;
                let local_name_ref = e.local_name();
                let local_name = local_name_ref.as_ref();
                validate_history_child(scopes.last_mut(), local_name, is_word_element)?;
                if local_name == b"numberingChange"
                    && matches!(&namespace, ResolveResult::Bound(Namespace(uri)) if *uri == crate::namespace::STRICT_WORDPROCESSINGML_NAMESPACE)
                {
                    return Err(Error::InvalidFormat(
                        "numberingChange is a Transitional-only revision".into(),
                    ));
                }
                let parent = scopes.last().copied().unwrap_or(RevisionScope::Other);
                if parent == RevisionScope::Marker {
                    return Err(Error::InvalidFormat(
                        "revision metadata marker must be childless".into(),
                    ));
                }
                let kind = revision_type(local_name, parent, is_word_element)?;
                scopes
                    .try_reserve(1)
                    .map_err(|_| Error::RevisionAllocation {
                        resource: "revision XML scopes",
                    })?;
                scopes.push(revision_scope(local_name, parent, is_word_element, kind));
                if let Some(kind) = kind {
                    check(
                        "records",
                        revisions.len().saturating_add(1),
                        limits.max_revisions,
                    )?;
                    let revision = revision_from_element(
                        &e,
                        kind,
                        reader.resolver(),
                        reader.decoder(),
                        legacy_attribute_prefix,
                        limits,
                        &mut metadata_bytes,
                    )?;
                    revisions
                        .try_reserve(1)
                        .map_err(|_| Error::RevisionAllocation {
                            resource: "revision records",
                        })?;
                    if is_inline_revision(kind) {
                        active
                            .try_reserve(1)
                            .map_err(|_| Error::RevisionAllocation {
                                resource: "active revisions",
                            })?;
                        active.push((revisions.len(), element_depth));
                    }
                    revisions.push(revision);
                } else if !active.is_empty()
                    && is_word_element
                    && matches!(local_name, b"t" | b"delText" | b"delInstrText")
                {
                    text_element_depth = Some(element_depth);
                } else if !active.is_empty()
                    && parent == RevisionScope::Other
                    && let Some(character) = revision_special_character(local_name, is_word_element)
                {
                    let mut encoded = [0; 4];
                    append_revision_text(
                        &mut revisions,
                        &active,
                        character.encode_utf8(&mut encoded),
                        &mut text_bytes,
                        limits,
                    )?;
                }
            },
            Event::Empty(e) => {
                if root_closed {
                    return Err(Error::InvalidFormat(
                        "revision XML has multiple root elements".into(),
                    ));
                }
                if !root_seen {
                    root_seen = true;
                }
                if text_element_depth.is_some() {
                    return Err(Error::InvalidFormat(
                        "revision text cannot contain child elements".into(),
                    ));
                }
                check("depth", element_depth.saturating_add(1), limits.max_depth)?;
                validate_history_child(
                    scopes.last_mut(),
                    e.local_name().as_ref(),
                    is_word_element,
                )?;
                if e.local_name().as_ref() == b"numberingChange"
                    && matches!(&namespace, ResolveResult::Bound(Namespace(uri)) if *uri == crate::namespace::STRICT_WORDPROCESSINGML_NAMESPACE)
                {
                    return Err(Error::InvalidFormat(
                        "numberingChange is a Transitional-only revision".into(),
                    ));
                }
                let parent = scopes.last().copied().unwrap_or(RevisionScope::Other);
                if parent == RevisionScope::Marker {
                    return Err(Error::InvalidFormat(
                        "revision metadata marker must be childless".into(),
                    ));
                }
                if let Some(kind) = revision_type(e.local_name().as_ref(), parent, is_word_element)?
                {
                    validate_history_end(revision_scope(
                        e.local_name().as_ref(),
                        parent,
                        is_word_element,
                        Some(kind),
                    ))?;
                    check(
                        "records",
                        revisions.len().saturating_add(1),
                        limits.max_revisions,
                    )?;
                    let revision = revision_from_element(
                        &e,
                        kind,
                        reader.resolver(),
                        reader.decoder(),
                        legacy_attribute_prefix,
                        limits,
                        &mut metadata_bytes,
                    )?;
                    revisions
                        .try_reserve(1)
                        .map_err(|_| Error::RevisionAllocation {
                            resource: "revision records",
                        })?;
                    revisions.push(revision);
                } else if !active.is_empty()
                    && parent == RevisionScope::Other
                    && let Some(character) =
                        revision_special_character(e.local_name().as_ref(), is_word_element)
                {
                    let mut encoded = [0; 4];
                    append_revision_text(
                        &mut revisions,
                        &active,
                        character.encode_utf8(&mut encoded),
                        &mut text_bytes,
                        limits,
                    )?;
                }
                if element_depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(e) if !root_seen || root_closed => {
                let text = e
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                if !text
                    .chars()
                    .all(|character| matches!(character, '\u{9}' | '\u{A}' | '\u{D}' | ' '))
                {
                    return Err(Error::InvalidFormat(
                        "revision XML has text outside its root element".into(),
                    ));
                }
            },
            Event::CData(_) if !root_seen || root_closed => {
                return Err(Error::InvalidFormat(
                    "revision XML has CDATA outside its root element".into(),
                ));
            },
            Event::Text(e)
                if scopes.last().is_some_and(|scope| {
                    matches!(
                        scope,
                        RevisionScope::Marker
                            | RevisionScope::History { .. }
                            | RevisionScope::OriginalProperties
                            | RevisionScope::OriginalPropertyContent
                    )
                }) =>
            {
                let text = e
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_marker_text(&text)?;
            },
            Event::CData(e)
                if scopes.last().is_some_and(|scope| {
                    matches!(
                        scope,
                        RevisionScope::Marker
                            | RevisionScope::History { .. }
                            | RevisionScope::OriginalProperties
                            | RevisionScope::OriginalPropertyContent
                    )
                }) =>
            {
                let text = e
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_marker_text(&text)?;
            },
            Event::Text(e) if text_element_depth.is_some() => {
                let encoded = e
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                let text = quick_xml::escape::unescape(&encoded)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                append_revision_text(&mut revisions, &active, &text, &mut text_bytes, limits)?;
            },
            Event::CData(e) if text_element_depth.is_some() => {
                let text = e
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                append_revision_text(&mut revisions, &active, &text, &mut text_bytes, limits)?;
            },
            Event::GeneralRef(reference) => {
                check(
                    "value bytes",
                    reference.as_ref().len(),
                    limits.max_value_bytes,
                )?;
                let text = decode_xml_reference(&reference)?;
                validate_xml_characters(&text)?;
                if !root_seen || root_closed {
                    return Err(Error::InvalidFormat(
                        "revision XML has a reference outside its root element".into(),
                    ));
                }
                if scopes.last().is_some_and(|scope| {
                    matches!(
                        scope,
                        RevisionScope::Marker
                            | RevisionScope::History { .. }
                            | RevisionScope::OriginalProperties
                            | RevisionScope::OriginalPropertyContent
                    )
                }) {
                    validate_marker_text(&text)?;
                }
                if text_element_depth.is_some() {
                    append_revision_text(&mut revisions, &active, &text, &mut text_bytes, limits)?;
                }
            },
            Event::Decl(_) => {
                if declaration_seen || events != 1 || root_seen {
                    return Err(Error::InvalidFormat(
                        "revision XML declaration must appear once at the beginning".into(),
                    ));
                }
                declaration_seen = true;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DTD declarations are not permitted in revision XML".into(),
                ));
            },
            Event::End(_) => {
                if active
                    .last()
                    .is_some_and(|(_, depth)| *depth == element_depth)
                {
                    active.pop();
                }
                if text_element_depth == Some(element_depth) {
                    text_element_depth = None;
                }
                let scope = scopes
                    .pop()
                    .ok_or_else(|| Error::InvalidFormat("invalid revision XML scope".into()))?;
                validate_history_end(scope)?;
                element_depth = element_depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid revision XML nesting".into()))?;
                if element_depth == 0 {
                    root_closed = true;
                }
            },
            Event::Eof => {
                if !root_seen {
                    return Err(Error::InvalidFormat(
                        "revision XML has no root element".into(),
                    ));
                }
                break;
            },
            _ => {},
        }
    }
    if element_depth != 0
        || !active.is_empty()
        || text_element_depth.is_some()
        || !scopes.is_empty()
    {
        return Err(Error::InvalidFormat(
            "truncated revision XML fragment".into(),
        ));
    }

    Ok(revisions)
}

fn validate_history_child(
    parent: Option<&mut RevisionScope>,
    local: &[u8],
    is_word: bool,
) -> Result<()> {
    let Some(parent) = parent else {
        return Ok(());
    };
    if is_word
        && matches!(
            parent,
            RevisionScope::OriginalProperties | RevisionScope::OriginalPropertyContent
        )
        && matches!(
            local,
            b"ins"
                | b"del"
                | b"moveFrom"
                | b"moveTo"
                | b"rPrChange"
                | b"pPrChange"
                | b"sectPrChange"
                | b"tblPrChange"
                | b"tblPrExChange"
                | b"tblGridChange"
                | b"trPrChange"
                | b"tcPrChange"
                | b"cellIns"
                | b"cellDel"
                | b"cellMerge"
                | b"numberingChange"
        )
    {
        return Err(Error::InvalidFormat(
            "original properties cannot contain revision metadata".into(),
        ));
    }
    if is_word {
        match parent {
            RevisionScope::NumberingProperties { phase } => {
                let next = match local {
                    b"ilvl" => 1,
                    b"numId" => 2,
                    b"numberingChange" => 3,
                    b"ins" => 4,
                    _ => 0,
                };
                if next != 0 {
                    if next <= *phase {
                        return Err(Error::InvalidFormat(
                            "numbering properties are duplicated or out of order".into(),
                        ));
                    }
                    *phase = next;
                }
            },
            RevisionScope::FieldCharacter { has_content }
                if matches!(local, b"fldData" | b"ffData" | b"numberingChange") =>
            {
                if *has_content {
                    return Err(Error::InvalidFormat(
                        "field character permits only one data or numbering child".into(),
                    ));
                }
                *has_content = true;
            },
            _ => {},
        }
    }
    let owner = match parent {
        RevisionScope::SectionProperties { changed } => Some((changed, b"sectPrChange".as_slice())),
        RevisionScope::TableGrid { changed } => Some((changed, b"tblGridChange".as_slice())),
        RevisionScope::TableExceptions { changed } => Some((changed, b"tblPrExChange".as_slice())),
        _ => None,
    };
    if let Some((changed, marker)) = owner {
        if is_word && *changed {
            return Err(Error::InvalidFormat(
                "property revision must occur once, after the current properties".into(),
            ));
        }
        if is_word && local == marker {
            *changed = true;
        }
        return Ok(());
    }
    let RevisionScope::History { kind, seen } = parent else {
        return Ok(());
    };
    let expected: &[u8] = match kind {
        RevisionType::SectionPropertiesChange => b"sectPr",
        RevisionType::TableGridChange => b"tblGrid",
        RevisionType::TablePropertyExceptionsChange => b"tblPrEx",
        _ => {
            return Err(Error::InvalidFormat(
                "invalid property-history scope".into(),
            ));
        },
    };
    if *seen || !is_word || local != expected {
        return Err(Error::InvalidFormat(
            "property revision must contain its single original property snapshot".into(),
        ));
    }
    *seen = true;
    Ok(())
}

fn validate_history_end(scope: RevisionScope) -> Result<()> {
    if let RevisionScope::History { kind, seen: false } = scope
        && kind != RevisionType::SectionPropertiesChange
    {
        return Err(Error::InvalidFormat(
            "property revision is missing its required original properties".into(),
        ));
    }
    Ok(())
}

fn validate_xml_characters(value: &str) -> Result<()> {
    if value.chars().all(|character| {
        matches!(
            character,
            '\u{9}' | '\u{a}' | '\u{d}' | '\u{20}'..='\u{d7ff}'
                | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}'
        )
    }) {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "revision XML contains an invalid XML character".into(),
        ))
    }
}

fn validate_marker_text(text: &str) -> Result<()> {
    if text
        .chars()
        .all(|character| matches!(character, ' ' | '\t' | '\r' | '\n'))
    {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "revision metadata marker must be childless".into(),
        ))
    }
}

fn append_revision_text(
    revisions: &mut [Revision],
    active: &[(usize, usize)],
    text: &str,
    text_bytes: &mut usize,
    limits: Limits,
) -> Result<()> {
    validate_xml_characters(text)?;
    let additional = text
        .len()
        .checked_mul(active.len())
        .ok_or(Error::RevisionLimit {
            resource: "text bytes",
            actual: usize::MAX,
            maximum: limits.max_text_bytes,
        })?;
    charge("text bytes", text_bytes, additional, limits.max_text_bytes)?;
    for &(index, _) in active {
        let revision = revisions.get_mut(index).ok_or_else(|| {
            Error::InvalidFormat("revision text scope has no owning record".into())
        })?;
        revision
            .text
            .try_reserve(text.len())
            .map_err(|source| Error::Allocation {
                resource: "revision text",
                source,
            })?;
        revision.text.push_str(text);
    }
    Ok(())
}

#[cfg(test)]
#[path = "revision/property_tests.rs"]
mod property_tests;

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Arc;

    fn paragraph_from_document(xml: String) -> crate::paragraph::Paragraph {
        let source = Arc::new(xml.into_bytes());
        let mut selected = None;
        crate::namespace::scan_word_element_ranges_with_context(
            source.as_slice(),
            &[],
            &[b"p".as_slice()],
            |_, start, length, namespaces| {
                selected = Some((start, length, namespaces));
                Ok(())
            },
        )
        .unwrap();
        let (start, length, namespaces) = selected.expect("paragraph range");
        crate::paragraph::Paragraph::from_arc_range_with_context(source, start, length, namespaces)
    }

    fn table_from_document(xml: String) -> crate::table::Table {
        let source = Arc::new(xml.into_bytes());
        let mut selected = None;
        crate::namespace::scan_word_element_ranges_with_context(
            source.as_slice(),
            &[],
            &[b"tbl".as_slice()],
            |_, start, length, namespaces| {
                selected = Some((start, length, namespaces));
                Ok(())
            },
        )
        .unwrap();
        let (start, length, namespaces) = selected.expect("table range");
        crate::table::Table::from_arc_range_with_context(source, start, length, namespaces)
    }

    #[test]
    fn test_revision_creation() {
        let rev = Revision::new(
            RevisionType::Insert,
            "John Doe".to_string(),
            Some("2024-11-05T10:30:00Z".to_string()),
            "0".to_string(),
        );

        assert_eq!(rev.revision_type(), RevisionType::Insert);
        assert_eq!(rev.author(), Some("John Doe"));
        assert_eq!(rev.date(), Some("2024-11-05T10:30:00Z"));
        assert_eq!(rev.date_utc(), None);
        assert_eq!(rev.id(), "0");
        assert_eq!(rev.text(), "");
    }

    #[test]
    fn test_parse_revisions_empty() {
        let xml = b"<w:p><w:r><w:t>Normal text</w:t></w:r></w:p>";
        let revisions = parse_revisions(xml).unwrap();
        assert_eq!(revisions.len(), 0);
    }

    #[test]
    fn test_parse_insert_revision() {
        let xml = br#"<w:p>
            <w:ins w:id="0" w:author="John Doe" w:date="2024-11-05T10:30:00Z">
                <w:r>
                    <w:t>inserted text</w:t>
                </w:r>
            </w:ins>
        </w:p>"#;

        let revisions = parse_revisions(xml).unwrap();
        assert_eq!(revisions.len(), 1);

        let rev = &revisions[0];
        assert_eq!(rev.revision_type(), RevisionType::Insert);
        assert_eq!(rev.author(), Some("John Doe"));
        assert_eq!(rev.date(), Some("2024-11-05T10:30:00Z"));
        assert_eq!(rev.id(), "0");
        assert_eq!(rev.text(), "inserted text");
    }

    #[test]
    fn test_parse_delete_revision() {
        let xml = br#"<w:p>
            <w:del w:id="1" w:author="Jane Smith" w:date="2024-11-05T11:00:00Z">
                <w:r>
                    <w:delText>deleted text</w:delText>
                </w:r>
            </w:del>
        </w:p>"#;

        let revisions = parse_revisions(xml).unwrap();
        assert_eq!(revisions.len(), 1);

        let rev = &revisions[0];
        assert_eq!(rev.revision_type(), RevisionType::Delete);
        assert_eq!(rev.author(), Some("Jane Smith"));
        assert_eq!(rev.text(), "deleted text");
    }

    #[test]
    fn revision_text_preserves_authored_edge_whitespace_and_cdata() {
        let xml = br#"<w:p>
            <w:ins w:id="1" w:author="Alice"><w:r><w:t xml:space="preserve">  before<![CDATA[ & inside ]]></w:t></w:r></w:ins>
        </w:p>"#;

        let revisions = parse_revisions(xml).unwrap();
        assert_eq!(revisions.len(), 1);
        assert_eq!(revisions[0].text(), "  before & inside ");
    }

    #[test]
    fn revision_text_decodes_references_and_xml_line_endings() {
        for (owner, text_element) in [
            ("ins", "t"),
            ("del", "delText"),
            ("moveFrom", "delInstrText"),
            ("moveTo", "t"),
        ] {
            let xml = format!(
                "<w:p><w:{owner} w:id=\"1\" w:author=\"Alice\"><w:r><w:{text_element}>  A&amp;&lt;&gt;&quot;&apos;&#65;&#x1F642;&#13;\r\nB\rC<![CDATA[ &amp;\r\n ]]></w:{text_element}></w:r></w:{owner}></w:p>"
            );
            let revisions = parse_revisions(xml.as_bytes()).unwrap();
            assert_eq!(revisions.len(), 1);
            assert_eq!(revisions[0].text(), "  A&<>\"'A🙂\r\nB\nC &amp;\n ");
        }
    }

    #[test]
    fn revision_text_rejects_invalid_references_and_xml_characters() {
        for text in ["&missing;", "&#xD800;", "&#0;", "&#1;", "\u{1}"] {
            let xml = format!(
                "<w:p><w:ins w:id=\"1\" w:author=\"Alice\"><w:r><w:t>{text}</w:t></w:r></w:ins></w:p>"
            );
            assert!(parse_revisions(xml.as_bytes()).is_err(), "{text:?}");
        }
    }

    #[test]
    fn revision_text_rejects_invalid_utf8_in_text_and_cdata() {
        for text in [
            b"bad\xfftext".as_slice(),
            b"<![CDATA[bad\xfftext]]>".as_slice(),
        ] {
            let mut xml = br#"<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:t>"#.to_vec();
            xml.extend_from_slice(text);
            xml.extend_from_slice(b"</w:t></w:r></w:ins></w:p>");
            assert!(parse_revisions(&xml).is_err());
        }
    }

    #[test]
    fn revision_fragment_rejects_dtd_declarations() {
        let xml = br#"<!DOCTYPE p [<!ENTITY value "replacement">]><w:p><w:ins w:id="1" w:author="Alice"><w:r><w:t>&value;</w:t></w:r></w:ins></w:p>"#;
        assert!(parse_revisions(xml).is_err());
    }

    #[test]
    fn truncated_revision_fragments_are_rejected() {
        for xml in [
            br#"<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:t>open"#.as_slice(),
            br#"<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:t>open</w:t></w:r>"#.as_slice(),
        ] {
            assert!(
                parse_revisions(xml).is_err(),
                "{}",
                String::from_utf8_lossy(xml)
            );
        }
    }

    #[test]
    fn test_parse_date_utc_with_bound_namespace_and_lexical_value() {
        let xml = format!(
            r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:du="{WORD_2023_DATE_UTC_NAMESPACE}">
                <w:ins w:id="1" w:author="Alice" du:dateUtc=" 2026-07-17T00:00:00.123456+00:00 ">
                    <w:r><w:t>inserted</w:t></w:r>
                </w:ins>
            </w:p>"#
        );

        let revisions = parse_revisions(xml.as_bytes()).unwrap();
        assert_eq!(revisions.len(), 1);
        assert_eq!(
            revisions[0].date_utc(),
            Some(" 2026-07-17T00:00:00.123456+00:00 ")
        );
        assert_eq!(revisions[0].text(), "inserted");
    }

    #[test]
    fn test_parse_property_revision_date_utc_and_ignore_wrong_namespace() {
        let xml = format!(
            r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:du="{WORD_2023_DATE_UTC_NAMESPACE}" xmlns:x="urn:foreign">
                <w:pPr><w:pPrChange w:id="2" w:author="Alice" du:dateUtc="2026-07-17T00:00:00Z"/></w:pPr>
                <x:ins x:id="3" x:author="Foreign" x:dateUtc="2026-07-17T00:00:00Z"/>
                <w:del w:id="4" w:author="Bob" w:dateUtc="2026-07-17T00:00:00Z"/>
            </w:p>"#
        );

        let revisions = parse_revisions(xml.as_bytes()).unwrap();
        assert_eq!(revisions.len(), 2);
        assert_eq!(revisions[0].revision_type(), RevisionType::FormatChange);
        assert_eq!(revisions[0].date_utc(), Some("2026-07-17T00:00:00Z"));
        assert_eq!(revisions[1].revision_type(), RevisionType::Delete);
        assert_eq!(revisions[1].date_utc(), None);
    }

    #[test]
    fn test_parse_date_utc_rejects_non_utc_lexical_value() {
        let xml = format!(
            r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:du="{WORD_2023_DATE_UTC_NAMESPACE}">
                <w:ins w:id="1" w:author="Alice" du:dateUtc="2026-07-17T00:00:00+01:00"/>
            </w:p>"#
        );

        assert!(parse_revisions(xml.as_bytes()).is_err());
    }

    #[test]
    fn test_parse_date_utc_requires_namespace_binding() {
        let xml = br#"<w:p><w:ins w:id="1" w:author="Alice" w16du:dateUtc="2026-07-17T00:00:00Z"/></w:p>"#;

        let revisions = parse_revisions(xml).unwrap();
        assert_eq!(revisions.len(), 1);
        assert_eq!(revisions[0].date_utc(), None);
    }

    #[test]
    fn paragraph_revisions_resolve_inherited_word_2023_namespace() {
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:du="{WORD_2023_DATE_UTC_NAMESPACE}"><w:body><w:p><w:ins w:id="1" w:author="Alice" du:dateUtc="2026-07-17T00:00:00Z"/></w:p></w:body></w:document>"#
        );
        let paragraph = paragraph_from_document(xml);

        let revisions = paragraph.revisions().unwrap();
        assert_eq!(revisions.len(), 1);
        assert_eq!(revisions[0].date_utc(), Some("2026-07-17T00:00:00Z"));
    }

    #[test]
    fn paragraph_revisions_reject_inherited_foreign_word_prefix() {
        let xml =
            r#"<w:p><w:ins w:id="1" w:author="Alice" du:dateUtc="2026-07-17T00:00:00Z"/></w:p>"#
                .to_owned();
        let source = Arc::new(xml.into_bytes());
        let namespaces: crate::namespace::NamespaceBindings = Arc::from(
            vec![
                (Some(b"w".to_vec()), b"urn:foreign".to_vec()),
                (
                    Some(b"du".to_vec()),
                    WORD_2023_DATE_UTC_NAMESPACE.as_bytes().to_vec(),
                ),
            ]
            .into_boxed_slice(),
        );
        let paragraph = crate::paragraph::Paragraph::from_arc_range_with_context(
            Arc::clone(&source),
            0,
            source.len() as u32,
            namespaces,
        );

        assert!(paragraph.revisions().unwrap().is_empty());
    }

    #[test]
    fn table_revisions_resolve_inherited_word_2023_namespace() {
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:du="{WORD_2023_DATE_UTC_NAMESPACE}"><w:body><w:tbl><w:tr><w:trPr><w:ins w:id="1" w:author="Alice" du:dateUtc="2026-07-17T00:00:00Z"/></w:trPr><w:tc><w:tcPr><w:cellIns w:id="2" w:author="Bob" du:dateUtc="2026-07-17T00:00:00.123Z"/></w:tcPr></w:tc></w:tr></w:tbl></w:body></w:document>"#
        );
        let table = table_from_document(xml);

        let revisions = table.revisions().unwrap();
        assert_eq!(revisions.len(), 2);
        assert_eq!(revisions[0].revision_type(), RevisionType::RowInsert);
        assert_eq!(revisions[0].date_utc(), Some("2026-07-17T00:00:00Z"));
        assert_eq!(revisions[1].revision_type(), RevisionType::CellInsert);
        assert_eq!(revisions[1].date_utc(), Some("2026-07-17T00:00:00.123Z"));
        let rows = table.rows().unwrap();
        let row_revisions = rows[0].revisions().unwrap();
        assert_eq!(row_revisions.len(), 2);
        let cells = rows[0].cells().unwrap();
        let cell_revisions = cells[0].revisions().unwrap();
        assert_eq!(cell_revisions.len(), 1);
        assert_eq!(
            cell_revisions[0].date_utc(),
            Some("2026-07-17T00:00:00.123Z")
        );
    }

    #[test]
    fn paragraph_revisions_ignore_inherited_foreign_date_utc_namespace() {
        let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w16du="urn:foreign"><w:body><w:p><w:ins w:id="1" w:author="Alice" w16du:dateUtc="2026-07-17T00:00:00Z"/></w:p></w:body></w:document>"#;
        let paragraph = paragraph_from_document(xml.to_owned());

        let revisions = paragraph.revisions().unwrap();
        assert_eq!(revisions.len(), 1);
        assert_eq!(revisions[0].date_utc(), None);
    }

    #[test]
    fn test_parse_multiple_revisions() {
        let xml = br#"<w:p>
            <w:ins w:id="0" w:author="Author1">
                <w:r><w:t>inserted</w:t></w:r>
            </w:ins>
            <w:del w:id="1" w:author="Author2">
                <w:r><w:delText>deleted</w:delText></w:r>
            </w:del>
            <w:moveFrom w:id="2" w:author="Author3">
                <w:r><w:t>moved</w:t></w:r>
            </w:moveFrom>
        </w:p>"#;

        let revisions = parse_revisions(xml).unwrap();
        assert_eq!(revisions.len(), 3);

        assert_eq!(revisions[0].revision_type(), RevisionType::Insert);
        assert_eq!(revisions[1].revision_type(), RevisionType::Delete);
        assert_eq!(revisions[2].revision_type(), RevisionType::MoveFrom);
    }

    #[test]
    fn nested_same_name_revision_closes_at_its_own_depth() {
        let xml = br#"<w:p>
            <w:ins w:id="1" w:author="Outer"><w:r><w:t>before</w:t></w:r>
                <w:ins w:id="2" w:author="Inner"><w:r><w:t>nested</w:t></w:r></w:ins>
                <w:r><w:t>after</w:t></w:r>
            </w:ins>
        </w:p>"#;

        let revisions = parse_revisions(xml).unwrap();
        assert_eq!(revisions.len(), 2);
        assert_eq!(revisions[0].id(), "1");
        assert_eq!(revisions[0].text(), "beforenestedafter");
        assert_eq!(revisions[1].id(), "2");
        assert_eq!(revisions[1].author(), Some("Inner"));
        assert_eq!(revisions[1].text(), "nested");
    }

    #[test]
    fn test_revision_type_display() {
        assert_eq!(format!("{}", RevisionType::Insert), "Insert");
        assert_eq!(format!("{}", RevisionType::Delete), "Delete");
        assert_eq!(format!("{}", RevisionType::MoveFrom), "Move From");
        assert_eq!(format!("{}", RevisionType::MoveTo), "Move To");
        assert_eq!(format!("{}", RevisionType::FormatChange), "Format Change");
    }
}
