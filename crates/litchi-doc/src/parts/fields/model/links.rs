use super::core::Field;
use super::mail_merge::MergeFieldSwitch;

/// The stored kind of a legacy Word DDE field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DdeFieldKind {
    /// A `DDE` field, which can request automatic updates with `\\a`.
    Dde,
    /// A `DDEAUTO` field, which declares automatic updates.
    DdeAuto,
}

/// One stored DDE result-representation switch.
///
/// This value describes a requested representation only. It never causes a
/// source to be contacted, converted, embedded, or displayed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DdeRepresentation {
    /// The `\\b` switch requests a bitmap representation.
    Bitmap,
    /// The `\\h` switch requests HTML-formatted text.
    Html,
    /// The `\\p` switch requests a picture representation.
    Picture,
    /// The `\\r` switch requests rich-text format.
    RichText,
    /// The `\\t` switch requests text-only format.
    Text,
    /// The `\\u` switch requests Unicode text.
    UnicodeText,
}

/// Typed, inert metadata for a legacy Word `DDE` or `DDEAUTO` field.
///
/// [MS-DOC] §2.9.90 identifies their native field-type bytes. Application,
/// source, item, representation, and storage switches remain stored metadata
/// only. This type never launches an application, initiates a DDE conversation,
/// opens a source, requests data, refreshes content, converts content, or
/// executes code.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DdeField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) kind: DdeFieldKind,
    pub(in crate::parts::fields) application: String,
    pub(in crate::parts::fields) source: String,
    pub(in crate::parts::fields) item: Option<String>,
    pub(in crate::parts::fields) automatic_updates: bool,
    pub(in crate::parts::fields) representation: Option<DdeRepresentation>,
    pub(in crate::parts::fields) omit_graphic_data: bool,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl DdeField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is a `DDE` or `DDEAUTO` field.
    pub fn kind(&self) -> DdeFieldKind {
        self.kind
    }

    /// Return the stored DDE application name without launching it.
    pub fn application(&self) -> &str {
        &self.application
    }

    /// Return the stored source identifier without opening or resolving it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional stored source item, such as a cell range or bookmark.
    pub fn item(&self) -> Option<&str> {
        self.item.as_deref()
    }

    /// Whether the stored instruction requests automatic DDE updates.
    ///
    /// This is metadata only. The API never performs an update.
    pub fn requests_automatic_updates(&self) -> bool {
        self.automatic_updates
    }

    /// Return the requested stored result representation, if present.
    ///
    /// This is metadata only and never triggers source access or conversion.
    pub fn representation(&self) -> Option<DdeRepresentation> {
        self.representation
    }

    /// Whether the stored `\\d` switch omits graphic data from the document.
    ///
    /// This is stored metadata only. The API never reads the source to obtain
    /// omitted data.
    pub fn omits_graphic_data(&self) -> bool {
        self.omit_graphic_data
    }

    /// Return unrecognized stored field switches in source order.
    ///
    /// These values are preserved as inert metadata and never interpreted.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from a DDE source.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// One stored result or storage switch for a Word `LINK` field.
///
/// These values describe a linked-object representation or whether graphic data
/// is stored. They never cause a source to be opened, contacted, converted, or
/// displayed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkResultOption {
    /// The `\\b` switch requests a bitmap representation.
    Bitmap,
    /// The `\\d` switch omits graphic data from the document.
    OmitGraphicData,
    /// The `\\h` switch requests HTML-formatted text.
    Html,
    /// The `\\p` switch requests a picture representation.
    Picture,
    /// The `\\r` switch requests rich-text format.
    RichText,
    /// The `\\t` switch requests text-only format.
    Text,
    /// The `\\u` switch requests Unicode text.
    UnicodeText,
}

/// One integral `LINK` `\\f` formatting mode.
///
/// ECMA-376 marks modes 1 and 3 unsupported. Those values, and values outside
/// its defined set, are retained as metadata without applying any formatting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkFormatting {
    /// 0: preserve formatting from the source file.
    Source,
    /// 2: match formatting in the destination document.
    Destination,
    /// 4: preserve source formatting for a SpreadsheetML workbook source.
    SpreadsheetSource,
    /// 5: match destination formatting for a SpreadsheetML workbook source.
    SpreadsheetDestination,
    /// An ECMA-376-unsupported or otherwise unrecognized integral mode.
    Unsupported(i64),
}

/// Typed, inert metadata for a legacy Word `LINK` field.
///
/// [MS-DOC] §2.9.90 identifies its native field-type byte. Application type,
/// source, item, and all result/formatting switches are retained as stored
/// field data. This type never activates an OLE server, launches an
/// application, opens a source, requests data, refreshes content, converts
/// content, or executes code.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LinkField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) application_type: String,
    pub(in crate::parts::fields) source: String,
    pub(in crate::parts::fields) item: Option<String>,
    pub(in crate::parts::fields) automatic_updates: bool,
    pub(in crate::parts::fields) result_options: Vec<LinkResultOption>,
    pub(in crate::parts::fields) formatting_modes: Vec<LinkFormatting>,
    pub(in crate::parts::fields) switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl LinkField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored linked-object application type.
    ///
    /// Word commonly stores an OLE Programmatic Identifier here. It is never
    /// looked up or activated by this API.
    pub fn application_type(&self) -> &str {
        &self.application_type
    }

    /// Return the stored source identifier without opening or resolving it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional stored source item, such as a cell range or bookmark.
    pub fn item(&self) -> Option<&str> {
        self.item.as_deref()
    }

    /// Whether the stored instruction requests automatic updates.
    ///
    /// This is metadata only. The API never performs an update.
    pub fn requests_automatic_updates(&self) -> bool {
        self.automatic_updates
    }

    /// Return recognized result and storage switches in stored source order.
    ///
    /// When several are present, `Self::effective_result_option` reflects
    /// Word's documented last-switch behavior. Neither method contacts the
    /// linked source.
    pub fn result_options(&self) -> &[LinkResultOption] {
        &self.result_options
    }

    /// Return the effective result or storage option under Word's documented
    /// last-switch behavior, if one was stored.
    pub fn effective_result_option(&self) -> Option<LinkResultOption> {
        self.result_options.last().copied()
    }

    /// Return integral `\\f` formatting modes in stored source order.
    ///
    /// These are metadata only; this API never formats linked content.
    pub fn formatting_modes(&self) -> &[LinkFormatting] {
        &self.formatting_modes
    }

    /// Return all stored field switches in source order.
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from a linked source.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The kind of externally sourced legacy Word include field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IncludeFieldKind {
    /// An `INCLUDETEXT` or historical `INCLUDE` field that stores a document or XML source.
    Text,
    /// An `INCLUDEPICTURE` or historical `IMPORT` field that stores an image source.
    Picture,
}

/// One recognized stored option of a legacy Word external-include field.
///
/// These values are configuration metadata only. This API never opens,
/// resolves, imports, transforms, or evaluates the referenced source.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ExternalIncludeOption {
    /// A document or graphics converter name from the `\\c` switch.
    Converter(String),
    /// A source encoding from the `INCLUDETEXT \\e` switch.
    Encoding(String),
    /// A source MIME type from the `INCLUDETEXT \\m` switch.
    MimeType(String),
    /// An XML namespace mapping from the `INCLUDETEXT \\n` switch.
    NamespaceMapping(String),
    /// An XSLT location from the `INCLUDETEXT \\t` switch.
    Xslt(String),
    /// An XPath expression from the `INCLUDETEXT \\x` switch.
    XPath(String),
}

/// Typed, inert metadata for a legacy Word external-include field.
///
/// ECMA-376 Part 1 §§17.16.5.27–28 defines these stored source operands and
/// switches. [MS-DOC] §2.9.90 also identifies historical `INCLUDE` and
/// `IMPORT` native aliases for `INCLUDETEXT` and `INCLUDEPICTURE`,
/// respectively. Source identifiers, bookmarks, options, and cached results
/// are retained as stored field data. This type never opens, resolves,
/// imports, fetches, refreshes, transforms, converts, evaluates, or executes
/// source content.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalIncludeField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) kind: IncludeFieldKind,
    pub(in crate::parts::fields) source: String,
    pub(in crate::parts::fields) bookmark: Option<String>,
    pub(in crate::parts::fields) suppress_nested_field_updates: bool,
    pub(in crate::parts::fields) omit_picture_data: bool,
    pub(in crate::parts::fields) options: Vec<ExternalIncludeOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl ExternalIncludeField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this stores a text or picture external-include field.
    pub fn kind(&self) -> IncludeFieldKind {
        self.kind
    }

    /// Return the stored source identifier without opening or resolving it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional stored text-include bookmark selector.
    ///
    /// Picture-include fields do not define a bookmark operand, so they
    /// always return `None` here.
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Return the optional stored `\\c` converter name.
    ///
    /// The converter is never looked up or invoked.
    pub fn converter(&self) -> Option<&str> {
        self.options.iter().find_map(|option| match option {
            ExternalIncludeOption::Converter(value) => Some(value.as_str()),
            _ => None,
        })
    }

    /// Return recognized converter and XML options in stored source order.
    ///
    /// All options are inert metadata. This method never resolves a converter,
    /// opens a source, runs XSLT, or evaluates XPath.
    pub fn options(&self) -> &[ExternalIncludeOption] {
        &self.options
    }

    /// Whether the stored text-include `\\!` switch suppresses nested field updates.
    ///
    /// This is stored metadata only. The API never performs an update.
    pub fn suppresses_nested_field_updates(&self) -> bool {
        self.suppress_nested_field_updates
    }

    /// Whether the stored picture-include `\\d` switch omits picture data.
    ///
    /// This is stored metadata only. The API never reads the source to obtain
    /// omitted picture data.
    pub fn omits_picture_data(&self) -> bool {
        self.omit_picture_data
    }

    /// Return unrecognized stored field switches in source order.
    ///
    /// These values are preserved as inert metadata and never interpreted.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from an external source.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}
