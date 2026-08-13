//! Bounded, non-mutating RTF validation and security inventory.
//!
//! The report deliberately contains no authored text, field instructions,
//! paths, object names, or payload bytes. It is a conservative inventory of
//! syntax that the parser has already proved to be present. Unknown retained
//! syntax is surfaced as `Unknown`; it is never treated as safe.

use crate::codec::error::RtfResult;
use crate::codec::limits::ParseLimits;
use crate::content::field::{FieldType, ParsedFieldCode};
use crate::model::document::RtfDocument;

/// Provenance emitted by the bounded lexer/parser after it has checked the
/// complete transport token stream.  Keeping this beside the retained model
/// prevents a report from inferring root validity merely from the presence of
/// a partially parsed document.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) struct ParseProvenance {
    pub(crate) syntax_valid: bool,
    pub(crate) root_valid: bool,
    pub(crate) document_valid: bool,
}

/// Compact parser-derived field safety classification.  It deliberately
/// contains no instruction text or target values; the report consumes this
/// cache instead of reparsing field instructions for every check.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) enum FieldSafety {
    Neutral,
    External,
    ExternalUnknown,
    Active,
    ActiveUnknown,
    ExternalAndActive,
    ExternalAndActiveUnknown,
}

/// Classify one field once at the parser/model boundary.  The typed accessors
/// are inert and perform no I/O, execution, fetching, or mutation.
pub(crate) fn classify_field(field: &crate::Field<'_>) -> FieldSafety {
    fn external_and_active(known: bool) -> FieldSafety {
        if known {
            FieldSafety::ExternalAndActive
        } else {
            FieldSafety::ExternalAndActiveUnknown
        }
    }

    fn active(known: bool) -> FieldSafety {
        if known {
            FieldSafety::Active
        } else {
            FieldSafety::ActiveUnknown
        }
    }

    match field.field_type {
        FieldType::Hyperlink => match field.parsed_code() {
            ParsedFieldCode::Hyperlink(code) if code.external_target.is_some() => {
                FieldSafety::External
            },
            ParsedFieldCode::Hyperlink(code) if code.bookmark.is_some() => FieldSafety::Neutral,
            ParsedFieldCode::Hyperlink(_) => FieldSafety::ExternalUnknown,
            ParsedFieldCode::Reference(_)
            | ParsedFieldCode::PageReference(_)
            | ParsedFieldCode::NoteReference(_)
            | ParsedFieldCode::Other { .. }
            | ParsedFieldCode::Malformed(_) => FieldSafety::ExternalUnknown,
        },
        FieldType::Dde | FieldType::DdeAuto => {
            if field.dde_link().is_some() {
                FieldSafety::ExternalAndActive
            } else {
                FieldSafety::ExternalAndActiveUnknown
            }
        },
        FieldType::Link => {
            if field.link_field().is_some() {
                FieldSafety::External
            } else {
                FieldSafety::ExternalUnknown
            }
        },
        FieldType::Include
        | FieldType::Import
        | FieldType::IncludeText
        | FieldType::IncludePicture => {
            if field.external_include().is_some() {
                FieldSafety::External
            } else {
                FieldSafety::ExternalUnknown
            }
        },
        FieldType::ReferencedDocument => {
            if field.referenced_document().is_some() {
                FieldSafety::External
            } else {
                FieldSafety::ExternalUnknown
            }
        },
        FieldType::MacroButton => {
            if field.macro_button().is_some() {
                FieldSafety::Active
            } else {
                FieldSafety::ActiveUnknown
            }
        },
        FieldType::Embed => {
            if field.embed_field().is_some() {
                FieldSafety::Active
            } else {
                FieldSafety::ActiveUnknown
            }
        },
        FieldType::AddIn | FieldType::Control | FieldType::HtmlControl => {
            if field.active_content_field().is_some() {
                FieldSafety::Active
            } else {
                FieldSafety::ActiveUnknown
            }
        },
        FieldType::Print => {
            if field.print_field().is_some() {
                FieldSafety::Active
            } else {
                FieldSafety::ActiveUnknown
            }
        },
        // Mail-merge fields are not resolved by the parser, but they are
        // executable data-source operations at the host boundary.  Keep the
        // classification conservative even for a cached result: a valid
        // instruction is an external and active surface, while a malformed
        // instruction is still unknown rather than neutral.
        FieldType::MergeField => external_and_active(field.merge_field().is_some()),
        FieldType::Database => external_and_active(field.database_field().is_some()),
        FieldType::MailMergeData => external_and_active(field.mail_merge_data().is_some()),
        FieldType::MergeRecord | FieldType::MergeSequence => {
            external_and_active(field.mail_merge_counter().is_some())
        },
        FieldType::MailMergeNext => external_and_active(field.mail_merge_next().is_some()),
        FieldType::MailMergeNextIf | FieldType::MailMergeSkipIf => {
            external_and_active(field.mail_merge_conditional_control().is_some())
        },
        FieldType::AddressBlock | FieldType::GreetingLine => {
            external_and_active(field.mail_merge_recipient_field().is_some())
        },
        // MERGEBARCODE consumes merge data and emits host-rendered content.
        // It does not itself identify a data source, so report it as active
        // without claiming an external link.
        FieldType::MergeBarcode => active(field.barcode_display_field().is_some()),
        FieldType::Unknown => FieldSafety::ExternalAndActiveUnknown,
        _ => FieldSafety::Neutral,
    }
}

/// The dependency needed to interpret one report check.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ValidationDependency {
    /// The bounded RTF lexer and group parser.
    Syntax,
    /// The RTF root-group and document-header checks.
    Root,
    /// The retained, fully parsed document model.
    Document,
    /// The bounded compressed-RTF transport decoder.
    CompressedTransport,
    /// The retained picture parser and payload limits.
    PictureParser,
    /// The retained object parser and payload limits.
    ObjectParser,
    /// The bounded opaque-syntax preservation store.
    OpaquePreservation,
    /// An application-supplied external-resource provider.
    ExternalProvider,
    /// An application-supplied execution provider.
    ExecutionProvider,
    /// An explicit repair planner, which this report does not provide.
    RepairPlanner,
    /// An application-supplied security policy.
    SecurityPolicy,
}

/// Content-free result state for one validation or security check.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ValidationStatus {
    /// The requested structural check completed successfully.
    Valid,
    /// A parser-proven feature or risk surface is present.
    Present,
    /// A parser-proven feature or risk surface is absent.
    Absent,
    /// The check is not meaningful for this input or this API boundary.
    NotApplicable,
    /// A known surface exists, but the required operation is intentionally not
    /// implemented by this crate.
    Unsupported,
    /// The bounded parser retained syntax that prevents a safe conclusion.
    Unknown,
}

/// One content-free validation result and its required dependency.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ValidationCheck {
    status: ValidationStatus,
    dependency: ValidationDependency,
}

impl ValidationCheck {
    const fn new(status: ValidationStatus, dependency: ValidationDependency) -> Self {
        Self { status, dependency }
    }

    /// Return the content-free result state.
    #[must_use]
    pub const fn status(self) -> ValidationStatus {
        self.status
    }

    /// Return the dependency that explains this check's scope.
    #[must_use]
    pub const fn dependency(self) -> ValidationDependency {
        self.dependency
    }

    /// Whether this check completed as a structural success.
    #[must_use]
    pub const fn is_valid(self) -> bool {
        matches!(self.status, ValidationStatus::Valid)
    }

    /// Whether a parser-proven feature is present.
    #[must_use]
    pub const fn is_present(self) -> bool {
        matches!(self.status, ValidationStatus::Present)
    }

    /// Whether the check is explicitly outside this API's scope.
    #[must_use]
    pub const fn is_not_applicable(self) -> bool {
        matches!(self.status, ValidationStatus::NotApplicable)
    }
}

/// Content-free bounded cardinalities retained by one report.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct ValidationCounts {
    source_bytes: usize,
    fields: usize,
    objects: usize,
    pictures: usize,
    form_fields: usize,
    opaque_nodes: usize,
    opaque_bytes: usize,
    unknown_syntax_markers: usize,
}

impl ValidationCounts {
    /// Encoded input bytes, including a compressed-RTF frame when present.
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.source_bytes
    }

    /// Number of retained generic field records.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Number of retained embedded or linked object records.
    #[must_use]
    pub const fn objects(self) -> usize {
        self.objects
    }

    /// Number of retained picture records.
    #[must_use]
    pub const fn pictures(self) -> usize {
        self.pictures
    }

    /// Number of retained legacy form-field records.
    #[must_use]
    pub const fn form_fields(self) -> usize {
        self.form_fields
    }

    /// Number of retained opaque syntax nodes.
    #[must_use]
    pub const fn opaque_nodes(self) -> usize {
        self.opaque_nodes
    }

    /// Aggregate bytes retained by opaque syntax nodes.
    #[must_use]
    pub const fn opaque_bytes(self) -> usize {
        self.opaque_bytes
    }

    /// Number of content-free markers retained for syntax the picture/object
    /// parsers could not safely interpret.
    #[must_use]
    pub const fn unknown_syntax_markers(self) -> usize {
        self.unknown_syntax_markers
    }
}

/// Hard parser bounds relevant to the report.
///
/// The transport byte and token values come from the caller-selected
/// [`ParseLimits`]. Group, object-count, picture-count, object-payload, and
/// picture-payload ceilings are parser hard limits and cannot be widened by
/// this report API.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ValidationLimits {
    parse: ParseLimits,
}

impl ValidationLimits {
    /// Hard maximum recursive RTF group depth.
    pub const MAX_GROUP_DEPTH: usize = crate::codec::parser::MAX_GROUP_NESTING_DEPTH;
    /// Hard maximum retained object records.
    pub const MAX_OBJECTS: usize = crate::codec::parser::MAX_OBJECTS;
    /// Hard maximum retained picture records implied by the token/resource
    /// ceilings. Individual picture payloads remain separately bounded.
    pub const MAX_PICTURES: usize = crate::picture::MAX_PICTURES;
    /// Hard maximum decoded `objdata` payload accepted by the parser.
    pub const MAX_OBJECT_PAYLOAD_BYTES: usize = crate::codec::parser::MAX_OBJECT_DATA_BYTES;
    /// Hard maximum decoded picture payload accepted by the parser.
    pub const MAX_PICTURE_PAYLOAD_BYTES: usize = crate::codec::parser::MAX_PICTURE_DATA_BYTES;

    const fn new(parse: ParseLimits) -> Self {
        Self { parse }
    }

    /// Return the byte/token/opaque limits used by the parser.
    pub const fn parse(self) -> ParseLimits {
        self.parse
    }

    /// Maximum encoded source bytes.
    #[must_use]
    pub const fn max_source_bytes(self) -> usize {
        self.parse.max_source_bytes()
    }

    /// Maximum decompressed compressed-RTF bytes.
    #[must_use]
    pub const fn max_decompressed_bytes(self) -> usize {
        self.parse.max_decompressed_bytes()
    }

    /// Maximum lexer tokens.
    #[must_use]
    pub const fn max_tokens(self) -> usize {
        self.parse.max_tokens()
    }

    /// Maximum aggregate `binN` bytes.
    #[must_use]
    pub const fn max_total_binary_bytes(self) -> usize {
        self.parse.max_total_binary_bytes()
    }

    /// Maximum bytes accepted in one `binN` payload.
    #[must_use]
    pub const fn max_binary_bytes(self) -> usize {
        self.parse.max_binary_bytes()
    }

    /// Maximum retained opaque nodes.
    #[must_use]
    pub const fn max_opaque_nodes(self) -> usize {
        self.parse.max_opaque_nodes()
    }

    /// Maximum bytes retained by one opaque syntax node.
    #[must_use]
    pub const fn max_opaque_node_bytes(self) -> usize {
        self.parse.max_opaque_node_bytes()
    }

    /// Maximum aggregate retained opaque bytes.
    #[must_use]
    pub const fn max_total_opaque_bytes(self) -> usize {
        self.parse.max_total_opaque_bytes()
    }

    /// Hard maximum recursive RTF group depth.
    #[must_use]
    pub const fn max_group_depth(self) -> usize {
        Self::MAX_GROUP_DEPTH
    }

    /// Hard maximum retained object records.
    #[must_use]
    pub const fn max_objects(self) -> usize {
        Self::MAX_OBJECTS
    }

    /// Hard maximum retained picture records.
    #[must_use]
    pub const fn max_pictures(self) -> usize {
        Self::MAX_PICTURES
    }

    /// Hard maximum decoded `objdata` payload accepted by the parser.
    #[must_use]
    pub const fn max_object_payload_bytes(self) -> usize {
        Self::MAX_OBJECT_PAYLOAD_BYTES
    }

    /// Hard maximum decoded picture payload accepted by the parser.
    #[must_use]
    pub const fn max_picture_payload_bytes(self) -> usize {
        Self::MAX_PICTURE_PAYLOAD_BYTES
    }
}

/// A content-free structural and security inventory for one parsed RTF.
///
/// The report is observational. It never mutates the document, follows a
/// target, opens a file, fetches a URL, evaluates a field, executes a macro or
/// object, repairs syntax, or serializes a candidate. A clean status means
/// "no parser-proven surface in the supported scope"; it is not a claim that
/// an unknown producer extension is harmless.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ValidationReport {
    syntax: ValidationCheck,
    root: ValidationCheck,
    document: ValidationCheck,
    compressed_transport: ValidationCheck,
    fields: ValidationCheck,
    external_links: ValidationCheck,
    objects: ValidationCheck,
    pictures: ValidationCheck,
    active_content: ValidationCheck,
    unsupported_syntax: ValidationCheck,
    external_resolution: ValidationCheck,
    execution: ValidationCheck,
    repair: ValidationCheck,
    security: ValidationCheck,
    counts: ValidationCounts,
    limits: ValidationLimits,
}

impl ValidationReport {
    /// Parse UTF-8 RTF once and produce its bounded report.
    #[allow(
        clippy::should_implement_trait,
        reason = "the inherent constructor keeps the validation API parallel with Document::parse"
    )]
    pub fn from_str(input: &str) -> RtfResult<Self> {
        Self::from_str_with_limits(input, ParseLimits::default())
    }

    /// Parse UTF-8 RTF once with an explicit finite profile.
    pub fn from_str_with_limits(input: &str, limits: ParseLimits) -> RtfResult<Self> {
        let document = crate::Document::parse_with_limits(input, limits)?;
        Ok(document.validation_report())
    }

    /// Parse original RTF bytes once, including bounded compressed-RTF
    /// decompression when the transport header is present.
    pub fn from_bytes(input: &[u8]) -> RtfResult<Self> {
        Self::from_bytes_with_limits(input, ParseLimits::default())
    }

    /// Parse original RTF bytes once with an explicit finite profile.
    pub fn from_bytes_with_limits(input: &[u8], limits: ParseLimits) -> RtfResult<Self> {
        let document = crate::Document::from_bytes_with_limits(input, limits)?;
        Ok(document.validation_report())
    }

    /// Build a report from an already parsed immutable facade without parsing
    /// or traversing the same document once per check.
    #[must_use]
    pub fn from_document(document: &crate::Document) -> Self {
        build_report(document.model(), document.limits())
    }

    /// Build a report from the advanced retained model without mutating it.
    #[must_use]
    pub fn from_raw(document: &crate::raw::Document<'_>) -> Self {
        build_report(document, document.parse_limits())
    }

    /// Structural lexer/token validation completed.
    #[must_use]
    pub const fn syntax(self) -> ValidationCheck {
        self.syntax
    }

    /// Root-group and `\\rtf` header validation completed.
    #[must_use]
    pub const fn root(self) -> ValidationCheck {
        self.root
    }

    /// The retained semantic document parse completed.
    #[must_use]
    pub const fn document(self) -> ValidationCheck {
        self.document
    }

    /// Whether a compressed-RTF transport frame was accepted.
    #[must_use]
    pub const fn compressed_transport(self) -> ValidationCheck {
        self.compressed_transport
    }

    /// Whether generic fields were retained.
    #[must_use]
    pub const fn fields(self) -> ValidationCheck {
        self.fields
    }

    /// Whether a parser-proven external hyperlink/reference surface exists.
    #[must_use]
    pub const fn external_links(self) -> ValidationCheck {
        self.external_links
    }

    /// Whether object destinations were retained.
    #[must_use]
    pub const fn objects(self) -> ValidationCheck {
        self.objects
    }

    /// Whether picture destinations were retained.
    #[must_use]
    pub const fn pictures(self) -> ValidationCheck {
        self.pictures
    }

    /// Whether parser-proven active-content surfaces were retained.
    #[must_use]
    pub const fn active_content(self) -> ValidationCheck {
        self.active_content
    }

    /// Whether unsupported syntax was retained as bounded inert nodes.
    #[must_use]
    pub const fn unsupported_syntax(self) -> ValidationCheck {
        self.unsupported_syntax
    }

    /// Whether an external-resource provider would be needed to continue.
    #[must_use]
    pub const fn external_resolution(self) -> ValidationCheck {
        self.external_resolution
    }

    /// Whether an execution provider would be needed to continue.
    #[must_use]
    pub const fn execution(self) -> ValidationCheck {
        self.execution
    }

    /// Repair is intentionally outside this report API.
    #[must_use]
    pub const fn repair(self) -> ValidationCheck {
        self.repair
    }

    /// Conservative aggregate security state.
    #[must_use]
    pub const fn security(self) -> ValidationCheck {
        self.security
    }

    /// Bounded, content-free cardinalities.
    #[must_use]
    pub const fn counts(self) -> ValidationCounts {
        self.counts
    }

    /// Byte/depth/group/object/picture limits governing the parse.
    #[must_use]
    pub const fn limits(self) -> ValidationLimits {
        self.limits
    }

    /// True only when no known or unknown security surface was retained.
    #[must_use]
    pub const fn is_conservatively_clean(self) -> bool {
        matches!(self.security.status, ValidationStatus::Absent)
    }
}

/// Validate original RTF bytes with the default finite profile.
pub fn validate(input: &[u8]) -> RtfResult<ValidationReport> {
    ValidationReport::from_bytes(input)
}

/// Validate original RTF bytes with an explicit finite profile.
pub fn validate_with_limits(input: &[u8], limits: ParseLimits) -> RtfResult<ValidationReport> {
    ValidationReport::from_bytes_with_limits(input, limits)
}

/// Validate UTF-8 RTF with the default finite profile.
pub fn validate_str(input: &str) -> RtfResult<ValidationReport> {
    ValidationReport::from_str(input)
}

/// Validate UTF-8 RTF with an explicit finite profile.
pub fn validate_str_with_limits(input: &str, limits: ParseLimits) -> RtfResult<ValidationReport> {
    ValidationReport::from_str_with_limits(input, limits)
}

fn build_report(document: &RtfDocument<'_>, parse_limits: ParseLimits) -> ValidationReport {
    let counts = ValidationCounts {
        source_bytes: document.preserved_source().map_or(0, <[u8]>::len),
        fields: document.fields().len(),
        objects: document.objects().len(),
        pictures: document.pictures().len(),
        form_fields: document.form_fields().len(),
        opaque_nodes: document.opaque_nodes().len(),
        opaque_bytes: document.opaque_nodes().iter().fold(0usize, |total, node| {
            total.saturating_add(node.source().len())
        }),
        unknown_syntax_markers: document.unknown_syntax_markers(),
    };

    let source_is_compressed = document
        .preserved_source()
        .is_some_and(crate::transport::is_compressed_rtf);
    let field_safety = document.field_safety();
    let unknown_field = field_safety.len() != document.fields().len()
        || document
            .fields()
            .iter()
            .any(|field| field.field_type == FieldType::Unknown);
    let unknown_picture = document
        .pictures()
        .iter()
        .any(|picture| picture.image_type == crate::ImageType::Unknown);
    let opaque = counts.opaque_nodes > 0 || counts.unknown_syntax_markers > 0;

    let mut external = !document.external_references().is_empty()
        || document.mail_merge().is_some()
        || document.xsl_transform().is_some()
        || document.xsl_transform_usage().is_requested();
    let mut active = !document.objects().is_empty()
        || !document.form_fields().is_empty()
        || document.mail_merge().is_some()
        || document.xsl_transform().is_some()
        || document.xsl_transform_usage().is_requested();
    let mut unknown_external = false;
    let mut unknown_active = false;
    let mut unknown_object = false;

    // Object destinations are inert records, but link-like and unrecognized
    // object kinds still affect the conservative external/active inventory.
    // No object payload, class name, or target text is exposed by this report.
    for object in document.objects() {
        match object.kind {
            crate::ObjectKind::Link
            | crate::ObjectKind::AutoLink
            | crate::ObjectKind::Subscriber
            | crate::ObjectKind::Publisher => {
                // `linkself` explicitly denotes an in-document object link;
                // only the remaining link-like modes prove an external
                // resource surface.
                if !object.link_self {
                    external = true;
                }
            },
            crate::ObjectKind::Unknown => {
                unknown_object = true;
                unknown_external = true;
                unknown_active = true;
            },
            crate::ObjectKind::Embedded
            | crate::ObjectKind::Html
            | crate::ObjectKind::InstallableCommand
            | crate::ObjectKind::OleControl => {},
        }
    }

    // Field semantics were classified once while the parser built the model.
    // A missing or out-of-sync cache is itself unsafe and never treated as a
    // clean result.
    if field_safety.len() == document.fields().len() {
        for safety in field_safety {
            match safety {
                FieldSafety::Neutral => {},
                FieldSafety::External => external = true,
                FieldSafety::ExternalUnknown => {
                    external = true;
                    unknown_external = true;
                },
                FieldSafety::Active => active = true,
                FieldSafety::ActiveUnknown => {
                    active = true;
                    unknown_active = true;
                },
                FieldSafety::ExternalAndActive => {
                    external = true;
                    active = true;
                },
                FieldSafety::ExternalAndActiveUnknown => {
                    external = true;
                    active = true;
                    unknown_external = true;
                    unknown_active = true;
                },
            }
        }
    } else {
        unknown_external = true;
        unknown_active = true;
    }

    let unsupported = if opaque {
        ValidationStatus::Present
    } else {
        ValidationStatus::Absent
    };
    let fields_status = if unknown_field {
        ValidationStatus::Unknown
    } else if counts.fields > 0 {
        ValidationStatus::Present
    } else {
        ValidationStatus::Absent
    };
    let external_status = surface_status(external, unknown_external || opaque);
    let active_status = surface_status(active, unknown_active || opaque);
    let object_status = surface_status(counts.objects > 0, opaque || unknown_object);
    let picture_status = surface_status(counts.pictures > 0, opaque || unknown_picture);
    let security_status = if unknown_external || unknown_active || opaque || unknown_picture {
        ValidationStatus::Unknown
    } else if external || active || counts.objects > 0 || counts.pictures > 0 {
        ValidationStatus::Present
    } else {
        ValidationStatus::Absent
    };

    ValidationReport {
        syntax: ValidationCheck::new(
            if document.parse_provenance().syntax_valid {
                ValidationStatus::Valid
            } else {
                ValidationStatus::Unknown
            },
            ValidationDependency::Syntax,
        ),
        root: ValidationCheck::new(
            if document.parse_provenance().root_valid {
                ValidationStatus::Valid
            } else {
                ValidationStatus::Unknown
            },
            ValidationDependency::Root,
        ),
        document: ValidationCheck::new(
            if document.parse_provenance().document_valid {
                ValidationStatus::Valid
            } else {
                ValidationStatus::Unknown
            },
            ValidationDependency::Document,
        ),
        compressed_transport: ValidationCheck::new(
            if source_is_compressed {
                ValidationStatus::Present
            } else {
                ValidationStatus::NotApplicable
            },
            ValidationDependency::CompressedTransport,
        ),
        fields: ValidationCheck::new(fields_status, ValidationDependency::Document),
        external_links: ValidationCheck::new(external_status, ValidationDependency::Document),
        objects: ValidationCheck::new(object_status, ValidationDependency::ObjectParser),
        pictures: ValidationCheck::new(picture_status, ValidationDependency::PictureParser),
        active_content: ValidationCheck::new(
            active_status,
            ValidationDependency::ExecutionProvider,
        ),
        unsupported_syntax: ValidationCheck::new(
            unsupported,
            ValidationDependency::OpaquePreservation,
        ),
        external_resolution: ValidationCheck::new(
            if matches!(external_status, ValidationStatus::Unknown) {
                ValidationStatus::Unknown
            } else if external {
                ValidationStatus::Unsupported
            } else {
                ValidationStatus::NotApplicable
            },
            ValidationDependency::ExternalProvider,
        ),
        execution: ValidationCheck::new(
            if matches!(active_status, ValidationStatus::Unknown) {
                ValidationStatus::Unknown
            } else if active {
                ValidationStatus::Unsupported
            } else {
                ValidationStatus::NotApplicable
            },
            ValidationDependency::ExecutionProvider,
        ),
        repair: ValidationCheck::new(
            ValidationStatus::NotApplicable,
            ValidationDependency::RepairPlanner,
        ),
        security: ValidationCheck::new(security_status, ValidationDependency::SecurityPolicy),
        counts,
        limits: ValidationLimits::new(parse_limits),
    }
}

const fn surface_status(present: bool, unknown: bool) -> ValidationStatus {
    if unknown {
        ValidationStatus::Unknown
    } else if present {
        ValidationStatus::Present
    } else {
        ValidationStatus::Absent
    }
}
