//! ODF metadata parsing functionality.
//!
//! This module provides comprehensive parsing of ODF metadata from meta.xml,
//! including document properties, statistics, and user information.

use chrono::{DateTime, Utc};
use litchi_core::{Error, Metadata, Result, xml::escape_xml};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::writer::Writer;
use std::collections::HashMap;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const META_NAMESPACE_STR: &str = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
const META_NAMESPACE: &[u8] = META_NAMESPACE_STR.as_bytes();
const DC_NAMESPACE_STR: &str = "http://purl.org/dc/elements/1.1/";
const DC_NAMESPACE: &[u8] = DC_NAMESPACE_STR.as_bytes();
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";

/// Comprehensive ODF metadata
#[derive(Debug, Clone, Default)]
pub struct OdfMetadata {
    /// Document title
    pub title: Option<String>,
    /// Document description
    pub description: Option<String>,
    /// Document subject
    pub subject: Option<String>,
    /// Document keywords
    pub keywords: Vec<String>,
    /// Document creator/author
    pub creator: Option<String>,
    /// Original document creator
    pub initial_creator: Option<String>,
    /// Person who last printed the document
    pub printed_by: Option<String>,
    /// Document language
    pub language: Option<String>,
    /// Entity responsible for making contributions to the document
    pub contributor: Option<String>,
    /// Entity responsible for making the document available
    pub publisher: Option<String>,
    /// Rights held in and over the document
    pub rights: Option<String>,
    /// Spatial or temporal topic coverage of the document
    pub coverage: Option<String>,
    /// File format or medium of the document
    pub format: Option<String>,
    /// Unambiguous reference to the document
    pub identifier: Option<String>,
    /// Related resource
    pub relation: Option<String>,
    /// Resource from which the document is derived
    pub source: Option<String>,
    /// Nature or genre of the document
    pub r#type: Option<String>,
    /// Creation date
    pub creation_date: Option<String>,
    /// Last modification date
    pub modification_date: Option<String>,
    /// Last print date
    pub print_date: Option<String>,
    /// Generator application
    pub generator: Option<String>,
    /// Exact non-negative editing-cycle count
    pub editing_cycles: Option<String>,
    /// Exact XML Schema duration spent editing
    pub editing_duration: Option<String>,
    /// Template reference, if present
    pub template: Option<TemplateMetadata>,
    /// Automatic reload behavior, if present
    pub auto_reload: Option<AutoReloadMetadata>,
    /// Default hyperlink behavior, if present
    pub hyperlink_behaviour: Option<HyperlinkBehaviourMetadata>,
    /// Document statistics
    pub statistics: DocumentStatistics,
    /// Custom properties
    pub custom_properties: HashMap<String, String>,
    /// Ordered, typed user-defined metadata, including duplicate names
    pub user_defined: Vec<UserDefinedMetadata>,
}

/// A `meta:user-defined` property with its exact lexical value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UserDefinedMetadata {
    /// Property name.
    pub name: String,
    /// Declared ODF value type. Missing declarations default to `string`.
    pub value_type: UserDefinedValueType,
    /// Exact decoded element text.
    pub value: String,
}

/// Standard ODF user-defined metadata value types.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UserDefinedValueType {
    /// XML Schema double.
    Float,
    /// XML Schema date or dateTime.
    Date,
    /// XML Schema duration.
    Time,
    /// XML Schema boolean.
    Boolean,
    /// String value.
    String,
}

/// Metadata describing the template used to create a document.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TemplateMetadata {
    /// Template URI.
    pub href: Option<String>,
    /// Human-readable template title.
    pub title: Option<String>,
    /// Template dateTime lexical value.
    pub date: Option<String>,
    /// XLink activation behavior.
    pub actuate: Option<String>,
}

/// Metadata describing automatic document reload behavior.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct AutoReloadMetadata {
    /// Reload URI.
    pub href: Option<String>,
    /// Exact XML Schema duration delay.
    pub delay: Option<String>,
    /// XLink show behavior.
    pub show: Option<String>,
    /// XLink activation behavior.
    pub actuate: Option<String>,
}

/// Metadata describing default hyperlink behavior.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct HyperlinkBehaviourMetadata {
    /// Target frame name.
    pub target_frame_name: Option<String>,
    /// XLink show behavior.
    pub show: Option<String>,
}

/// Document statistics from metadata
#[derive(Debug, Clone, Default)]
pub struct DocumentStatistics {
    /// Number of pages
    pub page_count: Option<String>,
    /// Number of paragraphs
    pub paragraph_count: Option<String>,
    /// Number of words
    pub word_count: Option<String>,
    /// Number of characters
    pub character_count: Option<String>,
    /// Number of tables
    pub table_count: Option<String>,
    /// Number of drawing objects
    pub draw_count: Option<String>,
    /// Number of images
    pub image_count: Option<String>,
    /// Number of embedded OLE objects
    pub ole_object_count: Option<String>,
    /// Number of objects
    pub object_count: Option<String>,
    /// Number of frames
    pub frame_count: Option<String>,
    /// Number of sentences
    pub sentence_count: Option<String>,
    /// Number of syllables
    pub syllable_count: Option<String>,
    /// Number of non-whitespace characters
    pub non_whitespace_character_count: Option<String>,
    /// Number of spreadsheet rows
    pub row_count: Option<String>,
    /// Number of spreadsheet cells
    pub cell_count: Option<String>,
}

impl OdfMetadata {
    /// Parse metadata from meta.xml content
    pub fn from_xml(xml_content: &str) -> Result<Self> {
        let mut reader = NsReader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut metadata = OdfMetadata::default();
        let mut depth = 0usize;
        let mut meta_depth = None;

        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buf)
                .map_err(metadata_xml_error)?;
            let namespace = known_namespace(&namespace);
            match event {
                Event::Start(ref element) => {
                    let local_name = element.local_name();
                    if meta_depth.is_none()
                        && namespace == KnownNamespace::Office
                        && local_name.as_ref() == b"meta"
                    {
                        depth = checked_depth_add(depth)?;
                        meta_depth = Some(depth);
                    } else if meta_depth == Some(depth) {
                        if let Some(field) = text_field(&namespace, local_name.as_ref()) {
                            let value = extract_text_content(&mut reader)?;
                            assign_text_field(&mut metadata, field, value);
                        } else if namespace == KnownNamespace::Meta
                            && local_name.as_ref() == b"user-defined"
                        {
                            let mut property =
                                parse_user_defined_property(element, &reader, String::new())?;
                            property.value = extract_text_content(&mut reader)?;
                            metadata
                                .custom_properties
                                .insert(property.name.clone(), property.value.clone());
                            metadata.user_defined.push(property);
                        } else {
                            parse_metadata_attributes(
                                &mut metadata,
                                &namespace,
                                local_name.as_ref(),
                                element,
                                &reader,
                            )?;
                            depth = checked_depth_add(depth)?;
                        }
                    } else {
                        depth = checked_depth_add(depth)?;
                    }
                },
                Event::Empty(ref element) => {
                    let local_name = element.local_name();
                    if meta_depth == Some(depth) {
                        if let Some(field) = text_field(&namespace, local_name.as_ref()) {
                            assign_text_field(&mut metadata, field, String::new());
                        } else if namespace == KnownNamespace::Meta
                            && local_name.as_ref() == b"user-defined"
                        {
                            let property =
                                parse_user_defined_property(element, &reader, String::new())?;
                            metadata
                                .custom_properties
                                .insert(property.name.clone(), property.value.clone());
                            metadata.user_defined.push(property);
                        } else {
                            parse_metadata_attributes(
                                &mut metadata,
                                &namespace,
                                local_name.as_ref(),
                                element,
                                &reader,
                            )?;
                        }
                    }
                },
                Event::End(ref element) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat(
                            "unexpected closing tag in OpenDocument metadata".to_string(),
                        )
                    })?;
                    if meta_depth == Some(depth + 1)
                        && namespace == KnownNamespace::Office
                        && element.local_name().as_ref() == b"meta"
                    {
                        meta_depth = None;
                    }
                },
                Event::Eof => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(metadata)
    }

    /// Parse a date string into DateTime<Utc>
    fn parse_date(date_str: Option<String>) -> Option<DateTime<Utc>> {
        date_str.and_then(|value| {
            crate::datatype::DateTime::decode(&value)
                .ok()
                .map(|date| date.with_timezone(&Utc))
        })
    }
}

#[derive(Clone, Copy)]
enum TextField {
    Title,
    Description,
    Subject,
    Keyword,
    Creator,
    Language,
    Contributor,
    Publisher,
    Rights,
    Coverage,
    Format,
    Identifier,
    Relation,
    Source,
    Type,
    Date,
    InitialCreator,
    PrintedBy,
    CreationDate,
    PrintDate,
    Generator,
    EditingCycles,
    EditingDuration,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum KnownNamespace {
    Office,
    Meta,
    Dc,
    Other,
}

fn known_namespace(namespace: &ResolveResult<'_>) -> KnownNamespace {
    if namespace_is(namespace, OFFICE_NAMESPACE) {
        KnownNamespace::Office
    } else if namespace_is(namespace, META_NAMESPACE) {
        KnownNamespace::Meta
    } else if namespace_is(namespace, DC_NAMESPACE) {
        KnownNamespace::Dc
    } else {
        KnownNamespace::Other
    }
}

fn namespace_is(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

fn text_field(namespace: &KnownNamespace, local_name: &[u8]) -> Option<TextField> {
    if *namespace == KnownNamespace::Dc {
        return match local_name {
            b"title" => Some(TextField::Title),
            b"description" => Some(TextField::Description),
            b"subject" => Some(TextField::Subject),
            b"creator" => Some(TextField::Creator),
            b"language" => Some(TextField::Language),
            b"contributor" => Some(TextField::Contributor),
            b"publisher" => Some(TextField::Publisher),
            b"rights" => Some(TextField::Rights),
            b"coverage" => Some(TextField::Coverage),
            b"format" => Some(TextField::Format),
            b"identifier" => Some(TextField::Identifier),
            b"relation" => Some(TextField::Relation),
            b"source" => Some(TextField::Source),
            b"type" => Some(TextField::Type),
            b"date" => Some(TextField::Date),
            _ => None,
        };
    }
    if *namespace == KnownNamespace::Meta {
        return match local_name {
            b"keyword" => Some(TextField::Keyword),
            b"initial-creator" => Some(TextField::InitialCreator),
            b"printed-by" => Some(TextField::PrintedBy),
            b"creation-date" => Some(TextField::CreationDate),
            b"print-date" => Some(TextField::PrintDate),
            b"generator" => Some(TextField::Generator),
            b"editing-cycles" => Some(TextField::EditingCycles),
            b"editing-duration" => Some(TextField::EditingDuration),
            _ => None,
        };
    }
    None
}

fn assign_text_field(metadata: &mut OdfMetadata, field: TextField, value: String) {
    match field {
        TextField::Title => metadata.title = Some(value),
        TextField::Description => metadata.description = Some(value),
        TextField::Subject => metadata.subject = Some(value),
        TextField::Keyword => metadata.keywords.push(value),
        TextField::Creator => metadata.creator = Some(value),
        TextField::Language => metadata.language = Some(value),
        TextField::Contributor => metadata.contributor = Some(value),
        TextField::Publisher => metadata.publisher = Some(value),
        TextField::Rights => metadata.rights = Some(value),
        TextField::Coverage => metadata.coverage = Some(value),
        TextField::Format => metadata.format = Some(value),
        TextField::Identifier => metadata.identifier = Some(value),
        TextField::Relation => metadata.relation = Some(value),
        TextField::Source => metadata.source = Some(value),
        TextField::Type => metadata.r#type = Some(value),
        TextField::Date => metadata.modification_date = Some(value),
        TextField::InitialCreator => metadata.initial_creator = Some(value),
        TextField::PrintedBy => metadata.printed_by = Some(value),
        TextField::CreationDate => metadata.creation_date = Some(value),
        TextField::PrintDate => metadata.print_date = Some(value),
        TextField::Generator => metadata.generator = Some(value),
        TextField::EditingCycles => metadata.editing_cycles = Some(value),
        TextField::EditingDuration => metadata.editing_duration = Some(value),
    }
}

/// Edit applied to a simple text metadata element when saving a mutable
/// document that retains its source meta.xml.
#[derive(Debug, Clone, PartialEq, Eq)]
enum MetaFieldEdit {
    /// Keep the source element, or its absence, untouched.
    Preserve,
    /// Set the element text, inserting the element when it is absent.
    Set(String),
    /// Remove the element from the source meta.xml.
    Remove,
}

impl MetaFieldEdit {
    /// Derive the edit between the value the source produced and the current
    /// mutable value. An unchanged value keeps the source element untouched.
    fn between(expected: Option<&str>, current: Option<&str>) -> Self {
        if expected == current {
            Self::Preserve
        } else if let Some(value) = current {
            Self::Set(value.to_string())
        } else {
            Self::Remove
        }
    }
}

/// Field-level edits applied to a retained source meta.xml during a mutable
/// save. Every element not named by an edit is preserved unchanged.
#[derive(Debug, Clone)]
pub(crate) struct MetaXmlPatch {
    generator: MetaFieldEdit,
    modification_date: MetaFieldEdit,
    title: MetaFieldEdit,
    creator: MetaFieldEdit,
    subject: MetaFieldEdit,
    description: MetaFieldEdit,
    keywords: MetaFieldEdit,
}

impl MetaXmlPatch {
    /// Start from a patch that preserves every source element.
    pub(crate) fn preserve_all() -> Self {
        Self {
            generator: MetaFieldEdit::Preserve,
            modification_date: MetaFieldEdit::Preserve,
            title: MetaFieldEdit::Preserve,
            creator: MetaFieldEdit::Preserve,
            subject: MetaFieldEdit::Preserve,
            description: MetaFieldEdit::Preserve,
            keywords: MetaFieldEdit::Preserve,
        }
    }

    /// Overwrite `meta:generator` and `dc:date`, matching the values a
    /// from-scratch save would emit.
    pub(crate) fn with_generator_and_modification_date(
        mut self,
        generator: &str,
        modification_date: String,
    ) -> Self {
        self.generator = MetaFieldEdit::Set(generator.to_string());
        self.modification_date = MetaFieldEdit::Set(modification_date);
        self
    }

    /// Derive edits for the simple mutable metadata fields by comparing the
    /// current mutable metadata against the metadata the source produced.
    /// Fields the user did not change keep their source element untouched.
    pub(crate) fn diff_simple_fields(mut self, source: &OdfMetadata, current: &Metadata) -> Self {
        let expected = Metadata::from(source.clone());
        self.title = MetaFieldEdit::between(expected.title.as_deref(), current.title.as_deref());
        self.creator =
            MetaFieldEdit::between(expected.author.as_deref(), current.author.as_deref());
        self.subject =
            MetaFieldEdit::between(expected.subject.as_deref(), current.subject.as_deref());
        self.description = MetaFieldEdit::between(
            expected.description.as_deref(),
            current.description.as_deref(),
        );
        self.keywords =
            MetaFieldEdit::between(expected.keywords.as_deref(), current.keywords.as_deref());
        self
    }

    fn edit(&self, field: PatchField) -> &MetaFieldEdit {
        match field {
            PatchField::Generator => &self.generator,
            PatchField::ModificationDate => &self.modification_date,
            PatchField::Title => &self.title,
            PatchField::Creator => &self.creator,
            PatchField::Subject => &self.subject,
            PatchField::Description => &self.description,
            PatchField::Keywords => &self.keywords,
        }
    }

    /// Take the pending edit for a field, leaving [`MetaFieldEdit::Preserve`].
    fn take(&mut self, field: PatchField) -> MetaFieldEdit {
        let edit = match field {
            PatchField::Generator => &mut self.generator,
            PatchField::ModificationDate => &mut self.modification_date,
            PatchField::Title => &mut self.title,
            PatchField::Creator => &mut self.creator,
            PatchField::Subject => &mut self.subject,
            PatchField::Description => &mut self.description,
            PatchField::Keywords => &mut self.keywords,
        };
        std::mem::replace(edit, MetaFieldEdit::Preserve)
    }

    /// Whether any field still waits for its element to be inserted.
    fn has_pending_insertions(&self) -> bool {
        PATCH_INSERTION_ORDER
            .iter()
            .any(|field| matches!(self.edit(*field), MetaFieldEdit::Set(_)))
    }

    /// Drop every pending insertion after it has been written.
    fn clear_pending_insertions(&mut self) {
        for field in PATCH_INSERTION_ORDER {
            if matches!(self.edit(field), MetaFieldEdit::Set(_)) {
                self.take(field);
            }
        }
    }
}

/// A simple text element of `office:meta` that a mutable save can patch.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PatchField {
    Generator,
    ModificationDate,
    Title,
    Creator,
    Subject,
    Description,
    Keywords,
}

/// Number of [`PatchField`] variants, sizing the consumed-edit bitmap.
const PATCH_FIELD_COUNT: usize = 7;

/// Insertion order for elements a patch appends to `office:meta`, following
/// the ODF schema sequence for simple metadata.
const PATCH_INSERTION_ORDER: [PatchField; PATCH_FIELD_COUNT] = [
    PatchField::Generator,
    PatchField::Title,
    PatchField::Description,
    PatchField::Subject,
    PatchField::Keywords,
    PatchField::Creator,
    PatchField::ModificationDate,
];

impl PatchField {
    const fn index(self) -> usize {
        self as usize
    }

    /// Namespace and local name of the field's element.
    const fn element(self) -> (KnownNamespace, &'static str) {
        match self {
            Self::Generator => (KnownNamespace::Meta, "generator"),
            Self::ModificationDate => (KnownNamespace::Dc, "date"),
            Self::Title => (KnownNamespace::Dc, "title"),
            Self::Creator => (KnownNamespace::Dc, "creator"),
            Self::Subject => (KnownNamespace::Dc, "subject"),
            Self::Description => (KnownNamespace::Dc, "description"),
            Self::Keywords => (KnownNamespace::Meta, "keyword"),
        }
    }
}

/// Map a direct `office:meta` child to the field a patch manages.
fn patch_field(namespace: &KnownNamespace, local_name: &[u8]) -> Option<PatchField> {
    match namespace {
        KnownNamespace::Dc => match local_name {
            b"title" => Some(PatchField::Title),
            b"creator" => Some(PatchField::Creator),
            b"subject" => Some(PatchField::Subject),
            b"description" => Some(PatchField::Description),
            b"date" => Some(PatchField::ModificationDate),
            _ => None,
        },
        KnownNamespace::Meta => match local_name {
            b"generator" => Some(PatchField::Generator),
            b"keyword" => Some(PatchField::Keywords),
            _ => None,
        },
        _ => None,
    }
}

/// Prefixes the source meta.xml root binds to the Dublin Core and ODF meta
/// namespaces, reused when a patch inserts new elements.
#[derive(Debug, Default)]
struct MetaNamespacePrefixes {
    dc: Option<String>,
    meta: Option<String>,
}

impl MetaNamespacePrefixes {
    /// Record the relevant bindings declared on the root element.
    fn observe(&mut self, element: &BytesStart<'_>, reader: &NsReader<&[u8]>) -> Result<()> {
        for attribute in element.attributes() {
            let attribute = attribute.map_err(metadata_xml_error)?;
            let bound_prefix = if attribute.key.as_ref() == b"xmlns" {
                String::new()
            } else if attribute
                .key
                .prefix()
                .is_some_and(|prefix| prefix.as_ref() == b"xmlns")
            {
                let local_name = attribute.key.local_name();
                match String::from_utf8(local_name.as_ref().to_vec()) {
                    Ok(prefix) => prefix,
                    Err(_) => continue,
                }
            } else {
                continue;
            };
            let uri = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(metadata_xml_error)?;
            if uri == DC_NAMESPACE_STR && self.dc.is_none() {
                self.dc = Some(bound_prefix);
            } else if uri == META_NAMESPACE_STR && self.meta.is_none() {
                self.meta = Some(bound_prefix);
            }
        }
        Ok(())
    }

    /// Qualified name for an inserted element, plus the namespace URI the
    /// element must declare locally when the root bound no prefix.
    fn qualified(
        &self,
        namespace: KnownNamespace,
        local_name: &str,
    ) -> (String, Option<&'static str>) {
        let (binding, fallback_prefix, namespace_uri) = match namespace {
            KnownNamespace::Dc => (&self.dc, "dc", DC_NAMESPACE_STR),
            KnownNamespace::Meta => (&self.meta, "meta", META_NAMESPACE_STR),
            _ => unreachable!("inserted metadata fields only use the dc and meta namespaces"),
        };
        match binding {
            Some(prefix) if prefix.is_empty() => (local_name.to_string(), None),
            Some(prefix) => (format!("{prefix}:{local_name}"), None),
            None => (
                format!("{fallback_prefix}:{local_name}"),
                Some(namespace_uri),
            ),
        }
    }
}

/// Apply field-level edits to a retained source meta.xml.
///
/// Elements and attributes not named by an edit are copied through unchanged
/// and in their original order. A `Set` edit replaces the text of an existing
/// element in place, or appends the element to `office:meta` when the source
/// does not contain it. A `Remove` edit drops the element.
///
/// Returns `Ok(None)` when the source has no `office:meta` to patch, so the
/// caller can fall back to generating meta.xml from scratch.
pub(crate) fn patch_meta_xml(source: &str, patch: &MetaXmlPatch) -> Result<Option<String>> {
    let mut reader = NsReader::from_str(source);
    let mut buffer = Vec::new();
    let mut writer = Writer::new(Vec::with_capacity(source.len()));
    let mut edits = patch.clone();
    let mut consumed = [false; PATCH_FIELD_COUNT];
    let mut prefixes = MetaNamespacePrefixes::default();
    let mut depth = 0usize;
    let mut meta_depth = None;
    let mut patched = false;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(metadata_xml_error)?;
        let namespace = known_namespace(&namespace);
        match event {
            Event::Start(ref element) => {
                let local_name = element.local_name();
                if depth == 0 {
                    prefixes.observe(element, &reader)?;
                }
                if meta_depth.is_none()
                    && namespace == KnownNamespace::Office
                    && local_name.as_ref() == b"meta"
                {
                    patched = true;
                    depth = checked_depth_add(depth)?;
                    meta_depth = Some(depth);
                    write_event(&mut writer, event.clone())?;
                } else if meta_depth == Some(depth)
                    && let Some(field) = patch_field(&namespace, local_name.as_ref())
                    && !matches!(edits.edit(field), MetaFieldEdit::Preserve)
                {
                    if consumed[field.index()] {
                        // Drop duplicate occurrences of a managed element.
                        skip_element_body(&mut reader)?;
                    } else {
                        consumed[field.index()] = true;
                        if let MetaFieldEdit::Set(value) = edits.take(field) {
                            write_event(&mut writer, event.clone())?;
                            write_text(&mut writer, &value)?;
                            let end = skip_element_body(&mut reader)?;
                            write_event(&mut writer, Event::End(end))?;
                        } else {
                            skip_element_body(&mut reader)?;
                        }
                    }
                } else {
                    depth = checked_depth_add(depth)?;
                    write_event(&mut writer, event.clone())?;
                }
            },
            Event::Empty(ref element) => {
                let local_name = element.local_name();
                if depth == 0 {
                    prefixes.observe(element, &reader)?;
                }
                if meta_depth.is_none()
                    && namespace == KnownNamespace::Office
                    && local_name.as_ref() == b"meta"
                {
                    patched = true;
                    if edits.has_pending_insertions() {
                        write_event(&mut writer, Event::Start(element.clone()))?;
                        write_pending_insertions(&mut writer, &edits, &prefixes)?;
                        edits.clear_pending_insertions();
                        let name = element_name(element)?;
                        write_event(&mut writer, Event::End(BytesEnd::new(name)))?;
                    } else {
                        write_event(&mut writer, event.clone())?;
                    }
                } else if meta_depth == Some(depth)
                    && let Some(field) = patch_field(&namespace, local_name.as_ref())
                    && !matches!(edits.edit(field), MetaFieldEdit::Preserve)
                {
                    if !consumed[field.index()] {
                        consumed[field.index()] = true;
                        if let MetaFieldEdit::Set(value) = edits.take(field) {
                            let name = element_name(element)?;
                            write_event(&mut writer, Event::Start(element.clone()))?;
                            write_text(&mut writer, &value)?;
                            write_event(&mut writer, Event::End(BytesEnd::new(name)))?;
                        }
                    }
                    // Removed fields and duplicates drop the empty element.
                } else {
                    write_event(&mut writer, event.clone())?;
                }
            },
            Event::End(ref element) => {
                let closes_meta = meta_depth == Some(depth)
                    && namespace == KnownNamespace::Office
                    && element.local_name().as_ref() == b"meta";
                if closes_meta {
                    write_pending_insertions(&mut writer, &edits, &prefixes)?;
                    edits.clear_pending_insertions();
                    meta_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat(
                        "unexpected closing tag in OpenDocument metadata".to_string(),
                    )
                })?;
                write_event(&mut writer, event.clone())?;
            },
            Event::Eof => break,
            event => write_event(&mut writer, event)?,
        }
        buffer.clear();
    }

    if !patched {
        return Ok(None);
    }
    String::from_utf8(writer.into_inner())
        .map(Some)
        .map_err(|error| Error::InvalidFormat(format!("patched metadata is not UTF-8: {error}")))
}

/// Write elements for every `Set` edit that no source element consumed,
/// following the ODF schema sequence for simple metadata.
fn write_pending_insertions(
    writer: &mut Writer<Vec<u8>>,
    edits: &MetaXmlPatch,
    prefixes: &MetaNamespacePrefixes,
) -> Result<()> {
    for field in PATCH_INSERTION_ORDER {
        if let MetaFieldEdit::Set(value) = edits.edit(field) {
            let (namespace, local_name) = field.element();
            let (name, declaration) = prefixes.qualified(namespace, local_name);
            write_text_element(writer, &name, declaration, value)?;
        }
    }
    Ok(())
}

/// Write a simple text element, declaring its namespace locally when the
/// source document bound no prefix for it.
fn write_text_element(
    writer: &mut Writer<Vec<u8>>,
    name: &str,
    namespace_declaration: Option<&'static str>,
    value: &str,
) -> Result<()> {
    let mut start = BytesStart::new(name);
    if let Some(namespace_uri) = namespace_declaration {
        let prefix = name.split(':').next().unwrap_or(name);
        let declaration_name = format!("xmlns:{prefix}");
        start.push_attribute((declaration_name.as_str(), namespace_uri));
    }
    write_event(writer, Event::Start(start))?;
    write_text(writer, value)?;
    write_event(writer, Event::End(BytesEnd::new(name)))
}

/// Write an already-escaped text node.
fn write_text(writer: &mut Writer<Vec<u8>>, value: &str) -> Result<()> {
    write_event(
        writer,
        Event::Text(BytesText::from_escaped(escape_xml(value))),
    )
}

/// Write an event unchanged.
fn write_event(writer: &mut Writer<Vec<u8>>, event: Event<'_>) -> Result<()> {
    writer.write_event(event).map_err(metadata_xml_error)
}

/// Skip events through the end tag of the element whose start tag was just
/// consumed, returning that end tag.
fn skip_element_body(reader: &mut NsReader<&[u8]>) -> Result<BytesEnd<'static>> {
    let mut buffer = Vec::new();
    let mut depth = 1usize;
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(metadata_xml_error)?;
        match event {
            Event::Start(_) => depth = checked_depth_add(depth)?,
            Event::End(end) => {
                depth -= 1;
                if depth == 0 {
                    return Ok(end.into_owned());
                }
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "unexpected end of OpenDocument metadata element".to_string(),
                ));
            },
            _ => {},
        }
        buffer.clear();
    }
}

/// Qualified name of an element as UTF-8.
fn element_name(element: &BytesStart<'_>) -> Result<String> {
    String::from_utf8(element.name().as_ref().to_vec())
        .map_err(|error| Error::InvalidFormat(format!("invalid metadata element name: {error}")))
}

fn extract_text_content(reader: &mut NsReader<&[u8]>) -> Result<String> {
    let mut buffer = Vec::new();
    let mut content = String::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(metadata_xml_error)?;
        match event {
            Event::Text(text) => content.push_str(
                &text
                    .xml_content(XmlVersion::Implicit1_0)
                    .map_err(metadata_xml_error)?,
            ),
            Event::CData(text) => content.push_str(
                &text
                    .xml_content(XmlVersion::Implicit1_0)
                    .map_err(metadata_xml_error)?,
            ),
            Event::GeneralRef(reference) => {
                if let Some(character) = reference.resolve_char_ref().map_err(metadata_xml_error)? {
                    content.push(character);
                } else {
                    let name = reference.decode().map_err(metadata_xml_error)?;
                    let replacement = quick_xml::escape::resolve_predefined_entity(&name)
                        .ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "unresolved entity '&{name};' in OpenDocument metadata"
                            ))
                        })?;
                    content.push_str(replacement);
                }
            },
            Event::End(_) => return Ok(content),
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "unexpected end of OpenDocument metadata text".to_string(),
                ));
            },
            _ => {
                return Err(Error::InvalidFormat(
                    "nested markup is not allowed in OpenDocument metadata values".to_string(),
                ));
            },
        }
        buffer.clear();
    }
}

fn parse_metadata_attributes(
    metadata: &mut OdfMetadata,
    namespace: &KnownNamespace,
    local_name: &[u8],
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
) -> Result<()> {
    if *namespace != KnownNamespace::Meta {
        return Ok(());
    }
    match local_name {
        b"document-statistic" => {
            metadata.statistics = parse_document_statistics(element, reader)?;
        },
        b"template" => {
            metadata.template = Some(TemplateMetadata {
                href: attribute_value(element, reader, XLINK_NAMESPACE, b"href")?,
                title: attribute_value(element, reader, XLINK_NAMESPACE, b"title")?,
                date: attribute_value(element, reader, META_NAMESPACE, b"date")?,
                actuate: attribute_value(element, reader, XLINK_NAMESPACE, b"actuate")?,
            });
        },
        b"auto-reload" => {
            metadata.auto_reload = Some(AutoReloadMetadata {
                href: attribute_value(element, reader, XLINK_NAMESPACE, b"href")?,
                delay: attribute_value(element, reader, META_NAMESPACE, b"delay")?,
                show: attribute_value(element, reader, XLINK_NAMESPACE, b"show")?,
                actuate: attribute_value(element, reader, XLINK_NAMESPACE, b"actuate")?,
            });
        },
        b"hyperlink-behaviour" => {
            metadata.hyperlink_behaviour = Some(HyperlinkBehaviourMetadata {
                target_frame_name: attribute_value(
                    element,
                    reader,
                    OFFICE_NAMESPACE,
                    b"target-frame-name",
                )?,
                show: attribute_value(element, reader, XLINK_NAMESPACE, b"show")?,
            });
        },
        _ => {},
    }
    Ok(())
}

fn parse_document_statistics(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
) -> Result<DocumentStatistics> {
    let mut statistics = DocumentStatistics::default();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(metadata_xml_error)?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if !namespace_is(&namespace, META_NAMESPACE) {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(metadata_xml_error)?
            .into_owned();
        if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit()) {
            return Err(Error::InvalidFormat(format!(
                "invalid non-negative ODF statistic '{value}'"
            )));
        }
        match local_name.as_ref() {
            b"page-count" => statistics.page_count = Some(value),
            b"paragraph-count" => statistics.paragraph_count = Some(value),
            b"word-count" => statistics.word_count = Some(value),
            b"character-count" => statistics.character_count = Some(value),
            b"table-count" => statistics.table_count = Some(value),
            b"draw-count" => statistics.draw_count = Some(value),
            b"image-count" => statistics.image_count = Some(value),
            b"ole-object-count" => statistics.ole_object_count = Some(value),
            b"object-count" => statistics.object_count = Some(value),
            b"frame-count" => statistics.frame_count = Some(value),
            b"sentence-count" => statistics.sentence_count = Some(value),
            b"syllable-count" => statistics.syllable_count = Some(value),
            b"non-whitespace-character-count" => {
                statistics.non_whitespace_character_count = Some(value);
            },
            b"row-count" => statistics.row_count = Some(value),
            b"cell-count" => statistics.cell_count = Some(value),
            _ => {},
        }
    }
    Ok(statistics)
}

fn parse_user_defined_property(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    value: String,
) -> Result<UserDefinedMetadata> {
    let name = attribute_value(element, reader, META_NAMESPACE, b"name")?.ok_or_else(|| {
        Error::InvalidFormat("meta:user-defined is missing meta:name".to_string())
    })?;
    let value_type = match attribute_value(element, reader, META_NAMESPACE, b"value-type")?
        .as_deref()
        .unwrap_or("string")
    {
        "float" => UserDefinedValueType::Float,
        "date" => UserDefinedValueType::Date,
        "time" => UserDefinedValueType::Time,
        "boolean" => UserDefinedValueType::Boolean,
        "string" => UserDefinedValueType::String,
        value_type => {
            return Err(Error::InvalidFormat(format!(
                "unsupported ODF metadata value type '{value_type}'"
            )));
        },
    };
    Ok(UserDefinedMetadata {
        name,
        value_type,
        value,
    })
}

fn attribute_value(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    expected_namespace: &[u8],
    expected_local_name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(metadata_xml_error)?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_is(&namespace, expected_namespace)
            && local_name.as_ref() == expected_local_name
        {
            return Ok(Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(metadata_xml_error)?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn checked_depth_add(depth: usize) -> Result<usize> {
    depth.checked_add(1).ok_or_else(|| {
        Error::InvalidFormat("OpenDocument metadata nesting depth overflow".to_string())
    })
}

fn metadata_xml_error(error: impl std::fmt::Display) -> Error {
    Error::InvalidFormat(format!("XML parsing error in metadata: {error}"))
}

impl From<OdfMetadata> for Metadata {
    fn from(odf_meta: OdfMetadata) -> Self {
        let author = odf_meta
            .initial_creator
            .clone()
            .or_else(|| odf_meta.creator.clone());
        Metadata {
            title: odf_meta.title,
            author,
            subject: odf_meta.subject,
            keywords: if odf_meta.keywords.is_empty() {
                None
            } else {
                Some(odf_meta.keywords.join(", "))
            },
            description: odf_meta.description,
            template: odf_meta.template.and_then(|template| template.href),
            last_modified_by: odf_meta.creator,
            revision: odf_meta.editing_cycles,
            created: OdfMetadata::parse_date(odf_meta.creation_date),
            modified: OdfMetadata::parse_date(odf_meta.modification_date),
            page_count: parse_u32_count(odf_meta.statistics.page_count),
            word_count: parse_u32_count(odf_meta.statistics.word_count),
            character_count: parse_u32_count(odf_meta.statistics.character_count),
            application: odf_meta.generator,
            last_printed_time: OdfMetadata::parse_date(odf_meta.print_date),
            ..Default::default()
        }
    }
}

fn parse_u32_count(value: Option<String>) -> Option<u32> {
    value.and_then(|value| value.parse().ok())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_odf_metadata_default() {
        let meta = OdfMetadata::default();
        assert!(meta.title.is_none());
        assert!(meta.description.is_none());
        assert!(meta.subject.is_none());
        assert!(meta.keywords.is_empty());
        assert!(meta.creator.is_none());
        assert!(meta.language.is_none());
        assert!(meta.creation_date.is_none());
        assert!(meta.modification_date.is_none());
        assert!(meta.generator.is_none());
        assert!(meta.custom_properties.is_empty());
    }

    #[test]
    fn test_odf_metadata_from_xml_empty() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
                      xmlns:dc="http://purl.org/dc/elements/1.1/">
    <office:meta>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert!(meta.title.is_none());
        assert!(meta.creator.is_none());
    }

    #[test]
    fn test_odf_metadata_from_xml_title() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:dc="http://purl.org/dc/elements/1.1/">
    <office:meta>
        <dc:title>Test Document</dc:title>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.title, Some("Test Document".to_string()));
    }

    #[test]
    fn test_odf_metadata_from_xml_creator() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:dc="http://purl.org/dc/elements/1.1/">
    <office:meta>
        <dc:creator>John Doe</dc:creator>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.creator, Some("John Doe".to_string()));
    }

    #[test]
    fn test_odf_metadata_from_xml_description() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:dc="http://purl.org/dc/elements/1.1/">
    <office:meta>
        <dc:description>This is a test document</dc:description>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(
            meta.description,
            Some("This is a test document".to_string())
        );
    }

    #[test]
    fn test_odf_metadata_from_xml_subject() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:dc="http://purl.org/dc/elements/1.1/">
    <office:meta>
        <dc:subject>Testing</dc:subject>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.subject, Some("Testing".to_string()));
    }

    #[test]
    fn test_odf_metadata_from_xml_keywords() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
    <office:meta>
        <meta:keyword>rust</meta:keyword>
        <meta:keyword>odf</meta:keyword>
        <meta:keyword>testing</meta:keyword>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.keywords, vec!["rust", "odf", "testing"]);
    }

    #[test]
    fn test_odf_metadata_from_xml_language() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:dc="http://purl.org/dc/elements/1.1/">
    <office:meta>
        <dc:language>en-US</dc:language>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.language, Some("en-US".to_string()));
    }

    #[test]
    fn test_odf_metadata_from_xml_dates() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:dc="http://purl.org/dc/elements/1.1/"
                      xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
    <office:meta>
        <meta:creation-date>2024-01-15T10:30:00Z</meta:creation-date>
        <dc:date>2024-03-20T14:45:00Z</dc:date>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.creation_date, Some("2024-01-15T10:30:00Z".to_string()));
        assert_eq!(
            meta.modification_date,
            Some("2024-03-20T14:45:00Z".to_string())
        );
    }

    #[test]
    fn test_odf_metadata_from_xml_generator() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
    <office:meta>
        <meta:generator>LibreOffice/7.0</meta:generator>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.generator, Some("LibreOffice/7.0".to_string()));
    }

    #[test]
    fn test_odf_metadata_from_xml_statistics() {
        // Note: The parser handles empty document-statistic elements
        // Statistics are parsed from attributes on the Start event
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
    <office:meta>
        <meta:document-statistic meta:page-count="5"
                                 meta:paragraph-count="42"
                                 meta:word-count="350"
                                 meta:character-count="2100"
                                 meta:table-count="3"
                                 meta:image-count="2"
                                 meta:object-count="1"></meta:document-statistic>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        // The statistics parsing happens on Start event with attributes
        assert_eq!(meta.statistics.page_count.as_deref(), Some("5"));
        assert_eq!(meta.statistics.paragraph_count.as_deref(), Some("42"));
        assert_eq!(meta.statistics.word_count.as_deref(), Some("350"));
        assert_eq!(meta.statistics.character_count.as_deref(), Some("2100"));
        assert_eq!(meta.statistics.table_count.as_deref(), Some("3"));
        assert_eq!(meta.statistics.image_count.as_deref(), Some("2"));
        assert_eq!(meta.statistics.object_count.as_deref(), Some("1"));
    }

    #[test]
    fn test_odf_metadata_from_xml_user_defined() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
    <office:meta>
        <meta:user-defined meta:name="Department">Engineering</meta:user-defined>
        <meta:user-defined meta:name="Project">Alpha</meta:user-defined>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(
            meta.custom_properties.get("Department"),
            Some(&"Engineering".to_string())
        );
        assert_eq!(
            meta.custom_properties.get("Project"),
            Some(&"Alpha".to_string())
        );
    }

    #[test]
    fn test_odf_metadata_from_xml_full() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:dc="http://purl.org/dc/elements/1.1/"
                      xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
    <office:meta>
        <dc:title>Full Test Document</dc:title>
        <dc:description>A comprehensive test</dc:description>
        <dc:subject>Testing</dc:subject>
        <dc:creator>Test Author</dc:creator>
        <dc:language>en</dc:language>
        <meta:creation-date>2024-01-01T00:00:00Z</meta:creation-date>
        <dc:date>2024-06-01T00:00:00Z</dc:date>
        <meta:generator>Test Generator</meta:generator>
        <meta:keyword>test</meta:keyword>
        <meta:document-statistic meta:page-count="10"></meta:document-statistic>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.title, Some("Full Test Document".to_string()));
        assert_eq!(meta.description, Some("A comprehensive test".to_string()));
        assert_eq!(meta.subject, Some("Testing".to_string()));
        assert_eq!(meta.creator, Some("Test Author".to_string()));
        assert_eq!(meta.language, Some("en".to_string()));
        assert_eq!(meta.creation_date, Some("2024-01-01T00:00:00Z".to_string()));
        assert_eq!(
            meta.modification_date,
            Some("2024-06-01T00:00:00Z".to_string())
        );
        assert_eq!(meta.generator, Some("Test Generator".to_string()));
        assert_eq!(meta.keywords, vec!["test"]);
        assert_eq!(meta.statistics.page_count.as_deref(), Some("10"));
    }

    #[test]
    fn parses_namespaces_entities_and_complete_metadata_without_annotation_leakage() {
        let xml = r#"<?xml version="1.0"?>
<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:m="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
            xmlns:d="http://purl.org/dc/elements/1.1/"
            xmlns:x="http://www.w3.org/1999/xlink">
  <o:meta>
    <d:title>R&amp;D &#x1F34B;</d:title>
    <d:creator>Last Editor</d:creator>
    <m:initial-creator>Original Author</m:initial-creator>
    <m:printed-by>Print Operator</m:printed-by>
    <m:print-date>2025-04-03T02:01:00</m:print-date>
    <m:editing-cycles>0000000000000000000000007</m:editing-cycles>
    <m:editing-duration>PT1H2M3.0000000001S</m:editing-duration>
    <m:template x:href="Templates/A&amp;B.ott" x:title="A &amp; B" m:date="2024-01-02T03:04:05" x:actuate="onRequest"/>
    <m:auto-reload x:href="next.fodt" m:delay="PT5M" x:show="replace" x:actuate="onLoad"/>
    <m:hyperlink-behaviour o:target-frame-name="_blank" x:show="new"/>
    <m:document-statistic m:page-count="184467440737095516160"
      m:table-count="1" m:draw-count="2" m:image-count="3"
      m:ole-object-count="4" m:object-count="5" m:paragraph-count="6"
      m:word-count="7" m:character-count="8" m:frame-count="9"
      m:sentence-count="10" m:syllable-count="11"
      m:non-whitespace-character-count="12" m:row-count="13" m:cell-count="14"/>
    <m:user-defined m:name="Flag" m:value-type="boolean">true</m:user-defined>
    <m:user-defined m:name="Flag" m:value-type="string"><![CDATA[A&B]]></m:user-defined>
  </o:meta>
  <o:body><o:text><o:annotation><d:creator>Annotation Author</d:creator></o:annotation></o:text></o:body>
</o:document>"#;

        let metadata = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(metadata.title.as_deref(), Some("R&D 🍋"));
        assert_eq!(metadata.creator.as_deref(), Some("Last Editor"));
        assert_eq!(metadata.initial_creator.as_deref(), Some("Original Author"));
        assert_eq!(metadata.printed_by.as_deref(), Some("Print Operator"));
        assert_eq!(
            metadata.editing_cycles.as_deref(),
            Some("0000000000000000000000007")
        );
        assert_eq!(
            metadata.editing_duration.as_deref(),
            Some("PT1H2M3.0000000001S")
        );

        let template = metadata.template.as_ref().unwrap();
        assert_eq!(template.href.as_deref(), Some("Templates/A&B.ott"));
        assert_eq!(template.title.as_deref(), Some("A & B"));
        assert_eq!(template.actuate.as_deref(), Some("onRequest"));
        let auto_reload = metadata.auto_reload.as_ref().unwrap();
        assert_eq!(auto_reload.href.as_deref(), Some("next.fodt"));
        assert_eq!(auto_reload.delay.as_deref(), Some("PT5M"));
        let hyperlink = metadata.hyperlink_behaviour.as_ref().unwrap();
        assert_eq!(hyperlink.target_frame_name.as_deref(), Some("_blank"));
        assert_eq!(hyperlink.show.as_deref(), Some("new"));

        assert_eq!(
            metadata.statistics.page_count.as_deref(),
            Some("184467440737095516160")
        );
        assert_eq!(metadata.statistics.draw_count.as_deref(), Some("2"));
        assert_eq!(metadata.statistics.ole_object_count.as_deref(), Some("4"));
        assert_eq!(metadata.statistics.frame_count.as_deref(), Some("9"));
        assert_eq!(metadata.statistics.sentence_count.as_deref(), Some("10"));
        assert_eq!(metadata.statistics.syllable_count.as_deref(), Some("11"));
        assert_eq!(
            metadata
                .statistics
                .non_whitespace_character_count
                .as_deref(),
            Some("12")
        );
        assert_eq!(metadata.statistics.row_count.as_deref(), Some("13"));
        assert_eq!(metadata.statistics.cell_count.as_deref(), Some("14"));

        assert_eq!(metadata.user_defined.len(), 2);
        assert_eq!(
            metadata.user_defined[0].value_type,
            UserDefinedValueType::Boolean
        );
        assert_eq!(metadata.user_defined[0].value, "true");
        assert_eq!(
            metadata.user_defined[1].value_type,
            UserDefinedValueType::String
        );
        assert_eq!(metadata.user_defined[1].value, "A&B");
        assert_eq!(
            metadata.custom_properties.get("Flag").map(String::as_str),
            Some("A&B")
        );

        let common: Metadata = metadata.into();
        assert_eq!(common.author.as_deref(), Some("Original Author"));
        assert_eq!(common.last_modified_by.as_deref(), Some("Last Editor"));
        assert_eq!(common.template.as_deref(), Some("Templates/A&B.ott"));
        assert_eq!(
            common.revision.as_deref(),
            Some("0000000000000000000000007")
        );
        assert_eq!(common.page_count, None);
        assert!(common.last_printed_time.is_some());
    }

    #[test]
    fn rejects_invalid_statistic_and_nested_simple_metadata() {
        for xml in [
            r#"<o:document-meta xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:m="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><o:meta><m:document-statistic m:page-count="-1"/></o:meta></o:document-meta>"#,
            r#"<o:document-meta xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="http://purl.org/dc/elements/1.1/"><o:meta><d:title>bad<d:subject>nested</d:subject></d:title></o:meta></o:document-meta>"#,
            r#"<o:document-meta xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:m="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><o:meta><m:user-defined>missing name</m:user-defined></o:meta></o:document-meta>"#,
        ] {
            assert!(OdfMetadata::from_xml(xml).is_err(), "accepted {xml}");
        }
    }

    #[test]
    fn test_document_statistics_default() {
        let stats = DocumentStatistics::default();
        assert!(stats.page_count.is_none());
        assert!(stats.paragraph_count.is_none());
        assert!(stats.word_count.is_none());
        assert!(stats.character_count.is_none());
        assert!(stats.table_count.is_none());
        assert!(stats.image_count.is_none());
        assert!(stats.object_count.is_none());
    }

    #[test]
    fn test_parse_date_iso8601() {
        let date = OdfMetadata::parse_date(Some("2024-03-15T14:30:00Z".to_string()));
        assert!(date.is_some());
    }

    #[test]
    fn test_parse_date_rfc3339() {
        let date = OdfMetadata::parse_date(Some("2024-03-15T00:00:00+00:00".to_string()));
        assert!(date.is_some());
    }

    #[test]
    fn test_parse_date_none() {
        let date = OdfMetadata::parse_date(None);
        assert!(date.is_none());
    }

    #[test]
    fn test_parse_date_invalid() {
        let date = OdfMetadata::parse_date(Some("not-a-date".to_string()));
        assert!(date.is_none());
    }

    #[test]
    fn test_into_metadata_empty() {
        let odf = OdfMetadata::default();
        let meta: Metadata = odf.into();
        assert!(meta.title.is_none());
        assert!(meta.author.is_none());
        assert!(meta.keywords.is_none());
    }

    #[test]
    fn test_into_metadata_with_data() {
        let odf = OdfMetadata {
            title: Some("Title".to_string()),
            creator: Some("Author".to_string()),
            subject: Some("Subject".to_string()),
            keywords: vec!["a".to_string(), "b".to_string()],
            description: Some("Desc".to_string()),
            creation_date: Some("2024-01-01T00:00:00Z".to_string()),
            modification_date: Some("2024-06-01T00:00:00Z".to_string()),
            generator: Some("App".to_string()),
            statistics: DocumentStatistics {
                page_count: Some("5".to_string()),
                word_count: Some("100".to_string()),
                character_count: Some("500".to_string()),
                ..Default::default()
            },
            ..Default::default()
        };

        let meta: Metadata = odf.into();
        assert_eq!(meta.title, Some("Title".to_string()));
        assert_eq!(meta.author, Some("Author".to_string()));
        assert_eq!(meta.subject, Some("Subject".to_string()));
        assert_eq!(meta.keywords, Some("a, b".to_string()));
        assert_eq!(meta.description, Some("Desc".to_string()));
        assert_eq!(meta.page_count, Some(5));
        assert_eq!(meta.word_count, Some(100));
        assert_eq!(meta.character_count, Some(500));
        assert_eq!(meta.application, Some("App".to_string()));
        assert!(meta.created.is_some());
        assert!(meta.modified.is_some());
    }

    #[test]
    fn test_into_metadata_no_keywords() {
        let odf = OdfMetadata {
            keywords: vec![],
            ..Default::default()
        };

        let meta: Metadata = odf.into();
        assert!(meta.keywords.is_none());
    }
}
