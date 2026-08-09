//! Bounded ODF metadata XML parsing and retained-source patching.

use super::model::{
    AutoReloadMetadata, DocumentStatistics, HyperlinkBehaviourMetadata, Metadata, TemplateMetadata,
    UserDefinedMetadata, UserDefinedValueType,
};
use litchi_core::{Error, Metadata as CoreMetadata, Result, xml::escape_xml};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::writer::Writer;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const META_NAMESPACE_STR: &str = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
const META_NAMESPACE: &[u8] = META_NAMESPACE_STR.as_bytes();
const DC_NAMESPACE_STR: &str = "http://purl.org/dc/elements/1.1/";
const DC_NAMESPACE: &[u8] = DC_NAMESPACE_STR.as_bytes();
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
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

impl Metadata {
    /// Parse metadata from `meta.xml` content.
    ///
    /// # Errors
    ///
    /// Returns an error when the XML is malformed or contains invalid ODF
    /// metadata values.
    pub fn from_xml(xml_content: &str) -> Result<Self> {
        let mut reader = NsReader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut metadata = Metadata::default();
        let mut depth = 0usize;
        let mut meta_depth = None;

        loop {
            let (resolved_namespace, event) = reader
                .read_resolved_event_into(&mut buf)
                .map_err(metadata_xml_error)?;
            let namespace = known_namespace(&resolved_namespace);
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
                        if let Some(field) = text_field(namespace, local_name.as_ref()) {
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
                                namespace,
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
                        if let Some(field) = text_field(namespace, local_name.as_ref()) {
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
                                namespace,
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
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {},
            }
            buf.clear();
        }

        Ok(metadata)
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

/// Field-level edits applied to a retained source meta.xml during a mutable
/// save. Every element not named by an edit is preserved unchanged.
#[derive(Debug, Clone)]
pub struct MetaXmlPatch {
    generator: MetaFieldEdit,
    modification_date: MetaFieldEdit,
    title: MetaFieldEdit,
    creator: MetaFieldEdit,
    subject: MetaFieldEdit,
    description: MetaFieldEdit,
    keywords: MetaFieldEdit,
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

/// Prefixes the source meta.xml root binds to the Dublin Core and ODF meta
/// namespaces, reused when a patch inserts new elements.
#[derive(Debug, Default)]
struct MetaNamespacePrefixes {
    dc: Option<String>,
    meta: Option<String>,
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

impl MetaXmlPatch {
    /// Start from a patch that preserves every source element.
    #[must_use]
    pub fn preserve_all() -> Self {
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
    #[must_use]
    pub fn with_generator_and_modification_date(
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
    #[must_use]
    pub fn diff_simple_fields(mut self, source: &Metadata, current: &CoreMetadata) -> Self {
        let expected = CoreMetadata::from(source.clone());
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

impl MetaNamespacePrefixes {
    /// Record the relevant bindings declared on the root element.
    fn observe(&mut self, element: &BytesStart<'_>, reader: &NsReader<&[u8]>) -> Result<()> {
        for raw_attribute in element.attributes() {
            let attribute = raw_attribute.map_err(metadata_xml_error)?;
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
            KnownNamespace::Office | KnownNamespace::Other => {
                return (local_name.to_string(), None);
            },
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

fn text_field(namespace: KnownNamespace, local_name: &[u8]) -> Option<TextField> {
    if namespace == KnownNamespace::Dc {
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
    if namespace == KnownNamespace::Meta {
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

fn assign_text_field(metadata: &mut Metadata, field: TextField, value: String) {
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

/// Map a direct `office:meta` child to the field a patch manages.
fn patch_field(namespace: KnownNamespace, local_name: &[u8]) -> Option<PatchField> {
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
        KnownNamespace::Office | KnownNamespace::Other => None,
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
///
/// # Errors
///
/// Returns an error when the source XML cannot be read or the rewritten XML
/// cannot be emitted as UTF-8.
pub fn patch_meta_xml(source: &str, patch: &MetaXmlPatch) -> Result<Option<String>> {
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
        let (resolved_namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(metadata_xml_error)?;
        let namespace = known_namespace(&resolved_namespace);
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
                    && let Some(field) = patch_field(namespace, local_name.as_ref())
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
                    && let Some(field) = patch_field(namespace, local_name.as_ref())
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
            other_event @ (Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_)) => write_event(&mut writer, other_event)?,
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
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
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
            Event::Start(_) | Event::Empty(_) | Event::Decl(_) | Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "nested markup is not allowed in OpenDocument metadata values".to_string(),
                ));
            },
        }
        buffer.clear();
    }
}

fn parse_metadata_attributes(
    metadata: &mut Metadata,
    namespace: KnownNamespace,
    local_name: &[u8],
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
) -> Result<()> {
    if namespace != KnownNamespace::Meta {
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
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute.map_err(metadata_xml_error)?;
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
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute.map_err(metadata_xml_error)?;
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
