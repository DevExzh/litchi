//! Source-bound policy checks for changed ordinary DOCX publication.
//!
//! The ordinary document transaction owns the main-document edit grammar.  It
//! must still inspect the optional settings part before a changed source is
//! published, because protection and tracked revisions are package policy,
//! rather than main-story syntax.  This module reads that part through its
//! exact source-preserving OPC view and never runs markup-compatibility
//! projection on it.

use litchi_core::{ExecutionContext, ExecutionError, Resource};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Error, Result};
use crate::settings::{STRICT_WORD_NAMESPACE, TRANSITIONAL_WORD_NAMESPACE};
use crate::variables::{self, SettingsDialect};

use super::Package;

const STRICT_SETTINGS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
const MAX_SETTINGS_SCAN_DEPTH: usize = 256;
const SETTINGS_SCAN_MEMORY_MULTIPLIER: u64 = 32;
const SETTINGS_SCAN_MEMORY_HEADROOM: u64 = 131_072;
const SETTINGS_SCAN_OBJECT_HEADROOM: u64 = 1_024;
const SETTINGS_TARGET_MEMORY_HEADROOM: u64 = 4_096;

/// Validate package policy which applies only to a real changed document
/// publication.
///
/// The optional settings relationship is resolved without materializing the
/// main document or any unrelated part.  The settings bytes remain borrowed
/// from the source-preserving OPC payload while the bounded XML guard and the
/// bounded policy scanner run.  Exact no-op callers deliberately do not invoke
/// this function.
pub(super) fn validate_changed_document(package: &Package) -> Result<()> {
    let version = package.package.source_version()?;
    let result = validate_changed_document_inner(package);
    finish_source_fence(package, version)?;
    result
}

fn validate_changed_document_inner(package: &Package) -> Result<()> {
    let context = package.package.execution_context();
    package.package.check_execution()?;
    let source_version = package.package.source_version()?;

    let main = package.package.main_document_part()?;
    let package_dialect = package_dialect(package)?;
    let mut settings_relationship = None;
    for relationship in main.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::SETTINGS | STRICT_SETTINGS_RELATIONSHIP
        )
    }) {
        if settings_relationship.replace(relationship).is_some() {
            return Err(Error::InvalidRelationship(
                "document has multiple settings relationships".into(),
            ));
        }
    }

    let Some(relationship) = settings_relationship else {
        finish_source_fence(package, source_version)?;
        return Ok(());
    };
    if relationship.is_external() {
        return Err(Error::InvalidRelationship(
            "settings relationship cannot be external".into(),
        ));
    }
    let settings_strict = relationship.reltype() == STRICT_SETTINGS_RELATIONSHIP;
    if package_dialect != Some(settings_strict) {
        return Err(Error::InvalidRelationship(
            "main document and settings relationship use mixed OOXML conformance families".into(),
        ));
    }

    // `target_partname` normalizes the relationship reference into an owned
    // PackURI.  Admit that small URI object before asking the OPC layer to
    // allocate it on a managed source.
    let _target_admission =
        admit_target_metadata(context.as_ref(), relationship.target_path().len())?;
    let target = relationship.target_partname()?;
    let settings_part = package.package.part(&target)?;
    if settings_part.content_type() != ct::WML_SETTINGS {
        return Err(Error::InvalidContentType {
            expected: ct::WML_SETTINGS.into(),
            got: settings_part.content_type().into(),
        });
    }

    // `source_xml` is the OPC source-preserving publication seam.  It also
    // applies the package-wide read-ahead transition before exact bytes are
    // fetched and lets the OPC owner handle signed-source policy.
    let source = settings_part.source_xml()?;
    let settings = inspect_settings_source(source.bytes(), settings_strict, context.as_ref())?;
    if settings.protected || settings.track_revisions {
        return Err(Error::UnsafeEdit {
            format: "DOCX",
            operation: "changed document publication",
            reason: "document or write protection, or tracked revisions, is enabled",
        });
    }

    finish_source_fence(package, source_version)
}

fn package_dialect(package: &Package) -> Result<Option<bool>> {
    let mut dialect = None;
    for relationship in package.package.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        )
    }) {
        if dialect
            .replace(relationship.reltype() == rt::STRICT_OFFICE_DOCUMENT)
            .is_some()
        {
            return Err(Error::InvalidRelationship(
                "package has multiple main-document relationships".into(),
            ));
        }
    }
    Ok(dialect)
}

fn finish_source_fence(package: &Package, expected: litchi_core::SourceVersion) -> Result<()> {
    let actual = package.package.source_version()?;
    if actual != expected {
        return Err(Error::Opc(litchi_opc::OpcError::SourceChanged {
            expected,
            actual,
        }));
    }
    package.package.check_execution()?;
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct SettingsPolicy {
    protected: bool,
    track_revisions: bool,
}

fn inspect_settings_source(
    xml: &[u8],
    expected_strict: bool,
    context: Option<&ExecutionContext>,
) -> Result<SettingsPolicy> {
    let _admission = admit_settings_workspace(xml, context)?;
    let facts = guard_settings_source(xml, context)?;

    if let Some(context) = context {
        // The source-policy pass clones the namespace resolver for each event
        // and can revisit the input-sized namespace buffer.  Charge its full
        // input-by-event envelope before it starts growing parser state.
        let policy_work = facts
            .event_count
            .checked_mul(facts.work_per_event)
            .ok_or_else(|| Error::InvalidFormat("settings validation work overflow".into()))?;
        context
            .consume(Resource::Work, policy_work)
            .map_err(map_execution_error)?;
    }

    // This existing source-policy inspection validates the root dialect and
    // the direct protection markers without creating a projected MCE buffer.
    let source_policy = variables::inspect_source_policy(xml)?;
    let source_strict = matches!(source_policy.dialect, SettingsDialect::Strict);
    if source_strict != expected_strict {
        return Err(Error::InvalidRelationship(
            "settings relationship and settings XML use mixed OOXML conformance families".into(),
        ));
    }

    let policy = SettingsPolicy {
        protected: source_policy.protected,
        track_revisions: facts.track_revisions,
    };
    if let Some(context) = context {
        context.check().map_err(map_execution_error)?;
    }
    Ok(policy)
}

fn admit_target_metadata(
    context: Option<&ExecutionContext>,
    target_bytes: usize,
) -> Result<Option<(litchi_core::Reservation, litchi_core::Reservation)>> {
    let Some(context) = context else {
        return Ok(None);
    };
    let target_bytes = u64::try_from(target_bytes)
        .map_err(|_| Error::InvalidFormat("settings relationship target is too large".into()))?;
    let memory = target_bytes
        .checked_add(SETTINGS_TARGET_MEMORY_HEADROOM)
        .ok_or_else(|| {
            Error::InvalidFormat("settings relationship target bound overflow".into())
        })?;
    let memory = context
        .reserve(Resource::Memory, memory)
        .map_err(map_execution_error)?;
    let objects = context
        .reserve(Resource::Objects, 2)
        .map_err(map_execution_error)?;
    Ok(Some((memory, objects)))
}

fn admit_settings_workspace(
    xml: &[u8],
    context: Option<&ExecutionContext>,
) -> Result<
    Option<(
        litchi_core::Reservation,
        litchi_core::Reservation,
        litchi_core::Reservation,
    )>,
> {
    let Some(context) = context else {
        return Ok(None);
    };
    let bytes = u64::try_from(xml.len())
        .map_err(|_| Error::InvalidFormat("settings XML byte count overflow".into()))?;
    let memory = bytes
        .checked_mul(SETTINGS_SCAN_MEMORY_MULTIPLIER)
        .and_then(|value| value.checked_add(SETTINGS_SCAN_MEMORY_HEADROOM))
        .ok_or_else(|| Error::InvalidFormat("settings XML memory bound overflow".into()))?;
    let objects = bytes
        .checked_add(SETTINGS_SCAN_OBJECT_HEADROOM)
        .ok_or_else(|| Error::InvalidFormat("settings XML object bound overflow".into()))?;
    let memory = context
        .reserve(Resource::Memory, memory)
        .map_err(map_execution_error)?;
    let objects = context
        .reserve(Resource::Objects, objects)
        .map_err(map_execution_error)?;
    let depth = context
        .reserve(Resource::Depth, MAX_SETTINGS_SCAN_DEPTH as u64)
        .map_err(map_execution_error)?;
    Ok(Some((memory, objects, depth)))
}

#[derive(Debug, Clone, Copy)]
struct SettingsGuardFacts {
    event_count: u64,
    work_per_event: u64,
    track_revisions: bool,
}

fn guard_settings_source(
    xml: &[u8],
    context: Option<&ExecutionContext>,
) -> Result<SettingsGuardFacts> {
    if xml.len() > crate::settings::MAX_SETTINGS_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "settings XML exceeds the {} byte source policy limit",
            crate::settings::MAX_SETTINGS_XML_BYTES
        )));
    }
    let work_per_event = u64::try_from(xml.len())
        .map_err(|_| Error::InvalidFormat("settings XML work bound overflow".into()))?
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("settings XML work bound overflow".into()))?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut events = 0u64;
    let mut saw_root = false;
    let mut root_closed = false;
    let mut saw_declaration = false;
    let mut track_revisions = None;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("settings XML event counter overflow".into()))?;
        if let Some(context) = context {
            context
                .consume(Resource::Work, work_per_event)
                .map_err(map_execution_error)?;
        }
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if saw_root || root_closed {
                        return Err(Error::InvalidFormat(
                            "settings XML has multiple roots".into(),
                        ));
                    }
                    saw_root = true;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("settings XML depth overflow".into()))?;
                if depth > MAX_SETTINGS_SCAN_DEPTH {
                    return Err(Error::InvalidFormat(
                        "settings XML exceeds the bounded source policy depth".into(),
                    ));
                }
                reject_settings_mce(
                    &resolver,
                    decoder,
                    &namespace,
                    &element,
                    context,
                    work_per_event,
                )?;
                inspect_track_revisions(
                    depth,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    &mut track_revisions,
                )?;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if saw_root || root_closed {
                        return Err(Error::InvalidFormat(
                            "settings XML has multiple roots".into(),
                        ));
                    }
                    saw_root = true;
                    root_closed = true;
                }
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("settings XML depth overflow".into()))?;
                if child_depth > MAX_SETTINGS_SCAN_DEPTH {
                    return Err(Error::InvalidFormat(
                        "settings XML exceeds the bounded source policy depth".into(),
                    ));
                }
                reject_settings_mce(
                    &resolver,
                    decoder,
                    &namespace,
                    &element,
                    context,
                    work_per_event,
                )?;
                inspect_track_revisions(
                    child_depth,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    &mut track_revisions,
                )?;
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("settings XML depth underflow".into()))?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Eof => {
                if !saw_root || depth != 0 {
                    return Err(Error::InvalidFormat(
                        "settings XML does not contain one complete root".into(),
                    ));
                }
                break;
            },
            Event::Text(text) => {
                if depth == 0 && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "settings XML has character data outside its root".into(),
                    ));
                }
            },
            Event::CData(_) if depth == 0 => {
                return Err(Error::InvalidFormat(
                    "settings XML has CDATA outside its root".into(),
                ));
            },
            Event::Comment(_) | Event::CData(_) => {},
            Event::Decl(_) => {
                if saw_declaration || saw_root || root_closed {
                    return Err(Error::InvalidFormat(
                        "settings XML declaration is not in its prolog".into(),
                    ));
                }
                saw_declaration = true;
            },
            Event::PI(_) => {
                return Err(Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "changed document publication",
                    reason: "processing instructions are not admitted in source-bound settings",
                });
            },
            Event::DocType(_) => {
                return Err(Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "changed document publication",
                    reason: "DTD declarations are not admitted in source-bound settings",
                });
            },
            Event::GeneralRef(reference) => {
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "settings XML has an entity reference outside its root".into(),
                    ));
                }
                if reference.is_char_ref() {
                    let value = reference
                        .resolve_char_ref()
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "settings XML character reference did not resolve".into(),
                            )
                        })?;
                    if !super::is_legal_xml_character(value) {
                        return Err(Error::InvalidFormat(
                            "settings XML character reference is not a legal XML character".into(),
                        ));
                    }
                } else {
                    let name = reference
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    if !matches!(name.as_ref(), "amp" | "apos" | "gt" | "lt" | "quot") {
                        return Err(Error::UnsafeEdit {
                            format: "DOCX",
                            operation: "changed document publication",
                            reason: "non-predefined entity references are not admitted in source-bound settings",
                        });
                    }
                }
            },
        }
    }
    Ok(SettingsGuardFacts {
        event_count: events,
        work_per_event,
        track_revisions: track_revisions.unwrap_or(false),
    })
}

fn inspect_track_revisions(
    depth: usize,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    track_revisions: &mut Option<bool>,
) -> Result<()> {
    if depth != 2
        || !is_wordprocessing_namespace(namespace)
        || element.local_name().as_ref() != b"trackRevisions"
    {
        return Ok(());
    }
    if track_revisions.is_some() {
        return Err(Error::InvalidFormat(
            "duplicate trackRevisions setting".into(),
        ));
    }
    let value = word_attribute_value(element, b"val", decoder, resolver)?;
    let enabled = match value.as_deref() {
        None => true,
        Some("true" | "1" | "on") => true,
        Some("false" | "0" | "off") => false,
        Some(value) => {
            return Err(Error::InvalidFormat(format!(
                "invalid Word on/off value '{value}'"
            )));
        },
    };
    *track_revisions = Some(enabled);
    Ok(())
}

fn is_wordprocessing_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == TRANSITIONAL_WORD_NAMESPACE || *value == STRICT_WORD_NAMESPACE
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
            return Err(Error::InvalidFormat(format!(
                "duplicate Word attribute '{}'",
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

fn reject_settings_mce(
    resolver: &NamespaceResolver,
    decoder: Decoder,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    context: Option<&ExecutionContext>,
    lookup_work: u64,
) -> Result<()> {
    if matches!(namespace, ResolveResult::Unknown(_)) {
        return Err(Error::InvalidFormat(
            "settings XML uses an undeclared element namespace prefix".into(),
        ));
    }
    if super::is_mce_namespace(namespace) {
        return Err(mce_refusal());
    }
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        super::validate_source_attribute_value(
            attribute.value.as_ref(),
            "changed document publication",
        )?;
        // NsReader retains lexical namespace values. An entity-escaped URI
        // could otherwise hide a Word policy marker or an MCE directive from
        // namespace comparisons while an Office consumer resolves that URI.
        // Refuse that unimplemented normalization case without changing bytes.
        if super::is_namespace_declaration(attribute.key)
            && attribute.value.as_ref().contains(&b'&')
        {
            return Err(mce_refusal());
        }
        let (attribute_namespace, _) = resolver.resolve_attribute(attribute.key);
        if !super::is_namespace_declaration(attribute.key)
            && matches!(attribute_namespace, ResolveResult::Unknown(_))
        {
            return Err(Error::InvalidFormat(
                "settings XML uses an undeclared attribute namespace prefix".into(),
            ));
        }
        if super::is_mce_namespace(&attribute_namespace) {
            // Ignoring a foreign settings namespace cannot introduce a direct
            // Word protection or revision-policy child. We inspect those direct
            // Word children in the exact source and preserve all settings bytes.
            // Other MCE directives, especially ProcessContent and alternate
            // branches, can change which policy children are visible and refuse.
            if attribute.key.local_name().as_ref() != b"Ignorable" {
                return Err(mce_refusal());
            }
            let prefixes = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?;
            for prefix in prefixes.split_ascii_whitespace() {
                if let Some(context) = context {
                    // One prefix resolution can visit the retained bindings;
                    // admit an input-sized lookup before each list item.
                    context
                        .consume(Resource::Work, lookup_work)
                        .map_err(map_execution_error)?;
                }
                if prefix.contains(':') {
                    return Err(mce_refusal());
                }
                let name = format!("{prefix}:_");
                let (namespace, _) = resolver.resolve_element(QName(name.as_bytes()));
                if !matches!(namespace, ResolveResult::Bound(_))
                    || is_wordprocessing_namespace(&namespace)
                    || super::is_mce_namespace(&namespace)
                {
                    return Err(mce_refusal());
                }
            }
        }
    }
    Ok(())
}

fn mce_refusal() -> Error {
    Error::UnsafeEdit {
        format: "DOCX",
        operation: "changed document publication",
        reason: "markup-compatibility projection or unknown source policy is not admitted",
    }
}

fn map_execution_error(error: ExecutionError) -> Error {
    match error {
        ExecutionError::Cancelled => Error::Opc(litchi_opc::OpcError::Cancelled),
        error => Error::Opc(litchi_opc::OpcError::Execution(error)),
    }
}
