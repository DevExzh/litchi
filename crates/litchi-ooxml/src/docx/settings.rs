//! Document settings and protection support.
//!
//! This module provides types and methods for accessing document settings
//! and protection status.

use crate::docx::mail_merge::{
    MailMergeSettings, parse_settings_mail_merge, validate_mail_merge_relationships,
};
use crate::docx::namespace::{is_wordprocessing_namespace, word_attribute_value};
use crate::docx::variables::DocumentVariables;
use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::XmlVersion;
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
        let xml = crate::common::mce::process_part(part)?;
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
        let mut saw_do_not_embed_smart_tags = false;
        let mut saw_attached_template = false;

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
                    } else if depth == 2 && saw_root && is_wordprocessing_namespace(&namespace) {
                        parse_setting(
                            &element,
                            decoder,
                            &resolver,
                            &mut settings,
                            &mut saw_do_not_embed_smart_tags,
                            &mut saw_attached_template,
                        )?;
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("Word settings XML nesting is too deep".into())
                    })?;
                    if child_depth == 1 {
                        validate_settings_root(&namespace, &element, saw_root)?;
                        saw_root = true;
                    } else if child_depth == 2
                        && saw_root
                        && is_wordprocessing_namespace(&namespace)
                    {
                        parse_setting(
                            &element,
                            decoder,
                            &resolver,
                            &mut settings,
                            &mut saw_do_not_embed_smart_tags,
                            &mut saw_attached_template,
                        )?;
                    }
                },
                Event::End(_) => {
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

fn parse_setting(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    settings: &mut DocumentSettings,
    saw_do_not_embed_smart_tags: &mut bool,
    saw_attached_template: &mut bool,
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
            if std::mem::replace(saw_do_not_embed_smart_tags, true) {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate doNotEmbedSmartTags setting".into(),
                ));
            }
            settings.do_not_embed_smart_tags = parse_on_off(element, decoder, resolver)?;
        },
        b"attachedTemplate" => {
            if std::mem::replace(saw_attached_template, true) {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate attachedTemplate setting".into(),
                ));
            }
            let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?
                .ok_or_else(|| {
                    OoxmlError::InvalidFormat(
                        "attachedTemplate relationship ID is required".into(),
                    )
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
        let slash = root.windows(2).rposition(|window| window == b"/>").ok_or_else(|| {
            OoxmlError::InvalidFormat("invalid empty settings root".into())
        })?;
        let mut output = Vec::with_capacity(xml.len() + replacement.len() + layout.root_qname.len() + 4);
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
    let insert_at = layout.mail_merge_insert_at.or(layout.root_end).ok_or_else(|| {
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
        let mut output = Vec::with_capacity(
            xml.len() + replacement.len() + layout.root_qname.len() + 4,
        );
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
        let mut output = Vec::with_capacity(xml.len() + replacement.len() + layout.root_qname.len() + 4);
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
                if depth == 2 && let Some(start) = attached_start.take() {
                    attached_template_range = Some(start..event_end);
                }
                if depth == 2 && let Some(start) = doc_vars_start.take() {
                    doc_vars_range = Some(start..event_end);
                }
                if depth == 2 && let Some(start) = mail_merge_start.take() {
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

fn document_variables_element(
    layout: &SettingsXmlLayout,
    variables: &DocumentVariables,
) -> String {
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

    fn attached_template_part(
        xml: &[u8],
        reltype: &str,
        target: &str,
        external: bool,
    ) -> BlobPart {
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
        let internal = attached_template_part(
            xml,
            ATTACHED_TEMPLATE_RELATIONSHIP,
            "template.dotx",
            false,
        );
        assert!(DocumentSettings::extract_from_part(&internal).is_err());
        let whitespace = attached_template_part(
            xml,
            ATTACHED_TEMPLATE_RELATIONSHIP,
            "file:///bad path.dotx",
            true,
        );
        assert!(DocumentSettings::extract_from_part(&whitespace).is_err());

        let mut duplicate = attached_template_part(
            xml,
            ATTACHED_TEMPLATE_RELATIONSHIP,
            "file:///a.dotx",
            true,
        );
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
        let inserted = String::from_utf8(patch_attached_template(empty, Some("rId7")).unwrap()).unwrap();
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
        let patched = String::from_utf8(patch_document_variables(xml, &variables).unwrap()).unwrap();
        assert_eq!(
            patched,
            r#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><!--before--><q:compat/><q:docVars><q:docVar q:name="Company &amp; Team" q:val="A &lt; B"/><q:docVar q:name="empty" q:val=""/></q:docVars><x:opaque><![CDATA[a < b]]></x:opaque><q:rsids/></q:settings>"#
        );

        variables.clear();
        let removed = String::from_utf8(
            patch_document_variables(patched.as_bytes(), &variables).unwrap(),
        )
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
