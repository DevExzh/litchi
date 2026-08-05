//! Document settings and protection support.
//!
//! This module provides types and methods for accessing document settings
//! and protection status.

use super::model::Settings;
use super::notes::{NoteNumberingProperties, NoteNumberingRestart, NotePosition};
use crate::Variables;
use crate::error::{Error, Result};
use crate::mail_merge::{
    Settings as MailMergeSettings, parse_settings_mail_merge, validate_mail_merge_relationships,
};
use crate::namespace::{
    STRICT_WORDPROCESSINGML_NAMESPACE, is_wordprocessing_namespace, word_attribute_value,
};
use crate::numbering::Format;
use litchi_opc::part::Part;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::ops::Range;

const TRANSITIONAL_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";

/// Relationship type required by Word for an attached document template.
pub(crate) const ATTACHED_TEMPLATE_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate";
/// Strict OOXML alias accepted when reading existing packages.
pub(crate) const STRICT_ATTACHED_TEMPLATE_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/attachedTemplate";
const MAX_ATTACHED_TEMPLATE_TARGET_LEN: usize = 32 * 1024;
const MAX_SETTINGS_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_SETTINGS_XML_NODES: usize = 250_000;
const MAX_SETTINGS_XML_DEPTH: usize = 256;

/// Decode the canonical DOCX document-variable collection from a settings part.
///
/// Markup-compatibility preprocessing and OPC ownership remain host concerns;
/// the validated model and XML codec belong to `litchi-docx`.
pub(crate) fn extract_document_variables(part: &dyn Part) -> Result<Variables> {
    let limit = crate::variables::MAX_DOCUMENT_VARIABLE_XML_BYTES;
    if part.blob().len() > limit {
        return Err(Error::InvalidFormat(format!(
            "settings XML exceeds the {limit} byte document-variable limit"
        )));
    }
    let xml = litchi_ooxml_common::mce::process_part(part)?;
    Ok(crate::parse_variables(xml.as_ref())?)
}

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
/// use litchi_docx::Package;
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
    /// Format-owned scalar settings parsed from `settings.xml`.
    values: OwnedSettings,
    /// Smart-tag type declarations.
    smart_tag_types: Vec<SmartTagType>,
    /// Whether applications should omit embedded smart-tag data when saving.
    do_not_embed_smart_tags: bool,
    /// Inert mail-merge connection and display metadata.
    mail_merge: Option<MailMergeSettings>,
    /// Inert external attached-template reference.
    attached_template: Option<AttachedTemplate>,
}

use crate::settings::{
    ColorSchemeIndex, ColorSchemeMapping, ColorSchemeSlot, CompatFlag, CompatibilityOption,
    CompatibilitySetting, ProofState, ProofingState, ProtectionType, SmartTagType,
    ThemeFontLanguages, View,
};

type OwnedSettings = Settings<Format>;
impl DocumentSettings {
    /// Create a new Settings value with default values.
    pub fn new() -> Self {
        Self {
            values: OwnedSettings::new(),
            smart_tag_types: Vec::new(),
            do_not_embed_smart_tags: false,
            mail_merge: None,
            attached_template: None,
        }
    }

    /// Check if the document is protected.
    #[inline]
    pub fn is_protected(&self) -> bool {
        self.values.is_protected()
    }

    /// Get the type of protection applied.
    #[inline]
    pub fn protection_type(&self) -> Option<ProtectionType> {
        self.values.protection_type()
    }

    /// Check if track revisions is enabled.
    #[inline]
    pub fn track_revisions(&self) -> bool {
        self.values.track_revisions()
    }

    /// Get the zoom percentage.
    #[inline]
    pub fn zoom_percent(&self) -> Option<u32> {
        self.values.zoom_percent()
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
        self.values.compatibility_options()
    }

    /// Return the `w:compatSetting` triples in document order.
    #[inline]
    pub fn compatibility_settings(&self) -> &[CompatibilitySetting] {
        self.values.compatibility_settings()
    }

    /// Look up a `w:compatSetting` triple by name and URI.
    pub fn compatibility_setting(&self, name: &str, uri: &str) -> Option<&CompatibilitySetting> {
        self.values.compatibility_setting(name, uri)
    }

    /// Return the Word compatibility mode (`compatibilityMode` value), when
    /// declared — for example `15` targets Word 2013 behavior.
    pub fn compatibility_mode(&self) -> Option<u32> {
        self.values.compatibility_mode()
    }

    /// Return the document-level footnote properties, if present.
    #[inline]
    pub fn footnote_properties(&self) -> Option<&NoteNumberingProperties<Format>> {
        self.values.footnote_properties()
    }

    /// Return the document-level endnote properties, if present.
    #[inline]
    pub fn endnote_properties(&self) -> Option<&NoteNumberingProperties<Format>> {
        self.values.endnote_properties()
    }

    /// Whether applications should recommend write protection for the
    /// document (`w:writeProtection`).
    #[inline]
    pub fn is_write_protected(&self) -> bool {
        self.values.is_write_protected()
    }

    /// Return the document view mode (`w:view`), when specified.
    #[inline]
    pub fn view(&self) -> Option<View> {
        self.values.view()
    }

    /// Return the proofing completion markers (`w:proofState`), if present.
    #[inline]
    pub fn proofing_state(&self) -> Option<&ProofingState> {
        self.values.proofing_state()
    }

    /// Return the default tab stop interval in twips (`w:defaultTabStop`),
    /// when specified.
    #[inline]
    pub fn default_tab_stop_twips(&self) -> Option<u32> {
        self.values.default_tab_stop_twips()
    }

    /// Return the theme font language defaults (`w:themeFontLang`), if
    /// present.
    #[inline]
    pub fn theme_font_languages(&self) -> Option<&ThemeFontLanguages> {
        self.values.theme_font_languages()
    }

    /// Return the theme color slot remapping (`w:clrSchemeMapping`), if
    /// present.
    #[inline]
    pub fn color_scheme_mapping(&self) -> Option<&ColorSchemeMapping> {
        self.values.color_scheme_mapping()
    }

    /// Serialize the editing view, proofing, and theme default elements
    /// (`w:writeProtection`, `w:view`, `w:proofState`, `w:defaultTabStop`,
    /// `w:themeFontLang`, `w:clrSchemeMapping`) in ECMA-376 schema order.
    pub fn to_editing_settings_xml(&self, prefix: &str) -> String {
        self.values.to_editing_settings_xml(prefix)
    }

    /// Extract settings from a settings.xml part.
    ///
    /// # Arguments
    ///
    /// * `part` - The settings part
    ///
    /// # Returns
    ///
    /// A Settings object
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
        let mut strict_wordprocessingml = false;
        let mut seen = SeenSettings::default();
        let mut saw_compat = false;
        let mut pending_group: Option<PendingGroup> = None;

        loop {
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
                        Error::InvalidFormat("Word settings XML nesting is too deep".into())
                    })?;
                    if depth == 1 {
                        validate_settings_root(&namespace, &element, saw_root)?;
                        strict_wordprocessingml = matches!(
                            namespace,
                            ResolveResult::Bound(Namespace(uri))
                                if uri == STRICT_WORDPROCESSINGML_NAMESPACE
                        );
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
                            parse_group_child(
                                group,
                                strict_wordprocessingml,
                                &element,
                                decoder,
                                &resolver,
                            )?;
                        }
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word settings XML nesting is too deep".into())
                    })?;
                    if child_depth == 1 {
                        validate_settings_root(&namespace, &element, saw_root)?;
                        strict_wordprocessingml = matches!(
                            namespace,
                            ResolveResult::Bound(Namespace(uri))
                                if uri == STRICT_WORDPROCESSINGML_NAMESPACE
                        );
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
                            parse_group_child(
                                group,
                                strict_wordprocessingml,
                                &element,
                                decoder,
                                &resolver,
                            )?;
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
                        Error::InvalidFormat("invalid Word settings XML nesting".into())
                    })?;
                },
                Event::Eof if depth != 0 => {
                    return Err(Error::InvalidFormat(
                        "unterminated Word settings XML".into(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        if !saw_root {
            return Err(Error::InvalidFormat(
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
        return Err(Error::InvalidFormat(
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
    FootnoteProperties(NoteNumberingProperties<Format>),
    EndnoteProperties(NoteNumberingProperties<Format>),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NoteKind {
    Footnote,
    Endnote,
}

fn begin_settings_group(
    element: &BytesStart<'_>,
    settings: &DocumentSettings,
    saw_compat: &mut bool,
) -> Result<Option<PendingGroup>> {
    match element.local_name().as_ref() {
        b"compat" => {
            if std::mem::replace(saw_compat, true) {
                return Err(Error::InvalidFormat(
                    "duplicate compat settings group".into(),
                ));
            }
            Ok(Some(PendingGroup::Compatibility {
                options: Vec::new(),
                settings: Vec::new(),
            }))
        },
        b"footnotePr" => {
            if settings.values.footnote_properties().is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate footnotePr settings group".into(),
                ));
            }
            Ok(Some(PendingGroup::FootnoteProperties(
                NoteNumberingProperties::<Format>::default(),
            )))
        },
        b"endnotePr" => {
            if settings.values.endnote_properties().is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate endnotePr settings group".into(),
                ));
            }
            Ok(Some(PendingGroup::EndnoteProperties(
                NoteNumberingProperties::<Format>::default(),
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
            settings.values.set_compatibility_options(options);
            settings.values.set_compatibility_settings(triples);
        },
        PendingGroup::FootnoteProperties(properties) => {
            if settings.values.footnote_properties().is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate footnotePr settings group".into(),
                ));
            }
            settings.values.set_footnote_properties(Some(properties));
        },
        PendingGroup::EndnoteProperties(properties) => {
            if settings.values.endnote_properties().is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate endnotePr settings group".into(),
                ));
            }
            settings.values.set_endnote_properties(Some(properties));
        },
    }
    Ok(())
}

fn parse_group_child(
    group: &mut PendingGroup,
    strict_wordprocessingml: bool,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    match group {
        PendingGroup::Compatibility { options, settings } => {
            if element.local_name().as_ref() == b"compatSetting" {
                settings.push(CompatibilitySetting::new(
                    required_attribute(element, b"name", decoder, resolver, "compatSetting name")?,
                    required_attribute(element, b"uri", decoder, resolver, "compatSetting URI")?,
                    required_attribute(element, b"val", decoder, resolver, "compatSetting value")?,
                ));
            } else {
                let local_name = element.local_name();
                let raw = std::str::from_utf8(local_name.as_ref()).map_err(|_| {
                    Error::InvalidFormat("compatibility flag name is not valid UTF-8".into())
                })?;
                let flag = raw.parse::<CompatFlag>().map_err(|_| {
                    Error::InvalidFormat(format!("invalid compatibility flag '{raw}'"))
                })?;
                if strict_wordprocessingml && !flag.is_strict() {
                    return Err(Error::InvalidFormat(format!(
                        "compatibility flag '{raw}' is not valid in Strict WordprocessingML"
                    )));
                }
                if options.iter().any(|option| option.flag() == flag) {
                    return Err(Error::InvalidFormat(format!(
                        "duplicate compatibility flag '{raw}'"
                    )));
                }
                options.push(CompatibilityOption::new(
                    flag,
                    parse_on_off(element, decoder, resolver)?,
                ));
            }
        },
        PendingGroup::FootnoteProperties(properties) => {
            parse_note_property_child(properties, NoteKind::Footnote, element, decoder, resolver)?;
        },
        PendingGroup::EndnoteProperties(properties) => {
            parse_note_property_child(properties, NoteKind::Endnote, element, decoder, resolver)?;
        },
    }
    Ok(())
}

fn parse_note_property_child(
    properties: &mut NoteNumberingProperties<Format>,
    kind: NoteKind,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    let mut position = properties.position();
    let mut format = properties.format();
    let mut start = properties.start();
    let mut restart = properties.restart();
    match element.local_name().as_ref() {
        b"pos" => {
            if position.is_some() {
                return Err(Error::InvalidFormat("duplicate note position".into()));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note position")?;
            let parsed_position = value
                .parse::<NotePosition>()
                .map_err(|_| Error::InvalidFormat(format!("invalid note position '{value}'")))?;
            if kind == NoteKind::Endnote && !parsed_position.valid_for_endnote() {
                return Err(Error::InvalidFormat(format!(
                    "position '{}' is not valid for an endnote",
                    parsed_position.as_str()
                )));
            }
            position = Some(parsed_position);
        },
        b"numFmt" => {
            if format.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate note numbering format".into(),
                ));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note numFmt")?;
            format = Some(value.parse().map_err(|_| {
                Error::InvalidFormat(format!("invalid note numbering format '{value}'"))
            })?);
        },
        b"numStart" => {
            if start.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate note numbering start".into(),
                ));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note numStart")?;
            start = Some(value.parse().map_err(|_| {
                Error::InvalidFormat(format!("invalid note numbering start '{value}'"))
            })?);
        },
        b"numRestart" => {
            if restart.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate note numbering restart".into(),
                ));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note numRestart")?;
            restart = Some(
                NoteNumberingRestart::from_xml(&value)
                    .map_err(|error| Error::InvalidFormat(error.to_string()))?,
            );
        },
        // `w:footnote`/`w:endnote` separator references carry no properties.
        _ => {},
    }
    *properties = NoteNumberingProperties::<Format>::from_parts(position, format, start, restart);
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
            settings.values.set_protected(true);
            if let Some(value) = word_attribute_value(element, b"edit", decoder, resolver)? {
                settings
                    .values
                    .set_protection_type(ProtectionType::from_xml(&value));
            }
            if let Some(value) = word_attribute_value(element, b"enforcement", decoder, resolver)? {
                settings.values.set_protected(parse_on_off_value(&value)?);
            }
        },
        b"trackRevisions" => {
            settings
                .values
                .set_track_revisions(parse_on_off(element, decoder, resolver)?);
        },
        b"zoom" => {
            if let Some(value) = word_attribute_value(element, b"percent", decoder, resolver)? {
                settings
                    .values
                    .set_zoom_percent(atoi_simd::parse::<u32, false, false>(value.as_bytes()).ok());
            }
        },
        b"smartTagType" => {
            let namespace_uri = required_attribute(
                element,
                b"namespaceuri",
                decoder,
                resolver,
                "smart-tag namespace URI",
            )?;
            let name =
                required_attribute(element, b"name", decoder, resolver, "smart-tag type name")?;
            let url = required_attribute(
                element,
                b"url",
                decoder,
                resolver,
                "smart-tag vocabulary URL",
            )?;
            settings
                .smart_tag_types
                .push(SmartTagType::new(namespace_uri, name, url).map_err(map_docx_error)?);
        },
        b"doNotEmbedSmartTags" => {
            if std::mem::replace(&mut seen.do_not_embed_smart_tags, true) {
                return Err(Error::InvalidFormat(
                    "duplicate doNotEmbedSmartTags setting".into(),
                ));
            }
            settings.do_not_embed_smart_tags = parse_on_off(element, decoder, resolver)?;
        },
        b"attachedTemplate" => {
            if std::mem::replace(&mut seen.attached_template, true) {
                return Err(Error::InvalidFormat(
                    "duplicate attachedTemplate setting".into(),
                ));
            }
            let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?
                .ok_or_else(|| {
                    Error::InvalidFormat("attachedTemplate relationship ID is required".into())
                })?;
            if relationship_id.is_empty() {
                return Err(Error::InvalidFormat(
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
                return Err(Error::InvalidFormat(
                    "duplicate writeProtection setting".into(),
                ));
            }
            settings
                .values
                .set_write_protected(parse_on_off(element, decoder, resolver)?);
        },
        b"view" => {
            if settings.values.view().is_some() {
                return Err(Error::InvalidFormat("duplicate view setting".into()));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "view mode")?;
            settings.values.set_view(Some(
                View::from_xml(&value).map_err(|error| Error::InvalidFormat(error.to_string()))?,
            ));
        },
        b"proofState" => {
            if settings.values.proofing_state().is_some() {
                return Err(Error::InvalidFormat("duplicate proofState setting".into()));
            }
            let mut state = ProofingState::new();
            if let Some(value) = word_attribute_value(element, b"spelling", decoder, resolver)? {
                state.set_spelling(Some(
                    ProofState::from_xml(&value)
                        .map_err(|error| Error::InvalidFormat(error.to_string()))?,
                ));
            }
            if let Some(value) = word_attribute_value(element, b"grammar", decoder, resolver)? {
                state.set_grammar(Some(
                    ProofState::from_xml(&value)
                        .map_err(|error| Error::InvalidFormat(error.to_string()))?,
                ));
            }
            settings.values.set_proofing_state(Some(state));
        },
        b"defaultTabStop" => {
            if settings.values.default_tab_stop_twips().is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate defaultTabStop setting".into(),
                ));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "default tab stop")?;
            settings
                .values
                .set_default_tab_stop_twips(Some(value.parse().map_err(|_| {
                    Error::InvalidFormat(format!("invalid default tab stop '{value}'"))
                })?));
        },
        b"themeFontLang" => {
            if settings.values.theme_font_languages().is_some() {
                return Err(Error::InvalidFormat(
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
            settings.values.set_theme_font_languages(Some(languages));
        },
        b"clrSchemeMapping" => {
            if settings.values.color_scheme_mapping().is_some() {
                return Err(Error::InvalidFormat(
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
                    mapping.set(
                        slot,
                        ColorSchemeIndex::from_xml(&value)
                            .map_err(|error| Error::InvalidFormat(error.to_string()))?,
                    );
                }
            }
            settings.values.set_color_scheme_mapping(Some(mapping));
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
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_relationship_attribute = matches!(
            namespace,
            ResolveResult::Bound(Namespace(uri))
                if uri == TRANSITIONAL_RELATIONSHIPS_NAMESPACE
                    || uri == STRICT_RELATIONSHIPS_NAMESPACE
        );
        if !is_relationship_attribute {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(
                "duplicate attachedTemplate relationship ID attribute".into(),
            ));
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

pub(crate) fn is_attached_template_relationship(value: &str) -> bool {
    matches!(
        value,
        ATTACHED_TEMPLATE_RELATIONSHIP | STRICT_ATTACHED_TEMPLATE_RELATIONSHIP
    )
}

pub(crate) fn validate_attached_template_target(target: &str) -> Result<()> {
    if target.is_empty() || target.len() > MAX_ATTACHED_TEMPLATE_TARGET_LEN {
        return Err(Error::InvalidFormat(
            "attached-template target must contain 1 to 32768 bytes".into(),
        ));
    }
    if target
        .chars()
        .any(|character| character.is_control() || character.is_whitespace())
    {
        return Err(Error::InvalidFormat(
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
        return Err(Error::InvalidFormat(
            "settings part has multiple attached-template relationships".into(),
        ));
    }

    let Some(attached_template) = settings.attached_template.as_mut() else {
        if matching.is_empty() {
            return Ok(());
        }
        return Err(Error::InvalidFormat(
            "settings part has an attached-template relationship without an attachedTemplate element"
                .into(),
        ));
    };
    let relationship = part
        .rels()
        .get(&attached_template.relationship_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "attachedTemplate references missing relationship {:?}",
                attached_template.relationship_id
            ))
        })?;
    if !is_attached_template_relationship(relationship.reltype()) {
        return Err(Error::InvalidFormat(
            "attachedTemplate relationship has the wrong type".into(),
        ));
    }
    if !relationship.is_external() {
        return Err(Error::InvalidFormat(
            "attachedTemplate relationship must use external target mode".into(),
        ));
    }
    validate_attached_template_target(relationship.target_ref())?;
    attached_template.target_uri = relationship.target_ref().to_owned();
    Ok(())
}

struct SettingsXmlLayout {
    #[cfg(any(feature = "fonts", test))]
    embed_true_type_fonts_range: Option<Range<usize>>,
    #[cfg(any(feature = "fonts", test))]
    embed_true_type_fonts_enabled: Option<bool>,
    #[cfg(any(feature = "fonts", test))]
    embed_true_type_fonts_insert_at: Option<usize>,
    #[cfg(any(feature = "fonts", test))]
    save_subset_fonts_range: Option<Range<usize>>,
    #[cfg(any(feature = "fonts", test))]
    save_subset_fonts_enabled: Option<bool>,
    #[cfg(any(feature = "fonts", test))]
    save_subset_fonts_insert_at: Option<usize>,
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

#[cfg(any(feature = "fonts", test))]
#[derive(Clone, Copy)]
enum FontFlag {
    EmbedTrueType,
    SaveSubset,
}

#[cfg(any(feature = "fonts", test))]
impl FontFlag {
    const fn local_name(self) -> &'static str {
        match self {
            Self::EmbedTrueType => "embedTrueTypeFonts",
            Self::SaveSubset => "saveSubsetFonts",
        }
    }
}

/// Losslessly enable Word font embedding and synchronize subset intent.
#[cfg(any(feature = "fonts", test))]
pub(crate) fn patch_font_embedding(xml: &[u8], subsetted: bool) -> Result<Vec<u8>> {
    let xml = patch_font_flag(xml, FontFlag::EmbedTrueType, true)?;
    patch_font_flag(&xml, FontFlag::SaveSubset, subsetted)
}

#[cfg(any(feature = "fonts", test))]
fn patch_font_flag(xml: &[u8], flag: FontFlag, enabled: bool) -> Result<Vec<u8>> {
    DocumentSettings::extract_from_xml(xml)?;
    let layout = scan_settings_xml_layout(xml)?;
    let (range, current, insert_at) = match flag {
        FontFlag::EmbedTrueType => (
            layout.embed_true_type_fonts_range.clone(),
            layout.embed_true_type_fonts_enabled,
            layout.embed_true_type_fonts_insert_at,
        ),
        FontFlag::SaveSubset => (
            layout.save_subset_fonts_range.clone(),
            layout.save_subset_fonts_enabled,
            layout.save_subset_fonts_insert_at,
        ),
    };
    if current == Some(enabled) || (!enabled && range.is_none()) {
        return Ok(xml.to_vec());
    }
    let replacement = if enabled {
        word_empty_element(&layout, flag)
    } else {
        String::new()
    };
    if let Some(range) = range {
        let capacity = xml
            .len()
            .checked_sub(range.len())
            .and_then(|size| size.checked_add(replacement.len()))
            .ok_or_else(|| Error::InvalidFormat("settings patch size overflow".into()))?;
        let mut output = settings_patch_buffer(capacity)?;
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
            .ok_or_else(|| Error::InvalidFormat("invalid empty settings root".into()))?;
        let capacity = xml
            .len()
            .checked_add(replacement.len())
            .and_then(|size| size.checked_add(layout.root_qname.len()))
            .and_then(|size| size.checked_add(4))
            .ok_or_else(|| Error::InvalidFormat("settings patch size overflow".into()))?;
        let mut output = settings_patch_buffer(capacity)?;
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
    let insert_at = insert_at
        .or(layout.root_end)
        .ok_or_else(|| Error::InvalidFormat("settings root has no insertion point".into()))?;
    let capacity = xml
        .len()
        .checked_add(replacement.len())
        .ok_or_else(|| Error::InvalidFormat("settings patch size overflow".into()))?;
    let mut output = settings_patch_buffer(capacity)?;
    output.extend_from_slice(&xml[..insert_at]);
    output.extend_from_slice(replacement.as_bytes());
    output.extend_from_slice(&xml[insert_at..]);
    Ok(output)
}

#[cfg(any(feature = "fonts", test))]
fn word_empty_element(layout: &SettingsXmlLayout, flag: FontFlag) -> String {
    let local_name = flag.local_name();
    match &layout.word_prefix {
        Some(prefix) => format!("<{}:{local_name}/>", String::from_utf8_lossy(prefix)),
        None => format!("<{local_name}/>"),
    }
}

#[cfg(any(feature = "fonts", test))]
fn settings_patch_buffer(capacity: usize) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "Word settings XML patch",
            source,
        })?;
    Ok(output)
}

pub(crate) fn patch_mail_merge(
    xml: &[u8],
    mail_merge: Option<&MailMergeSettings>,
    conformance: crate::mail_merge::Conformance,
) -> Result<Vec<u8>> {
    DocumentSettings::extract_from_xml(xml)?;
    let layout = scan_settings_xml_layout(xml)?;
    let replacement = mail_merge
        .map(|value| {
            value
                .to_xml(conformance)
                .map_err(crate::mail_merge::map_docx_error)
        })
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
            .ok_or_else(|| Error::InvalidFormat("invalid empty settings root".into()))?;
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
            Error::InvalidFormat("settings root has no mailMerge insertion point".into())
        })?;
    let mut output = Vec::with_capacity(xml.len() + replacement.len());
    output.extend_from_slice(&xml[..insert_at]);
    output.extend_from_slice(replacement.as_bytes());
    output.extend_from_slice(&xml[insert_at..]);
    Ok(output)
}

pub(crate) fn patch_document_variables(xml: &[u8], variables: &Variables) -> Result<Vec<u8>> {
    variables.validate()?;
    Variables::from_xml(xml)?;
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
            .ok_or_else(|| Error::InvalidFormat("invalid empty settings root".into()))?;
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
        .ok_or_else(|| Error::InvalidFormat("settings root has no insertion point".into()))?;
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
            .ok_or_else(|| Error::InvalidFormat("invalid empty settings root".into()))?;
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
        .ok_or_else(|| Error::InvalidFormat("settings root has no closing element".into()))?;
    let mut output = Vec::with_capacity(xml.len() + replacement.len());
    output.extend_from_slice(&xml[..insert_at]);
    output.extend_from_slice(replacement.as_bytes());
    output.extend_from_slice(&xml[insert_at..]);
    Ok(output)
}

fn scan_settings_xml_layout(xml: &[u8]) -> Result<SettingsXmlLayout> {
    if xml.len() > MAX_SETTINGS_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "settings XML exceeds {MAX_SETTINGS_XML_BYTES} bytes"
        )));
    }
    std::str::from_utf8(xml).map_err(|_| {
        Error::InvalidFormat("lossless settings mutation currently requires UTF-8 XML".into())
    })?;
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_qname = None;
    let mut word_prefix = None;
    let mut relationship_prefix = None;
    let mut strict = false;
    let mut root_empty_range = None;
    let mut root_end = None;
    #[cfg(any(feature = "fonts", test))]
    let mut embed_true_type_fonts_range = None;
    #[cfg(any(feature = "fonts", test))]
    let mut embed_true_type_fonts_enabled = None;
    #[cfg(any(feature = "fonts", test))]
    let mut embed_true_type_fonts_start = None;
    #[cfg(any(feature = "fonts", test))]
    let mut embed_true_type_fonts_insert_at = None;
    #[cfg(any(feature = "fonts", test))]
    let mut save_subset_fonts_range = None;
    #[cfg(any(feature = "fonts", test))]
    let mut save_subset_fonts_enabled = None;
    #[cfg(any(feature = "fonts", test))]
    let mut save_subset_fonts_start = None;
    #[cfg(any(feature = "fonts", test))]
    let mut save_subset_fonts_insert_at = None;
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
            .map_err(|_| Error::InvalidFormat("settings XML offset is too large".into()))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("settings XML offset is too large".into()))?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        if !matches!(&event, Event::Eof) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormat("settings XML node count overflow".into()))?;
            if nodes > MAX_SETTINGS_XML_NODES {
                return Err(Error::InvalidFormat(format!(
                    "settings XML exceeds {MAX_SETTINGS_XML_NODES} nodes"
                )));
            }
        }

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("settings XML nesting is too deep".into())
                })?;
                if depth > MAX_SETTINGS_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "settings XML exceeds depth {MAX_SETTINGS_XML_DEPTH}"
                    )));
                }
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
                #[cfg(any(feature = "fonts", test))]
                if depth == 2 && is_wordprocessing_namespace(&namespace) {
                    let local = element.local_name();
                    let local = local.as_ref();
                    if local == b"embedTrueTypeFonts" {
                        if embed_true_type_fonts_start.is_some()
                            || embed_true_type_fonts_range.is_some()
                        {
                            return Err(Error::InvalidFormat(
                                "settings has multiple embedTrueTypeFonts elements".into(),
                            ));
                        }
                        let enabled = parse_on_off(&element, reader.decoder(), &resolver)?;
                        embed_true_type_fonts_enabled = Some(enabled);
                        embed_true_type_fonts_start = Some(event_start);
                    } else if local == b"saveSubsetFonts" {
                        if save_subset_fonts_start.is_some() || save_subset_fonts_range.is_some() {
                            return Err(Error::InvalidFormat(
                                "settings has multiple saveSubsetFonts elements".into(),
                            ));
                        }
                        let enabled = parse_on_off(&element, reader.decoder(), &resolver)?;
                        save_subset_fonts_enabled = Some(enabled);
                        save_subset_fonts_start = Some(event_start);
                    }
                    if embed_true_type_fonts_insert_at.is_none()
                        && local != b"embedTrueTypeFonts"
                        && !is_before_embed_true_type_fonts(local)
                    {
                        embed_true_type_fonts_insert_at = Some(event_start);
                    }
                    if save_subset_fonts_insert_at.is_none()
                        && local != b"saveSubsetFonts"
                        && !is_before_save_subset_fonts(local)
                    {
                        save_subset_fonts_insert_at = Some(event_start);
                    }
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("settings XML nesting is too deep".into())
                })?;
                if child_depth > MAX_SETTINGS_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "settings XML exceeds depth {MAX_SETTINGS_XML_DEPTH}"
                    )));
                }
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
                #[cfg(any(feature = "fonts", test))]
                if child_depth == 2 && is_wordprocessing_namespace(&namespace) {
                    let local = element.local_name();
                    let local = local.as_ref();
                    if local == b"embedTrueTypeFonts" {
                        if embed_true_type_fonts_start.is_some()
                            || embed_true_type_fonts_range.is_some()
                        {
                            return Err(Error::InvalidFormat(
                                "settings has multiple embedTrueTypeFonts elements".into(),
                            ));
                        }
                        embed_true_type_fonts_enabled =
                            Some(parse_on_off(&element, reader.decoder(), &resolver)?);
                        embed_true_type_fonts_range = Some(event_start..event_end);
                    } else if local == b"saveSubsetFonts" {
                        if save_subset_fonts_start.is_some() || save_subset_fonts_range.is_some() {
                            return Err(Error::InvalidFormat(
                                "settings has multiple saveSubsetFonts elements".into(),
                            ));
                        }
                        save_subset_fonts_enabled =
                            Some(parse_on_off(&element, reader.decoder(), &resolver)?);
                        save_subset_fonts_range = Some(event_start..event_end);
                    }
                    if embed_true_type_fonts_insert_at.is_none()
                        && local != b"embedTrueTypeFonts"
                        && !is_before_embed_true_type_fonts(local)
                    {
                        embed_true_type_fonts_insert_at = Some(event_start);
                    }
                    if save_subset_fonts_insert_at.is_none()
                        && local != b"saveSubsetFonts"
                        && !is_before_save_subset_fonts(local)
                    {
                        save_subset_fonts_insert_at = Some(event_start);
                    }
                }
            },
            Event::End(_) => {
                #[cfg(any(feature = "fonts", test))]
                if depth == 2
                    && let Some(start) = embed_true_type_fonts_start.take()
                {
                    embed_true_type_fonts_range = Some(start..event_end);
                }
                #[cfg(any(feature = "fonts", test))]
                if depth == 2
                    && let Some(start) = save_subset_fonts_start.take()
                {
                    save_subset_fonts_range = Some(start..event_end);
                }
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
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid settings XML nesting".into()))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(SettingsXmlLayout {
        #[cfg(any(feature = "fonts", test))]
        embed_true_type_fonts_range,
        #[cfg(any(feature = "fonts", test))]
        embed_true_type_fonts_enabled,
        #[cfg(any(feature = "fonts", test))]
        embed_true_type_fonts_insert_at,
        #[cfg(any(feature = "fonts", test))]
        save_subset_fonts_range,
        #[cfg(any(feature = "fonts", test))]
        save_subset_fonts_enabled,
        #[cfg(any(feature = "fonts", test))]
        save_subset_fonts_insert_at,
        attached_template_range,
        doc_vars_range,
        doc_vars_insert_at,
        mail_merge_range,
        mail_merge_insert_at,
        root_empty_range,
        root_end,
        root_qname: root_qname
            .ok_or_else(|| Error::InvalidFormat("settings root is missing".into()))?,
        word_prefix,
        relationship_prefix,
        strict,
    })
}

#[cfg(any(feature = "fonts", test))]
fn is_before_embed_true_type_fonts(local_name: &[u8]) -> bool {
    matches!(
        local_name,
        b"writeProtection"
            | b"view"
            | b"zoom"
            | b"linkStyles"
            | b"removePersonalInformation"
            | b"removeDateAndTime"
            | b"doNotDisplayPageBoundaries"
            | b"displayBackgroundShape"
            | b"printPostScriptOverText"
            | b"printFractionalCharacterWidth"
            | b"printFormsData"
    )
}

#[cfg(any(feature = "fonts", test))]
fn is_before_save_subset_fonts(local_name: &[u8]) -> bool {
    is_before_embed_true_type_fonts(local_name)
        || matches!(local_name, b"embedTrueTypeFonts" | b"embedSystemFonts")
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

fn document_variables_element(layout: &SettingsXmlLayout, variables: &Variables) -> String {
    let prefix = layout
        .word_prefix
        .as_deref()
        .map(String::from_utf8_lossy)
        .unwrap_or_else(|| "w".into());
    let mut output = format!("<{prefix}:docVars");
    if layout.word_prefix.is_none() {
        let namespace = if layout.strict {
            STRICT_WORDPROCESSINGML_NAMESPACE
        } else {
            crate::namespace::WORDPROCESSINGML_NAMESPACE
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
        ResolveResult::Bound(Namespace(uri))
            if *uri == STRICT_WORDPROCESSINGML_NAMESPACE
    );
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        let Some(prefix) = key.strip_prefix(b"xmlns:") else {
            continue;
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
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
    word_attribute_value(element, name, decoder, resolver)?
        .ok_or_else(|| Error::InvalidFormat(format!("Word {description} attribute is required")))
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
        _ => Err(Error::InvalidFormat(format!(
            "invalid Word on/off value '{value}'"
        ))),
    }
}

fn map_docx_error(error: Error) -> Error {
    match error {
        Error::Invalid(message) => Error::InvalidFormat(message),
        other => other,
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
    use crate::settings::{COMPATIBILITY_MODE_SETTING_NAME, COMPATIBILITY_SETTING_URI};
    use litchi_opc::PackURI;
    use litchi_opc::part::{BlobPart, Part};
    use std::mem::size_of;

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

        let oversized_name = "n".repeat(crate::settings::MAX_SMART_TAG_NAME_CHARS + 1);
        let oversized = format!(
            r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:smartTagType w:namespaceuri="urn:test" w:name="{oversized_name}" w:url="https://example.test"/></w:settings>"#
        );
        assert!(matches!(
            DocumentSettings::extract_from_xml(oversized.as_bytes()),
            Err(Error::InvalidFormat(message)) if message.contains("smart-tag name")
        ));

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
        assert_eq!(
            settings.compatibility_options()[0].flag(),
            CompatFlag::UseFarEastLayout
        );
        assert!(settings.compatibility_options()[0].is_enabled());
        assert_eq!(
            settings.compatibility_options()[1].flag(),
            CompatFlag::DoNotExpandShiftReturn
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
    fn strict_and_transitional_compatibility_flag_domains_are_exhaustive() {
        const TRANSITIONAL_TOKENS: [&str; 65] = [
            "useSingleBorderforContiguousCells",
            "wpJustification",
            "noTabHangInd",
            "noLeading",
            "spaceForUL",
            "noColumnBalance",
            "balanceSingleByteDoubleByteWidth",
            "noExtraLineSpacing",
            "doNotLeaveBackslashAlone",
            "ulTrailSpace",
            "doNotExpandShiftReturn",
            "spacingInWholePoints",
            "lineWrapLikeWord6",
            "printBodyTextBeforeHeader",
            "printColBlack",
            "wpSpaceWidth",
            "showBreaksInFrames",
            "subFontBySize",
            "suppressBottomSpacing",
            "suppressTopSpacing",
            "suppressSpacingAtTopOfPage",
            "suppressTopSpacingWP",
            "suppressSpBfAfterPgBrk",
            "swapBordersFacingPages",
            "convMailMergeEsc",
            "truncateFontHeightsLikeWP6",
            "mwSmallCaps",
            "usePrinterMetrics",
            "doNotSuppressParagraphBorders",
            "wrapTrailSpaces",
            "footnoteLayoutLikeWW8",
            "shapeLayoutLikeWW8",
            "alignTablesRowByRow",
            "forgetLastTabAlignment",
            "adjustLineHeightInTable",
            "autoSpaceLikeWord95",
            "noSpaceRaiseLower",
            "doNotUseHTMLParagraphAutoSpacing",
            "layoutRawTableWidth",
            "layoutTableRowsApart",
            "useWord97LineBreakRules",
            "doNotBreakWrappedTables",
            "doNotSnapToGridInCell",
            "selectFldWithFirstOrLastChar",
            "applyBreakingRules",
            "doNotWrapTextWithPunct",
            "doNotUseEastAsianBreakRules",
            "useWord2002TableStyleRules",
            "growAutofit",
            "useFELayout",
            "useNormalStyleForList",
            "doNotUseIndentAsNumberingTabStop",
            "useAltKinsokuLineBreakRules",
            "allowSpaceOfSameStyleInTable",
            "doNotSuppressIndentation",
            "doNotAutofitConstrainedTables",
            "autofitToFirstFixedWidthCell",
            "underlineTabInNumList",
            "displayHangulFixedWidth",
            "splitPgBreakAndParaMark",
            "doNotVertAlignCellWithSp",
            "doNotBreakConstrainedForcedTable",
            "doNotVertAlignInTxbx",
            "useAnsiKerningPairs",
            "cachedColBalance",
        ];
        const STRICT_TOKENS: [&str; 7] = [
            "spaceForUL",
            "balanceSingleByteDoubleByteWidth",
            "doNotLeaveBackslashAlone",
            "ulTrailSpace",
            "doNotExpandShiftReturn",
            "adjustLineHeightInTable",
            "applyBreakingRules",
        ];

        assert_eq!(CompatFlag::ALL.len(), TRANSITIONAL_TOKENS.len());
        let mut flags = std::collections::HashSet::new();
        for (flag, raw) in CompatFlag::ALL.iter().copied().zip(TRANSITIONAL_TOKENS) {
            assert!(flags.insert(flag), "duplicate compatibility flag {raw}");
            assert_eq!(raw.parse(), Ok(flag));
            assert_eq!(flag.as_str(), raw);
            assert_eq!(flag.to_string(), raw);
            assert_eq!(flag.is_strict(), STRICT_TOKENS.contains(&raw));

            let transitional = format!(
                r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:{raw}/></w:compat></w:settings>"#
            );
            let parsed = DocumentSettings::extract_from_xml(transitional.as_bytes()).unwrap();
            assert_eq!(parsed.compatibility_options()[0].flag(), flag);

            let strict = format!(
                r#"<w:settings xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:compat><w:{raw}/></w:compat></w:settings>"#
            );
            let parsed = DocumentSettings::extract_from_xml(strict.as_bytes());
            if flag.is_strict() {
                assert_eq!(parsed.unwrap().compatibility_options()[0].flag(), flag);
            } else {
                assert!(parsed.is_err(), "Strict accepted Transitional-only {raw}");
            }
        }
        assert_eq!(flags.len(), 65);
        assert_eq!(CompatFlag::CachedColumnBalance as u8, 64);
        assert_eq!(size_of::<CompatFlag>(), 1);
        assert!("vendorCompat".parse::<CompatFlag>().is_err());
        assert!("UseFELayout".parse::<CompatFlag>().is_err());

        for namespace in [
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
        ] {
            let unknown = format!(
                r#"<w:settings xmlns:w="{namespace}"><w:compat><w:vendorCompat/></w:compat></w:settings>"#
            );
            assert!(DocumentSettings::extract_from_xml(unknown.as_bytes()).is_err());

            let duplicate = format!(
                r#"<w:settings xmlns:w="{namespace}"><w:compat><w:spaceForUL/><w:spaceForUL/></w:compat></w:settings>"#
            );
            assert!(DocumentSettings::extract_from_xml(duplicate.as_bytes()).is_err());
        }
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
        assert_eq!(footnotes.position(), Some(NotePosition::PageBottom));
        assert_eq!(footnotes.format(), Some(Format::LowerRoman));
        assert_eq!(footnotes.start(), Some(2));
        assert_eq!(footnotes.restart(), Some(NoteNumberingRestart::EachPage));

        let endnotes = settings.endnote_properties().unwrap();
        assert_eq!(endnotes.position(), Some(NotePosition::DocumentEnd));
        assert_eq!(endnotes.format(), Some(Format::UpperLetter));
        assert_eq!(endnotes.start(), None);
        assert_eq!(endnotes.restart(), None);
    }

    #[test]
    fn strict_and_transitional_note_position_domains_are_closed() {
        for (raw, expected) in [
            ("pageBottom", NotePosition::PageBottom),
            ("beneathText", NotePosition::BeneathText),
            ("sectEnd", NotePosition::SectionEnd),
            ("docEnd", NotePosition::DocumentEnd),
        ] {
            assert_eq!(raw.parse(), Ok(expected));
            assert_eq!(expected.as_str(), raw);
            assert_eq!(expected.to_string(), raw);

            for namespace in [
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "http://purl.oclc.org/ooxml/wordprocessingml/main",
            ] {
                let xml = format!(
                    r#"<w:settings xmlns:w="{namespace}"><w:footnotePr><w:pos w:val="{raw}"/></w:footnotePr></w:settings>"#
                );
                assert_eq!(
                    DocumentSettings::extract_from_xml(xml.as_bytes())
                        .unwrap()
                        .footnote_properties()
                        .unwrap()
                        .position(),
                    Some(expected)
                );
            }
        }
        assert!("vendorPosition".parse::<NotePosition>().is_err());
        assert!("PageBottom".parse::<NotePosition>().is_err());
        assert_eq!(size_of::<NotePosition>(), 1);

        for (raw, expected) in [
            ("sectEnd", NotePosition::SectionEnd),
            ("docEnd", NotePosition::DocumentEnd),
        ] {
            for namespace in [
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "http://purl.oclc.org/ooxml/wordprocessingml/main",
            ] {
                let xml = format!(
                    r#"<w:settings xmlns:w="{namespace}"><w:endnotePr><w:pos w:val="{raw}"/></w:endnotePr></w:settings>"#
                );
                assert_eq!(
                    DocumentSettings::extract_from_xml(xml.as_bytes())
                        .unwrap()
                        .endnote_properties()
                        .unwrap()
                        .position(),
                    Some(expected)
                );
            }
        }

        for raw in ["pageBottom", "beneathText"] {
            for namespace in [
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "http://purl.oclc.org/ooxml/wordprocessingml/main",
            ] {
                let xml = format!(
                    r#"<w:settings xmlns:w="{namespace}"><w:endnotePr><w:pos w:val="{raw}"/></w:endnotePr></w:settings>"#
                );
                assert!(DocumentSettings::extract_from_xml(xml.as_bytes()).is_err());
            }
        }

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

        let bad_position = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnotePr><w:pos w:val="vendorPosition"/></w:footnotePr></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(bad_position).is_err());

        let footnote_only_position = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnotePr><w:pos w:val="pageBottom"/></w:endnotePr></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(footnote_only_position).is_err());

        let bad_format = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnotePr><w:numFmt w:val="vendorNumbering"/></w:footnotePr></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(bad_format).is_err());

        let duplicate_child = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnotePr><w:numFmt w:val="decimal"/><w:numFmt w:val="bullet"/></w:footnotePr></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(duplicate_child).is_err());
    }

    #[test]
    fn parses_view_proofing_and_theme_defaults() {
        let xml = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:writeProtection/><w:view w:val="print"/><w:proofState w:spelling="clean" w:grammar="dirty"/><w:defaultTabStop w:val="720"/><w:themeFontLang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:hyperlink="hyperlink"/></w:settings>"#;
        let settings = DocumentSettings::extract_from_xml(xml).unwrap();
        assert!(settings.is_write_protected());
        assert_eq!(settings.view(), Some(View::Print));
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
        assert_eq!(settings.view(), Some(View::Web));
        let proofing = settings.proofing_state().unwrap();
        assert_eq!(proofing.spelling(), None);
        assert_eq!(proofing.grammar(), None);
    }

    #[test]
    fn editing_settings_enums_round_trip() {
        for (raw, expected) in [
            ("none", View::None),
            ("print", View::Print),
            ("outline", View::Outline),
            ("masterPages", View::MasterPages),
            ("normal", View::Normal),
            ("web", View::Web),
        ] {
            assert_eq!(View::from_xml(raw).unwrap(), expected);
            assert_eq!(expected.as_str(), raw);
        }
        assert!(View::from_xml("immersive").is_err());

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
    fn font_embedding_inserts_each_missing_flag_in_schema_order() {
        let missing_embed = br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:printFormsData/><q:saveSubsetFonts q:val="on"/><q:saveFormsData/></q:settings>"#;
        let patched = patch_font_embedding(missing_embed, true).unwrap();
        assert_eq!(
            patched,
            br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:printFormsData/><q:embedTrueTypeFonts/><q:saveSubsetFonts q:val="on"/><q:saveFormsData/></q:settings>"#
        );

        let missing_subset = br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:printFormsData/><q:embedTrueTypeFonts q:val="true"/><q:embedSystemFonts/><q:saveFormsData/></q:settings>"#;
        let patched = patch_font_embedding(missing_subset, true).unwrap();
        assert_eq!(
            patched,
            br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:printFormsData/><q:embedTrueTypeFonts q:val="true"/><q:embedSystemFonts/><q:saveSubsetFonts/><q:saveFormsData/></q:settings>"#
        );
    }

    #[test]
    fn font_embedding_rewrites_false_word_flags_without_touching_foreign_twins() {
        let xml = br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:not-wordprocessingml"><x:embedTrueTypeFonts x:val="off"/><q:embedTrueTypeFonts q:val="false" x:val="true"><x:keep/></q:embedTrueTypeFonts><q:saveSubsetFonts q:val="0"/><x:saveSubsetFonts/></q:settings>"#;
        let enabled = patch_font_embedding(xml, true).unwrap();
        assert_eq!(
            enabled,
            br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:not-wordprocessingml"><x:embedTrueTypeFonts x:val="off"/><q:embedTrueTypeFonts/><q:saveSubsetFonts/><x:saveSubsetFonts/></q:settings>"#
        );

        let full_font = patch_font_embedding(&enabled, false).unwrap();
        assert_eq!(
            full_font,
            br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:not-wordprocessingml"><x:embedTrueTypeFonts x:val="off"/><q:embedTrueTypeFonts/><x:saveSubsetFonts/></q:settings>"#
        );
    }

    #[test]
    fn font_embedding_expands_self_closing_strict_root() {
        let xml = br#"<?xml version="1.0"?><s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"/>"#;
        assert_eq!(
            patch_font_embedding(xml, true).unwrap(),
            br#"<?xml version="1.0"?><s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:embedTrueTypeFonts/><s:saveSubsetFonts/></s:settings>"#
        );
        assert_eq!(
            patch_font_embedding(xml, false).unwrap(),
            br#"<?xml version="1.0"?><s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:embedTrueTypeFonts/></s:settings>"#
        );

        let default_namespace =
            br#"<settings xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
        assert_eq!(
            patch_font_embedding(default_namespace, true).unwrap(),
            br#"<settings xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><embedTrueTypeFonts/><saveSubsetFonts/></settings>"#
        );
    }

    #[test]
    fn font_embedding_matching_flags_are_an_exact_namespace_aware_noop() {
        let xml = br#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:not-wordprocessingml"><!--keep--><x:embedTrueTypeFonts x:val="off"/><q:embedTrueTypeFonts x:val="off"/><q:saveSubsetFonts q:val="on" x:val="off"/><x:saveSubsetFonts/></q:settings>"#;
        assert_eq!(patch_font_embedding(xml, true).unwrap(), xml);

        let explicit_false = br#"<q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:embedTrueTypeFonts/><q:saveSubsetFonts q:val="off"/></q:settings>"#;
        assert_eq!(
            patch_font_embedding(explicit_false, false).unwrap(),
            explicit_false
        );
    }

    #[test]
    fn font_embedding_rejects_duplicate_or_invalid_word_flags() {
        let duplicate = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:embedTrueTypeFonts/><w:embedTrueTypeFonts/></w:settings>"#;
        assert!(patch_font_embedding(duplicate, true).is_err());

        let invalid = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:embedTrueTypeFonts w:val="maybe"/></w:settings>"#;
        assert!(patch_font_embedding(invalid, true).is_err());

        let mut utf16 = vec![0xFF, 0xFE];
        for unit in r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#.encode_utf16() {
            utf16.extend_from_slice(&unit.to_le_bytes());
        }
        assert!(patch_font_embedding(&utf16, true).is_err());
    }

    #[test]
    fn parses_bundled_settings_resource() {
        let settings =
            DocumentSettings::extract_from_xml(include_bytes!("../resources/settings.xml"))
                .unwrap();
        assert_eq!(settings.compatibility_mode(), Some(14));
        assert_eq!(settings.compatibility_settings().len(), 4);
        assert!(settings.compatibility_options().iter().any(|option| {
            option.flag() == CompatFlag::UseFarEastLayout && option.is_enabled()
        }));
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
        let mut variables = Variables::new();
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
        let mut variables = Variables::new();
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
