//! Bounded WordprocessingML settings XML model codec and snapshot edits.

use super::super::notes::{NoteNumberingProperties, NoteNumberingRestart, NotePosition};
use super::model::{AttachedTemplate, DocumentSettings};
use crate::Variables;
use crate::error::{Error, Result};
use crate::mail_merge::{Settings as MailMergeSettings, parse_settings_mail_merge};
use crate::namespace::{
    STRICT_WORDPROCESSINGML_NAMESPACE, is_wordprocessing_namespace, word_attribute_value,
};
use crate::numbering::Format;
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
const MAX_SETTINGS_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_SETTINGS_XML_NODES: usize = 250_000;
const MAX_SETTINGS_XML_DEPTH: usize = 256;

use crate::settings::{
    ColorSchemeIndex, ColorSchemeMapping, ColorSchemeSlot, CompatFlag, CompatibilityOption,
    CompatibilitySetting, Extensions, ProofState, ProofingState, ProtectionType, SmartTagType,
    ThemeFontLanguages, View,
};

impl DocumentSettings {
    pub(super) fn extract_from_xml(xml_bytes: &[u8]) -> Result<Self> {
        let mut reader = NsReader::from_reader(xml_bytes);

        let mut settings = Self::new();
        settings
            .values
            .set_extensions(Extensions::parse(xml_bytes)?);
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
