#![expect(
    clippy::format_push_string,
    reason = "serialization preserves the established byte-emission path"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::struct_excessive_bools,
    reason = "the public model preserves independent OOXML flags"
)]
//! Namespace-aware bounded XML codec for `WordprocessingML` document settings.

use super::colors::{ColorSchemeIndex, ColorSchemeMapping, ColorSchemeSlot};
use super::compatibility::{CompatFlag, CompatibilityOption, CompatibilitySetting};
use super::editing::{ProofState, ProofingState, ProtectionType, ThemeFontLanguages, View};
use super::extensions::Extensions;
use super::model::Settings;
use super::notes::{NoteNumberFormat, NoteNumberingProperties, NoteNumberingRestart, NotePosition};
use super::support::{invalid, reserve_one, xml_error};
use super::{
    MAX_SETTINGS_XML_BYTES, MAX_SETTINGS_XML_DEPTH, MAX_SETTINGS_XML_NODES, STRICT_WORD_NAMESPACE,
    TRANSITIONAL_WORD_NAMESPACE,
};
use crate::Result;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

impl<F: Copy> Settings<F> {
    /// Serialize the editing/view/theme settings in schema order.
    pub fn to_editing_settings_xml(&self, prefix: &str) -> String {
        let mut xml = String::new();
        if self.write_protection {
            xml.push_str(&format!("<{prefix}:writeProtection/>"));
        }
        if let Some(view) = self.view {
            xml.push_str(&format!(
                "<{prefix}:view {prefix}:val=\"{}\"/>",
                view.as_str()
            ));
        }
        if let Some(state) = &self.proofing_state {
            xml.push_str(&state.to_xml(prefix));
        }
        if let Some(twips) = self.default_tab_stop_twips {
            xml.push_str(&format!(
                "<{prefix}:defaultTabStop {prefix}:val=\"{twips}\"/>"
            ));
        }
        if let Some(languages) = &self.theme_font_languages {
            xml.push_str(&languages.to_xml(prefix));
        }
        if let Some(mapping) = &self.color_scheme_mapping {
            xml.push_str(&mapping.to_xml(prefix));
        }
        xml
    }
}

impl Settings<NoteNumberFormat> {
    /// Parse the bounded format-owned portion of a `settings.xml` document.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_SETTINGS_XML_BYTES {
            return Err(invalid(format!(
                "settings XML exceeds {MAX_SETTINGS_XML_BYTES} bytes"
            )));
        }

        let mut reader = NsReader::from_reader(xml);
        let mut settings = Self::new();
        settings.set_extensions(Extensions::parse(xml)?);
        let mut depth = 0usize;
        let mut nodes = 0usize;
        let mut saw_root = false;
        let mut strict_wordprocessingml = false;
        let mut seen = SeenSettings::default();
        let mut saw_compat = false;
        let mut pending_group: Option<PendingGroup> = None;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| xml_error(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);

            if matches!(event, Event::Start(_) | Event::Empty(_)) {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("settings XML node counter overflow"))?;
                if nodes > MAX_SETTINGS_XML_NODES {
                    return Err(invalid(format!(
                        "settings XML exceeds {MAX_SETTINGS_XML_NODES} nodes"
                    )));
                }
            }

            match event {
                Event::Start(element) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("Word settings XML nesting is too deep"))?;
                    if depth > MAX_SETTINGS_XML_DEPTH {
                        return Err(invalid(format!(
                            "settings XML exceeds depth {MAX_SETTINGS_XML_DEPTH}"
                        )));
                    }
                    if depth == 1 {
                        validate_settings_root(&namespace, &element, saw_root)?;
                        strict_wordprocessingml = matches!(
                            namespace,
                            ResolveResult::Bound(Namespace(uri)) if uri == STRICT_WORD_NAMESPACE
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
                    let child_depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("Word settings XML nesting is too deep"))?;
                    if child_depth > MAX_SETTINGS_XML_DEPTH {
                        return Err(invalid(format!(
                            "settings XML exceeds depth {MAX_SETTINGS_XML_DEPTH}"
                        )));
                    }
                    if child_depth == 1 {
                        validate_settings_root(&namespace, &element, saw_root)?;
                        strict_wordprocessingml = matches!(
                            namespace,
                            ResolveResult::Bound(Namespace(uri)) if uri == STRICT_WORD_NAMESPACE
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
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("invalid Word settings XML nesting"))?;
                },
                Event::Eof if depth != 0 => {
                    return Err(invalid("unterminated Word settings XML"));
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

        if !saw_root {
            return Err(invalid("settings part has no settings root"));
        }
        Ok(settings)
    }

    /// Serialize the complete modeled format-owned settings fragment.
    #[must_use]
    pub fn to_xml(&self, prefix: &str) -> String {
        let mut xml = String::new();
        if self.protected {
            xml.push_str(&format!("<{prefix}:documentProtection"));
            if let Some(protection_type) = self.protection_type {
                xml.push_str(&format!(" {prefix}:edit=\"{}\"", protection_type.to_xml()));
            }
            xml.push_str(&format!(" {prefix}:enforcement=\"on\"/>"));
        }
        if self.track_revisions {
            xml.push_str(&format!("<{prefix}:trackRevisions/>"));
        }
        if let Some(percent) = self.zoom_percent {
            xml.push_str(&format!("<{prefix}:zoom {prefix}:percent=\"{percent}\"/>"));
        }
        if !self.compatibility_options.is_empty() || !self.compatibility_settings.is_empty() {
            xml.push_str(&format!("<{prefix}:compat>"));
            for option in &self.compatibility_options {
                xml.push_str(&option.to_xml(prefix));
            }
            for setting in &self.compatibility_settings {
                xml.push_str(&setting.to_xml(prefix));
            }
            xml.push_str(&format!("</{prefix}:compat>"));
        }
        if let Some(properties) = &self.footnote_properties {
            xml.push_str(&properties.to_xml(prefix, "footnotePr"));
        }
        if let Some(properties) = &self.endnote_properties {
            xml.push_str(&properties.to_xml(prefix, "endnotePr"));
        }
        xml.push_str(&self.to_editing_settings_xml(prefix));
        xml.push_str(&self.extensions.to_xml(prefix));
        xml
    }
}

#[derive(Debug, Default)]
struct SeenSettings {
    write_protection: bool,
    view: bool,
    proofing_state: bool,
    default_tab_stop: bool,
    theme_font_languages: bool,
    color_scheme_mapping: bool,
}

enum PendingGroup {
    Compatibility {
        options: Vec<CompatibilityOption>,
        settings: Vec<CompatibilitySetting>,
    },
    FootnoteProperties(NoteNumberingProperties),
    EndnoteProperties(NoteNumberingProperties),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NoteKind {
    Footnote,
    Endnote,
}

fn begin_settings_group(
    element: &BytesStart<'_>,
    settings: &Settings,
    saw_compat: &mut bool,
) -> Result<Option<PendingGroup>> {
    match element.local_name().as_ref() {
        b"compat" => {
            if std::mem::replace(saw_compat, true) {
                return Err(invalid("duplicate compat settings group"));
            }
            Ok(Some(PendingGroup::Compatibility {
                options: Vec::new(),
                settings: Vec::new(),
            }))
        },
        b"footnotePr" => {
            if settings.footnote_properties.is_some() {
                return Err(invalid("duplicate footnotePr settings group"));
            }
            Ok(Some(PendingGroup::FootnoteProperties(
                NoteNumberingProperties::default(),
            )))
        },
        b"endnotePr" => {
            if settings.endnote_properties.is_some() {
                return Err(invalid("duplicate endnotePr settings group"));
            }
            Ok(Some(PendingGroup::EndnoteProperties(
                NoteNumberingProperties::default(),
            )))
        },
        _ => Ok(None),
    }
}

fn finish_settings_group(settings: &mut Settings, group: PendingGroup) -> Result<()> {
    match group {
        PendingGroup::Compatibility {
            options,
            settings: triples,
        } => {
            settings.compatibility_options = options;
            settings.compatibility_settings = triples;
        },
        PendingGroup::FootnoteProperties(properties) => {
            if settings.footnote_properties.replace(properties).is_some() {
                return Err(invalid("duplicate footnotePr settings group"));
            }
        },
        PendingGroup::EndnoteProperties(properties) => {
            if settings.endnote_properties.replace(properties).is_some() {
                return Err(invalid("duplicate endnotePr settings group"));
            }
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
                reserve_one(settings, "Word compatibility settings")?;
                settings.push(CompatibilitySetting::new(
                    required_attribute(element, b"name", decoder, resolver, "compatSetting name")?,
                    required_attribute(element, b"uri", decoder, resolver, "compatSetting URI")?,
                    required_attribute(element, b"val", decoder, resolver, "compatSetting value")?,
                ));
            } else {
                let local_name = element.local_name();
                let raw = std::str::from_utf8(local_name.as_ref()).map_err(|_source_error| {
                    invalid("compatibility flag name is not valid UTF-8")
                })?;
                let flag = raw.parse::<CompatFlag>().map_err(|_source_error| {
                    invalid(format!("invalid compatibility flag '{raw}'"))
                })?;
                if strict_wordprocessingml && !flag.is_strict() {
                    return Err(invalid(format!(
                        "compatibility flag '{raw}' is not valid in Strict WordprocessingML"
                    )));
                }
                if options.iter().any(|option| option.flag() == flag) {
                    return Err(invalid(format!("duplicate compatibility flag '{raw}'")));
                }
                reserve_one(options, "Word compatibility options")?;
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
    properties: &mut NoteNumberingProperties,
    kind: NoteKind,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"pos" => {
            if properties.position.is_some() {
                return Err(invalid("duplicate note position"));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note position")?;
            let position = value
                .parse::<NotePosition>()
                .map_err(|_source_error| invalid(format!("invalid note position '{value}'")))?;
            if kind == NoteKind::Endnote && !position.valid_for_endnote() {
                return Err(invalid(format!(
                    "position '{}' is not valid for an endnote",
                    position.as_str()
                )));
            }
            properties.position = Some(position);
        },
        b"numFmt" => {
            if properties.format.is_some() {
                return Err(invalid("duplicate note numbering format"));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note numFmt")?;
            properties.format = Some(value.parse().map_err(|_source_error| {
                invalid(format!("invalid note numbering format '{value}'"))
            })?);
        },
        b"numStart" => {
            if properties.start.is_some() {
                return Err(invalid("duplicate note numbering start"));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note numStart")?;
            properties.start = Some(value.parse().map_err(|_source_error| {
                invalid(format!("invalid note numbering start '{value}'"))
            })?);
        },
        b"numRestart" => {
            if properties.restart.is_some() {
                return Err(invalid("duplicate note numbering restart"));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "note numRestart")?;
            properties.restart = Some(NoteNumberingRestart::from_xml(&value)?);
        },
        // `w:footnote`/`w:endnote` separator references carry no properties.
        _ => {},
    }
    Ok(())
}

fn parse_setting(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    settings: &mut Settings,
    seen: &mut SeenSettings,
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
                settings.zoom_percent = value.parse::<u32>().ok();
            }
        },
        b"writeProtection" => {
            if std::mem::replace(&mut seen.write_protection, true) {
                return Err(invalid("duplicate writeProtection setting"));
            }
            settings.write_protection = parse_on_off(element, decoder, resolver)?;
        },
        b"view" => {
            if std::mem::replace(&mut seen.view, true) {
                return Err(invalid("duplicate view setting"));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "view mode")?;
            settings.view = Some(View::from_xml(&value)?);
        },
        b"proofState" => {
            if std::mem::replace(&mut seen.proofing_state, true) {
                return Err(invalid("duplicate proofState setting"));
            }
            let mut state = ProofingState::new();
            if let Some(value) = word_attribute_value(element, b"spelling", decoder, resolver)? {
                state.set_spelling(Some(ProofState::from_xml(&value)?));
            }
            if let Some(value) = word_attribute_value(element, b"grammar", decoder, resolver)? {
                state.set_grammar(Some(ProofState::from_xml(&value)?));
            }
            settings.proofing_state = Some(state);
        },
        b"defaultTabStop" => {
            if std::mem::replace(&mut seen.default_tab_stop, true) {
                return Err(invalid("duplicate defaultTabStop setting"));
            }
            let value = required_attribute(element, b"val", decoder, resolver, "default tab stop")?;
            settings.default_tab_stop_twips =
                Some(value.parse().map_err(|_source_error| {
                    invalid(format!("invalid default tab stop '{value}'"))
                })?);
        },
        b"themeFontLang" => {
            if std::mem::replace(&mut seen.theme_font_languages, true) {
                return Err(invalid("duplicate themeFontLang setting"));
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
            settings.theme_font_languages = Some(languages);
        },
        b"clrSchemeMapping" => {
            if std::mem::replace(&mut seen.color_scheme_mapping, true) {
                return Err(invalid("duplicate clrSchemeMapping setting"));
            }
            let mut mapping = ColorSchemeMapping::new();
            for slot in ColorSchemeSlot::ALL {
                if let Some(value) = word_attribute_value(
                    element,
                    slot.attribute_name().as_bytes(),
                    decoder,
                    resolver,
                )? {
                    mapping.set(slot, ColorSchemeIndex::from_xml(&value)?);
                }
            }
            settings.color_scheme_mapping = Some(mapping);
        },
        _ => {},
    }
    Ok(())
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
        return Err(invalid(
            "settings part has an invalid or trailing root element",
        ));
    }
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
        let attribute = attribute.map_err(|error| xml_error(error.to_string()))?;
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
            return Err(invalid(format!(
                "duplicate Word attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| xml_error(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

fn required_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<String> {
    word_attribute_value(element, name, decoder, resolver)?
        .ok_or_else(|| invalid(format!("Word {description} attribute is required")))
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
        _ => Err(invalid(format!("invalid Word on/off value '{value}'"))),
    }
}
