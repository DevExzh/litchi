#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::cast_possible_truncation,
    reason = "OOXML numeric values are bounded before conversion"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Namespace-aware XML codec for `WordprocessingML` numbering.
//!
//! OPC parts and markup-compatibility preprocessing stay in the host crate;
//! this layer consumes the resulting XML bytes only.

use super::model::{
    Collection, Definition, Instance, Level, Override, PictureBullet, Restart, Suffix,
};
use super::transaction::Change;
use super::validation;
use crate::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::borrow::Cow;

const TRANSITIONAL_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const VML_NAMESPACE: &[u8] = b"urn:schemas-microsoft-com:vml";
const DRAWINGML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const TRANSITIONAL_WORDPROCESSING_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORDPROCESSING_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";

struct PendingDefinition {
    depth: usize,
    value: Definition,
}

struct PendingInstance {
    depth: usize,
    id: u32,
    abstract_num_id: Option<u32>,
    overrides: Vec<Override>,
}

struct PendingOverride {
    depth: usize,
    level: u8,
    start_override: Option<i64>,
    definition: Option<Level>,
}

struct PendingLevel {
    depth: usize,
    value: Level,
    in_override: bool,
}

struct PendingPictureBullet {
    depth: usize,
    id: u32,
    image_relationship_id: Option<String>,
}

impl Collection {
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        validation::validate_xml(xml)?;
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().check_end_names = true;
        let mut result = Self::new();
        let mut abstract_num: Option<PendingDefinition> = None;
        let mut num: Option<PendingInstance> = None;
        let mut level_override: Option<PendingOverride> = None;
        let mut level: Option<PendingLevel> = None;
        let mut picture_bullet: Option<PendingPictureBullet> = None;
        let mut ignorable = Vec::<Vec<Vec<u8>>>::new();
        let mut depth = 0usize;
        let mut nodes = 0usize;

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
                    nodes = nodes
                        .checked_add(1)
                        .ok_or_else(|| invalid("numbering XML element counter overflow"))?;
                    if nodes > validation::MAX_XML_NODES {
                        return Err(invalid(&format!(
                            "numbering XML exceeds {} elements",
                            validation::MAX_XML_NODES
                        )));
                    }
                    depth = depth.checked_add(1).ok_or_else(too_deep)?;
                    if depth > validation::MAX_XML_DEPTH {
                        return Err(invalid(&format!(
                            "numbering XML nesting exceeds {}",
                            validation::MAX_XML_DEPTH
                        )));
                    }
                    let effective_ignorable =
                        effective_ignorable(&element, decoder, &resolver, ignorable.last())?;
                    begin_element(
                        &namespace,
                        &element,
                        decoder,
                        &resolver,
                        depth,
                        &mut abstract_num,
                        &mut num,
                        &mut level_override,
                        &mut level,
                        &mut picture_bullet,
                        &effective_ignorable,
                    )?;
                    ignorable.push(effective_ignorable);
                },
                Event::Empty(element) => {
                    nodes = nodes
                        .checked_add(1)
                        .ok_or_else(|| invalid("numbering XML element counter overflow"))?;
                    if nodes > validation::MAX_XML_NODES {
                        return Err(invalid(&format!(
                            "numbering XML exceeds {} elements",
                            validation::MAX_XML_NODES
                        )));
                    }
                    let child_depth = depth.checked_add(1).ok_or_else(too_deep)?;
                    if child_depth > validation::MAX_XML_DEPTH {
                        return Err(invalid(&format!(
                            "numbering XML nesting exceeds {}",
                            validation::MAX_XML_DEPTH
                        )));
                    }
                    let effective_ignorable =
                        effective_ignorable(&element, decoder, &resolver, ignorable.last())?;
                    empty_element(
                        &namespace,
                        &element,
                        decoder,
                        &resolver,
                        child_depth,
                        &mut result,
                        &mut abstract_num,
                        &mut num,
                        &mut level_override,
                        &mut level,
                        &mut picture_bullet,
                        &effective_ignorable,
                    )?;
                },
                Event::End(element) => {
                    if is_wordprocessing_namespace(&namespace) {
                        match element.local_name().as_ref() {
                            b"numPicBullet"
                                if picture_bullet
                                    .as_ref()
                                    .is_some_and(|value| value.depth == depth) =>
                            {
                                let pending = picture_bullet.take().ok_or_else(|| {
                                    invalid("picture bullet state missing at end")
                                })?;
                                push_picture_bullet(
                                    &mut result.picture_bullets,
                                    PictureBullet {
                                        id: pending.id,
                                        image_relationship_id: pending.image_relationship_id,
                                    },
                                )?;
                            },
                            b"lvl" if level.as_ref().is_some_and(|value| value.depth == depth) => {
                                let pending = level
                                    .take()
                                    .ok_or_else(|| invalid("level state missing at end"))?;
                                finish_level(&mut abstract_num, &mut level_override, pending)?;
                            },
                            b"lvlOverride"
                                if level_override
                                    .as_ref()
                                    .is_some_and(|value| value.depth == depth) =>
                            {
                                if level.is_some() {
                                    return Err(invalid("unterminated level in lvlOverride"));
                                }
                                let pending = level_override
                                    .take()
                                    .ok_or_else(|| invalid("override state missing at end"))?;
                                finish_override(&mut num, pending)?;
                            },
                            b"abstractNum"
                                if abstract_num
                                    .as_ref()
                                    .is_some_and(|value| value.depth == depth) =>
                            {
                                if level.is_some() {
                                    return Err(invalid("unterminated abstract numbering level"));
                                }
                                let pending = abstract_num.take().ok_or_else(|| {
                                    invalid("abstract numbering state missing at end")
                                })?;
                                push_abstract(&mut result.abstract_nums, pending.value)?;
                            },
                            b"num" if num.as_ref().is_some_and(|value| value.depth == depth) => {
                                if level.is_some() || level_override.is_some() {
                                    return Err(invalid("unterminated numbering override"));
                                }
                                let pending = num
                                    .take()
                                    .ok_or_else(|| invalid("num state missing at end"))?;
                                let abstract_num_id = pending.abstract_num_id.ok_or_else(|| {
                                    invalid(&format!(
                                        "numbering instance {} is missing abstractNumId",
                                        pending.id
                                    ))
                                })?;
                                push_num(
                                    &mut result.nums,
                                    Instance {
                                        id: pending.id,
                                        abstract_num_id,
                                        overrides: pending.overrides,
                                    },
                                )?;
                            },
                            _ => {},
                        }
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("invalid numbering XML nesting"))?;
                    ignorable
                        .pop()
                        .ok_or_else(|| invalid("numbering XML scope stack underflow"))?;
                },
                Event::Eof
                    if depth != 0
                        || abstract_num.is_some()
                        || num.is_some()
                        || level_override.is_some()
                        || level.is_some()
                        || picture_bullet.is_some()
                        || !ignorable.is_empty() =>
                {
                    return Err(invalid("unterminated numbering XML"));
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

        for value in &result.nums {
            if result.get_abstract_num(value.abstract_num_id).is_none() {
                return Err(invalid(&format!(
                    "numbering instance {} references missing abstractNum {}",
                    value.id, value.abstract_num_id
                )));
            }
        }
        Ok(result)
    }
}

#[allow(
    clippy::too_many_arguments,
    reason = "signature mirrors the corresponding OOXML record"
)]
fn begin_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    depth: usize,
    abstract_num: &mut Option<PendingDefinition>,
    num: &mut Option<PendingInstance>,
    level_override: &mut Option<PendingOverride>,
    level: &mut Option<PendingLevel>,
    picture_bullet: &mut Option<PendingPictureBullet>,
    ignorable: &[Vec<u8>],
) -> Result<()> {
    if let Some(pending) = picture_bullet.as_mut() {
        return capture_picture_bullet_image(namespace, element, decoder, resolver, pending);
    }
    if !is_wordprocessing_namespace(namespace) {
        return Ok(());
    }
    let name = element.local_name();
    match name.as_ref() {
        b"numPicBullet" if abstract_num.is_none() && num.is_none() => {
            *picture_bullet = Some(PendingPictureBullet {
                depth,
                id: required_u32(element, b"numPicBulletId", decoder, resolver)?,
                image_relationship_id: None,
            });
        },
        b"abstractNum" => {
            if abstract_num.is_some() || num.is_some() {
                return Err(invalid("nested numbering definitions are invalid"));
            }
            *abstract_num = Some(PendingDefinition {
                depth,
                value: Definition {
                    id: required_u32(element, b"abstractNumId", decoder, resolver)?,
                    num_type: None,
                    num_style_link: None,
                    style_link: None,
                    restart_numbering_after_break: restart_numbering_after_break(
                        element, decoder, resolver, ignorable,
                    )?,
                    levels: Vec::new(),
                },
            });
        },
        b"num" => {
            if abstract_num.is_some() || num.is_some() {
                return Err(invalid("nested numbering definitions are invalid"));
            }
            *num = Some(PendingInstance {
                depth,
                id: required_u32(element, b"numId", decoder, resolver)?,
                abstract_num_id: None,
                overrides: Vec::new(),
            });
        },
        b"lvl"
            if level.is_none()
                && level_override
                    .as_ref()
                    .is_some_and(|value| depth == value.depth + 1) =>
        {
            *level = Some(PendingLevel {
                depth,
                value: Level::new(required_level(element, decoder, resolver)?),
                in_override: true,
            });
        },
        b"lvl"
            if level.is_none()
                && abstract_num
                    .as_ref()
                    .is_some_and(|value| depth == value.depth + 1) =>
        {
            *level = Some(PendingLevel {
                depth,
                value: Level::new(required_level(element, decoder, resolver)?),
                in_override: false,
            });
        },
        b"lvlOverride"
            if level_override.is_none()
                && num.as_ref().is_some_and(|value| depth == value.depth + 1) =>
        {
            *level_override = Some(PendingOverride {
                depth,
                level: required_level(element, decoder, resolver)?,
                start_override: None,
                definition: None,
            });
        },
        _ => apply_child(
            element,
            decoder,
            resolver,
            depth,
            abstract_num,
            num,
            level_override,
            level,
        )?,
    }
    Ok(())
}

#[allow(
    clippy::too_many_arguments,
    reason = "signature mirrors the corresponding OOXML record"
)]
fn empty_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    depth: usize,
    result: &mut Collection,
    abstract_num: &mut Option<PendingDefinition>,
    num: &mut Option<PendingInstance>,
    level_override: &mut Option<PendingOverride>,
    level: &mut Option<PendingLevel>,
    picture_bullet: &mut Option<PendingPictureBullet>,
    ignorable: &[Vec<u8>],
) -> Result<()> {
    if let Some(pending) = picture_bullet.as_mut() {
        return capture_picture_bullet_image(namespace, element, decoder, resolver, pending);
    }
    if !is_wordprocessing_namespace(namespace) {
        return Ok(());
    }
    match element.local_name().as_ref() {
        b"numPicBullet" if abstract_num.is_none() && num.is_none() => {
            push_picture_bullet(
                &mut result.picture_bullets,
                PictureBullet {
                    id: required_u32(element, b"numPicBulletId", decoder, resolver)?,
                    image_relationship_id: None,
                },
            )?;
        },
        b"abstractNum" => {
            if abstract_num.is_some() || num.is_some() {
                return Err(invalid("nested numbering definitions are invalid"));
            }
            push_abstract(
                &mut result.abstract_nums,
                Definition {
                    id: required_u32(element, b"abstractNumId", decoder, resolver)?,
                    num_type: None,
                    num_style_link: None,
                    style_link: None,
                    restart_numbering_after_break: restart_numbering_after_break(
                        element, decoder, resolver, ignorable,
                    )?,
                    levels: Vec::new(),
                },
            )?;
        },
        b"num" => return Err(invalid("numbering instance is missing abstractNumId")),
        b"lvl"
            if level_override
                .as_ref()
                .is_some_and(|value| depth == value.depth + 1) =>
        {
            finish_level(
                abstract_num,
                level_override,
                PendingLevel {
                    depth,
                    value: Level::new(required_level(element, decoder, resolver)?),
                    in_override: true,
                },
            )?;
        },
        b"lvl"
            if abstract_num
                .as_ref()
                .is_some_and(|value| depth == value.depth + 1) =>
        {
            finish_level(
                abstract_num,
                level_override,
                PendingLevel {
                    depth,
                    value: Level::new(required_level(element, decoder, resolver)?),
                    in_override: false,
                },
            )?;
        },
        b"lvlOverride" if num.as_ref().is_some_and(|value| depth == value.depth + 1) => {
            finish_override(
                num,
                PendingOverride {
                    depth,
                    level: required_level(element, decoder, resolver)?,
                    start_override: None,
                    definition: None,
                },
            )?;
        },
        _ => apply_child(
            element,
            decoder,
            resolver,
            depth,
            abstract_num,
            num,
            level_override,
            level,
        )?,
    }
    Ok(())
}

#[allow(
    clippy::too_many_arguments,
    reason = "signature mirrors the corresponding OOXML record"
)]
fn apply_child(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    depth: usize,
    abstract_num: &mut Option<PendingDefinition>,
    num: &mut Option<PendingInstance>,
    level_override: &mut Option<PendingOverride>,
    level: &mut Option<PendingLevel>,
) -> Result<()> {
    if let Some(value) = level.as_mut().filter(|value| depth == value.depth + 1) {
        match element.local_name().as_ref() {
            b"start" => value.value.start = required_i64(element, b"val", decoder, resolver)?,
            b"numFmt" => {
                let raw = required_string(element, b"val", decoder, resolver)?;
                value.value.format = raw.parse().map_err(|_source_error| {
                    invalid(&format!("invalid numbering format '{raw}'"))
                })?;
                value.value.custom_format =
                    word_attribute_value(element, b"format", decoder, resolver)?;
            },
            b"lvlText" => {
                value.value.level_text = Some(
                    word_attribute_value(element, b"val", decoder, resolver)?.unwrap_or_default(),
                );
            },
            b"suff" => {
                value.value.suffix = match required_string(element, b"val", decoder, resolver)?
                    .as_str()
                {
                    "tab" => Suffix::Tab,
                    "space" => Suffix::Space,
                    "nothing" => Suffix::Nothing,
                    other => return Err(invalid(&format!("invalid numbering suffix '{other}'"))),
                }
            },
            b"lvlRestart" => {
                let raw = required_u32(element, b"val", decoder, resolver)?;
                value.value.restart = match raw {
                    0 => Restart::Never,
                    1..=9 => Restart::After((raw - 1) as u8),
                    _ => return Err(invalid(&format!("invalid lvlRestart '{raw}'"))),
                };
            },
            b"isLgl" => value.value.legal = on_off(element, decoder, resolver)?,
            b"pStyle" => {
                value.value.paragraph_style =
                    Some(required_string(element, b"val", decoder, resolver)?);
            },
            b"lvlPicBulletId" => {
                value.value.picture_bullet_id =
                    Some(required_u32(element, b"val", decoder, resolver)?);
            },
            _ => {},
        }
        return Ok(());
    }
    if let Some(value) = level_override
        .as_mut()
        .filter(|value| depth == value.depth + 1)
    {
        if element.local_name().as_ref() == b"startOverride" {
            if value.start_override.is_some() {
                return Err(invalid("duplicate startOverride"));
            }
            value.start_override = Some(required_i64(element, b"val", decoder, resolver)?);
        }
        return Ok(());
    }
    if let Some(value) = abstract_num
        .as_mut()
        .filter(|value| depth == value.depth + 1)
    {
        match element.local_name().as_ref() {
            b"multiLevelType" => {
                let raw = required_string(element, b"val", decoder, resolver)?;
                if value.value.num_type.is_some() {
                    return Err(invalid("duplicate multiLevelType"));
                }
                value.value.num_type = Some(raw.parse().map_err(|_source_error| {
                    invalid(&format!("invalid multiLevelType '{raw}'"))
                })?);
            },
            b"numStyleLink" => {
                let raw = required_string(element, b"val", decoder, resolver)?;
                set_once(&mut value.value.num_style_link, raw, "numStyleLink")?;
            },
            b"styleLink" => {
                let raw = required_string(element, b"val", decoder, resolver)?;
                set_once(&mut value.value.style_link, raw, "styleLink")?;
            },
            _ => {},
        }
    }
    if let Some(value) = num.as_mut().filter(|value| depth == value.depth + 1)
        && element.local_name().as_ref() == b"abstractNumId"
    {
        if value.abstract_num_id.is_some() {
            return Err(invalid("duplicate abstractNumId"));
        }
        value.abstract_num_id = Some(required_u32(element, b"val", decoder, resolver)?);
    }
    Ok(())
}

fn finish_level(
    abstract_num: &mut Option<PendingDefinition>,
    level_override: &mut Option<PendingOverride>,
    pending: PendingLevel,
) -> Result<()> {
    if pending.in_override {
        let target = level_override
            .as_mut()
            .ok_or_else(|| invalid("level outside lvlOverride"))?;
        if target.definition.is_some() {
            return Err(invalid("duplicate override level definition"));
        }
        if pending.value.level != target.level {
            return Err(invalid("override level indices do not match"));
        }
        target.definition = Some(pending.value);
    } else {
        let target = abstract_num
            .as_mut()
            .ok_or_else(|| invalid("level outside abstractNum"))?;
        if target
            .value
            .levels
            .iter()
            .any(|value| value.level == pending.value.level)
        {
            return Err(invalid("duplicate abstract numbering level"));
        }
        target.value.levels.push(pending.value);
    }
    Ok(())
}

fn finish_override(num: &mut Option<PendingInstance>, pending: PendingOverride) -> Result<()> {
    let target = num
        .as_mut()
        .ok_or_else(|| invalid("lvlOverride outside num"))?;
    if target
        .overrides
        .iter()
        .any(|value| value.level == pending.level)
    {
        return Err(invalid("duplicate lvlOverride"));
    }
    target.overrides.push(Override {
        level: pending.level,
        start_override: pending.start_override,
        definition: pending.definition,
    });
    Ok(())
}

fn set_once(slot: &mut Option<String>, value: String, name: &str) -> Result<()> {
    if slot.is_some() {
        return Err(invalid(&format!("duplicate {name}")));
    }
    *slot = Some(value);
    Ok(())
}

fn push_abstract(values: &mut Vec<Definition>, value: Definition) -> Result<()> {
    if values.iter().any(|item| item.id == value.id) {
        return Err(invalid(&format!(
            "duplicate abstract numbering ID {}",
            value.id
        )));
    }
    values.push(value);
    Ok(())
}

fn push_num(values: &mut Vec<Instance>, value: Instance) -> Result<()> {
    if values.iter().any(|item| item.id == value.id) {
        return Err(invalid(&format!(
            "duplicate numbering instance ID {}",
            value.id
        )));
    }
    values.push(value);
    Ok(())
}

fn push_picture_bullet(values: &mut Vec<PictureBullet>, value: PictureBullet) -> Result<()> {
    if values.iter().any(|item| item.id() == value.id()) {
        return Err(invalid(&format!(
            "duplicate picture bullet ID {}",
            value.id()
        )));
    }
    values.push(value);
    Ok(())
}

/// Capture the first image relationship inside a `w:numPicBullet` definition.
///
/// Word writes the bullet picture either as VML (`v:imagedata r:id`) or as
/// `DrawingML` (`a:blip r:embed`/`a:link`); everything else inside `w:pict` is
/// inert shape geometry and is ignored.
fn capture_picture_bullet_image(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    pending: &mut PendingPictureBullet,
) -> Result<()> {
    if pending.image_relationship_id.is_some() {
        return Ok(());
    }
    let names: &[&[u8]] = match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == VML_NAMESPACE => {
            match element.local_name().as_ref() {
                b"imagedata" => &[b"id"],
                _ => return Ok(()),
            }
        },
        ResolveResult::Bound(Namespace(uri))
            if *uri == DRAWINGML_NAMESPACE || *uri == STRICT_DRAWINGML_NAMESPACE =>
        {
            match element.local_name().as_ref() {
                b"blip" => &[b"embed", b"link"],
                _ => return Ok(()),
            }
        },
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
            return Ok(());
        },
    };
    pending.image_relationship_id = relationship_attribute(element, names, decoder, resolver)?;
    Ok(())
}

fn relationship_attribute(
    element: &BytesStart<'_>,
    names: &[&[u8]],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if !names.contains(&attribute.key.local_name().as_ref()) {
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
            return Err(invalid("duplicate picture bullet image relationship"));
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

fn required_string(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<String> {
    word_attribute_value(element, name, decoder, resolver)?.ok_or_else(|| {
        invalid(&format!(
            "Word numbering element is missing required '{}' attribute",
            String::from_utf8_lossy(name)
        ))
    })
}

fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<u32> {
    let value = required_string(element, name, decoder, resolver)?;
    value
        .parse()
        .map_err(|_source_error| invalid(&format!("invalid Word numbering integer '{value}'")))
}

fn required_i64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<i64> {
    let value = required_string(element, name, decoder, resolver)?;
    value
        .parse()
        .map_err(|_source_error| invalid(&format!("invalid Word numbering integer '{value}'")))
}

fn required_level(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<u8> {
    let value = required_u32(element, b"ilvl", decoder, resolver)?;
    u8::try_from(value)
        .ok()
        .filter(|value| *value <= 8)
        .ok_or_else(|| invalid(&format!("invalid numbering level '{value}'")))
}

fn effective_ignorable(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    inherited: Option<&Vec<Vec<u8>>>,
) -> Result<Vec<Vec<u8>>> {
    let mut direct = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"Ignorable" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_markup_compatibility_namespace(&namespace) {
            continue;
        }
        if direct.is_some() {
            return Err(invalid("duplicate numbering mc:Ignorable attributes"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        direct = Some(validation::parse_ignorable(&value)?);
    }
    Ok(direct.or_else(|| inherited.cloned()).unwrap_or_default())
}

fn restart_numbering_after_break(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    ignorable: &[Vec<u8>],
) -> Result<Option<bool>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"restartNumberingAfterBreak" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_word_2012_namespace(&namespace) {
            continue;
        }
        let prefix = attribute
            .key
            .prefix()
            .map(|prefix| prefix.into_inner().to_vec())
            .filter(|prefix| !prefix.is_empty())
            .ok_or_else(|| {
                invalid("restartNumberingAfterBreak must use a prefixed Word 2012 attribute")
            })?;
        if !validation::has_ignorable_prefix(ignorable, &prefix) {
            return Err(invalid(
                "restartNumberingAfterBreak namespace is not listed in numbering mc:Ignorable",
            ));
        }
        if value.is_some() {
            return Err(invalid("duplicate restartNumberingAfterBreak attributes"));
        }
        let raw = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        value = Some(validation::parse_on_off(&raw)?);
    }
    Ok(value)
}

fn on_off(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<bool> {
    Ok(
        match word_attribute_value(element, b"val", decoder, resolver)?.as_deref() {
            None | Some("1" | "true" | "on") => true,
            Some("0" | "false" | "off") => false,
            Some(value) => return Err(invalid(&format!("invalid on/off value '{value}'"))),
        },
    )
}

fn too_deep() -> Error {
    invalid("numbering XML nesting is too deep")
}
fn invalid(message: &str) -> Error {
    Error::Invalid(message.to_owned())
}

#[derive(Debug, Clone)]
struct RawAttribute {
    name: Vec<u8>,
    name_start: usize,
    value_start: usize,
    value_end: usize,
}

#[derive(Debug, Clone)]
struct AttributeRange {
    name_start: usize,
    value_start: usize,
    value_end: usize,
}

#[derive(Debug, Clone, Default)]
struct Scope {
    namespaces: Vec<(Vec<u8>, Vec<u8>)>,
    ignorable: Vec<Vec<u8>>,
}

impl Scope {
    fn set_namespace(&mut self, prefix: Vec<u8>, namespace: Vec<u8>) {
        if let Some((_, value)) = self
            .namespaces
            .iter_mut()
            .find(|(candidate, _)| *candidate == prefix)
        {
            *value = namespace;
        } else {
            self.namespaces.push((prefix, namespace));
        }
    }

    fn namespace(&self, prefix: &[u8]) -> Option<&[u8]> {
        self.namespaces
            .iter()
            .find(|(candidate, _)| candidate.as_slice() == prefix)
            .map(|(_, namespace)| namespace.as_slice())
    }

    fn prefix_for(&self, namespace: &[u8]) -> Option<Vec<u8>> {
        self.namespaces
            .iter()
            .find(|(prefix, candidate)| !prefix.is_empty() && candidate.as_slice() == namespace)
            .map(|(prefix, _)| prefix.clone())
    }
}

#[derive(Debug, Clone)]
struct DefinitionLocation {
    id: u32,
    start: usize,
    tag_end: usize,
    extension: Option<AttributeRange>,
    ignorable: Option<AttributeRange>,
    scope: Scope,
}

#[derive(Debug, Default)]
struct Layout {
    definitions: Vec<DefinitionLocation>,
}

#[derive(Debug, Clone)]
struct Edit {
    start: usize,
    end: usize,
    replacement: Vec<u8>,
}

/// Rewrite only the Word 2012 restart attribute seams in `numbering.xml`.
/// Every byte outside an edited attribute or the owning `abstractNum` start
/// tag is copied from the authored source.
pub(crate) fn rewrite_restart_numbering_after_break(
    xml: &[u8],
    changes: &[Change],
) -> Result<Vec<u8>> {
    validation::validate_xml(xml)?;
    let layout = locate_definitions(xml)?;
    let mut edits = Vec::new();

    for change in changes {
        let location = layout
            .definitions
            .iter()
            .find(|definition| definition.id == change.abstract_num_id)
            .ok_or_else(|| {
                invalid(&format!(
                    "abstract numbering definition {} does not exist",
                    change.abstract_num_id
                ))
            })?;

        let current = extension_value(xml, location)?;
        if current != change.before {
            return Err(invalid(&format!(
                "abstract numbering definition {} does not match its source policy",
                change.abstract_num_id
            )));
        }
        if change.before == change.after {
            continue;
        }

        match (location.extension.as_ref(), change.after) {
            (Some(attribute), Some(value)) => edits.push(Edit {
                start: attribute.value_start,
                end: attribute.value_end,
                replacement: on_off_lexical(value).as_bytes().to_vec(),
            }),
            (Some(attribute), None) => edits.push(Edit {
                start: attribute.name_start,
                end: attribute.value_end,
                replacement: Vec::new(),
            }),
            (None, Some(value)) => {
                insert_restart_attribute(xml, location, value, &mut edits)?;
            },
            (None, None) => {
                return Err(invalid(
                    "numbering edit tried to remove an absent restart attribute",
                ));
            },
        }
    }

    if edits.is_empty() {
        return Ok(xml.to_vec());
    }
    edits.sort_by(|left, right| {
        right
            .start
            .cmp(&left.start)
            .then_with(|| right.end.cmp(&left.end))
    });
    for pair in edits.windows(2) {
        if pair[0].start < pair[1].end {
            return Err(invalid("overlapping numbering XML edits"));
        }
    }
    let mut output = xml.to_vec();
    for edit in edits {
        if edit.end > output.len() || edit.start > edit.end {
            return Err(invalid("numbering XML edit range is outside its source"));
        }
        output.splice(edit.start..edit.end, edit.replacement);
    }
    Ok(output)
}

fn extension_value(xml: &[u8], location: &DefinitionLocation) -> Result<Option<bool>> {
    let Some(attribute) = location.extension.as_ref() else {
        return Ok(None);
    };
    let value = xml
        .get(attribute.value_start..attribute.value_end)
        .ok_or_else(|| invalid("restart attribute value range is outside numbering XML"))?;
    let value = std::str::from_utf8(value)
        .map_err(|error| invalid(&format!("restart attribute is not UTF-8: {error}")))?;
    validation::parse_on_off(value).map(Some)
}

fn insert_restart_attribute(
    xml: &[u8],
    location: &DefinitionLocation,
    value: bool,
    edits: &mut Vec<Edit>,
) -> Result<()> {
    let extension_prefix = location
        .scope
        .prefix_for(validation::WORD_2012_NAMESPACE)
        .unwrap_or_else(|| choose_prefix(&location.scope, b"w15"));
    let extension_prefix_is_bound = location
        .scope
        .namespace(&extension_prefix)
        .is_some_and(|namespace| namespace == validation::WORD_2012_NAMESPACE);
    let extension_name = qualified_name(&extension_prefix, b"restartNumberingAfterBreak");

    let mut insertion = Vec::new();
    if !extension_prefix_is_bound {
        append_attribute(
            &mut insertion,
            &qualified_name(b"xmlns", &extension_prefix),
            validation::WORD_2012_NAMESPACE,
        );
    }

    if let Some(ignorable) = location.ignorable.as_ref() {
        let raw = xml
            .get(ignorable.value_start..ignorable.value_end)
            .ok_or_else(|| invalid("numbering mc:Ignorable range is outside its source"))?;
        let raw = std::str::from_utf8(raw)
            .map_err(|error| invalid(&format!("numbering mc:Ignorable is not UTF-8: {error}")))?;
        if !validation::has_ignorable_prefix(&location.scope.ignorable, &extension_prefix) {
            let replacement = append_ignorable_token(raw, &extension_prefix);
            edits.push(Edit {
                start: ignorable.value_start,
                end: ignorable.value_end,
                replacement,
            });
        }
    } else if !validation::has_ignorable_prefix(&location.scope.ignorable, &extension_prefix) {
        let compatibility_prefix = location
            .scope
            .prefix_for(validation::MC_NAMESPACE)
            .unwrap_or_else(|| choose_prefix(&location.scope, b"mc"));
        let compatibility_is_bound = location
            .scope
            .namespace(&compatibility_prefix)
            .is_some_and(|namespace| namespace == validation::MC_NAMESPACE);
        if !compatibility_is_bound {
            append_attribute(
                &mut insertion,
                &qualified_name(b"xmlns", &compatibility_prefix),
                validation::MC_NAMESPACE,
            );
        }
        let ignorable = render_ignorable(&location.scope.ignorable, &extension_prefix);
        append_attribute(
            &mut insertion,
            &qualified_name(&compatibility_prefix, b"Ignorable"),
            ignorable.as_bytes(),
        );
    }

    append_attribute(
        &mut insertion,
        &extension_name,
        on_off_lexical(value).as_bytes(),
    );
    let tag = xml
        .get(location.start..location.tag_end)
        .ok_or_else(|| invalid("abstractNum start tag range is outside numbering XML"))?;
    let offset = attribute_insert_offset(tag)
        .ok_or_else(|| invalid("abstractNum start tag has no attribute insertion point"))?;
    edits.push(Edit {
        start: location.start + offset,
        end: location.start + offset,
        replacement: insertion,
    });
    Ok(())
}

fn append_attribute(output: &mut Vec<u8>, name: &[u8], value: &[u8]) {
    output.push(b' ');
    output.extend_from_slice(name);
    output.extend_from_slice(b"=\"");
    output.extend_from_slice(value);
    output.push(b'\"');
}

fn qualified_name(prefix: &[u8], local: &[u8]) -> Vec<u8> {
    let mut name = Vec::with_capacity(prefix.len() + 1 + local.len());
    name.extend_from_slice(prefix);
    name.push(b':');
    name.extend_from_slice(local);
    name
}

fn on_off_lexical(value: bool) -> &'static str {
    if value { "true" } else { "false" }
}

fn append_ignorable_token(value: &str, prefix: &[u8]) -> Vec<u8> {
    let prefix = String::from_utf8_lossy(prefix);
    if value.trim().is_empty() {
        prefix.into_owned().into_bytes()
    } else {
        format!("{value} {prefix}").into_bytes()
    }
}

fn render_ignorable(value: &[Vec<u8>], added: &[u8]) -> String {
    let mut output = value
        .iter()
        .map(|prefix| String::from_utf8_lossy(prefix).into_owned())
        .collect::<Vec<_>>();
    if !validation::has_ignorable_prefix(value, added) {
        output.push(String::from_utf8_lossy(added).into_owned());
    }
    output.join(" ")
}

fn choose_prefix(scope: &Scope, preferred: &[u8]) -> Vec<u8> {
    if scope.namespace(preferred).is_none() {
        return preferred.to_vec();
    }
    for suffix in 2..=1024 {
        let mut candidate = preferred.to_vec();
        candidate.extend_from_slice(suffix.to_string().as_bytes());
        if scope.namespace(&candidate).is_none() {
            return candidate;
        }
    }
    // The input is bounded and the loop above provides ample room for a
    // valid NCName. This branch is retained as a checked failure boundary.
    preferred.to_vec()
}

fn attribute_insert_offset(tag: &[u8]) -> Option<usize> {
    let close = tag.iter().rposition(|byte| *byte == b'>')?;
    let mut before_close = close;
    while before_close > 0 && tag[before_close - 1].is_ascii_whitespace() {
        before_close -= 1;
    }
    if before_close > 0 && tag[before_close - 1] == b'/' {
        Some(before_close - 1)
    } else {
        Some(close)
    }
}

fn locate_definitions(xml: &[u8]) -> Result<Layout> {
    validation::validate_xml(xml)?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut layout = Layout::default();
    let mut scopes = Vec::<Scope>::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;

    loop {
        let event_start = position(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let event_end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let decoder = reader.decoder();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("numbering XML element counter overflow"))?;
                if nodes > validation::MAX_XML_NODES {
                    return Err(invalid(&format!(
                        "numbering XML exceeds {} elements",
                        validation::MAX_XML_NODES
                    )));
                }
                depth = depth.checked_add(1).ok_or_else(too_deep)?;
                if depth > validation::MAX_XML_DEPTH {
                    return Err(invalid(&format!(
                        "numbering XML nesting exceeds {}",
                        validation::MAX_XML_DEPTH
                    )));
                }
                let tag = xml
                    .get(event_start..event_end)
                    .ok_or_else(|| invalid("numbering element range is outside its source"))?;
                let scope = scope_for(scopes.last(), &element, tag, decoder, &resolver)?;
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"abstractNum"
                {
                    layout.definitions.push(definition_location(
                        xml,
                        event_start,
                        event_end,
                        &element,
                        decoder,
                        &resolver,
                        &scope,
                    )?);
                }
                scopes.push(scope);
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("numbering XML element counter overflow"))?;
                if nodes > validation::MAX_XML_NODES {
                    return Err(invalid(&format!(
                        "numbering XML exceeds {} elements",
                        validation::MAX_XML_NODES
                    )));
                }
                let child_depth = depth.checked_add(1).ok_or_else(too_deep)?;
                if child_depth > validation::MAX_XML_DEPTH {
                    return Err(invalid(&format!(
                        "numbering XML nesting exceeds {}",
                        validation::MAX_XML_DEPTH
                    )));
                }
                let tag = xml
                    .get(event_start..event_end)
                    .ok_or_else(|| invalid("numbering element range is outside its source"))?;
                let parent = scopes.last();
                let scope = scope_for(parent, &element, tag, decoder, &resolver)?;
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"abstractNum"
                {
                    layout.definitions.push(definition_location(
                        xml,
                        event_start,
                        event_end,
                        &element,
                        decoder,
                        &resolver,
                        &scope,
                    )?);
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid numbering XML nesting"))?;
                scopes
                    .pop()
                    .ok_or_else(|| invalid("numbering XML scope stack underflow"))?;
            },
            Event::Eof if depth != 0 || !scopes.is_empty() => {
                return Err(invalid("unterminated numbering XML"));
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
    Ok(layout)
}

fn scope_for(
    inherited: Option<&Scope>,
    element: &BytesStart<'_>,
    tag: &[u8],
    decoder: Decoder,
    _resolver: &NamespaceResolver,
) -> Result<Scope> {
    let mut scope = inherited.cloned().unwrap_or_default();
    let raw_attributes = scan_attributes(tag)?;
    for attribute in &raw_attributes {
        let prefix = if attribute.name == b"xmlns" {
            Some(Vec::new())
        } else {
            attribute.name.strip_prefix(b"xmlns:").map(<[u8]>::to_vec)
        };
        let Some(prefix) = prefix else {
            continue;
        };
        let value = decoded_attribute_value(element, &attribute.name, decoder)?;
        scope.set_namespace(prefix, value.into_bytes());
    }

    let mut direct_ignorable = None;
    for attribute in &raw_attributes {
        let Some((prefix, local)) = split_qualified_name(&attribute.name) else {
            continue;
        };
        if local != b"Ignorable" || scope.namespace(prefix) != Some(validation::MC_NAMESPACE) {
            continue;
        }
        if direct_ignorable.is_some() {
            return Err(invalid("duplicate numbering mc:Ignorable attributes"));
        }
        let value = decoded_attribute_value(element, &attribute.name, decoder)?;
        direct_ignorable = Some(validation::parse_ignorable(&value)?);
    }
    if let Some(value) = direct_ignorable {
        scope.ignorable = value;
    }
    Ok(scope)
}

fn definition_location(
    xml: &[u8],
    start: usize,
    tag_end: usize,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    scope: &Scope,
) -> Result<DefinitionLocation> {
    let id = required_u32(element, b"abstractNumId", decoder, resolver)?;
    let tag = xml
        .get(start..tag_end)
        .ok_or_else(|| invalid("abstractNum range is outside numbering XML"))?;
    let raw_attributes = scan_attributes(tag)?;
    let mut extension = None;
    for attribute in &raw_attributes {
        let Some((prefix, local)) = split_qualified_name(&attribute.name) else {
            continue;
        };
        if local != b"restartNumberingAfterBreak" {
            continue;
        }
        if scope.namespace(prefix) != Some(validation::WORD_2012_NAMESPACE) {
            continue;
        }
        if !validation::has_ignorable_prefix(&scope.ignorable, prefix) {
            return Err(invalid(
                "restartNumberingAfterBreak namespace is not listed in numbering mc:Ignorable",
            ));
        }
        if extension.is_some() {
            return Err(invalid("duplicate restartNumberingAfterBreak attributes"));
        }
        extension = Some(AttributeRange {
            name_start: start + attribute.name_start,
            value_start: start + attribute.value_start,
            value_end: start + attribute.value_end,
        });
        let value = xml
            .get(start + attribute.value_start..start + attribute.value_end)
            .ok_or_else(|| invalid("restart attribute value range is outside numbering XML"))?;
        let value = std::str::from_utf8(value)
            .map_err(|error| invalid(&format!("restart attribute is not UTF-8: {error}")))?;
        validation::parse_on_off(value)?;
    }
    let ignorable = raw_attributes.iter().find_map(|attribute| {
        let (prefix, local) = split_qualified_name(&attribute.name)?;
        if local == b"Ignorable" && scope.namespace(prefix) == Some(validation::MC_NAMESPACE) {
            Some(AttributeRange {
                name_start: start + attribute.name_start,
                value_start: start + attribute.value_start,
                value_end: start + attribute.value_end,
            })
        } else {
            None
        }
    });
    Ok(DefinitionLocation {
        id,
        start,
        tag_end,
        extension,
        ignorable,
        scope: scope.clone(),
    })
}

fn decoded_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<String> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() != name {
            continue;
        }
        return attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map(Cow::into_owned)
            .map_err(|error| Error::Xml(error.to_string()));
    }
    Err(invalid("numbering namespace attribute value is missing"))
}

fn split_qualified_name(name: &[u8]) -> Option<(&[u8], &[u8])> {
    let index = name.iter().rposition(|byte| *byte == b':')?;
    let (prefix, local) = name.split_at(index);
    let local = local.get(1..)?;
    (!prefix.is_empty() && !local.is_empty()).then_some((prefix, local))
}

fn scan_attributes(tag: &[u8]) -> Result<Vec<RawAttribute>> {
    let mut index = 0usize;
    if tag.first() != Some(&b'<') {
        return Err(invalid("numbering start tag does not begin with '<'"));
    }
    index += 1;
    skip_name(tag, &mut index)?;
    let mut attributes = Vec::new();
    loop {
        while tag.get(index).is_some_and(u8::is_ascii_whitespace) {
            index += 1;
        }
        if index >= tag.len() || tag[index] == b'>' {
            break;
        }
        if tag[index] == b'/' {
            index += 1;
            while tag.get(index).is_some_and(u8::is_ascii_whitespace) {
                index += 1;
            }
            if tag.get(index) == Some(&b'>') {
                break;
            }
            return Err(invalid("numbering empty start tag has invalid close"));
        }
        let name_start = index;
        skip_name(tag, &mut index)?;
        let name = tag[name_start..index].to_vec();
        while tag.get(index).is_some_and(u8::is_ascii_whitespace) {
            index += 1;
        }
        if tag.get(index) != Some(&b'=') {
            return Err(invalid("numbering attribute has no value"));
        }
        index += 1;
        while tag.get(index).is_some_and(u8::is_ascii_whitespace) {
            index += 1;
        }
        let quote = *tag
            .get(index)
            .ok_or_else(|| invalid("numbering attribute value is truncated"))?;
        if !matches!(quote, b'\'' | b'\"') {
            return Err(invalid("numbering attribute value is not quoted"));
        }
        index += 1;
        let value_start = index;
        while tag.get(index).is_some_and(|byte| *byte != quote) {
            index += 1;
        }
        let value_end = index;
        if tag.get(index) != Some(&quote) {
            return Err(invalid("numbering attribute value is unterminated"));
        }
        index += 1;
        attributes.push(RawAttribute {
            name,
            name_start,
            value_start,
            value_end,
        });
    }
    Ok(attributes)
}

fn skip_name(tag: &[u8], index: &mut usize) -> Result<()> {
    let start = *index;
    while tag
        .get(*index)
        .is_some_and(|byte| !byte.is_ascii_whitespace() && !matches!(*byte, b'=' | b'/' | b'>'))
    {
        *index += 1;
    }
    if *index == start {
        return Err(invalid("numbering XML name is missing"));
    }
    Ok(())
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_source_error| invalid("numbering XML offset does not fit usize"))
}

/// Parse a standalone `WordprocessingML` numbering payload.
///
/// # Errors
///
/// Returns an error if the operation cannot be completed.
pub fn parse_numbering(xml: &[u8]) -> Result<Collection> {
    Collection::from_xml(xml)
}

fn is_wordprocessing_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == TRANSITIONAL_WORDPROCESSING_NAMESPACE
                || *value == STRICT_WORDPROCESSING_NAMESPACE
    )
}

fn is_word_2012_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == validation::WORD_2012_NAMESPACE
    )
}

fn is_markup_compatibility_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == validation::MC_NAMESPACE
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
            return Err(Error::Invalid(format!(
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
