//! Namespace-aware XML codec for WordprocessingML numbering.
//!
//! OPC parts and markup-compatibility preprocessing stay in the host crate;
//! this layer consumes the resulting XML bytes only.

use super::model::{
    Collection, Definition, Instance, Level, Override, PictureBullet, Restart, Suffix,
};
use crate::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

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
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = NsReader::from_reader(xml);
        let mut result = Self::new();
        let mut abstract_num: Option<PendingDefinition> = None;
        let mut num: Option<PendingInstance> = None;
        let mut level_override: Option<PendingOverride> = None;
        let mut level: Option<PendingLevel> = None;
        let mut picture_bullet: Option<PendingPictureBullet> = None;
        let mut depth = 0usize;

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
                    depth = depth.checked_add(1).ok_or_else(too_deep)?;
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
                    )?;
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(too_deep)?;
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
                },
                Event::Eof
                    if depth != 0
                        || abstract_num.is_some()
                        || num.is_some()
                        || level_override.is_some()
                        || level.is_some()
                        || picture_bullet.is_some() =>
                {
                    return Err(invalid("unterminated numbering XML"));
                },
                Event::Eof => break,
                _ => {},
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

#[allow(clippy::too_many_arguments)]
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

#[allow(clippy::too_many_arguments)]
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

#[allow(clippy::too_many_arguments)]
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
                value.value.format = raw
                    .parse()
                    .map_err(|_| invalid(&format!("invalid numbering format '{raw}'")))?;
                value.value.custom_format =
                    word_attribute_value(element, b"format", decoder, resolver)?;
            },
            b"lvlText" => {
                value.value.level_text = Some(
                    word_attribute_value(element, b"val", decoder, resolver)?.unwrap_or_default(),
                )
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
                    Some(required_string(element, b"val", decoder, resolver)?)
            },
            b"lvlPicBulletId" => {
                value.value.picture_bullet_id =
                    Some(required_u32(element, b"val", decoder, resolver)?)
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
                value.value.num_type = Some(
                    raw.parse()
                        .map_err(|_| invalid(&format!("invalid multiLevelType '{raw}'")))?,
                );
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
/// DrawingML (`a:blip r:embed`/`a:link`); everything else inside `w:pict` is
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
        _ => return Ok(()),
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
        .map_err(|_| invalid(&format!("invalid Word numbering integer '{value}'")))
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
        .map_err(|_| invalid(&format!("invalid Word numbering integer '{value}'")))
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

/// Parse a standalone WordprocessingML numbering payload.
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
