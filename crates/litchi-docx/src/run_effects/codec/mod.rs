//! XML codec for the Word 2010 `w:rPr` effect children.

use std::fmt::Write as FmtWrite;

use crate::error::{Error, Result};
use litchi_core::xml::escape_xml;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{
    Bevel, BevelPreset, Camera, Color, ColorTransform, CompoundLine, Effect, EffectKind, Fill,
    Glow, Gradient, GradientStop, LightRig, LightRigDirection, LightRigType, LineCap, LineDash,
    LineJoin, PathKind, PenAlignment, PresetCamera, PresetMaterial, Props3d, RectAlignment,
    Reflection, RelativeRect, RgbColor, Scene3d, SchemeColor, SchemeColorValue, Shade, Shadow,
    SphereCoords, TextFill, TextOutline,
};

/// Word 2010 run-effect namespace.
pub const NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/word/2010/wordml";
const MAX_XML_BYTES: usize = 4 * 1024 * 1024;
const MAX_DEPTH: usize = 64;
const MAX_NODES: usize = 8_192;

/// Parse a complete `w:r` or `w:rPr` fragment.
pub fn parse(xml: &[u8]) -> Result<super::model::Effects> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "Word run effects XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut rpr_depth = None::<usize>;
    let mut stack = Vec::<Frame>::new();
    let mut effects = super::model::Effects::new();

    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("Word run effects XML offset overflow".into()))?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let is_w14 = is_word2010(&namespace);
        let event = event.into_owned();
        match event {
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("Word run XML node counter overflow".into()))?;
                if nodes > MAX_NODES {
                    return Err(Error::Invalid(format!(
                        "Word run effects exceed {MAX_NODES} XML nodes"
                    )));
                }
                let local = element.local_name().as_ref().to_vec();
                if !root_seen {
                    root_seen = true;
                    if local.as_slice() != b"r" && local.as_slice() != b"rPr" {
                        return Err(Error::InvalidFormat(
                            "Word run effects XML must have r or rPr root".into(),
                        ));
                    }
                }
                if depth >= MAX_DEPTH {
                    return Err(Error::InvalidFormat(
                        "Word run effects XML is nested too deeply".into(),
                    ));
                }
                let direct = rpr_depth == Some(depth);
                let next_depth = depth + 1;
                if rpr_depth.is_none() && local.as_slice() == b"rPr" {
                    rpr_depth = Some(next_depth);
                }
                stack.push(Frame {
                    start,
                    depth: next_depth,
                    local,
                    is_w14,
                    direct,
                });
                depth = next_depth;
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("Word run XML node counter overflow".into()))?;
                if nodes > MAX_NODES {
                    return Err(Error::Invalid(format!(
                        "Word run effects exceed {MAX_NODES} XML nodes"
                    )));
                }
                let local = element.local_name().as_ref().to_vec();
                if !root_seen {
                    root_seen = true;
                    if local.as_slice() != b"r" && local.as_slice() != b"rPr" {
                        return Err(Error::InvalidFormat(
                            "Word run effects XML must have r or rPr root".into(),
                        ));
                    }
                } else if rpr_depth == Some(depth) {
                    let end = usize::try_from(reader.buffer_position()).map_err(|_| {
                        Error::InvalidFormat("Word run effects XML offset overflow".into())
                    })?;
                    consume_direct(&mut effects, &local, is_w14, &xml[start..end])?;
                }
            },
            Event::End(element) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| Error::InvalidFormat("invalid Word run XML nesting".into()))?;
                if frame.depth != depth || frame.local != element.local_name().as_ref() {
                    return Err(Error::InvalidFormat("mismatched Word run XML end".into()));
                }
                depth -= 1;
                let end = usize::try_from(reader.buffer_position()).map_err(|_| {
                    Error::InvalidFormat("Word run effects XML offset overflow".into())
                })?;
                if frame.direct {
                    consume_direct(
                        &mut effects,
                        &frame.local,
                        frame.is_w14,
                        &xml[frame.start..end],
                    )?;
                }
                if rpr_depth == Some(frame.depth) && frame.local.as_slice() == b"rPr" {
                    rpr_depth = None;
                }
            },
            Event::Text(text) if rpr_depth.is_some() => {
                if !text.into_inner().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "Word run effects XML cannot contain non-whitespace text".into(),
                    ));
                }
            },
            Event::CData(text) if rpr_depth.is_some() => {
                if !text.into_inner().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "Word run effects XML cannot contain non-whitespace text".into(),
                    ));
                }
            },
            Event::Eof => break,
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "Word run effects XML cannot contain declarations or processing instructions"
                        .into(),
                ));
            },
            _ => {},
        }
    }
    if !root_seen || depth != 0 || !stack.is_empty() {
        return Err(Error::InvalidFormat(
            "Word run effects XML is not one complete element".into(),
        ));
    }
    effects.validate()?;
    Ok(effects)
}

/// Write direct effects as `w14:*` children of an existing `w:rPr` element.
pub(crate) fn write(value: &super::model::Effects, output: &mut String) -> Result<()> {
    value.validate()?;
    for effect in value.iter() {
        match effect {
            Effect::Glow(value) => write_glow(value, output)?,
            Effect::Shadow(value) => write_shadow(value, output)?,
            Effect::Reflection(value) => write_reflection(value, output)?,
            Effect::TextOutline(value) => write_text_outline(value, output)?,
            Effect::TextFill(value) => write_text_fill(value, output)?,
            Effect::Scene3d(value) => write_scene3d(value, output)?,
            Effect::Props3d(value) => write_props3d(value, output)?,
            Effect::Unknown(value) => {
                output.push_str(std::str::from_utf8(value.as_bytes()).map_err(|error| {
                    Error::InvalidFormat(format!("opaque run effect is not UTF-8: {error}"))
                })?)
            },
        }
    }
    Ok(())
}

#[derive(Debug)]
struct Frame {
    start: usize,
    depth: usize,
    local: Vec<u8>,
    is_w14: bool,
    direct: bool,
}

fn is_word2010(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == NAMESPACE)
        || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"w14")
}

fn consume_direct(
    value: &mut super::model::Effects,
    local: &[u8],
    is_w14: bool,
    xml: &[u8],
) -> Result<()> {
    let kind = match (is_w14, local) {
        (true, b"glow") => Some(EffectKind::Glow),
        (true, b"shadow") => Some(EffectKind::Shadow),
        (true, b"reflection") => Some(EffectKind::Reflection),
        (true, b"textOutline") => Some(EffectKind::TextOutline),
        (true, b"textFill") => Some(EffectKind::TextFill),
        (true, b"scene3d") => Some(EffectKind::Scene3d),
        (true, b"props3d") => Some(EffectKind::Props3d),
        _ => None,
    };
    let effect = if let Some(kind) = kind {
        parse_known(kind, xml)?
    } else {
        Effect::Unknown(super::model::OpaqueExtension::new(xml.to_vec())?)
    };
    value.push(effect)?;
    Ok(())
}

#[derive(Debug, Clone)]
struct Node {
    local: String,
    attrs: Vec<(String, String, bool)>,
    children: Vec<Node>,
}

fn parse_node(xml: &[u8]) -> Result<Node> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut stack = Vec::<Node>::new();
    let mut root = None;
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
                require_word2010(&namespace, element.local_name().as_ref())?;
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(
                        "run effect XML is nested too deeply".into(),
                    ));
                }
                stack.push(node_from_start(&element, decoder, &resolver)?);
            },
            Event::Empty(element) => {
                require_word2010(&namespace, element.local_name().as_ref())?;
                let node = node_from_start(&element, decoder, &resolver)?;
                attach_node(&mut stack, &mut root, node)?;
            },
            Event::End(_) => {
                let node = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("run effect XML has an unexpected end".into())
                })?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("run effect XML nesting underflow".into())
                })?;
                attach_node(&mut stack, &mut root, node)?;
            },
            Event::Text(text) => {
                if !text.into_inner().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "run effect XML has unexpected text".into(),
                    ));
                }
            },
            Event::CData(text) => {
                if !text.into_inner().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "run effect XML has unexpected text".into(),
                    ));
                }
            },
            Event::Eof => break,
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "run effect XML has forbidden prolog content".into(),
                ));
            },
            _ => {},
        }
    }
    if !stack.is_empty() || depth != 0 {
        return Err(Error::InvalidFormat(
            "run effect XML is unterminated".into(),
        ));
    }
    root.ok_or_else(|| Error::InvalidFormat("run effect XML has no root".into()))
}

fn node_from_start(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<Node> {
    let mut attrs = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        let local = String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
        let is_namespace = key == b"xmlns" || key.starts_with(b"xmlns:");
        if !is_namespace {
            let (namespace, _) = resolver.resolve_attribute(attribute.key);
            require_word2010(&namespace, attribute.key.local_name().as_ref())?;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        attrs.push((local, value, is_namespace));
    }
    Ok(Node {
        local: String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
        attrs,
        children: Vec::new(),
    })
}

fn require_word2010(namespace: &ResolveResult<'_>, local: &[u8]) -> Result<()> {
    if is_word2010(namespace) {
        return Ok(());
    }
    Err(Error::InvalidFormat(format!(
        "run effect element or attribute '{}' is not in the Word 2010 namespace",
        String::from_utf8_lossy(local)
    )))
}

fn attach_node(stack: &mut [Node], root: &mut Option<Node>, node: Node) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(Error::InvalidFormat(
            "run effect XML has multiple roots".into(),
        ));
    }
    Ok(())
}

fn parse_known(kind: EffectKind, xml: &[u8]) -> Result<Effect> {
    let node = parse_node(xml)?;
    let effect = match kind {
        EffectKind::Glow => Effect::Glow(parse_glow(&node)?),
        EffectKind::Shadow => Effect::Shadow(parse_shadow(&node)?),
        EffectKind::Reflection => Effect::Reflection(parse_reflection(&node)?),
        EffectKind::TextOutline => Effect::TextOutline(parse_text_outline(&node)?),
        EffectKind::TextFill => Effect::TextFill(parse_text_fill(&node)?),
        EffectKind::Scene3d => Effect::Scene3d(parse_scene3d(&node)?),
        EffectKind::Props3d => Effect::Props3d(parse_props3d(&node)?),
        EffectKind::Unknown => {
            return Err(Error::Invalid(
                "opaque run effects must use the unknown-child path".into(),
            ));
        },
    };
    Ok(effect)
}

fn parse_glow(node: &Node) -> Result<Glow> {
    expect(node, "glow")?;
    let color = parse_one_color(node)?;
    Ok(Glow {
        color: Some(color),
        radius: attr_u64(node, "rad")?,
    })
}

fn parse_shadow(node: &Node) -> Result<Shadow> {
    expect(node, "shadow")?;
    Ok(Shadow {
        color: Some(parse_one_color(node)?),
        blur_radius: attr_u64(node, "blurRad")?,
        distance: attr_u64(node, "dist")?,
        direction: attr_u32(node, "dir")?,
        scale_x: attr_i32(node, "sx")?,
        scale_y: attr_i32(node, "sy")?,
        skew_x: attr_i32(node, "kx")?,
        skew_y: attr_i32(node, "ky")?,
        alignment: attr_token(node, "algn")?.map(parse_alignment).transpose()?,
    })
}

fn parse_reflection(node: &Node) -> Result<Reflection> {
    expect(node, "reflection")?;
    reject_children(node)?;
    Ok(Reflection {
        blur_radius: attr_u64(node, "blurRad")?,
        start_alpha: attr_u32(node, "stA")?,
        start_position: attr_u32(node, "stPos")?,
        end_alpha: attr_u32(node, "endA")?,
        end_position: attr_u32(node, "endPos")?,
        distance: attr_u64(node, "dist")?,
        direction: attr_u32(node, "dir")?,
        fade_direction: attr_u32(node, "fadeDir")?,
        scale_x: attr_i32(node, "sx")?,
        scale_y: attr_i32(node, "sy")?,
        skew_x: attr_i32(node, "kx")?,
        skew_y: attr_i32(node, "ky")?,
        alignment: attr_token(node, "algn")?.map(parse_alignment).transpose()?,
    })
}

fn parse_text_fill(node: &Node) -> Result<TextFill> {
    expect(node, "textFill")?;
    Ok(TextFill {
        fill: parse_fill(node)?,
    })
}

fn parse_text_outline(node: &Node) -> Result<TextOutline> {
    expect(node, "textOutline")?;
    let mut fill = None;
    let mut dash = None;
    let mut join = None;
    for child in &node.children {
        match child.local.as_str() {
            "noFill" | "solidFill" | "gradFill" if fill.is_none() => {
                fill = Some(parse_fill_node(child)?);
            },
            "noFill" | "solidFill" | "gradFill" => duplicate("text outline fill")?,
            "prstDash" if dash.is_none() => {
                dash = Some(parse_dash(child)?);
            },
            "prstDash" => duplicate("text outline dash")?,
            "round" if join.is_none() => join = Some(LineJoin::Round),
            "bevel" if join.is_none() => join = Some(LineJoin::Bevel),
            "miter" if join.is_none() => {
                join = Some(LineJoin::Miter {
                    limit: attr_u32(child, "lim")?,
                });
            },
            "round" | "bevel" | "miter" => duplicate("text outline join")?,
            _ => return Err(unknown_child(&node.local, &child.local)),
        }
    }
    Ok(TextOutline {
        fill,
        dash,
        join,
        width: attr_u64(node, "w")?,
        cap: attr_token(node, "cap")?.map(parse_cap).transpose()?,
        compound: attr_token(node, "cmpd")?.map(parse_compound).transpose()?,
        alignment: attr_token(node, "algn")?
            .map(parse_pen_alignment)
            .transpose()?,
    })
}

fn parse_scene3d(node: &Node) -> Result<Scene3d> {
    expect(node, "scene3d")?;
    if node.children.len() != 2
        || node.children[0].local != "camera"
        || node.children[1].local != "lightRig"
    {
        return Err(Error::InvalidFormat(
            "scene3d requires camera followed by lightRig".into(),
        ));
    }
    let camera = Camera {
        preset: PresetCamera::parse(attr_required(&node.children[0], "prst")?)?,
    };
    let light = &node.children[1];
    let rotation = match light.children.as_slice() {
        [] => None,
        [child] if child.local == "rot" => Some(parse_sphere(child)?),
        _ => return Err(Error::InvalidFormat("lightRig has invalid children".into())),
    };
    Ok(Scene3d {
        camera,
        light_rig: LightRig {
            rig: LightRigType::parse(attr_required(light, "rig")?)?,
            direction: LightRigDirection::parse(attr_required(light, "dir")?)?,
            rotation,
        },
    })
}

fn parse_props3d(node: &Node) -> Result<Props3d> {
    expect(node, "props3d")?;
    let mut value = Props3d {
        extrusion_height: attr_u64(node, "extrusionH")?,
        contour_width: attr_u64(node, "contourW")?,
        material: attr_token(node, "prstMaterial")?
            .map(PresetMaterial::parse)
            .transpose()?,
        ..Props3d::default()
    };
    for child in &node.children {
        match child.local.as_str() {
            "bevelT" if value.bevel_top.is_none() => value.bevel_top = Some(parse_bevel(child)?),
            "bevelB" if value.bevel_bottom.is_none() => {
                value.bevel_bottom = Some(parse_bevel(child)?);
            },
            "extrusionClr" if value.extrusion_color.is_none() => {
                value.extrusion_color = Some(parse_one_color(child)?);
            },
            "contourClr" if value.contour_color.is_none() => {
                value.contour_color = Some(parse_one_color(child)?);
            },
            "bevelT" | "bevelB" | "extrusionClr" | "contourClr" => duplicate("props3d child")?,
            _ => return Err(unknown_child(&node.local, &child.local)),
        }
    }
    Ok(value)
}

fn parse_bevel(node: &Node) -> Result<Bevel> {
    Ok(Bevel {
        width: attr_u64(node, "w")?,
        height: attr_u64(node, "h")?,
        preset: attr_token(node, "prst")?
            .map(BevelPreset::parse)
            .transpose()?,
    })
}

fn parse_sphere(node: &Node) -> Result<SphereCoords> {
    Ok(SphereCoords {
        latitude: attr_required(node, "lat")?.parse().map_err(number("lat"))?,
        longitude: attr_required(node, "lon")?.parse().map_err(number("lon"))?,
        revolution: attr_required(node, "rev")?.parse().map_err(number("rev"))?,
    })
}

fn parse_one_color(node: &Node) -> Result<Color> {
    let colors: Vec<_> = node
        .children
        .iter()
        .filter(|child| child.local == "srgbClr" || child.local == "schemeClr")
        .collect();
    if colors.len() != 1 || node.children.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "{} requires exactly one color child",
            node.local
        )));
    }
    parse_color(colors[0])
}

fn parse_color(node: &Node) -> Result<Color> {
    let transforms = node
        .children
        .iter()
        .map(parse_transform)
        .collect::<Result<Vec<_>>>()?;
    if node.local == "srgbClr" {
        let value = attr_required(node, "val")?;
        if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(Error::InvalidFormat(format!("invalid RGB color '{value}'")));
        }
        let mut rgb = [0; 3];
        for (index, slot) in rgb.iter_mut().enumerate() {
            *slot = u8::from_str_radix(&value[index * 2..index * 2 + 2], 16)
                .map_err(number("RGB component"))?;
        }
        Ok(Color::Rgb(RgbColor {
            value: rgb,
            transforms,
        }))
    } else if node.local == "schemeClr" {
        let value = SchemeColorValue::from_xml(attr_required(node, "val")?)
            .ok_or_else(|| Error::InvalidFormat("invalid scheme color value".into()))?;
        Ok(Color::Scheme(SchemeColor { value, transforms }))
    } else {
        Err(unknown_child("color", &node.local))
    }
}

fn parse_transform(node: &Node) -> Result<ColorTransform> {
    let value = attr_required(node, "val")?;
    Ok(match node.local.as_str() {
        "tint" => ColorTransform::Tint(value.parse().map_err(number("tint"))?),
        "shade" => ColorTransform::Shade(value.parse().map_err(number("shade"))?),
        "alpha" => ColorTransform::Alpha(value.parse().map_err(number("alpha"))?),
        "hueMod" => ColorTransform::HueMod(value.parse().map_err(number("hueMod"))?),
        "sat" => ColorTransform::Saturation(value.parse().map_err(number("sat"))?),
        "satOff" => ColorTransform::SaturationOffset(value.parse().map_err(number("satOff"))?),
        "satMod" => ColorTransform::SaturationMod(value.parse().map_err(number("satMod"))?),
        "lum" => ColorTransform::Luminance(value.parse().map_err(number("lum"))?),
        "lumOff" => ColorTransform::LuminanceOffset(value.parse().map_err(number("lumOff"))?),
        "lumMod" => ColorTransform::LuminanceMod(value.parse().map_err(number("lumMod"))?),
        _ => return Err(unknown_child("color", &node.local)),
    })
}

fn parse_fill(node: &Node) -> Result<Option<Fill>> {
    if node.children.is_empty() {
        return Ok(None);
    }
    if node.children.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "{} has multiple fill children",
            node.local
        )));
    }
    Ok(Some(parse_fill_node(&node.children[0])?))
}

fn parse_fill_node(node: &Node) -> Result<Fill> {
    Ok(match node.local.as_str() {
        "noFill" => {
            reject_children(node)?;
            Fill::NoFill
        },
        "solidFill" => {
            let color = if node.children.is_empty() {
                None
            } else {
                Some(parse_one_color(node)?)
            };
            Fill::Solid(color)
        },
        "gradFill" => Fill::Gradient(parse_gradient(node)?),
        _ => return Err(unknown_child("fill", &node.local)),
    })
}

fn parse_gradient(node: &Node) -> Result<Gradient> {
    let stop_list = node.children.iter().find(|child| child.local == "gsLst");
    let stops = if let Some(list) = stop_list {
        if list.children.len() < 2 || list.children.len() > 10 {
            return Err(Error::InvalidFormat("gradient needs 2..=10 stops".into()));
        }
        list.children
            .iter()
            .map(|stop| {
                if stop.local != "gs" {
                    return Err(unknown_child("gsLst", &stop.local));
                }
                Ok(GradientStop {
                    position: attr_required(stop, "pos")?.parse().map_err(number("pos"))?,
                    color: parse_one_color(stop)?,
                })
            })
            .collect::<Result<Vec<_>>>()?
    } else {
        Vec::new()
    };
    let shade = node
        .children
        .iter()
        .find(|child| child.local == "lin" || child.local == "path");
    let shade = shade.map(parse_shade).transpose()?;
    Ok(Gradient { stops, shade })
}

fn parse_shade(node: &Node) -> Result<Shade> {
    if node.local == "lin" {
        reject_children(node)?;
        return Ok(Shade::Linear {
            angle: attr_u32(node, "ang")?,
            scaled: attr_token(node, "scaled")?.map(parse_on_off).transpose()?,
        });
    }
    Ok(Shade::Path {
        path: attr_token(node, "path")?.map(parse_path).transpose()?,
        fill_to: node
            .children
            .iter()
            .find(|child| child.local == "fillToRect")
            .map(parse_rect)
            .transpose()?,
    })
}

fn parse_rect(node: &Node) -> Result<RelativeRect> {
    Ok(RelativeRect {
        left: attr_i32(node, "l")?,
        top: attr_i32(node, "t")?,
        right: attr_i32(node, "r")?,
        bottom: attr_i32(node, "b")?,
    })
}

fn parse_dash(node: &Node) -> Result<LineDash> {
    match attr_token(node, "val")? {
        Some(value) => parse_dash_token(value),
        None => Ok(LineDash::Solid),
    }
}

fn parse_dash_token(value: &str) -> Result<LineDash> {
    Ok(match value {
        "solid" => LineDash::Solid,
        "dot" => LineDash::Dot,
        "sysDot" => LineDash::SysDot,
        "dash" => LineDash::Dash,
        "sysDash" => LineDash::SysDash,
        "lgDash" => LineDash::LargeDash,
        "dashDot" => LineDash::DashDot,
        "sysDashDot" => LineDash::SysDashDot,
        "lgDashDot" => LineDash::LargeDashDot,
        "lgDashDotDot" => LineDash::LargeDashDotDot,
        "sysDashDotDot" => LineDash::SysDashDotDot,
        _ => {
            return Err(Error::InvalidFormat(format!(
                "invalid outline dash '{value}'"
            )));
        },
    })
}

fn parse_path(value: &str) -> Result<PathKind> {
    Ok(match value {
        "shape" => PathKind::Shape,
        "circle" => PathKind::Circle,
        "rect" => PathKind::Rect,
        _ => {
            return Err(Error::InvalidFormat(format!(
                "invalid gradient path '{value}'"
            )));
        },
    })
}

fn parse_on_off(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid on/off value '{value}'"
        ))),
    }
}

fn parse_alignment(value: &str) -> Result<RectAlignment> {
    Ok(match value {
        "none" => RectAlignment::None,
        "tl" => RectAlignment::TopLeft,
        "t" => RectAlignment::Top,
        "tr" => RectAlignment::TopRight,
        "l" => RectAlignment::Left,
        "ctr" => RectAlignment::Center,
        "r" => RectAlignment::Right,
        "bl" => RectAlignment::BottomLeft,
        "b" => RectAlignment::Bottom,
        "br" => RectAlignment::BottomRight,
        _ => {
            return Err(Error::InvalidFormat(format!(
                "invalid rectangle alignment '{value}'"
            )));
        },
    })
}

fn parse_cap(value: &str) -> Result<LineCap> {
    Ok(match value {
        "flat" => LineCap::Flat,
        "rnd" => LineCap::Round,
        "sq" => LineCap::Square,
        _ => return Err(Error::InvalidFormat(format!("invalid line cap '{value}'"))),
    })
}

fn parse_compound(value: &str) -> Result<CompoundLine> {
    Ok(match value {
        "sng" => CompoundLine::Single,
        "dbl" => CompoundLine::Double,
        "thickThin" => CompoundLine::ThickThin,
        "thinThick" => CompoundLine::ThinThick,
        "tri" => CompoundLine::Triple,
        _ => {
            return Err(Error::InvalidFormat(format!(
                "invalid compound line '{value}'"
            )));
        },
    })
}

fn parse_pen_alignment(value: &str) -> Result<PenAlignment> {
    match value {
        "ctr" => Ok(PenAlignment::Center),
        "in" => Ok(PenAlignment::Inside),
        _ => Err(Error::InvalidFormat(format!(
            "invalid pen alignment '{value}'"
        ))),
    }
}

fn expect(node: &Node, local: &str) -> Result<()> {
    if node.local == local {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "expected {local}, found {}",
            node.local
        )))
    }
}

fn reject_children(node: &Node) -> Result<()> {
    if node.children.is_empty() {
        Ok(())
    } else {
        Err(unknown_child(&node.local, &node.children[0].local))
    }
}

fn unknown_child(parent: &str, child: &str) -> Error {
    Error::InvalidFormat(format!(
        "unsupported child {parent}/{child} in typed run effect"
    ))
}

fn duplicate(name: &str) -> Result<()> {
    Err(Error::InvalidFormat(format!("duplicate {name}")))
}

fn attr_required<'a>(node: &'a Node, name: &str) -> Result<&'a str> {
    attr_token(node, name)?.ok_or_else(|| Error::InvalidFormat(format!("missing {name} attribute")))
}

fn attr_token<'a>(node: &'a Node, name: &str) -> Result<Option<&'a str>> {
    let mut value = None;
    for (local, current, is_namespace) in &node.attrs {
        if *is_namespace || local != name {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(format!("duplicate {name} attribute")));
        }
        value = Some(current.as_str());
    }
    Ok(value)
}

fn attr_u64(node: &Node, name: &str) -> Result<Option<u64>> {
    attr_token(node, name)?
        .map(|value| parse_number(value, name))
        .transpose()
}

fn attr_u32(node: &Node, name: &str) -> Result<Option<u32>> {
    attr_token(node, name)?
        .map(|value| parse_number(value, name))
        .transpose()
}

fn attr_i32(node: &Node, name: &str) -> Result<Option<i32>> {
    attr_token(node, name)?
        .map(|value| parse_number(value, name))
        .transpose()
}

fn parse_number<T: std::str::FromStr>(value: &str, name: &str) -> Result<T> {
    value
        .parse()
        .map_err(|_| Error::InvalidFormat(format!("invalid numeric {name} attribute")))
}

fn number<T: std::fmt::Display>(name: &'static str) -> impl FnOnce(T) -> Error {
    move |_| Error::InvalidFormat(format!("invalid numeric {name} attribute"))
}

fn write_glow(value: &Glow, xml: &mut String) -> Result<()> {
    xml.push_str("<w14:glow xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"");
    attr_u64_out(xml, "rad", value.radius)?;
    xml.push('>');
    write_color_children(xml, value.color.as_ref())?;
    xml.push_str("</w14:glow>");
    Ok(())
}

fn write_shadow(value: &Shadow, xml: &mut String) -> Result<()> {
    xml.push_str("<w14:shadow xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"");
    attr_u64_out(xml, "blurRad", value.blur_radius)?;
    attr_u64_out(xml, "dist", value.distance)?;
    attr_u32_out(xml, "dir", value.direction)?;
    attr_i32_out(xml, "sx", value.scale_x)?;
    attr_i32_out(xml, "sy", value.scale_y)?;
    attr_i32_out(xml, "kx", value.skew_x)?;
    attr_i32_out(xml, "ky", value.skew_y)?;
    attr_str_out(xml, "algn", value.alignment.map(alignment_str))?;
    xml.push('>');
    write_color_children(xml, value.color.as_ref())?;
    xml.push_str("</w14:shadow>");
    Ok(())
}

fn write_reflection(value: &Reflection, xml: &mut String) -> Result<()> {
    xml.push_str(
        "<w14:reflection xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"",
    );
    attr_u64_out(xml, "blurRad", value.blur_radius)?;
    attr_u32_out(xml, "stA", value.start_alpha)?;
    attr_u32_out(xml, "stPos", value.start_position)?;
    attr_u32_out(xml, "endA", value.end_alpha)?;
    attr_u32_out(xml, "endPos", value.end_position)?;
    attr_u64_out(xml, "dist", value.distance)?;
    attr_u32_out(xml, "dir", value.direction)?;
    attr_u32_out(xml, "fadeDir", value.fade_direction)?;
    attr_i32_out(xml, "sx", value.scale_x)?;
    attr_i32_out(xml, "sy", value.scale_y)?;
    attr_i32_out(xml, "kx", value.skew_x)?;
    attr_i32_out(xml, "ky", value.skew_y)?;
    attr_str_out(xml, "algn", value.alignment.map(alignment_str))?;
    xml.push_str("/>");
    Ok(())
}

fn write_text_fill(value: &TextFill, xml: &mut String) -> Result<()> {
    xml.push_str(
        "<w14:textFill xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">",
    );
    if let Some(fill) = &value.fill {
        write_fill(fill, xml)?;
    }
    xml.push_str("</w14:textFill>");
    Ok(())
}

fn write_text_outline(value: &TextOutline, xml: &mut String) -> Result<()> {
    xml.push_str(
        "<w14:textOutline xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"",
    );
    attr_u64_out(xml, "w", value.width)?;
    attr_str_out(xml, "cap", value.cap.map(cap_str))?;
    attr_str_out(xml, "cmpd", value.compound.map(compound_str))?;
    attr_str_out(xml, "algn", value.alignment.map(pen_alignment_str))?;
    if value.fill.is_none() && value.dash.is_none() && value.join.is_none() {
        xml.push_str("/>");
        return Ok(());
    }
    xml.push('>');
    if let Some(fill) = &value.fill {
        write_fill(fill, xml)?;
    }
    if let Some(dash) = value.dash {
        write!(xml, "<w14:prstDash w14:val=\"{}\"/>", dash_str(dash))?;
    }
    if let Some(join) = &value.join {
        match join {
            LineJoin::Round => xml.push_str("<w14:round/>"),
            LineJoin::Bevel => xml.push_str("<w14:bevel/>"),
            LineJoin::Miter { limit } => {
                xml.push_str("<w14:miter");
                attr_u32_out(xml, "lim", *limit)?;
                xml.push_str("/>");
            },
        }
    }
    xml.push_str("</w14:textOutline>");
    Ok(())
}

fn write_scene3d(value: &Scene3d, xml: &mut String) -> Result<()> {
    xml.push_str(
        "<w14:scene3d xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">",
    );
    write!(
        xml,
        "<w14:camera w14:prst=\"{}\"/>",
        escape_xml(value.camera.preset.as_str())
    )?;
    write!(
        xml,
        "<w14:lightRig w14:rig=\"{}\" w14:dir=\"{}\"",
        escape_xml(value.light_rig.rig.as_str()),
        escape_xml(value.light_rig.direction.as_str())
    )?;
    if let Some(rotation) = value.light_rig.rotation {
        write!(
            xml,
            "><w14:rot w14:lat=\"{}\" w14:lon=\"{}\" w14:rev=\"{}\"/></w14:lightRig>",
            rotation.latitude, rotation.longitude, rotation.revolution
        )?;
    } else {
        xml.push_str("/></w14:lightRig>");
    }
    xml.push_str("</w14:scene3d>");
    Ok(())
}

fn write_props3d(value: &Props3d, xml: &mut String) -> Result<()> {
    xml.push_str("<w14:props3d xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"");
    attr_u64_out(xml, "extrusionH", value.extrusion_height)?;
    attr_u64_out(xml, "contourW", value.contour_width)?;
    attr_str_out(
        xml,
        "prstMaterial",
        value.material.as_ref().map(PresetMaterial::as_str),
    )?;
    if value.bevel_top.is_none()
        && value.bevel_bottom.is_none()
        && value.extrusion_color.is_none()
        && value.contour_color.is_none()
    {
        xml.push_str("/>");
        return Ok(());
    }
    xml.push('>');
    if let Some(bevel) = &value.bevel_top {
        write_bevel("bevelT", bevel, xml)?;
    }
    if let Some(bevel) = &value.bevel_bottom {
        write_bevel("bevelB", bevel, xml)?;
    }
    if let Some(color) = &value.extrusion_color {
        write_color_element("extrusionClr", color, xml)?;
    }
    if let Some(color) = &value.contour_color {
        write_color_element("contourClr", color, xml)?;
    }
    xml.push_str("</w14:props3d>");
    Ok(())
}

fn write_bevel(name: &str, value: &Bevel, xml: &mut String) -> Result<()> {
    write!(xml, "<w14:{name}")?;
    attr_u64_out(xml, "w", value.width)?;
    attr_u64_out(xml, "h", value.height)?;
    attr_str_out(xml, "prst", value.preset.as_ref().map(BevelPreset::as_str))?;
    xml.push_str("/>");
    Ok(())
}

fn write_fill(value: &Fill, xml: &mut String) -> Result<()> {
    match value {
        Fill::NoFill => xml.push_str("<w14:noFill/>"),
        Fill::Solid(color) => {
            xml.push_str("<w14:solidFill>");
            write_color_children(xml, color.as_ref())?;
            xml.push_str("</w14:solidFill>");
        },
        Fill::Gradient(value) => {
            xml.push_str("<w14:gradFill>");
            if !value.stops.is_empty() {
                xml.push_str("<w14:gsLst>");
                for stop in &value.stops {
                    write!(xml, "<w14:gs w14:pos=\"{}\">", stop.position)?;
                    write_color_children(xml, Some(&stop.color))?;
                    xml.push_str("</w14:gs>");
                }
                xml.push_str("</w14:gsLst>");
            }
            if let Some(shade) = &value.shade {
                write_shade(shade, xml)?;
            }
            xml.push_str("</w14:gradFill>");
        },
    }
    Ok(())
}

fn write_shade(value: &Shade, xml: &mut String) -> Result<()> {
    match value {
        Shade::Linear { angle, scaled } => {
            xml.push_str("<w14:lin");
            attr_u32_out(xml, "ang", *angle)?;
            attr_str_out(xml, "scaled", scaled.map(on_off_str))?;
            xml.push_str("/>");
        },
        Shade::Path { path, fill_to } => {
            xml.push_str("<w14:path");
            attr_str_out(xml, "path", path.map(path_str))?;
            if let Some(rect) = fill_to {
                xml.push('>');
                write!(xml, "<w14:fillToRect")?;
                attr_i32_out(xml, "l", rect.left)?;
                attr_i32_out(xml, "t", rect.top)?;
                attr_i32_out(xml, "r", rect.right)?;
                attr_i32_out(xml, "b", rect.bottom)?;
                xml.push_str("/></w14:path>");
            } else {
                xml.push_str("/>");
            }
        },
    }
    Ok(())
}

fn write_color_element(name: &str, color: &Color, xml: &mut String) -> Result<()> {
    write!(xml, "<w14:{name}>")?;
    write_color_children(xml, Some(color))?;
    write!(xml, "</w14:{name}>")?;
    Ok(())
}

fn write_color_children(xml: &mut String, color: Option<&Color>) -> Result<()> {
    let Some(color) = color else { return Ok(()) };
    match color {
        Color::Rgb(value) => {
            write!(
                xml,
                "<w14:srgbClr w14:val=\"{:02X}{:02X}{:02X}\">",
                value.value[0], value.value[1], value.value[2]
            )?;
            write_transforms(&value.transforms, xml)?;
            xml.push_str("</w14:srgbClr>");
        },
        Color::Scheme(value) => {
            write!(xml, "<w14:schemeClr w14:val=\"{}\">", value.value.as_str())?;
            write_transforms(&value.transforms, xml)?;
            xml.push_str("</w14:schemeClr>");
        },
    }
    Ok(())
}

fn write_transforms(transforms: &[ColorTransform], xml: &mut String) -> Result<()> {
    for transform in transforms {
        let (name, value) = match transform {
            ColorTransform::Tint(value) => ("tint", value.to_string()),
            ColorTransform::Shade(value) => ("shade", value.to_string()),
            ColorTransform::Alpha(value) => ("alpha", value.to_string()),
            ColorTransform::HueMod(value) => ("hueMod", value.to_string()),
            ColorTransform::Saturation(value) => ("sat", value.to_string()),
            ColorTransform::SaturationOffset(value) => ("satOff", value.to_string()),
            ColorTransform::SaturationMod(value) => ("satMod", value.to_string()),
            ColorTransform::Luminance(value) => ("lum", value.to_string()),
            ColorTransform::LuminanceOffset(value) => ("lumOff", value.to_string()),
            ColorTransform::LuminanceMod(value) => ("lumMod", value.to_string()),
        };
        write!(xml, "<w14:{name} w14:val=\"{value}\"/>")?;
    }
    Ok(())
}

fn attr_u64_out(xml: &mut String, name: &str, value: Option<u64>) -> Result<()> {
    if let Some(value) = value {
        write!(xml, " w14:{name}=\"{value}\"")?;
    }
    Ok(())
}

fn attr_u32_out(xml: &mut String, name: &str, value: Option<u32>) -> Result<()> {
    if let Some(value) = value {
        write!(xml, " w14:{name}=\"{value}\"")?;
    }
    Ok(())
}

fn attr_i32_out(xml: &mut String, name: &str, value: Option<i32>) -> Result<()> {
    if let Some(value) = value {
        write!(xml, " w14:{name}=\"{value}\"")?;
    }
    Ok(())
}

fn attr_str_out(xml: &mut String, name: &str, value: Option<&str>) -> Result<()> {
    if let Some(value) = value {
        write!(xml, " w14:{name}=\"{}\"", escape_xml(value))?;
    }
    Ok(())
}

fn alignment_str(value: RectAlignment) -> &'static str {
    match value {
        RectAlignment::None => "none",
        RectAlignment::TopLeft => "tl",
        RectAlignment::Top => "t",
        RectAlignment::TopRight => "tr",
        RectAlignment::Left => "l",
        RectAlignment::Center => "ctr",
        RectAlignment::Right => "r",
        RectAlignment::BottomLeft => "bl",
        RectAlignment::Bottom => "b",
        RectAlignment::BottomRight => "br",
    }
}

fn cap_str(value: LineCap) -> &'static str {
    match value {
        LineCap::Flat => "flat",
        LineCap::Round => "rnd",
        LineCap::Square => "sq",
    }
}

fn compound_str(value: CompoundLine) -> &'static str {
    match value {
        CompoundLine::Single => "sng",
        CompoundLine::Double => "dbl",
        CompoundLine::ThickThin => "thickThin",
        CompoundLine::ThinThick => "thinThick",
        CompoundLine::Triple => "tri",
    }
}

fn pen_alignment_str(value: PenAlignment) -> &'static str {
    match value {
        PenAlignment::Center => "ctr",
        PenAlignment::Inside => "in",
    }
}

fn dash_str(value: LineDash) -> &'static str {
    match value {
        LineDash::Solid => "solid",
        LineDash::Dot => "dot",
        LineDash::SysDot => "sysDot",
        LineDash::Dash => "dash",
        LineDash::SysDash => "sysDash",
        LineDash::LargeDash => "lgDash",
        LineDash::DashDot => "dashDot",
        LineDash::SysDashDot => "sysDashDot",
        LineDash::LargeDashDot => "lgDashDot",
        LineDash::LargeDashDotDot => "lgDashDotDot",
        LineDash::SysDashDotDot => "sysDashDotDot",
    }
}

fn path_str(value: PathKind) -> &'static str {
    match value {
        PathKind::Shape => "shape",
        PathKind::Circle => "circle",
        PathKind::Rect => "rect",
    }
}

fn on_off_str(value: bool) -> &'static str {
    if value { "1" } else { "0" }
}
