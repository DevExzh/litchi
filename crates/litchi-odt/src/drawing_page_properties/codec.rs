//! Bounded XML parsing and lossless mutation for drawing-page styles.

use super::model::*;
use litchi_core::{Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{QName, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const SMIL: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";
const SVG: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const XML: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const MAX_XML: usize = 32 * 1024 * 1024;
const MAX_ATTRIBUTES: usize = 64;
const MAX_DEPTH: usize = 128;
const MAX_STYLES: usize = 65_536;
const MAX_TOTAL: usize = 16 * 1024 * 1024;
const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const DRAW_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const SMIL_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";
const SVG_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const XLINK_NS: &str = "http://www.w3.org/1999/xlink";

impl Sound {
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            r#"<presentation:sound xmlns:presentation="{PRESENTATION_NS}" xmlns:xlink="{XLINK_NS}" xlink:type="simple" xlink:href="{}""#,
            escape_xml(&self.href)
        );
        if self.actuate_on_request {
            xml.push_str(r#" xlink:actuate="onRequest""#);
        }
        if let Some(show) = self.show {
            xml.push_str(&format!(r#" xlink:show="{}""#, show.xml()));
        }
        if let Some(value) = self.play_full {
            xml.push_str(&format!(r#" presentation:play-full="{value}""#));
        }
        if let Some(id) = &self.xml_id {
            xml.push_str(&format!(r#" xml:id="{}""#, escape_xml(id)));
        }
        xml.push_str("/>");
        Ok(xml)
    }
}

impl StyleProperties {
    pub fn from_xml_fragment(fragment: &str) -> Result<Self> {
        let xml = format!(
            r#"<office:document xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}"><office:styles><style:style style:name="fragment" style:family="drawing-page">{fragment}</style:style></office:styles></office:document>"#
        );
        let mut set = parse_drawing_page_style_properties(&xml)?;
        set.styles
            .pop()
            .and_then(|style| style.properties)
            .ok_or_else(|| bad("fragment does not contain style:drawing-page-properties"))
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            r#"<style:drawing-page-properties xmlns:style="{STYLE_NS}" xmlns:draw="{DRAW_NS}" xmlns:presentation="{PRESENTATION_NS}" xmlns:smil="{SMIL_NS}" xmlns:svg="{SVG_NS}" xmlns:xlink="{XLINK_NS}""#
        );
        macro_rules! attr {
            ($field:expr,$name:literal,$value:expr) => {
                if let Some(value) = $field {
                    xml.push_str(&format!(concat!(" ", $name, "=\"{}\""), $value(value)))
                }
            };
        }
        attr!(self.fill, "draw:fill", |v: Fill| v.xml());
        attr!(self.fill_color.as_ref(), "draw:fill-color", |v: &Color| v
            .as_str()
            .to_owned());
        attr!(
            self.secondary_fill_color.as_ref(),
            "draw:secondary-fill-color",
            |v: &Color| v.as_str().to_owned()
        );
        attr!(
            self.fill_gradient_name.as_ref(),
            "draw:fill-gradient-name",
            |v: &StyleNameRef| v.as_str().to_owned()
        );
        attr!(
            self.gradient_step_count.as_ref(),
            "draw:gradient-step-count",
            |v: &NonNegativeInteger| v.as_str().to_owned()
        );
        attr!(
            self.fill_hatch_name.as_ref(),
            "draw:fill-hatch-name",
            |v: &StyleNameRef| v.as_str().to_owned()
        );
        attr!(
            self.fill_hatch_solid,
            "draw:fill-hatch-solid",
            |v: bool| if v { "true" } else { "false" }
        );
        attr!(
            self.fill_image_name.as_ref(),
            "draw:fill-image-name",
            |v: &StyleNameRef| v.as_str().to_owned()
        );
        attr!(self.repeat, "style:repeat", |v: Repeat| v.xml());
        attr!(
            self.fill_image_width.as_ref(),
            "draw:fill-image-width",
            |v: &LengthOrPercent| v.as_str().to_owned()
        );
        attr!(
            self.fill_image_height.as_ref(),
            "draw:fill-image-height",
            |v: &LengthOrPercent| v.as_str().to_owned()
        );
        attr!(
            self.fill_image_ref_point_x.as_ref(),
            "draw:fill-image-ref-point-x",
            |v: &Percent| v.as_str().to_owned()
        );
        attr!(
            self.fill_image_ref_point_y.as_ref(),
            "draw:fill-image-ref-point-y",
            |v: &Percent| v.as_str().to_owned()
        );
        attr!(
            self.fill_image_ref_point,
            "draw:fill-image-ref-point",
            |v: ImageRefPoint| v.xml()
        );
        if let Some(value) = &self.tile_repeat_offset {
            xml.push_str(&format!(
                r#" draw:tile-repeat-offset="{} {}""#,
                value.percentage,
                value.direction.xml()
            ));
        }
        attr!(self.opacity.as_ref(), "draw:opacity", |v: &Percent| v
            .as_str()
            .to_owned());
        attr!(
            self.opacity_name.as_ref(),
            "draw:opacity-name",
            |v: &StyleNameRef| v.as_str().to_owned()
        );
        attr!(self.fill_rule, "svg:fill-rule", |v: FillRule| v.xml());
        attr!(
            self.transition_type,
            "presentation:transition-type",
            |v: TransitionType| v.xml()
        );
        attr!(
            self.transition_style.as_ref(),
            "presentation:transition-style",
            |v: &TransitionStyle| v.as_str().to_owned()
        );
        attr!(
            self.transition_speed,
            "presentation:transition-speed",
            |v: TransitionSpeed| v.xml()
        );
        if let Some(value) = &self.smil_type {
            xml.push_str(&format!(r#" smil:type="{}""#, escape_xml(value)));
        }
        if let Some(value) = &self.smil_subtype {
            xml.push_str(&format!(r#" smil:subtype="{}""#, escape_xml(value)));
        }
        attr!(
            self.direction,
            "smil:direction",
            |v: TransitionDirection| v.xml()
        );
        attr!(self.fade_color.as_ref(), "smil:fadeColor", |v: &Color| v
            .as_str()
            .to_owned());
        attr!(
            self.duration.as_ref(),
            "presentation:duration",
            |v: &Duration| v.as_str().to_owned()
        );
        attr!(
            self.visibility,
            "presentation:visibility",
            |v: Visibility| v.xml()
        );
        attr!(
            self.background_size,
            "draw:background-size",
            |v: BackgroundSize| v.xml()
        );
        attr!(
            self.background_objects_visible,
            "presentation:background-objects-visible",
            |v: bool| if v { "true" } else { "false" }
        );
        attr!(
            self.background_visible,
            "presentation:background-visible",
            |v: bool| if v { "true" } else { "false" }
        );
        attr!(
            self.display_header,
            "presentation:display-header",
            |v: bool| if v { "true" } else { "false" }
        );
        attr!(
            self.display_footer,
            "presentation:display-footer",
            |v: bool| if v { "true" } else { "false" }
        );
        attr!(
            self.display_page_number,
            "presentation:display-page-number",
            |v: bool| if v { "true" } else { "false" }
        );
        attr!(
            self.display_date_time,
            "presentation:display-date-time",
            |v: bool| if v { "true" } else { "false" }
        );
        if let Some(sound) = &self.sound {
            xml.push('>');
            xml.push_str(&sound.to_xml_fragment()?);
            xml.push_str("</style:drawing-page-properties>");
        } else {
            xml.push_str("/>");
        }
        Ok(xml)
    }
}

impl Style {
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let tag = if self.is_default_style {
            "default-style"
        } else {
            "style"
        };
        let mut xml =
            format!(r#"<style:{tag} xmlns:style="{STYLE_NS}" style:family="drawing-page""#);
        if let Some(value) = &self.name {
            xml.push_str(&format!(r#" style:name="{}""#, escape_xml(value)));
        }
        if let Some(value) = &self.parent_style_name {
            xml.push_str(&format!(
                r#" style:parent-style-name="{}""#,
                escape_xml(value)
            ));
        }
        if let Some(value) = &self.properties {
            xml.push('>');
            xml.push_str(&value.to_xml_fragment()?);
            xml.push_str(&format!("</style:{tag}>"));
        } else {
            xml.push_str("/>");
        }
        Ok(xml)
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum Ns {
    Office,
    Style,
    Draw,
    Presentation,
    Smil,
    Svg,
    Xlink,
    Xml,
    Other,
}
fn ns(value: ResolveResult<'_>) -> Ns {
    match value {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE => Ns::Office,
        ResolveResult::Bound(value) if value.as_ref() == STYLE => Ns::Style,
        ResolveResult::Bound(value) if value.as_ref() == DRAW => Ns::Draw,
        ResolveResult::Bound(value) if value.as_ref() == PRESENTATION => Ns::Presentation,
        ResolveResult::Bound(value) if value.as_ref() == SMIL => Ns::Smil,
        ResolveResult::Bound(value) if value.as_ref() == SVG => Ns::Svg,
        ResolveResult::Bound(value) if value.as_ref() == XLINK => Ns::Xlink,
        ResolveResult::Bound(value) if value.as_ref() == XML => Ns::Xml,
        _ => Ns::Other,
    }
}
fn element(reader: &NsReader<&[u8]>, name: QName<'_>) -> (Ns, Vec<u8>) {
    let (namespace, local) = reader.resolver().resolve_element(name);
    (ns(namespace), local.as_ref().to_vec())
}
fn attrs(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Vec<(Ns, Vec<u8>, String)>> {
    let mut out = Vec::new();
    let mut seen = HashSet::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| bad(format!("invalid drawing-page property attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if out.len() >= MAX_ATTRIBUTES {
            return Err(bad("too many drawing-page property attributes"));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let key = (ns(namespace), local.as_ref().to_vec());
        if !seen.insert(key.clone()) {
            return Err(bad("duplicate drawing-page property attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid drawing-page property value: {error}")))?
            .into_owned();
        safe(&value, "drawing-page property value", true)?;
        out.push((key.0, key.1, value));
    }
    Ok(out)
}
fn take(attrs: &mut Vec<(Ns, Vec<u8>, String)>, namespace: Ns, local: &[u8]) -> Option<String> {
    attrs
        .iter()
        .position(|value| value.0 == namespace && value.1 == local)
        .map(|index| attrs.remove(index).2)
}
fn boolean(value: &str) -> Result<bool> {
    match value {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(bad("ODF boolean must be true or false")),
    }
}
fn enum_value<T>(value: Option<String>, parse: fn(&str) -> Result<T>) -> Result<Option<T>> {
    value.map(|value| parse(&value)).transpose()
}
fn header(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    default: bool,
) -> Result<Option<Style>> {
    let mut attrs = attrs(reader, version, start)?;
    if take(&mut attrs, Ns::Style, b"family").as_deref() != Some("drawing-page") {
        return Ok(None);
    }
    let style = Style {
        name: take(&mut attrs, Ns::Style, b"name"),
        parent_style_name: take(&mut attrs, Ns::Style, b"parent-style-name"),
        is_default_style: default,
        properties: None,
    };
    style.validate()?;
    Ok(Some(style))
}
fn parse_properties(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<StyleProperties> {
    let mut a = attrs(reader, version, start)?;
    let value = StyleProperties {
        fill: enum_value(take(&mut a, Ns::Draw, b"fill"), Fill::parse)?,
        fill_color: take(&mut a, Ns::Draw, b"fill-color")
            .map(Color::new)
            .transpose()?,
        secondary_fill_color: take(&mut a, Ns::Draw, b"secondary-fill-color")
            .map(Color::new)
            .transpose()?,
        fill_gradient_name: take(&mut a, Ns::Draw, b"fill-gradient-name")
            .map(StyleNameRef::new)
            .transpose()?,
        gradient_step_count: take(&mut a, Ns::Draw, b"gradient-step-count")
            .map(NonNegativeInteger::new)
            .transpose()?,
        fill_hatch_name: take(&mut a, Ns::Draw, b"fill-hatch-name")
            .map(StyleNameRef::new)
            .transpose()?,
        fill_hatch_solid: take(&mut a, Ns::Draw, b"fill-hatch-solid")
            .map(|v| boolean(&v))
            .transpose()?,
        fill_image_name: take(&mut a, Ns::Draw, b"fill-image-name")
            .map(StyleNameRef::new)
            .transpose()?,
        repeat: enum_value(take(&mut a, Ns::Style, b"repeat"), Repeat::parse)?,
        fill_image_width: take(&mut a, Ns::Draw, b"fill-image-width")
            .map(LengthOrPercent::new)
            .transpose()?,
        fill_image_height: take(&mut a, Ns::Draw, b"fill-image-height")
            .map(LengthOrPercent::new)
            .transpose()?,
        fill_image_ref_point_x: take(&mut a, Ns::Draw, b"fill-image-ref-point-x")
            .map(Percent::new)
            .transpose()?,
        fill_image_ref_point_y: take(&mut a, Ns::Draw, b"fill-image-ref-point-y")
            .map(Percent::new)
            .transpose()?,
        fill_image_ref_point: enum_value(
            take(&mut a, Ns::Draw, b"fill-image-ref-point"),
            ImageRefPoint::parse,
        )?,
        tile_repeat_offset: take(&mut a, Ns::Draw, b"tile-repeat-offset")
            .map(|v| {
                let mut parts = v.split_ascii_whitespace();
                let percentage = parts
                    .next()
                    .ok_or_else(|| bad("invalid draw:tile-repeat-offset"))?;
                let direction = parts
                    .next()
                    .ok_or_else(|| bad("invalid draw:tile-repeat-offset"))?;
                if parts.next().is_some() {
                    return Err(bad("invalid draw:tile-repeat-offset"));
                }
                TileRepeatOffset::new(percentage, TileDirection::parse(direction)?)
            })
            .transpose()?,
        opacity: take(&mut a, Ns::Draw, b"opacity")
            .map(Percent::new)
            .transpose()?,
        opacity_name: take(&mut a, Ns::Draw, b"opacity-name")
            .map(StyleNameRef::new)
            .transpose()?,
        fill_rule: enum_value(take(&mut a, Ns::Svg, b"fill-rule"), FillRule::parse)?,
        transition_type: enum_value(
            take(&mut a, Ns::Presentation, b"transition-type"),
            TransitionType::parse,
        )?,
        transition_style: take(&mut a, Ns::Presentation, b"transition-style")
            .map(TransitionStyle::new)
            .transpose()?,
        transition_speed: enum_value(
            take(&mut a, Ns::Presentation, b"transition-speed"),
            TransitionSpeed::parse,
        )?,
        smil_type: take(&mut a, Ns::Smil, b"type"),
        smil_subtype: take(&mut a, Ns::Smil, b"subtype"),
        direction: enum_value(
            take(&mut a, Ns::Smil, b"direction"),
            TransitionDirection::parse,
        )?,
        fade_color: take(&mut a, Ns::Smil, b"fadeColor")
            .map(Color::new)
            .transpose()?,
        duration: take(&mut a, Ns::Presentation, b"duration")
            .map(Duration::new)
            .transpose()?,
        visibility: enum_value(
            take(&mut a, Ns::Presentation, b"visibility"),
            Visibility::parse,
        )?,
        background_size: enum_value(
            take(&mut a, Ns::Draw, b"background-size"),
            BackgroundSize::parse,
        )?,
        background_objects_visible: take(&mut a, Ns::Presentation, b"background-objects-visible")
            .map(|v| boolean(&v))
            .transpose()?,
        background_visible: take(&mut a, Ns::Presentation, b"background-visible")
            .map(|v| boolean(&v))
            .transpose()?,
        display_header: take(&mut a, Ns::Presentation, b"display-header")
            .map(|v| boolean(&v))
            .transpose()?,
        display_footer: take(&mut a, Ns::Presentation, b"display-footer")
            .map(|v| boolean(&v))
            .transpose()?,
        display_page_number: take(&mut a, Ns::Presentation, b"display-page-number")
            .map(|v| boolean(&v))
            .transpose()?,
        display_date_time: take(&mut a, Ns::Presentation, b"display-date-time")
            .map(|v| boolean(&v))
            .transpose()?,
        sound: None,
    };
    if !a.is_empty() {
        return Err(bad(
            "unknown style:drawing-page-properties attribute or wrong namespace",
        ));
    }
    value.validate()?;
    Ok(value)
}
fn parse_sound(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Sound> {
    let mut a = attrs(reader, version, start)?;
    if take(&mut a, Ns::Xlink, b"type").as_deref() != Some("simple") {
        return Err(bad("presentation:sound requires xlink:type=\"simple\""));
    }
    let href = take(&mut a, Ns::Xlink, b"href")
        .ok_or_else(|| bad("presentation:sound requires xlink:href"))?;
    let actuate = take(&mut a, Ns::Xlink, b"actuate");
    if actuate.as_deref().is_some_and(|v| v != "onRequest") {
        return Err(bad("invalid presentation:sound xlink:actuate"));
    }
    let show = enum_value(take(&mut a, Ns::Xlink, b"show"), SoundShow::parse)?;
    let play_full = take(&mut a, Ns::Presentation, b"play-full")
        .map(|v| boolean(&v))
        .transpose()?;
    let xml_id = take(&mut a, Ns::Xml, b"id");
    if !a.is_empty() {
        return Err(bad(
            "unknown presentation:sound attribute or wrong namespace",
        ));
    }
    let value = Sound {
        href,
        play_full,
        actuate_on_request: actuate.is_some(),
        show,
        xml_id,
    };
    value.validate()?;
    Ok(value)
}

struct Active {
    depth: usize,
    style: Style,
    seen: bool,
    properties_depth: Option<usize>,
    sound_depth: Option<usize>,
}
fn push_style(out: &mut Vec<Style>, style: Style, total: &mut usize) -> Result<()> {
    if out.len() >= MAX_STYLES
        || out.iter().any(|value| {
            value.name == style.name && value.is_default_style == style.is_default_style
        })
    {
        return Err(bad("duplicate or excessive drawing-page style"));
    }
    *total += style.to_xml_fragment()?.len();
    if *total > MAX_TOTAL {
        return Err(bad("drawing-page style data is too large"));
    }
    out.push(style);
    Ok(())
}
/// Parse direct drawing-page styles from `office:styles` and `office:automatic-styles`.
pub fn parse_drawing_page_style_properties(xml: &str) -> Result<Styles> {
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut active: Option<Active> = None;
    let mut out = Vec::new();
    let mut total = 0;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("styles XML nesting is too deep"));
                }
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|value| {
                    value.0 == Ns::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if active.is_some() {
                        return Err(bad("nested drawing-page style"));
                    }
                    if let Some(style) =
                        header(&reader, version, &start, current.1 == b"default-style")?
                    {
                        active = Some(Active {
                            depth,
                            style,
                            seen: false,
                            properties_depth: None,
                            sound_depth: None,
                        });
                    }
                    continue;
                }
                if let Some(value) = active.as_mut() {
                    if depth == value.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"drawing-page-properties"
                    {
                        if value.seen {
                            return Err(bad("duplicate style:drawing-page-properties"));
                        }
                        value.seen = true;
                        value.style.properties = Some(parse_properties(&reader, version, &start)?);
                        value.properties_depth = Some(depth);
                    } else if current.1 == b"drawing-page-properties" {
                        return Err(bad(
                            "style:drawing-page-properties has invalid namespace or parent",
                        ));
                    } else if value.properties_depth.is_some_and(|p| depth == p + 1)
                        && current.0 == Ns::Presentation
                        && current.1 == b"sound"
                    {
                        if value
                            .style
                            .properties
                            .as_ref()
                            .ok_or_else(|| bad("missing drawing-page properties"))?
                            .sound
                            .is_some()
                        {
                            return Err(bad("duplicate presentation:sound"));
                        }
                        value
                            .style
                            .properties
                            .as_mut()
                            .ok_or_else(|| bad("missing drawing-page properties"))?
                            .sound = Some(parse_sound(&reader, version, &start)?);
                        value.sound_depth = Some(depth);
                    } else if value.properties_depth.is_some_and(|p| depth > p) {
                        return Err(bad("unexpected style:drawing-page-properties child"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let current = element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                let direct = parent.is_some_and(|value| {
                    value.0 == Ns::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                if direct {
                    if let Some(style) =
                        header(&reader, version, &start, current.1 == b"default-style")?
                    {
                        push_style(&mut out, style, &mut total)?;
                    }
                    continue;
                }
                if let Some(value) = active.as_mut() {
                    if depth == value.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"drawing-page-properties"
                    {
                        if value.seen {
                            return Err(bad("duplicate style:drawing-page-properties"));
                        }
                        value.seen = true;
                        value.style.properties = Some(parse_properties(&reader, version, &start)?);
                    } else if current.1 == b"drawing-page-properties" {
                        return Err(bad(
                            "style:drawing-page-properties has invalid namespace or parent",
                        ));
                    } else if value.properties_depth.is_some_and(|p| depth == p + 1)
                        && current.0 == Ns::Presentation
                        && current.1 == b"sound"
                    {
                        if value
                            .style
                            .properties
                            .as_ref()
                            .ok_or_else(|| bad("missing drawing-page properties"))?
                            .sound
                            .is_some()
                        {
                            return Err(bad("duplicate presentation:sound"));
                        }
                        value
                            .style
                            .properties
                            .as_mut()
                            .ok_or_else(|| bad("missing drawing-page properties"))?
                            .sound = Some(parse_sound(&reader, version, &start)?);
                    } else if value.properties_depth.is_some_and(|p| depth > p) {
                        return Err(bad("unexpected style:drawing-page-properties child"));
                    }
                }
            },
            Ok(Event::Text(text)) => {
                let bytes: &[u8] = text.as_ref();
                if active
                    .as_ref()
                    .is_some_and(|value| value.properties_depth.is_some())
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(bad("unexpected text in style:drawing-page-properties"));
                }
            },
            Ok(Event::CData(text)) => {
                let bytes: &[u8] = text.as_ref();
                if active
                    .as_ref()
                    .is_some_and(|value| value.properties_depth.is_some())
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(bad("unexpected text in style:drawing-page-properties"));
                }
            },
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if let Some(value) = active.as_mut() {
                    if value.sound_depth == Some(depth) {
                        value.sound_depth = None;
                    }
                    if value.properties_depth == Some(depth) {
                        value.properties_depth = None;
                    }
                }
                if active.as_ref().is_some_and(|value| value.depth == depth) {
                    let value = active
                        .take()
                        .ok_or_else(|| bad("missing active drawing-page style"))?;
                    push_style(&mut out, value.style, &mut total)?;
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    if !stack.is_empty() || active.is_some() {
        return Err(bad("truncated styles XML"));
    }
    Ok(Styles { styles: out })
}

#[derive(Default)]
struct Span {
    start: usize,
    end: usize,
    end_start: usize,
    qname: String,
    empty: bool,
}
#[derive(Default)]
struct TargetSpans {
    style: Span,
    properties: Option<Span>,
}
fn boundary(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| bad("invalid XML event boundary"))
}
fn replace_span(xml: &str, span: &Span, value: &str) -> String {
    format!("{}{}{}", &xml[..span.start], value, &xml[span.end..])
}
fn expand_span(xml: &str, span: &Span, value: &str) -> Result<String> {
    let raw = &xml[span.start..span.end];
    let slash = raw
        .rfind("/>")
        .ok_or_else(|| bad("invalid empty element"))?;
    Ok(replace_span(
        xml,
        span,
        &format!("{}>{value}</{}>", &raw[..slash], span.qname),
    ))
}
/// Losslessly replace, insert, or remove one existing drawing-page style's property element.
pub fn set_drawing_page_style_properties_xml(xml: &str, requested: &Style) -> Result<String> {
    requested.validate()?;
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut target_depth = None;
    let mut active: Option<TargetSpans> = None;
    let mut found = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|value| {
                    value.0 == Ns::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) =
                        header(&reader, version, &start, current.1 == b"default-style")?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target drawing-page style"));
                        }
                        target_depth = Some(depth);
                        active = Some(TargetSpans {
                            style: Span {
                                start: begin,
                                qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                                ..Default::default()
                            },
                            ..Default::default()
                        });
                    }
                } else if target_depth.is_some_and(|value| depth == value + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"drawing-page-properties"
                {
                    let span = Span {
                        start: begin,
                        qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                        ..Default::default()
                    };
                    if active
                        .as_mut()
                        .ok_or_else(|| bad("missing target drawing-page style"))?
                        .properties
                        .replace(span)
                        .is_some()
                    {
                        return Err(bad("duplicate style:drawing-page-properties"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                let direct = parent.is_some_and(|value| {
                    value.0 == Ns::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                let span = Span {
                    start: begin,
                    end,
                    end_start: begin,
                    qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                    empty: true,
                };
                if direct {
                    if let Some(style) =
                        header(&reader, version, &start, current.1 == b"default-style")?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target drawing-page style"));
                        }
                        found = Some(TargetSpans {
                            style: span,
                            ..Default::default()
                        });
                    }
                } else if target_depth.is_some_and(|value| depth == value + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"drawing-page-properties"
                    && active
                        .as_mut()
                        .ok_or_else(|| bad("missing target drawing-page style"))?
                        .properties
                        .replace(span)
                        .is_some()
                {
                    return Err(bad("duplicate style:drawing-page-properties"));
                }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let depth = stack.len();
                if let Some(spans) = active.as_mut() {
                    if spans.properties.as_ref().is_some_and(|span| span.end == 0)
                        && target_depth.is_some_and(|value| depth == value + 1)
                    {
                        let span = spans
                            .properties
                            .as_mut()
                            .ok_or_else(|| bad("missing target drawing-page properties span"))?;
                        span.end_start = begin;
                        span.end = end;
                    }
                    if target_depth == Some(depth) {
                        spans.style.end_start = begin;
                        spans.style.end = end;
                        found = active.take();
                        target_depth = None;
                    }
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    let spans = found.ok_or_else(|| bad("target drawing-page style does not exist"))?;
    let replacement = requested
        .properties
        .as_ref()
        .map(StyleProperties::to_xml_fragment)
        .transpose()?;
    if let Some(properties) = &spans.properties {
        return Ok(replace_span(
            xml,
            properties,
            replacement.as_deref().unwrap_or(""),
        ));
    }
    let Some(replacement) = replacement else {
        return Ok(xml.to_owned());
    };
    if spans.style.empty {
        return expand_span(xml, &spans.style, &replacement);
    }
    let mut out = xml.to_owned();
    out.insert_str(spans.style.end_start, &replacement);
    Ok(out)
}
