//! Complete typed ODF `style:drawing-page-properties` support.

use crate::{FlatOpenDocument, OpenDocumentPackage};
use litchi_core::{Error, Result, xml::escape_xml};
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
const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const DRAW_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const SMIL_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";
const SVG_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
const MAX_XML: usize = 32 * 1024 * 1024;
const MAX_VALUE: usize = 1024 * 1024;
const MAX_ATTRIBUTES: usize = 64;
const MAX_DEPTH: usize = 128;
const MAX_STYLES: usize = 65_536;
const MAX_TOTAL: usize = 16 * 1024 * 1024;

fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn safe(value: &str, name: &str, empty: bool) -> Result<()> {
    if (!empty && value.is_empty()) || value.len() > MAX_VALUE || value.chars().any(|c| matches!(c, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}' | '\u{fffe}' | '\u{ffff}')) {
        return Err(bad(format!("invalid {name}")));
    }
    Ok(())
}
fn ncname(value: &str, name: &str, empty: bool) -> Result<()> {
    safe(value, name, empty)?;
    if value.is_empty() && empty {
        return Ok(());
    }
    let mut chars = value.chars();
    let first = chars.next().ok_or_else(|| bad(format!("invalid {name}")))?;
    if !(first == '_' || first.is_alphabetic())
        || chars.any(|c| !(c == '_' || c == '-' || c == '.' || c.is_alphanumeric()))
    {
        return Err(bad(format!("invalid {name}")));
    }
    Ok(())
}
fn decimal(value: &str, signed: bool) -> bool {
    let value = if signed {
        value.strip_prefix('-').unwrap_or(value)
    } else {
        value
    };
    if value.is_empty() {
        return false;
    }
    let mut split = value.split('.');
    let left = split.next().unwrap_or_default();
    let right = split.next();
    if split.next().is_some() {
        return false;
    }
    match right {
        None => !left.is_empty() && left.bytes().all(|b| b.is_ascii_digit()),
        Some(right) => {
            (!left.is_empty() || !right.is_empty())
                && left.bytes().all(|b| b.is_ascii_digit())
                && right.bytes().all(|b| b.is_ascii_digit())
        },
    }
}
fn percent(value: &str) -> bool {
    value.strip_suffix('%').is_some_and(|x| decimal(x, true))
}
fn zero_to_hundred_percent(value: &str) -> bool {
    let Some(number) = value.strip_suffix('%') else {
        return false;
    };
    if !decimal(number, false) {
        return false;
    }
    number.parse::<f64>().is_ok_and(|number| number <= 100.0)
}
fn length(value: &str) -> bool {
    ["cm", "mm", "in", "pt", "pc", "px"]
        .iter()
        .any(|unit| value.strip_suffix(unit).is_some_and(|x| decimal(x, true)))
}
fn duration(value: &str) -> bool {
    let value = value.strip_prefix('-').unwrap_or(value);
    let Some(mut rest) = value.strip_prefix('P') else {
        return false;
    };
    if rest.is_empty() {
        return false;
    }
    let mut any = false;
    let mut time = false;
    let mut last = 0u8;
    while !rest.is_empty() {
        if rest.starts_with('T') {
            if time {
                return false;
            }
            time = true;
            last = 0;
            rest = &rest[1..];
            if rest.is_empty() {
                return false;
            }
            continue;
        }
        let number_end = rest
            .find(|c: char| !c.is_ascii_digit() && c != '.')
            .unwrap_or(rest.len());
        if number_end == 0 || number_end == rest.len() {
            return false;
        }
        let number = &rest[..number_end];
        let marker = rest.as_bytes()[number_end] as char;
        let rank = match (time, marker) {
            (false, 'Y') => 1,
            (false, 'M') => 2,
            (false, 'D') => 3,
            (true, 'H') => 1,
            (true, 'M') => 2,
            (true, 'S') => 3,
            _ => return false,
        };
        if rank <= last
            || (marker != 'S'
                && (number.contains('.') || !number.bytes().all(|b| b.is_ascii_digit())))
            || (marker == 'S' && !decimal(number, false))
        {
            return false;
        }
        last = rank;
        any = true;
        rest = &rest[number_end + 1..];
    }
    any
}

macro_rules! lexical {
    ($name:ident, $validator:expr, $label:literal, $empty:expr) => {
        #[derive(Debug, Clone, PartialEq, Eq)]
        pub struct $name(String);
        impl $name {
            pub fn new(value: impl Into<String>) -> Result<Self> {
                let value = value.into();
                safe(&value, $label, $empty)?;
                if !($validator)(&value) {
                    return Err(bad(concat!("invalid ", $label)));
                }
                Ok(Self(value))
            }
            pub fn as_str(&self) -> &str {
                &self.0
            }
        }
    };
}

lexical!(
    DrawingPageColor,
    |x: &str| x.len() == 7 && x.starts_with('#') && x[1..].bytes().all(|b| b.is_ascii_hexdigit()),
    "ODF color",
    false
);
lexical!(DrawingPagePercent, percent, "ODF percentage", false);
lexical!(
    DrawingPageLengthOrPercent,
    |x: &str| length(x) || percent(x),
    "ODF length or percentage",
    false
);
lexical!(
    DrawingPageDuration,
    duration,
    "presentation:duration",
    false
);
lexical!(
    DrawingPageStyleNameRef,
    |x: &str| ncname(x, "style name reference", true).is_ok(),
    "style name reference",
    true
);
lexical!(
    DrawingPageNonNegativeInteger,
    |x: &str| !x.is_empty() && x.bytes().all(|b| b.is_ascii_digit()),
    "non-negative integer",
    false
);

macro_rules! keyword_enum {
    ($name:ident { $($variant:ident => $value:literal),+ $(,)? }) => {
        #[derive(Debug, Clone, Copy, PartialEq, Eq)] pub enum $name { $($variant),+ }
        impl $name {
            fn parse(value:&str)->Result<Self>{match value{$($value=>Ok(Self::$variant),)+_=>Err(bad(concat!("invalid ",stringify!($name))))}}
            fn xml(self)->&'static str{match self{$(Self::$variant=>$value),+}}
        }
    };
}
keyword_enum!(DrawingPageFill { None=>"none", Solid=>"solid", Bitmap=>"bitmap", Gradient=>"gradient", Hatch=>"hatch" });
keyword_enum!(DrawingPageRepeat { NoRepeat=>"no-repeat", Repeat=>"repeat", Stretch=>"stretch" });
keyword_enum!(DrawingPageImageRefPoint { TopLeft=>"top-left", Top=>"top", TopRight=>"top-right", Left=>"left", Center=>"center", Right=>"right", BottomLeft=>"bottom-left", Bottom=>"bottom", BottomRight=>"bottom-right" });
keyword_enum!(DrawingPageTileDirection { Horizontal=>"horizontal", Vertical=>"vertical" });
keyword_enum!(DrawingPageFillRule { Nonzero=>"nonzero", Evenodd=>"evenodd" });
keyword_enum!(DrawingPageTransitionType { Manual=>"manual", Automatic=>"automatic", SemiAutomatic=>"semi-automatic" });
keyword_enum!(DrawingPageTransitionSpeed { Slow=>"slow", Medium=>"medium", Fast=>"fast" });
keyword_enum!(DrawingPageTransitionDirection { Forward=>"forward", Reverse=>"reverse" });
keyword_enum!(DrawingPageVisibility { Visible=>"visible", Hidden=>"hidden" });
keyword_enum!(DrawingPageBackgroundSize { Full=>"full", Border=>"border" });
keyword_enum!(DrawingPageSoundShow { New=>"new", Replace=>"replace" });

const TRANSITION_STYLES: &[&str] = &[
    "none",
    "fade-from-left",
    "fade-from-top",
    "fade-from-right",
    "fade-from-bottom",
    "fade-from-upperleft",
    "fade-from-upperright",
    "fade-from-lowerleft",
    "fade-from-lowerright",
    "move-from-left",
    "move-from-top",
    "move-from-right",
    "move-from-bottom",
    "move-from-upperleft",
    "move-from-upperright",
    "move-from-lowerleft",
    "move-from-lowerright",
    "uncover-to-left",
    "uncover-to-top",
    "uncover-to-right",
    "uncover-to-bottom",
    "uncover-to-upperleft",
    "uncover-to-upperright",
    "uncover-to-lowerleft",
    "uncover-to-lowerright",
    "fade-to-center",
    "fade-from-center",
    "vertical-stripes",
    "horizontal-stripes",
    "clockwise",
    "counterclockwise",
    "open-vertical",
    "open-horizontal",
    "close-vertical",
    "close-horizontal",
    "wavyline-from-left",
    "wavyline-from-top",
    "wavyline-from-right",
    "wavyline-from-bottom",
    "spiralin-left",
    "spiralin-right",
    "spiralout-left",
    "spiralout-right",
    "roll-from-top",
    "roll-from-left",
    "roll-from-right",
    "roll-from-bottom",
    "stretch-from-left",
    "stretch-from-top",
    "stretch-from-right",
    "stretch-from-bottom",
    "vertical-lines",
    "horizontal-lines",
    "dissolve",
    "random",
    "vertical-checkerboard",
    "horizontal-checkerboard",
    "interlocking-horizontal-left",
    "interlocking-horizontal-right",
    "interlocking-vertical-top",
    "interlocking-vertical-bottom",
    "fly-away",
    "open",
    "close",
    "melt",
];
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DrawingPageTransitionStyle(String);
impl DrawingPageTransitionStyle {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if !TRANSITION_STYLES.contains(&value.as_str()) {
            return Err(bad("invalid presentation:transition-style"));
        }
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DrawingPageTileRepeatOffset {
    pub percentage: String,
    pub direction: DrawingPageTileDirection,
}
impl DrawingPageTileRepeatOffset {
    pub fn new(percentage: impl Into<String>, direction: DrawingPageTileDirection) -> Result<Self> {
        let percentage = percentage.into();
        if !zero_to_hundred_percent(&percentage) {
            return Err(bad("invalid draw:tile-repeat-offset percentage"));
        }
        Ok(Self {
            percentage,
            direction,
        })
    }
}

/// Inert `presentation:sound`; no link is loaded or executed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DrawingPageSound {
    pub href: String,
    pub play_full: Option<bool>,
    pub actuate_on_request: bool,
    pub show: Option<DrawingPageSoundShow>,
    pub xml_id: Option<String>,
}
impl DrawingPageSound {
    pub fn new(href: impl Into<String>) -> Result<Self> {
        let value = Self {
            href: href.into(),
            play_full: None,
            actuate_on_request: false,
            show: None,
            xml_id: None,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn validate(&self) -> Result<()> {
        safe(&self.href, "xlink:href", true)?;
        if let Some(id) = &self.xml_id {
            ncname(id, "xml:id", false)?
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            r#"<presentation:sound xmlns:presentation="{PRESENTATION_NS}" xmlns:xlink="{XLINK_NS}" xlink:type="simple" xlink:href="{}""#,
            escape_xml(&self.href)
        );
        if self.actuate_on_request {
            xml.push_str(r#" xlink:actuate="onRequest""#)
        }
        if let Some(show) = self.show {
            xml.push_str(&format!(r#" xlink:show="{}""#, show.xml()))
        }
        if let Some(value) = self.play_full {
            xml.push_str(&format!(r#" presentation:play-full="{value}""#))
        }
        if let Some(id) = &self.xml_id {
            xml.push_str(&format!(r#" xml:id="{}""#, escape_xml(id)))
        }
        xml.push_str("/>");
        Ok(xml)
    }
}

/// Complete `style:drawing-page-properties` value.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DrawingPageStyleProperties {
    pub fill: Option<DrawingPageFill>,
    pub fill_color: Option<DrawingPageColor>,
    pub secondary_fill_color: Option<DrawingPageColor>,
    pub fill_gradient_name: Option<DrawingPageStyleNameRef>,
    pub gradient_step_count: Option<DrawingPageNonNegativeInteger>,
    pub fill_hatch_name: Option<DrawingPageStyleNameRef>,
    pub fill_hatch_solid: Option<bool>,
    pub fill_image_name: Option<DrawingPageStyleNameRef>,
    pub repeat: Option<DrawingPageRepeat>,
    pub fill_image_width: Option<DrawingPageLengthOrPercent>,
    pub fill_image_height: Option<DrawingPageLengthOrPercent>,
    pub fill_image_ref_point_x: Option<DrawingPagePercent>,
    pub fill_image_ref_point_y: Option<DrawingPagePercent>,
    pub fill_image_ref_point: Option<DrawingPageImageRefPoint>,
    pub tile_repeat_offset: Option<DrawingPageTileRepeatOffset>,
    pub opacity: Option<DrawingPagePercent>,
    pub opacity_name: Option<DrawingPageStyleNameRef>,
    pub fill_rule: Option<DrawingPageFillRule>,
    pub transition_type: Option<DrawingPageTransitionType>,
    pub transition_style: Option<DrawingPageTransitionStyle>,
    pub transition_speed: Option<DrawingPageTransitionSpeed>,
    pub smil_type: Option<String>,
    pub smil_subtype: Option<String>,
    pub direction: Option<DrawingPageTransitionDirection>,
    pub fade_color: Option<DrawingPageColor>,
    pub duration: Option<DrawingPageDuration>,
    pub visibility: Option<DrawingPageVisibility>,
    pub background_size: Option<DrawingPageBackgroundSize>,
    pub background_objects_visible: Option<bool>,
    pub background_visible: Option<bool>,
    pub display_header: Option<bool>,
    pub display_footer: Option<bool>,
    pub display_page_number: Option<bool>,
    pub display_date_time: Option<bool>,
    pub sound: Option<DrawingPageSound>,
}
impl DrawingPageStyleProperties {
    pub fn validate(&self) -> Result<()> {
        if let Some(value) = &self.smil_type {
            safe(value, "smil:type", true)?
        }
        if let Some(value) = &self.smil_subtype {
            safe(value, "smil:subtype", true)?
        }
        if let Some(value) = &self.sound {
            value.validate()?
        }
        Ok(())
    }
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
        attr!(self.fill, "draw:fill", |v: DrawingPageFill| v.xml());
        attr!(
            self.fill_color.as_ref(),
            "draw:fill-color",
            |v: &DrawingPageColor| v.as_str().to_owned()
        );
        attr!(
            self.secondary_fill_color.as_ref(),
            "draw:secondary-fill-color",
            |v: &DrawingPageColor| v.as_str().to_owned()
        );
        attr!(
            self.fill_gradient_name.as_ref(),
            "draw:fill-gradient-name",
            |v: &DrawingPageStyleNameRef| v.as_str().to_owned()
        );
        attr!(
            self.gradient_step_count.as_ref(),
            "draw:gradient-step-count",
            |v: &DrawingPageNonNegativeInteger| v.as_str().to_owned()
        );
        attr!(
            self.fill_hatch_name.as_ref(),
            "draw:fill-hatch-name",
            |v: &DrawingPageStyleNameRef| v.as_str().to_owned()
        );
        attr!(
            self.fill_hatch_solid,
            "draw:fill-hatch-solid",
            |v: bool| if v { "true" } else { "false" }
        );
        attr!(
            self.fill_image_name.as_ref(),
            "draw:fill-image-name",
            |v: &DrawingPageStyleNameRef| v.as_str().to_owned()
        );
        attr!(self.repeat, "style:repeat", |v: DrawingPageRepeat| v.xml());
        attr!(
            self.fill_image_width.as_ref(),
            "draw:fill-image-width",
            |v: &DrawingPageLengthOrPercent| v.as_str().to_owned()
        );
        attr!(
            self.fill_image_height.as_ref(),
            "draw:fill-image-height",
            |v: &DrawingPageLengthOrPercent| v.as_str().to_owned()
        );
        attr!(
            self.fill_image_ref_point_x.as_ref(),
            "draw:fill-image-ref-point-x",
            |v: &DrawingPagePercent| v.as_str().to_owned()
        );
        attr!(
            self.fill_image_ref_point_y.as_ref(),
            "draw:fill-image-ref-point-y",
            |v: &DrawingPagePercent| v.as_str().to_owned()
        );
        attr!(
            self.fill_image_ref_point,
            "draw:fill-image-ref-point",
            |v: DrawingPageImageRefPoint| v.xml()
        );
        if let Some(value) = &self.tile_repeat_offset {
            xml.push_str(&format!(
                r#" draw:tile-repeat-offset="{} {}""#,
                value.percentage,
                value.direction.xml()
            ))
        }
        attr!(
            self.opacity.as_ref(),
            "draw:opacity",
            |v: &DrawingPagePercent| v.as_str().to_owned()
        );
        attr!(
            self.opacity_name.as_ref(),
            "draw:opacity-name",
            |v: &DrawingPageStyleNameRef| v.as_str().to_owned()
        );
        attr!(self.fill_rule, "svg:fill-rule", |v: DrawingPageFillRule| v
            .xml());
        attr!(
            self.transition_type,
            "presentation:transition-type",
            |v: DrawingPageTransitionType| v.xml()
        );
        attr!(
            self.transition_style.as_ref(),
            "presentation:transition-style",
            |v: &DrawingPageTransitionStyle| v.as_str().to_owned()
        );
        attr!(
            self.transition_speed,
            "presentation:transition-speed",
            |v: DrawingPageTransitionSpeed| v.xml()
        );
        if let Some(value) = &self.smil_type {
            xml.push_str(&format!(r#" smil:type="{}""#, escape_xml(value)))
        }
        if let Some(value) = &self.smil_subtype {
            xml.push_str(&format!(r#" smil:subtype="{}""#, escape_xml(value)))
        }
        attr!(
            self.direction,
            "smil:direction",
            |v: DrawingPageTransitionDirection| v.xml()
        );
        attr!(
            self.fade_color.as_ref(),
            "smil:fadeColor",
            |v: &DrawingPageColor| v.as_str().to_owned()
        );
        attr!(
            self.duration.as_ref(),
            "presentation:duration",
            |v: &DrawingPageDuration| v.as_str().to_owned()
        );
        attr!(
            self.visibility,
            "presentation:visibility",
            |v: DrawingPageVisibility| v.xml()
        );
        attr!(
            self.background_size,
            "draw:background-size",
            |v: DrawingPageBackgroundSize| v.xml()
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
            xml.push_str("</style:drawing-page-properties>")
        } else {
            xml.push_str("/>")
        }
        Ok(xml)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DrawingPageStyle {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub properties: Option<DrawingPageStyleProperties>,
}
impl DrawingPageStyle {
    pub fn named(
        name: impl Into<String>,
        properties: Option<DrawingPageStyleProperties>,
    ) -> Result<Self> {
        let value = Self {
            name: Some(name.into()),
            parent_style_name: None,
            is_default_style: false,
            properties,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn default_style(properties: Option<DrawingPageStyleProperties>) -> Self {
        Self {
            name: None,
            parent_style_name: None,
            is_default_style: true,
            properties,
        }
    }
    pub fn validate(&self) -> Result<()> {
        match (&self.name, self.is_default_style) {
            (Some(value), false) => ncname(value, "drawing-page style name", false)?,
            (None, true) => {},
            _ => return Err(bad("invalid drawing-page style identity")),
        }
        if let Some(value) = &self.parent_style_name {
            if self.is_default_style {
                return Err(bad("default drawing-page style cannot have a parent"));
            }
            ncname(value, "parent drawing-page style name", false)?
        }
        if let Some(value) = &self.properties {
            value.validate()?
        }
        Ok(())
    }
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
            xml.push_str(&format!(r#" style:name="{}""#, escape_xml(value)))
        }
        if let Some(value) = &self.parent_style_name {
            xml.push_str(&format!(
                r#" style:parent-style-name="{}""#,
                escape_xml(value)
            ))
        }
        if let Some(value) = &self.properties {
            xml.push('>');
            xml.push_str(&value.to_xml_fragment()?);
            xml.push_str(&format!("</style:{tag}>"))
        } else {
            xml.push_str("/>")
        }
        Ok(xml)
    }
}
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DrawingPageStyleSet {
    pub styles: Vec<DrawingPageStyle>,
}
impl DrawingPageStyleSet {
    pub fn get(&self, name: &str) -> Option<&DrawingPageStyle> {
        self.styles
            .iter()
            .find(|style| style.name.as_deref() == Some(name))
    }
    pub fn default_style(&self) -> Option<&DrawingPageStyle> {
        self.styles.iter().find(|style| style.is_default_style)
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
        out.push((key.0, key.1, value))
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
) -> Result<Option<DrawingPageStyle>> {
    let mut attrs = attrs(reader, version, start)?;
    if take(&mut attrs, Ns::Style, b"family").as_deref() != Some("drawing-page") {
        return Ok(None);
    }
    let style = DrawingPageStyle {
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
) -> Result<DrawingPageStyleProperties> {
    let mut a = attrs(reader, version, start)?;
    let value = DrawingPageStyleProperties {
        fill: enum_value(take(&mut a, Ns::Draw, b"fill"), DrawingPageFill::parse)?,
        fill_color: take(&mut a, Ns::Draw, b"fill-color")
            .map(DrawingPageColor::new)
            .transpose()?,
        secondary_fill_color: take(&mut a, Ns::Draw, b"secondary-fill-color")
            .map(DrawingPageColor::new)
            .transpose()?,
        fill_gradient_name: take(&mut a, Ns::Draw, b"fill-gradient-name")
            .map(DrawingPageStyleNameRef::new)
            .transpose()?,
        gradient_step_count: take(&mut a, Ns::Draw, b"gradient-step-count")
            .map(DrawingPageNonNegativeInteger::new)
            .transpose()?,
        fill_hatch_name: take(&mut a, Ns::Draw, b"fill-hatch-name")
            .map(DrawingPageStyleNameRef::new)
            .transpose()?,
        fill_hatch_solid: take(&mut a, Ns::Draw, b"fill-hatch-solid")
            .map(|v| boolean(&v))
            .transpose()?,
        fill_image_name: take(&mut a, Ns::Draw, b"fill-image-name")
            .map(DrawingPageStyleNameRef::new)
            .transpose()?,
        repeat: enum_value(take(&mut a, Ns::Style, b"repeat"), DrawingPageRepeat::parse)?,
        fill_image_width: take(&mut a, Ns::Draw, b"fill-image-width")
            .map(DrawingPageLengthOrPercent::new)
            .transpose()?,
        fill_image_height: take(&mut a, Ns::Draw, b"fill-image-height")
            .map(DrawingPageLengthOrPercent::new)
            .transpose()?,
        fill_image_ref_point_x: take(&mut a, Ns::Draw, b"fill-image-ref-point-x")
            .map(DrawingPagePercent::new)
            .transpose()?,
        fill_image_ref_point_y: take(&mut a, Ns::Draw, b"fill-image-ref-point-y")
            .map(DrawingPagePercent::new)
            .transpose()?,
        fill_image_ref_point: enum_value(
            take(&mut a, Ns::Draw, b"fill-image-ref-point"),
            DrawingPageImageRefPoint::parse,
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
                DrawingPageTileRepeatOffset::new(
                    percentage,
                    DrawingPageTileDirection::parse(direction)?,
                )
            })
            .transpose()?,
        opacity: take(&mut a, Ns::Draw, b"opacity")
            .map(DrawingPagePercent::new)
            .transpose()?,
        opacity_name: take(&mut a, Ns::Draw, b"opacity-name")
            .map(DrawingPageStyleNameRef::new)
            .transpose()?,
        fill_rule: enum_value(
            take(&mut a, Ns::Svg, b"fill-rule"),
            DrawingPageFillRule::parse,
        )?,
        transition_type: enum_value(
            take(&mut a, Ns::Presentation, b"transition-type"),
            DrawingPageTransitionType::parse,
        )?,
        transition_style: take(&mut a, Ns::Presentation, b"transition-style")
            .map(DrawingPageTransitionStyle::new)
            .transpose()?,
        transition_speed: enum_value(
            take(&mut a, Ns::Presentation, b"transition-speed"),
            DrawingPageTransitionSpeed::parse,
        )?,
        smil_type: take(&mut a, Ns::Smil, b"type"),
        smil_subtype: take(&mut a, Ns::Smil, b"subtype"),
        direction: enum_value(
            take(&mut a, Ns::Smil, b"direction"),
            DrawingPageTransitionDirection::parse,
        )?,
        fade_color: take(&mut a, Ns::Smil, b"fadeColor")
            .map(DrawingPageColor::new)
            .transpose()?,
        duration: take(&mut a, Ns::Presentation, b"duration")
            .map(DrawingPageDuration::new)
            .transpose()?,
        visibility: enum_value(
            take(&mut a, Ns::Presentation, b"visibility"),
            DrawingPageVisibility::parse,
        )?,
        background_size: enum_value(
            take(&mut a, Ns::Draw, b"background-size"),
            DrawingPageBackgroundSize::parse,
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
) -> Result<DrawingPageSound> {
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
    let show = enum_value(
        take(&mut a, Ns::Xlink, b"show"),
        DrawingPageSoundShow::parse,
    )?;
    let play_full = take(&mut a, Ns::Presentation, b"play-full")
        .map(|v| boolean(&v))
        .transpose()?;
    let xml_id = take(&mut a, Ns::Xml, b"id");
    if !a.is_empty() {
        return Err(bad(
            "unknown presentation:sound attribute or wrong namespace",
        ));
    }
    let value = DrawingPageSound {
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
    style: DrawingPageStyle,
    seen: bool,
    properties_depth: Option<usize>,
    sound_depth: Option<usize>,
}
fn push_style(
    out: &mut Vec<DrawingPageStyle>,
    style: DrawingPageStyle,
    total: &mut usize,
) -> Result<()> {
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
pub fn parse_drawing_page_style_properties(xml: &str) -> Result<DrawingPageStyleSet> {
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
                        })
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
                        value.properties_depth = Some(depth)
                    } else if current.1 == b"drawing-page-properties" {
                        return Err(bad(
                            "style:drawing-page-properties has invalid namespace or parent",
                        ));
                    } else if value.properties_depth.is_some_and(|p| depth == p + 1)
                        && current.0 == Ns::Presentation
                        && current.1 == b"sound"
                    {
                        if value.style.properties.as_ref().unwrap().sound.is_some() {
                            return Err(bad("duplicate presentation:sound"));
                        }
                        value.style.properties.as_mut().unwrap().sound =
                            Some(parse_sound(&reader, version, &start)?);
                        value.sound_depth = Some(depth)
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
                        push_style(&mut out, style, &mut total)?
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
                        value.style.properties = Some(parse_properties(&reader, version, &start)?)
                    } else if current.1 == b"drawing-page-properties" {
                        return Err(bad(
                            "style:drawing-page-properties has invalid namespace or parent",
                        ));
                    } else if value.properties_depth.is_some_and(|p| depth == p + 1)
                        && current.0 == Ns::Presentation
                        && current.1 == b"sound"
                    {
                        if value.style.properties.as_ref().unwrap().sound.is_some() {
                            return Err(bad("duplicate presentation:sound"));
                        }
                        value.style.properties.as_mut().unwrap().sound =
                            Some(parse_sound(&reader, version, &start)?)
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
                        value.sound_depth = None
                    }
                    if value.properties_depth == Some(depth) {
                        value.properties_depth = None
                    }
                }
                if active.as_ref().is_some_and(|value| value.depth == depth) {
                    push_style(&mut out, active.take().unwrap().style, &mut total)?
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
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
    Ok(DrawingPageStyleSet { styles: out })
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
pub fn set_drawing_page_style_properties_xml(
    xml: &str,
    requested: &DrawingPageStyle,
) -> Result<String> {
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
                        })
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
                    if active.as_mut().unwrap().properties.replace(span).is_some() {
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
                        })
                    }
                } else if target_depth.is_some_and(|value| depth == value + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"drawing-page-properties"
                {
                    if active.as_mut().unwrap().properties.replace(span).is_some() {
                        return Err(bad("duplicate style:drawing-page-properties"));
                    }
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
                        let span = spans.properties.as_mut().unwrap();
                        span.end_start = begin;
                        span.end = end
                    }
                    if target_depth == Some(depth) {
                        spans.style.end_start = begin;
                        spans.style.end = end;
                        found = active.take();
                        target_depth = None
                    }
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
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
        .map(DrawingPageStyleProperties::to_xml_fragment)
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

impl OpenDocumentPackage {
    pub fn drawing_page_style_properties(&self) -> Result<DrawingPageStyleSet> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |xml| parse_drawing_page_style_properties(&xml),
        )
    }
}
impl FlatOpenDocument {
    pub fn drawing_page_style_properties(&self) -> Result<DrawingPageStyleSet> {
        parse_drawing_page_style_properties(self.xml())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    const HEAD: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:automatic-styles>"#;
    fn doc(body: &str) -> String {
        format!("{HEAD}{body}</office:automatic-styles></office:document>")
    }
    #[test]
    fn complete_family_round_trips() {
        let xml = doc(
            r##"<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties draw:fill="gradient" draw:fill-color="#A0b1C2" draw:secondary-fill-color="#010203" draw:fill-gradient-name="g1" draw:gradient-step-count="00012" draw:fill-hatch-name="h1" draw:fill-hatch-solid="true" draw:fill-image-name="" style:repeat="stretch" draw:fill-image-width="0cm" draw:fill-image-height="-.5%" draw:fill-image-ref-point-x="-12.5%" draw:fill-image-ref-point-y=".5%" draw:fill-image-ref-point="bottom-right" draw:tile-repeat-offset="100.0% vertical" draw:opacity="120%" draw:opacity-name="o1" svg:fill-rule="evenodd" presentation:transition-type="semi-automatic" presentation:transition-style="melt" presentation:transition-speed="fast" smil:type="fade &amp; dissolve" smil:subtype="crossfade" smil:direction="reverse" smil:fadeColor="#aB09fF" presentation:duration="P1Y2M3DT4H5M6.50S" presentation:visibility="hidden" draw:background-size="border" presentation:background-objects-visible="false" presentation:background-visible="true" presentation:display-header="false" presentation:display-footer="true" presentation:display-page-number="false" presentation:display-date-time="true"><presentation:sound xlink:type="simple" xlink:href="https://example.invalid/a&amp;b.ogg" xlink:actuate="onRequest" xlink:show="new" presentation:play-full="true" xml:id="sound_1"/></style:drawing-page-properties></style:style>"##,
        );
        let set = parse_drawing_page_style_properties(&xml).unwrap();
        let p = set.get("dp1").unwrap().properties.as_ref().unwrap();
        assert_eq!(p.gradient_step_count.as_ref().unwrap().as_str(), "00012");
        assert_eq!(
            p.sound.as_ref().unwrap().href,
            "https://example.invalid/a&b.ogg"
        );
        let fragment = p.to_xml_fragment().unwrap();
        assert_eq!(
            DrawingPageStyleProperties::from_xml_fragment(&fragment).unwrap(),
            *p
        )
    }
    #[test]
    fn parses_real_libreoffice_remote_background_without_loading_it() {
        let xml = include_str!(
            "../../../test-data/libreoffice-core/sd/qa/unit/tiledrendering/data/slide-background-link.fodp"
        );
        let set = parse_drawing_page_style_properties(xml).unwrap();
        let p = set.get("dp1").unwrap().properties.as_ref().unwrap();
        assert_eq!(p.fill, Some(DrawingPageFill::Bitmap));
        assert_eq!(p.fill_image_name.as_ref().unwrap().as_str(), "remote_bg");
        assert_eq!(p.repeat, Some(DrawingPageRepeat::Stretch))
    }
    #[test]
    fn lossless_replace_insert_and_remove() {
        let original = doc(
            "<!--keep--><style:style style:name=\"a\" style:family=\"drawing-page\"><x:keep xmlns:x=\"urn:keep\"/></style:style><style:style style:name=\"b\" style:family=\"drawing-page\"><style:drawing-page-properties draw:fill=\"none\"/></style:style>",
        );
        let mut a = DrawingPageStyle::named(
            "a",
            Some(DrawingPageStyleProperties {
                fill: Some(DrawingPageFill::Solid),
                ..Default::default()
            }),
        )
        .unwrap();
        let inserted = set_drawing_page_style_properties_xml(&original, &a).unwrap();
        assert!(inserted.contains("<!--keep--><style:style"));
        assert!(inserted.contains("<x:keep xmlns:x=\"urn:keep\"/><style:drawing-page-properties"));
        a.properties = None;
        let removed_a = set_drawing_page_style_properties_xml(&inserted, &a).unwrap();
        assert_eq!(removed_a, original);
        let b = DrawingPageStyle::named("b", None).unwrap();
        let removed = set_drawing_page_style_properties_xml(&removed_a, &b).unwrap();
        assert!(!removed.contains("draw:fill=\"none\""));
        assert!(removed.contains("<!--keep-->"))
    }
    #[test]
    fn rejects_malformed_namespaces_lexicals_duplicates_and_children() {
        let cases = [
            r#"<style:style style:name="a" style:family="drawing-page"><style:drawing-page-properties presentation:display-header="1"/></style:style>"#,
            r##"<style:style style:name="a" style:family="drawing-page"><style:drawing-page-properties draw:fill-color="#fff"/></style:style>"##,
            r#"<style:style style:name="a" style:family="drawing-page"><style:drawing-page-properties draw:tile-repeat-offset="101% horizontal"/></style:style>"#,
            r#"<style:style style:name="a" style:family="drawing-page"><style:drawing-page-properties draw:fill="none" draw:fill="solid"/></style:style>"#,
            r#"<style:style style:name="a" style:family="drawing-page"><style:drawing-page-properties><draw:sound/></style:drawing-page-properties></style:style>"#,
            r#"<style:style style:name="a" style:family="drawing-page"><draw:drawing-page-properties/></style:style>"#,
            r#"<style:style style:name="a" style:family="drawing-page"><style:drawing-page-properties><presentation:sound xlink:type="simple" xlink:href="x"><presentation:sound xlink:type="simple" xlink:href="y"/></presentation:sound></style:drawing-page-properties></style:style>"#,
        ];
        for case in cases {
            assert!(
                parse_drawing_page_style_properties(&doc(case)).is_err(),
                "accepted {case}"
            )
        }
    }
}
