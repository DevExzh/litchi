//! Typed ODF drawing-page property values and lexical invariants.

use litchi_core::{Error, Result};

const MAX_VALUE: usize = 1024 * 1024;

pub(super) fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
pub(super) fn safe(value: &str, name: &str, empty: bool) -> Result<()> {
    if (!empty && value.is_empty()) || value.len() > MAX_VALUE || value.chars().any(|c| matches!(c, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}' | '\u{fffe}' | '\u{ffff}')) {
        return Err(bad(format!("invalid {name}")));
    }
    Ok(())
}
pub(super) fn ncname(value: &str, name: &str, empty: bool) -> Result<()> {
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
    Color,
    |x: &str| x.len() == 7 && x.starts_with('#') && x[1..].bytes().all(|b| b.is_ascii_hexdigit()),
    "ODF color",
    false
);
lexical!(Percent, percent, "ODF percentage", false);
lexical!(
    LengthOrPercent,
    |x: &str| length(x) || percent(x),
    "ODF length or percentage",
    false
);
lexical!(Duration, duration, "presentation:duration", false);
lexical!(
    StyleNameRef,
    |x: &str| ncname(x, "style name reference", true).is_ok(),
    "style name reference",
    true
);
lexical!(
    NonNegativeInteger,
    |x: &str| !x.is_empty() && x.bytes().all(|b| b.is_ascii_digit()),
    "non-negative integer",
    false
);

macro_rules! keyword_enum {
    ($name:ident { $($variant:ident => $value:literal),+ $(,)? }) => {
        #[derive(Debug, Clone, Copy, PartialEq, Eq)] pub enum $name { $($variant),+ }
        impl $name {
            pub(super) fn parse(value:&str)->Result<Self>{match value{$($value=>Ok(Self::$variant),)+_=>Err(bad(concat!("invalid ",stringify!($name))))}}
            pub(super) fn xml(self)->&'static str{match self{$(Self::$variant=>$value),+}}
        }
    };
}
keyword_enum!(Fill { None=>"none", Solid=>"solid", Bitmap=>"bitmap", Gradient=>"gradient", Hatch=>"hatch" });
keyword_enum!(Repeat { NoRepeat=>"no-repeat", Repeat=>"repeat", Stretch=>"stretch" });
keyword_enum!(ImageRefPoint { TopLeft=>"top-left", Top=>"top", TopRight=>"top-right", Left=>"left", Center=>"center", Right=>"right", BottomLeft=>"bottom-left", Bottom=>"bottom", BottomRight=>"bottom-right" });
keyword_enum!(TileDirection { Horizontal=>"horizontal", Vertical=>"vertical" });
keyword_enum!(FillRule { Nonzero=>"nonzero", Evenodd=>"evenodd" });
keyword_enum!(TransitionType { Manual=>"manual", Automatic=>"automatic", SemiAutomatic=>"semi-automatic" });
keyword_enum!(TransitionSpeed { Slow=>"slow", Medium=>"medium", Fast=>"fast" });
keyword_enum!(TransitionDirection { Forward=>"forward", Reverse=>"reverse" });
keyword_enum!(Visibility { Visible=>"visible", Hidden=>"hidden" });
keyword_enum!(BackgroundSize { Full=>"full", Border=>"border" });
keyword_enum!(SoundShow { New=>"new", Replace=>"replace" });

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
pub struct TransitionStyle(String);
impl TransitionStyle {
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
pub struct TileRepeatOffset {
    pub percentage: String,
    pub direction: TileDirection,
}
impl TileRepeatOffset {
    pub fn new(percentage: impl Into<String>, direction: TileDirection) -> Result<Self> {
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
pub struct Sound {
    pub href: String,
    pub play_full: Option<bool>,
    pub actuate_on_request: bool,
    pub show: Option<SoundShow>,
    pub xml_id: Option<String>,
}
impl Sound {
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
            ncname(id, "xml:id", false)?;
        }
        Ok(())
    }
}

/// Complete `style:drawing-page-properties` value.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct StyleProperties {
    pub fill: Option<Fill>,
    pub fill_color: Option<Color>,
    pub secondary_fill_color: Option<Color>,
    pub fill_gradient_name: Option<StyleNameRef>,
    pub gradient_step_count: Option<NonNegativeInteger>,
    pub fill_hatch_name: Option<StyleNameRef>,
    pub fill_hatch_solid: Option<bool>,
    pub fill_image_name: Option<StyleNameRef>,
    pub repeat: Option<Repeat>,
    pub fill_image_width: Option<LengthOrPercent>,
    pub fill_image_height: Option<LengthOrPercent>,
    pub fill_image_ref_point_x: Option<Percent>,
    pub fill_image_ref_point_y: Option<Percent>,
    pub fill_image_ref_point: Option<ImageRefPoint>,
    pub tile_repeat_offset: Option<TileRepeatOffset>,
    pub opacity: Option<Percent>,
    pub opacity_name: Option<StyleNameRef>,
    pub fill_rule: Option<FillRule>,
    pub transition_type: Option<TransitionType>,
    pub transition_style: Option<TransitionStyle>,
    pub transition_speed: Option<TransitionSpeed>,
    pub smil_type: Option<String>,
    pub smil_subtype: Option<String>,
    pub direction: Option<TransitionDirection>,
    pub fade_color: Option<Color>,
    pub duration: Option<Duration>,
    pub visibility: Option<Visibility>,
    pub background_size: Option<BackgroundSize>,
    pub background_objects_visible: Option<bool>,
    pub background_visible: Option<bool>,
    pub display_header: Option<bool>,
    pub display_footer: Option<bool>,
    pub display_page_number: Option<bool>,
    pub display_date_time: Option<bool>,
    pub sound: Option<Sound>,
}
impl StyleProperties {
    pub fn validate(&self) -> Result<()> {
        if let Some(value) = &self.smil_type {
            safe(value, "smil:type", true)?;
        }
        if let Some(value) = &self.smil_subtype {
            safe(value, "smil:subtype", true)?;
        }
        if let Some(value) = &self.sound {
            value.validate()?;
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Style {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub properties: Option<StyleProperties>,
}
impl Style {
    pub fn named(name: impl Into<String>, properties: Option<StyleProperties>) -> Result<Self> {
        let value = Self {
            name: Some(name.into()),
            parent_style_name: None,
            is_default_style: false,
            properties,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn default_style(properties: Option<StyleProperties>) -> Self {
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
            ncname(value, "parent drawing-page style name", false)?;
        }
        if let Some(value) = &self.properties {
            value.validate()?;
        }
        Ok(())
    }
}
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Styles {
    pub styles: Vec<Style>,
}
impl Styles {
    pub fn get(&self, name: &str) -> Option<&Style> {
        self.styles
            .iter()
            .find(|style| style.name.as_deref() == Some(name))
    }
    pub fn default_style(&self) -> Option<&Style> {
        self.styles.iter().find(|style| style.is_default_style)
    }
}
