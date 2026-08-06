//! Typed semantic values for Word 2010 run-property effects.
//!
//! The owner deliberately models only the seven visual effects from
//! `[MS-DOCX]` §2.2.1.  OpenType run extensions and later Word namespaces are
//! retained as bounded opaque children by [`Effect::Unknown`].

use crate::error::{Error, Result};

use super::validation;

/// Maximum number of direct effect children retained on one `w:rPr`.
pub const MAX_EFFECTS: usize = 64;
/// Maximum bytes retained for one unsupported extension element.
pub const MAX_OPAQUE_BYTES: usize = 256 * 1024;
/// Maximum number of color transforms on one DrawingML color.
pub const MAX_COLOR_TRANSFORMS: usize = 32;
/// Maximum gradient stops accepted by the Word 2010 schema.
pub const MAX_GRADIENT_STOPS: usize = 10;

/// Ordered Word 2010 run-property effects.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[must_use]
pub struct RunEffects {
    pub(super) values: Vec<Effect>,
}

impl RunEffects {
    /// Create an empty effect collection.
    #[inline]
    pub const fn new() -> Self {
        Self { values: Vec::new() }
    }

    /// Number of typed and opaque direct children.
    #[inline]
    pub fn len(&self) -> usize {
        self.values.len()
    }

    /// Whether the collection has no direct children.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.values.is_empty()
    }

    /// Iterate effects in source order.
    #[inline]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Effect> {
        self.values.iter()
    }

    /// Return the first glow effect, if present.
    pub fn glow(&self) -> Option<&Glow> {
        self.values.iter().find_map(|effect| match effect {
            Effect::Glow(value) => Some(value),
            _ => None,
        })
    }

    /// Return the first shadow effect, if present.
    pub fn shadow(&self) -> Option<&Shadow> {
        self.values.iter().find_map(|effect| match effect {
            Effect::Shadow(value) => Some(value),
            _ => None,
        })
    }

    /// Return the first reflection effect, if present.
    pub fn reflection(&self) -> Option<&Reflection> {
        self.values.iter().find_map(|effect| match effect {
            Effect::Reflection(value) => Some(value),
            _ => None,
        })
    }

    /// Return the first text outline effect, if present.
    pub fn text_outline(&self) -> Option<&TextOutline> {
        self.values.iter().find_map(|effect| match effect {
            Effect::TextOutline(value) => Some(value),
            _ => None,
        })
    }

    /// Return the first text fill effect, if present.
    pub fn text_fill(&self) -> Option<&TextFill> {
        self.values.iter().find_map(|effect| match effect {
            Effect::TextFill(value) => Some(value),
            _ => None,
        })
    }

    /// Return the first 3-D scene effect, if present.
    pub fn scene3d(&self) -> Option<&Scene3d> {
        self.values.iter().find_map(|effect| match effect {
            Effect::Scene3d(value) => Some(value),
            _ => None,
        })
    }

    /// Return the first 3-D text-properties effect, if present.
    pub fn props3d(&self) -> Option<&Props3d> {
        self.values.iter().find_map(|effect| match effect {
            Effect::Props3d(value) => Some(value),
            _ => None,
        })
    }

    /// Iterate unsupported or foreign extension children in source order.
    pub fn unknown(&self) -> impl Iterator<Item = &OpaqueExtension> {
        self.values.iter().filter_map(|effect| match effect {
            Effect::Unknown(value) => Some(value),
            _ => None,
        })
    }

    /// Append an effect after validating cardinality and duplicate rules.
    pub fn push(&mut self, effect: Effect) -> Result<&mut Self> {
        if self.values.len() >= MAX_EFFECTS {
            return Err(Error::Invalid(format!(
                "Word run effects exceed {MAX_EFFECTS} children"
            )));
        }
        if effect.kind().is_known()
            && self
                .values
                .iter()
                .any(|existing| existing.kind() == effect.kind())
        {
            return Err(Error::Invalid(format!(
                "duplicate Word run effect '{}'",
                effect.kind().as_str()
            )));
        }
        effect.validate()?;
        self.values.push(effect);
        Ok(self)
    }

    /// Append a bounded unsupported extension child.
    pub fn push_unknown(&mut self, value: OpaqueExtension) -> Result<&mut Self> {
        self.push(Effect::Unknown(value))
    }

    /// Replace or remove one typed effect without disturbing opaque order.
    pub fn set_glow(&mut self, value: Option<Glow>) -> Result<&mut Self> {
        self.replace(EffectKind::Glow, value.map(Effect::Glow))
    }

    /// Replace or remove one typed effect without disturbing opaque order.
    pub fn set_shadow(&mut self, value: Option<Shadow>) -> Result<&mut Self> {
        self.replace(EffectKind::Shadow, value.map(Effect::Shadow))
    }

    /// Replace or remove one typed effect without disturbing opaque order.
    pub fn set_reflection(&mut self, value: Option<Reflection>) -> Result<&mut Self> {
        self.replace(EffectKind::Reflection, value.map(Effect::Reflection))
    }

    /// Replace or remove one typed effect without disturbing opaque order.
    pub fn set_text_outline(&mut self, value: Option<TextOutline>) -> Result<&mut Self> {
        self.replace(EffectKind::TextOutline, value.map(Effect::TextOutline))
    }

    /// Replace or remove one typed effect without disturbing opaque order.
    pub fn set_text_fill(&mut self, value: Option<TextFill>) -> Result<&mut Self> {
        self.replace(EffectKind::TextFill, value.map(Effect::TextFill))
    }

    /// Replace or remove one typed effect without disturbing opaque order.
    pub fn set_scene3d(&mut self, value: Option<Scene3d>) -> Result<&mut Self> {
        self.replace(EffectKind::Scene3d, value.map(Effect::Scene3d))
    }

    /// Replace or remove one typed effect without disturbing opaque order.
    pub fn set_props3d(&mut self, value: Option<Props3d>) -> Result<&mut Self> {
        self.replace(EffectKind::Props3d, value.map(Effect::Props3d))
    }

    /// Validate every typed and opaque child.
    pub fn validate(&self) -> Result<()> {
        validation::validate(self)
    }

    /// Parse a complete `w:r` or `w:rPr` XML fragment.
    pub fn parse(xml: &[u8]) -> Result<Self> {
        super::codec::parse(xml)
    }

    fn replace(&mut self, kind: EffectKind, value: Option<Effect>) -> Result<&mut Self> {
        if let Some(index) = self
            .values
            .iter()
            .position(|existing| existing.kind() == kind)
        {
            if let Some(value) = value {
                value.validate()?;
                self.values[index] = value;
            } else {
                self.values.remove(index);
            }
            return Ok(self);
        }
        if let Some(value) = value {
            self.push(value)?;
        }
        Ok(self)
    }
}

/// One typed or opaque direct child of `w:rPr`.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Effect {
    /// `w14:glow`.
    Glow(Glow),
    /// `w14:shadow`.
    Shadow(Shadow),
    /// `w14:reflection`.
    Reflection(Reflection),
    /// `w14:textOutline`.
    TextOutline(TextOutline),
    /// `w14:textFill`.
    TextFill(TextFill),
    /// `w14:scene3d`.
    Scene3d(Scene3d),
    /// `w14:props3d`.
    Props3d(Props3d),
    /// A bounded unsupported or foreign extension element.
    Unknown(OpaqueExtension),
}

impl Effect {
    /// Return the contextual effect kind.
    pub const fn kind(&self) -> EffectKind {
        match self {
            Self::Glow(_) => EffectKind::Glow,
            Self::Shadow(_) => EffectKind::Shadow,
            Self::Reflection(_) => EffectKind::Reflection,
            Self::TextOutline(_) => EffectKind::TextOutline,
            Self::TextFill(_) => EffectKind::TextFill,
            Self::Scene3d(_) => EffectKind::Scene3d,
            Self::Props3d(_) => EffectKind::Props3d,
            Self::Unknown(_) => EffectKind::Unknown,
        }
    }

    pub(crate) fn validate(&self) -> Result<()> {
        match self {
            Self::Glow(value) => value.validate(),
            Self::Shadow(value) => value.validate(),
            Self::Reflection(value) => value.validate(),
            Self::TextOutline(value) => value.validate(),
            Self::TextFill(value) => value.validate(),
            Self::Scene3d(value) => value.validate(),
            Self::Props3d(value) => value.validate(),
            Self::Unknown(value) => validation::validate_opaque(value),
        }
    }
}

/// The exclusive direct-effect kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EffectKind {
    Glow,
    Shadow,
    Reflection,
    TextOutline,
    TextFill,
    Scene3d,
    Props3d,
    Unknown,
}

impl EffectKind {
    pub const fn is_known(self) -> bool {
        !matches!(self, Self::Unknown)
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Glow => "glow",
            Self::Shadow => "shadow",
            Self::Reflection => "reflection",
            Self::TextOutline => "textOutline",
            Self::TextFill => "textFill",
            Self::Scene3d => "scene3d",
            Self::Props3d => "props3d",
            Self::Unknown => "unknown",
        }
    }
}

/// Bounded raw XML for an unsupported direct extension child.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct OpaqueExtension {
    xml: Box<[u8]>,
}

impl OpaqueExtension {
    /// Retain one complete XML element after applying the byte bound.
    pub fn new(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        if xml.is_empty() {
            return Err(Error::Invalid("opaque run effect cannot be empty".into()));
        }
        if xml.len() > MAX_OPAQUE_BYTES {
            return Err(Error::Invalid(format!(
                "opaque run effect exceeds {MAX_OPAQUE_BYTES} bytes"
            )));
        }
        Ok(Self {
            xml: xml.into_boxed_slice(),
        })
    }

    /// Borrow the exact retained XML bytes.
    #[inline]
    pub fn as_bytes(&self) -> &[u8] {
        &self.xml
    }
}

/// A Word/DrawingML RGB or theme-bound color.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Color {
    /// `srgbClr@val`.
    Rgb(RgbColor),
    /// `schemeClr@val`.
    Scheme(SchemeColor),
}

/// An RGB color and its ordered DrawingML transforms.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RgbColor {
    pub value: [u8; 3],
    pub transforms: Vec<ColorTransform>,
}

impl RgbColor {
    /// Construct an RGB color without transforms.
    pub const fn new(value: [u8; 3]) -> Self {
        Self {
            value,
            transforms: Vec::new(),
        }
    }
}

/// A scheme color and its ordered DrawingML transforms.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SchemeColor {
    pub value: SchemeColorValue,
    pub transforms: Vec<ColorTransform>,
}

/// The closed `ST_SchemeColorVal` domain used by Word 2010 effects.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SchemeColorValue {
    Background1,
    Text1,
    Background2,
    Text2,
    Dark1,
    Light1,
    Dark2,
    Light2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
    Placeholder,
}

impl SchemeColorValue {
    pub(crate) fn from_xml(value: &str) -> Option<Self> {
        Some(match value {
            "bg1" => Self::Background1,
            "tx1" => Self::Text1,
            "bg2" => Self::Background2,
            "tx2" => Self::Text2,
            "dk1" => Self::Dark1,
            "lt1" => Self::Light1,
            "dk2" => Self::Dark2,
            "lt2" => Self::Light2,
            "accent1" => Self::Accent1,
            "accent2" => Self::Accent2,
            "accent3" => Self::Accent3,
            "accent4" => Self::Accent4,
            "accent5" => Self::Accent5,
            "accent6" => Self::Accent6,
            "hlink" => Self::Hyperlink,
            "folHlink" => Self::FollowedHyperlink,
            "phClr" => Self::Placeholder,
            _ => return None,
        })
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Background1 => "bg1",
            Self::Text1 => "tx1",
            Self::Background2 => "bg2",
            Self::Text2 => "tx2",
            Self::Dark1 => "dk1",
            Self::Light1 => "lt1",
            Self::Dark2 => "dk2",
            Self::Light2 => "lt2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hlink",
            Self::FollowedHyperlink => "folHlink",
            Self::Placeholder => "phClr",
        }
    }
}

/// One DrawingML color transform, represented in thousandths of a percent.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColorTransform {
    Tint(u32),
    Shade(u32),
    Alpha(u32),
    HueMod(u32),
    Saturation(i32),
    SaturationOffset(i32),
    SaturationMod(u32),
    Luminance(i32),
    LuminanceOffset(i32),
    LuminanceMod(u32),
}

/// Glow around text.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Glow {
    pub color: Option<Color>,
    pub radius: Option<u64>,
}

/// Shadow behind text.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Shadow {
    pub color: Option<Color>,
    pub blur_radius: Option<u64>,
    pub distance: Option<u64>,
    pub direction: Option<u32>,
    pub scale_x: Option<i32>,
    pub scale_y: Option<i32>,
    pub skew_x: Option<i32>,
    pub skew_y: Option<i32>,
    pub alignment: Option<RectAlignment>,
}

/// Reflection below or beside text.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Reflection {
    pub blur_radius: Option<u64>,
    pub start_alpha: Option<u32>,
    pub start_position: Option<u32>,
    pub end_alpha: Option<u32>,
    pub end_position: Option<u32>,
    pub distance: Option<u64>,
    pub direction: Option<u32>,
    pub fade_direction: Option<u32>,
    pub scale_x: Option<i32>,
    pub scale_y: Option<i32>,
    pub skew_x: Option<i32>,
    pub skew_y: Option<i32>,
    pub alignment: Option<RectAlignment>,
}

/// Text fill effect.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TextFill {
    /// `None` represents an empty `textFill`, whose schema default is black.
    pub fill: Option<Fill>,
}

/// Text outline effect.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TextOutline {
    pub fill: Option<Fill>,
    pub dash: Option<LineDash>,
    pub join: Option<LineJoin>,
    pub width: Option<u64>,
    pub cap: Option<LineCap>,
    pub compound: Option<CompoundLine>,
    pub alignment: Option<PenAlignment>,
}

/// Fill choice shared by text fill and text outline.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Fill {
    NoFill,
    Solid(Option<Color>),
    Gradient(Gradient),
}

/// Gradient stop list and shading mode.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Gradient {
    pub stops: Vec<GradientStop>,
    pub shade: Option<Shade>,
}

/// One gradient stop, with a 0..100000 position.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GradientStop {
    pub position: u32,
    pub color: Color,
}

/// Gradient shading mode.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Shade {
    Linear {
        angle: Option<u32>,
        scaled: Option<bool>,
    },
    Path {
        path: Option<PathKind>,
        fill_to: Option<RelativeRect>,
    },
}

/// Relative gradient focus rectangle.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct RelativeRect {
    pub left: Option<i32>,
    pub top: Option<i32>,
    pub right: Option<i32>,
    pub bottom: Option<i32>,
}

/// Preset path used by a path gradient.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PathKind {
    Shape,
    Circle,
    Rect,
}

/// Preset line dash.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineDash {
    Solid,
    Dot,
    SysDot,
    Dash,
    SysDash,
    LargeDash,
    DashDot,
    SysDashDot,
    LargeDashDot,
    LargeDashDotDot,
    SysDashDotDot,
}

/// Line join choice for a text outline.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LineJoin {
    Round,
    Bevel,
    Miter { limit: Option<u32> },
}

/// Line ending cap.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineCap {
    Flat,
    Round,
    Square,
}

/// Compound line choice.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CompoundLine {
    Single,
    Double,
    ThickThin,
    ThinThick,
    Triple,
}

/// Pen alignment for an outline.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PenAlignment {
    Center,
    Inside,
}

/// Closed rectangle-alignment token used by shadow and reflection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RectAlignment {
    None,
    TopLeft,
    Top,
    TopRight,
    Left,
    Center,
    Right,
    BottomLeft,
    Bottom,
    BottomRight,
}

/// A 3-D scene containing a required camera and light rig.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Scene3d {
    pub camera: Camera,
    pub light_rig: LightRig,
}

/// A camera preset from the Word 2010 closed domain.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Camera {
    pub preset: PresetCamera,
}

/// A light rig and optional spherical rotation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LightRig {
    pub rig: LightRigType,
    pub direction: LightRigDirection,
    pub rotation: Option<SphereCoords>,
}

/// Required light-rig rotation coordinates.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SphereCoords {
    pub latitude: u32,
    pub longitude: u32,
    pub revolution: u32,
}

/// 3-D extrusion, contour, bevel, and material settings.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Props3d {
    pub bevel_top: Option<Bevel>,
    pub bevel_bottom: Option<Bevel>,
    pub extrusion_color: Option<Color>,
    pub contour_color: Option<Color>,
    pub extrusion_height: Option<u64>,
    pub contour_width: Option<u64>,
    pub material: Option<PresetMaterial>,
}

/// One text bevel.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Bevel {
    pub width: Option<u64>,
    pub height: Option<u64>,
    pub preset: Option<BevelPreset>,
}

macro_rules! token_type {
    ($name:ident, $values:ident) => {
        #[derive(Debug, Clone, PartialEq, Eq)]
        pub struct $name(Box<str>);

        impl $name {
            pub fn new(value: impl AsRef<str>) -> Result<Self> {
                let value = value.as_ref();
                if !$values.contains(&value) {
                    return Err(Error::Invalid(format!(
                        "invalid {} token '{value}'",
                        stringify!($name)
                    )));
                }
                Ok(Self(value.into()))
            }

            pub(crate) fn parse(value: &str) -> Result<Self> {
                Self::new(value)
            }

            #[inline]
            pub fn as_str(&self) -> &str {
                &self.0
            }
        }
    };
}

const CAMERA_PRESETS: &[&str] = &[
    "legacyObliqueTopLeft",
    "legacyObliqueTop",
    "legacyObliqueTopRight",
    "legacyObliqueLeft",
    "legacyObliqueFront",
    "legacyObliqueRight",
    "legacyObliqueBottomLeft",
    "legacyObliqueBottom",
    "legacyObliqueBottomRight",
    "legacyPerspectiveTopLeft",
    "legacyPerspectiveTop",
    "legacyPerspectiveTopRight",
    "legacyPerspectiveLeft",
    "legacyPerspectiveFront",
    "legacyPerspectiveRight",
    "legacyPerspectiveBottomLeft",
    "legacyPerspectiveBottom",
    "legacyPerspectiveBottomRight",
    "orthographicFront",
    "isometricTopUp",
    "isometricTopDown",
    "isometricBottomUp",
    "isometricBottomDown",
    "isometricLeftUp",
    "isometricLeftDown",
    "isometricRightUp",
    "isometricRightDown",
    "isometricOffAxis1Left",
    "isometricOffAxis1Right",
    "isometricOffAxis1Top",
    "isometricOffAxis2Left",
    "isometricOffAxis2Right",
    "isometricOffAxis2Top",
    "isometricOffAxis3Left",
    "isometricOffAxis3Right",
    "isometricOffAxis3Bottom",
    "isometricOffAxis4Left",
    "isometricOffAxis4Right",
    "isometricOffAxis4Bottom",
    "obliqueTopLeft",
    "obliqueTop",
    "obliqueTopRight",
    "obliqueLeft",
    "obliqueRight",
    "obliqueBottomLeft",
    "obliqueBottom",
    "obliqueBottomRight",
    "perspectiveFront",
    "perspectiveLeft",
    "perspectiveRight",
    "perspectiveAbove",
    "perspectiveBelow",
    "perspectiveAboveLeftFacing",
    "perspectiveAboveRightFacing",
    "perspectiveContrastingLeftFacing",
    "perspectiveContrastingRightFacing",
    "perspectiveHeroicLeftFacing",
    "perspectiveHeroicRightFacing",
    "perspectiveHeroicExtremeLeftFacing",
    "perspectiveHeroicExtremeRightFacing",
    "perspectiveRelaxed",
    "perspectiveRelaxedModerately",
];
const LIGHT_RIGS: &[&str] = &[
    "legacyFlat1",
    "legacyFlat2",
    "legacyFlat3",
    "legacyFlat4",
    "legacyNormal1",
    "legacyNormal2",
    "legacyNormal3",
    "legacyNormal4",
    "legacyHarsh1",
    "legacyHarsh2",
    "legacyHarsh3",
    "legacyHarsh4",
    "threePt",
    "balanced",
    "soft",
    "harsh",
    "flood",
    "contrasting",
    "morning",
    "sunrise",
    "sunset",
    "chilly",
    "freezing",
    "flat",
    "twoPt",
    "glow",
    "brightRoom",
];
const LIGHT_DIRECTIONS: &[&str] = &["tl", "t", "tr", "l", "r", "bl", "b", "br"];
const MATERIALS: &[&str] = &[
    "legacyMatte",
    "legacyPlastic",
    "legacyMetal",
    "legacyWireframe",
    "matte",
    "plastic",
    "metal",
    "warmMatte",
    "translucentPowder",
    "powder",
    "dkEdge",
    "softEdge",
    "clear",
    "flat",
    "softmetal",
    "none",
];
const BEVEL_PRESETS: &[&str] = &[
    "relaxedInset",
    "circle",
    "slope",
    "cross",
    "angle",
    "softRound",
    "convex",
    "coolSlant",
    "divot",
    "riblet",
    "hardEdge",
    "artDeco",
];

token_type!(PresetCamera, CAMERA_PRESETS);
token_type!(LightRigType, LIGHT_RIGS);
token_type!(LightRigDirection, LIGHT_DIRECTIONS);
token_type!(PresetMaterial, MATERIALS);
token_type!(BevelPreset, BEVEL_PRESETS);

impl Color {
    pub(crate) fn validate(&self) -> Result<()> {
        let transforms = match self {
            Self::Rgb(value) => &value.transforms,
            Self::Scheme(value) => &value.transforms,
        };
        if transforms.len() > MAX_COLOR_TRANSFORMS {
            return Err(Error::Invalid(format!(
                "too many DrawingML color transforms (maximum {MAX_COLOR_TRANSFORMS})"
            )));
        }
        for transform in transforms {
            match transform {
                ColorTransform::Tint(value)
                | ColorTransform::Shade(value)
                | ColorTransform::Alpha(value)
                | ColorTransform::SaturationMod(value)
                | ColorTransform::LuminanceMod(value)
                    if *value > 100_000 =>
                {
                    return Err(Error::Invalid(
                        "positive DrawingML color transform exceeds 100000".into(),
                    ));
                },
                ColorTransform::HueMod(value) if *value > 1_000_000 => {
                    return Err(Error::Invalid("hueMod exceeds 1000000".into()));
                },
                ColorTransform::Saturation(value)
                | ColorTransform::SaturationOffset(value)
                | ColorTransform::Luminance(value)
                | ColorTransform::LuminanceOffset(value)
                    if !(-100_000..=100_000).contains(value) =>
                {
                    return Err(Error::Invalid(
                        "signed DrawingML color transform is outside -100000..=100000".into(),
                    ));
                },
                _ => {},
            }
        }
        Ok(())
    }
}

impl Glow {
    pub(crate) fn validate(&self) -> Result<()> {
        validate_required_color(&self.color, "glow")?;
        validate_coordinate(self.radius, "glow radius")
    }
}

impl Shadow {
    pub(crate) fn validate(&self) -> Result<()> {
        validate_required_color(&self.color, "shadow")?;
        validate_coordinate(self.blur_radius, "shadow blur radius")?;
        validate_coordinate(self.distance, "shadow distance")?;
        validate_angle(self.direction, "shadow direction")?;
        validate_signed_percentage(self.scale_x, "shadow sx")?;
        validate_signed_percentage(self.scale_y, "shadow sy")?;
        Ok(())
    }
}

impl Reflection {
    pub(crate) fn validate(&self) -> Result<()> {
        validate_coordinate(self.blur_radius, "reflection blur radius")?;
        for (value, name) in [
            (self.start_alpha, "reflection stA"),
            (self.start_position, "reflection stPos"),
            (self.end_alpha, "reflection endA"),
            (self.end_position, "reflection endPos"),
        ] {
            validate_fixed_percentage(value, name)?;
        }
        validate_coordinate(self.distance, "reflection distance")?;
        validate_angle(self.direction, "reflection direction")?;
        validate_angle(self.fade_direction, "reflection fade direction")?;
        validate_signed_percentage(self.scale_x, "reflection sx")?;
        validate_signed_percentage(self.scale_y, "reflection sy")?;
        Ok(())
    }
}

impl TextFill {
    pub(crate) fn validate(&self) -> Result<()> {
        if let Some(fill) = &self.fill {
            fill.validate()
        } else {
            Ok(())
        }
    }
}

impl TextOutline {
    pub(crate) fn validate(&self) -> Result<()> {
        if let Some(fill) = &self.fill {
            fill.validate()?
        }
        if let Some(LineJoin::Miter { limit }) = &self.join {
            validate_positive_percentage(*limit, "miter limit")?;
        }
        validate_coordinate(self.width, "outline width")
    }
}

impl Fill {
    fn validate(&self) -> Result<()> {
        match self {
            Self::NoFill => Ok(()),
            Self::Solid(color) => validate_color(color),
            Self::Gradient(value) => value.validate(),
        }
    }
}

impl Gradient {
    fn validate(&self) -> Result<()> {
        if self.stops.len() < 2 || self.stops.len() > MAX_GRADIENT_STOPS {
            return Err(Error::Invalid(format!(
                "gradient stop count {} is outside 2..={MAX_GRADIENT_STOPS}",
                self.stops.len()
            )));
        }
        let mut previous = None;
        for stop in &self.stops {
            if stop.position > 100_000 || previous.is_some_and(|value| value >= stop.position) {
                return Err(Error::Invalid(
                    "gradient stops must be strictly increasing in 0..=100000".into(),
                ));
            }
            stop.color.validate()?;
            previous = Some(stop.position);
        }
        if let Some(shade) = &self.shade {
            match shade {
                Shade::Linear { angle, .. } => validate_angle(*angle, "gradient angle")?,
                Shade::Path { fill_to, .. } => {
                    if let Some(rect) = fill_to {
                        for (value, name) in [
                            (rect.left, "gradient left"),
                            (rect.top, "gradient top"),
                            (rect.right, "gradient right"),
                            (rect.bottom, "gradient bottom"),
                        ] {
                            validate_signed_percentage(value, name)?;
                        }
                    }
                },
            }
        }
        Ok(())
    }
}

impl Scene3d {
    pub(crate) fn validate(&self) -> Result<()> {
        if let Some(rotation) = &self.light_rig.rotation {
            for (value, name) in [
                (rotation.latitude, "scene latitude"),
                (rotation.longitude, "scene longitude"),
                (rotation.revolution, "scene revolution"),
            ] {
                validate_angle(Some(value), name)?;
            }
        }
        Ok(())
    }
}

impl Props3d {
    pub(crate) fn validate(&self) -> Result<()> {
        for bevel in [&self.bevel_top, &self.bevel_bottom].into_iter().flatten() {
            validate_coordinate(bevel.width, "bevel width")?;
            validate_coordinate(bevel.height, "bevel height")?;
        }
        validate_color(&self.extrusion_color)?;
        validate_color(&self.contour_color)?;
        validate_coordinate(self.extrusion_height, "extrusion height")?;
        validate_coordinate(self.contour_width, "contour width")
    }
}

fn validate_color(value: &Option<Color>) -> Result<()> {
    if let Some(value) = value {
        value.validate()?;
    }
    Ok(())
}

fn validate_required_color(value: &Option<Color>, name: &str) -> Result<()> {
    let Some(value) = value else {
        return Err(Error::Invalid(format!("{name} requires a color choice")));
    };
    value.validate()
}

fn validate_coordinate(value: Option<u64>, name: &str) -> Result<()> {
    if value.is_some_and(|value| value > i32::MAX as u64) {
        return Err(Error::Invalid(format!(
            "{name} exceeds the Word coordinate bound"
        )));
    }
    Ok(())
}

fn validate_angle(value: Option<u32>, name: &str) -> Result<()> {
    if value.is_some_and(|value| value > i32::MAX as u32) {
        return Err(Error::Invalid(format!(
            "{name} exceeds the XML angle bound"
        )));
    }
    Ok(())
}

fn validate_fixed_percentage(value: Option<u32>, name: &str) -> Result<()> {
    if value.is_some_and(|value| value > 100_000) {
        return Err(Error::Invalid(format!("{name} exceeds 100000")));
    }
    Ok(())
}

fn validate_positive_percentage(value: Option<u32>, name: &str) -> Result<()> {
    if value.is_some_and(|value| value > 100_000) {
        return Err(Error::Invalid(format!("{name} exceeds 100000")));
    }
    Ok(())
}

fn validate_signed_percentage(value: Option<i32>, name: &str) -> Result<()> {
    let _ = (value, name);
    Ok(())
}
