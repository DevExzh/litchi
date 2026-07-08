//! Escher shape property parsing (Opt record).
//!
//! Properties control shape appearance: position, size, colors, rotation, etc.
//! Based on MS-ODRAW specification section 2.3.
//!
//! # Complex Properties
//!
//! Properties can be simple (4-byte value) or complex (variable-length data).
//! Complex properties use a two-pass parsing approach:
//! 1. First pass: Parse all 6-byte property headers
//! 2. Second pass: Read complex data that follows the headers
//!
//! # Performance
//!
//! - Two-pass parsing minimizes data copying
//! - Zero-copy for complex data (borrows from source)
//! - HashMap for O(1) property lookup
//! - Pre-allocated capacity based on property count

use super::container::EscherContainer;
use super::record::EscherRecord;
use super::types::EscherRecordType;
use std::collections::HashMap;

const IS_COMPLEX: u16 = 0x8000;
const PROPERTY_ID_MASK: u16 = 0x3FFF;

/// Comprehensive Escher property ID enumeration from MS-ODRAW specification.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum EscherPropertyId {
    Rotation = 0x0004,
    LockRotation = 0x0077,
    LockAspectRatio = 0x0078,
    LockPosition = 0x0079,
    LockAgainstSelect = 0x007A,
    LockCropping = 0x007B,
    LockVertices = 0x007C,
    LockText = 0x007D,
    LockAdjustHandles = 0x007E,
    LockAgainstGrouping = 0x007F,
    TextId = 0x0080,
    TextLeft = 0x0081,
    TextTop = 0x0082,
    TextRight = 0x0083,
    TextBottom = 0x0084,
    WrapText = 0x0085,
    ScaleText = 0x0086,
    AnchorText = 0x0087,
    TextFlow = 0x0088,
    FontRotation = 0x0089,
    IdOfNextShape = 0x008A,
    TextBidi = 0x008B,
    SingleClickSelects = 0x00BB,
    UseHostMargins = 0x00BC,
    RotateTextWithShape = 0x00BD,
    SizeShapeToFitText = 0x00BE,
    SizeTextToFitShape = 0x00BF,
    GeoTextUnicode = 0x00C0,
    GeoTextRtf = 0x00C1,
    GeoTextAlignmentOnCurve = 0x00C2,
    GeoTextDefaultPointSize = 0x00C3,
    GeoTextSpacing = 0x00C4,
    GeoTextFontFamilyName = 0x00C5,
    GeoTextBoldFont = 0x00FA,
    GeoTextItalicFont = 0x00FB,
    GeoTextUnderlineFont = 0x00FC,
    GeoTextShadowFont = 0x00FD,
    GeoTextSmallCapsFont = 0x00FE,
    GeoTextStrikethroughFont = 0x00FF,
    BlipCropFromTop = 0x0100,
    BlipCropFromBottom = 0x0101,
    BlipCropFromLeft = 0x0102,
    BlipCropFromRight = 0x0103,
    BlipToDisplay = 0x0104,
    PictureFileName = 0x0105,
    BlipFlags = 0x0106,
    TransparentColor = 0x0107,
    PictureContrast = 0x0108,
    PictureBrightness = 0x0109,
    PictureGamma = 0x010A,
    PictureId = 0x010B,
    DoubleMod = 0x010C,
    PictureFillMod = 0x010D,
    PictureLine = 0x010E,
    PrintBlip = 0x010F,
    PrintBlipFilename = 0x0110,
    PrintFlags = 0x0111,
    NoHitTestPicture = 0x013C,
    PictureGray = 0x013D,
    PictureBilevel = 0x013E,
    PictureActive = 0x013F,
    GeomLeft = 0x0140,
    GeomTop = 0x0141,
    GeomRight = 0x0142,
    GeomBottom = 0x0143,
    ShapePath = 0x0144,
    Vertices = 0x0145,
    SegmentInfo = 0x0146,
    AdjustValue = 0x0147,
    Adjust2Value = 0x0148,
    Adjust3Value = 0x0149,
    Adjust4Value = 0x014A,
    Adjust5Value = 0x014B,
    Adjust6Value = 0x014C,
    Adjust7Value = 0x014D,
    Adjust8Value = 0x014E,
    Adjust9Value = 0x014F,
    Adjust10Value = 0x0150,
    ConnectionSites = 0x0151,
    ConnectionSitesDir = 0x0152,
    XLimo = 0x0153,
    YLimo = 0x0154,
    AdjustHandles = 0x0155,
    Guides = 0x0156,
    Inscribe = 0x0157,
    Cxk = 0x0158,
    Fragments = 0x0159,
    ShadowOk = 0x017A,
    ThreeDOk = 0x017B,
    LineOk = 0x017C,
    GeoTextOk = 0x017D,
    FillShadeShapeOk = 0x017E,
    FillOk = 0x017F,
    FillType = 0x0180,
    FillColor = 0x0181,
    FillOpacity = 0x0182,
    FillBackColor = 0x0183,
    FillBackOpacity = 0x0184,
    FillCrMod = 0x0185,
    FillBlip = 0x0186,
    FillBlipName = 0x0187,
    FillBlipFlags = 0x0188,
    FillWidth = 0x0189,
    FillHeight = 0x018A,
    FillAngle = 0x018B,
    FillFocus = 0x018C,
    FillToLeft = 0x018D,
    FillToTop = 0x018E,
    FillToRight = 0x018F,
    FillToBottom = 0x0190,
    FillRectLeft = 0x0191,
    FillRectTop = 0x0192,
    FillRectRight = 0x0193,
    FillRectBottom = 0x0194,
    FillDzType = 0x0195,
    FillShadePreset = 0x0196,
    FillShadeColors = 0x0197,
    FillOriginX = 0x0198,
    FillOriginY = 0x0199,
    FillShapeOriginX = 0x019A,
    FillShapeOriginY = 0x019B,
    FillShadeType = 0x019C,
    Filled = 0x01BB,
    HitTestFill = 0x01BC,
    FillShape = 0x01BD,
    UseRect = 0x01BE,
    NoFillHitTest = 0x01BF,
    LineColor = 0x01C0,
    LineOpacity = 0x01C1,
    LineBackColor = 0x01C2,
    LineCrMod = 0x01C3,
    LineType = 0x01C4,
    LineFillBlip = 0x01C5,
    LineFillBlipName = 0x01C6,
    LineFillBlipFlags = 0x01C7,
    LineFillWidth = 0x01C8,
    LineFillHeight = 0x01C9,
    LineFillDzType = 0x01CA,
    LineWidth = 0x01CB,
    LineMiterLimit = 0x01CC,
    LineStyle = 0x01CD,
    LineDashing = 0x01CE,
    LineDashStyle = 0x01CF,
    LineStartArrowhead = 0x01D0,
    LineEndArrowhead = 0x01D1,
    LineStartArrowWidth = 0x01D2,
    LineStartArrowLength = 0x01D3,
    LineEndArrowWidth = 0x01D4,
    LineEndArrowLength = 0x01D5,
    LineJoinStyle = 0x01D6,
    LineEndCapStyle = 0x01D7,
    ArrowheadsOk = 0x01FB,
    AnyLine = 0x01FC,
    HitTestLine = 0x01FD,
    LineFillShape = 0x01FE,
    NoLineDrawDash = 0x01FF,
    ShadowType = 0x0200,
    ShadowColor = 0x0201,
    ShadowHighlight = 0x0202,
    ShadowCrMod = 0x0203,
    ShadowOpacity = 0x0204,
    ShadowOffsetX = 0x0205,
    ShadowOffsetY = 0x0206,
    ShadowSecondOffsetX = 0x0207,
    ShadowSecondOffsetY = 0x0208,
    ShadowScaleXToX = 0x0209,
    ShadowScaleYToX = 0x020A,
    ShadowScaleXToY = 0x020B,
    ShadowScaleYToY = 0x020C,
    ShadowPerspectiveX = 0x020D,
    ShadowPerspectiveY = 0x020E,
    ShadowWeight = 0x020F,
    ShadowOriginX = 0x0210,
    ShadowOriginY = 0x0211,
    Shadow = 0x023E,
    ShadowObscured = 0x023F,
    PerspectiveType = 0x0240,
    PerspectiveOffsetX = 0x0241,
    PerspectiveOffsetY = 0x0242,
    PerspectiveScaleXToX = 0x0243,
    PerspectiveScaleYToX = 0x0244,
    PerspectiveScaleXToY = 0x0245,
    PerspectiveScaleYToY = 0x0246,
    PerspectivePerspectiveX = 0x0247,
    PerspectivePerspectiveY = 0x0248,
    PerspectiveWeight = 0x0249,
    PerspectiveOriginX = 0x024A,
    PerspectiveOriginY = 0x024B,
    PerspectiveOn = 0x027F,
    ThreeDSpecularAmount = 0x0280,
    ThreeDDiffuseAmount = 0x0281,
    ThreeDShininess = 0x0282,
    ThreeDEdgeThickness = 0x0283,
    ThreeDExtrudeForward = 0x0284,
    ThreeDExtrudeBackward = 0x0285,
    ThreeDExtrusionColor = 0x0287,
    ThreeDCrMod = 0x0288,
    ThreeDExtrusionColorExt = 0x0289,
    ThreeDEffect = 0x02BC,
    ThreeDMetallic = 0x02BD,
    ThreeDUseExtrusionColor = 0x02BE,
    ThreeDLightFace = 0x02BF,
    ThreeDStyleYRotationAngle = 0x02C0,
    ThreeDStyleXRotationAngle = 0x02C1,
    ThreeDStyleRotationAxisX = 0x02C2,
    ThreeDStyleRotationAxisY = 0x02C3,
    ThreeDStyleRotationAxisZ = 0x02C4,
    ThreeDStyleRotationAngle = 0x02C5,
    ThreeDStyleRotationCenterX = 0x02C6,
    ThreeDStyleRotationCenterY = 0x02C7,
    ThreeDStyleRotationCenterZ = 0x02C8,
    ThreeDStyleRenderMode = 0x02C9,
    ThreeDStyleTolerance = 0x02CA,
    ThreeDStyleXViewpoint = 0x02CB,
    ThreeDStyleYViewpoint = 0x02CC,
    ThreeDStyleZViewpoint = 0x02CD,
    ThreeDStyleOriginX = 0x02CE,
    ThreeDStyleOriginY = 0x02CF,
    ThreeDStyleSkewAngle = 0x02D0,
    ThreeDStyleSkewAmount = 0x02D1,
    ThreeDStyleAmbientIntensity = 0x02D2,
    ThreeDStyleKeyX = 0x02D3,
    ThreeDStyleKeyY = 0x02D4,
    ThreeDStyleKeyZ = 0x02D5,
    ThreeDStyleKeyIntensity = 0x02D6,
    ThreeDStyleFillX = 0x02D7,
    ThreeDStyleFillY = 0x02D8,
    ThreeDStyleFillZ = 0x02D9,
    ThreeDStyleFillIntensity = 0x02DA,
    ShapeMaster = 0x0301,
    ShapeConnectorStyle = 0x0303,
    ShapeBlackAndWhiteSettings = 0x0304,
    ShapeWModePureBw = 0x0305,
    ShapeWModeBw = 0x0306,
    ShapeOleIcon = 0x033A,
    ShapePreferRelativeResize = 0x033B,
    ShapeLockShapeType = 0x033C,
    ShapeDeleteAttachedObject = 0x033E,
    ShapeBackgroundShape = 0x033F,
    CalloutType = 0x0340,
    CalloutXYGap = 0x0341,
    CalloutAngle = 0x0342,
    CalloutDropType = 0x0343,
    CalloutDrop = 0x0344,
    CalloutLength = 0x0345,
    GroupName = 0x0380,
    GroupDescription = 0x0381,
    Hyperlink = 0x0382,
    DiagramType = 0x0500,
    DiagramStyle = 0x0501,
    Unknown = 0xFFFF,
}

impl From<u16> for EscherPropertyId {
    fn from(value: u16) -> Self {
        let prop_num = value & PROPERTY_ID_MASK;
        match prop_num {
            0x0004 => Self::Rotation,
            0x0077 => Self::LockRotation,
            0x0078 => Self::LockAspectRatio,
            0x0079 => Self::LockPosition,
            0x007A => Self::LockAgainstSelect,
            0x007B => Self::LockCropping,
            0x007C => Self::LockVertices,
            0x007D => Self::LockText,
            0x007E => Self::LockAdjustHandles,
            0x007F => Self::LockAgainstGrouping,
            0x0080 => Self::TextId,
            0x0081 => Self::TextLeft,
            0x0082 => Self::TextTop,
            0x0083 => Self::TextRight,
            0x0084 => Self::TextBottom,
            0x0085 => Self::WrapText,
            0x0086 => Self::ScaleText,
            0x0087 => Self::AnchorText,
            0x0088 => Self::TextFlow,
            0x0089 => Self::FontRotation,
            0x008A => Self::IdOfNextShape,
            0x008B => Self::TextBidi,
            0x00BB => Self::SingleClickSelects,
            0x00BC => Self::UseHostMargins,
            0x00BD => Self::RotateTextWithShape,
            0x00BE => Self::SizeShapeToFitText,
            0x00BF => Self::SizeTextToFitShape,
            0x00C0 => Self::GeoTextUnicode,
            0x00C1 => Self::GeoTextRtf,
            0x00C2 => Self::GeoTextAlignmentOnCurve,
            0x00C3 => Self::GeoTextDefaultPointSize,
            0x00C4 => Self::GeoTextSpacing,
            0x00C5 => Self::GeoTextFontFamilyName,
            0x00FA => Self::GeoTextBoldFont,
            0x00FB => Self::GeoTextItalicFont,
            0x00FC => Self::GeoTextUnderlineFont,
            0x00FD => Self::GeoTextShadowFont,
            0x00FE => Self::GeoTextSmallCapsFont,
            0x00FF => Self::GeoTextStrikethroughFont,
            0x0100 => Self::BlipCropFromTop,
            0x0101 => Self::BlipCropFromBottom,
            0x0102 => Self::BlipCropFromLeft,
            0x0103 => Self::BlipCropFromRight,
            0x0104 => Self::BlipToDisplay,
            0x0105 => Self::PictureFileName,
            0x0106 => Self::BlipFlags,
            0x0107 => Self::TransparentColor,
            0x0108 => Self::PictureContrast,
            0x0109 => Self::PictureBrightness,
            0x010A => Self::PictureGamma,
            0x010B => Self::PictureId,
            0x010C => Self::DoubleMod,
            0x010D => Self::PictureFillMod,
            0x010E => Self::PictureLine,
            0x010F => Self::PrintBlip,
            0x0110 => Self::PrintBlipFilename,
            0x0111 => Self::PrintFlags,
            0x013C => Self::NoHitTestPicture,
            0x013D => Self::PictureGray,
            0x013E => Self::PictureBilevel,
            0x013F => Self::PictureActive,
            0x0140 => Self::GeomLeft,
            0x0141 => Self::GeomTop,
            0x0142 => Self::GeomRight,
            0x0143 => Self::GeomBottom,
            0x0144 => Self::ShapePath,
            0x0145 => Self::Vertices,
            0x0146 => Self::SegmentInfo,
            0x0147 => Self::AdjustValue,
            0x0148 => Self::Adjust2Value,
            0x0149 => Self::Adjust3Value,
            0x014A => Self::Adjust4Value,
            0x014B => Self::Adjust5Value,
            0x014C => Self::Adjust6Value,
            0x014D => Self::Adjust7Value,
            0x014E => Self::Adjust8Value,
            0x014F => Self::Adjust9Value,
            0x0150 => Self::Adjust10Value,
            0x0151 => Self::ConnectionSites,
            0x0152 => Self::ConnectionSitesDir,
            0x0153 => Self::XLimo,
            0x0154 => Self::YLimo,
            0x0155 => Self::AdjustHandles,
            0x0156 => Self::Guides,
            0x0157 => Self::Inscribe,
            0x0158 => Self::Cxk,
            0x0159 => Self::Fragments,
            0x017A => Self::ShadowOk,
            0x017B => Self::ThreeDOk,
            0x017C => Self::LineOk,
            0x017D => Self::GeoTextOk,
            0x017E => Self::FillShadeShapeOk,
            0x017F => Self::FillOk,
            0x0180 => Self::FillType,
            0x0181 => Self::FillColor,
            0x0182 => Self::FillOpacity,
            0x0183 => Self::FillBackColor,
            0x0184 => Self::FillBackOpacity,
            0x0185 => Self::FillCrMod,
            0x0186 => Self::FillBlip,
            0x0187 => Self::FillBlipName,
            0x0188 => Self::FillBlipFlags,
            0x0189 => Self::FillWidth,
            0x018A => Self::FillHeight,
            0x018B => Self::FillAngle,
            0x018C => Self::FillFocus,
            0x018D => Self::FillToLeft,
            0x018E => Self::FillToTop,
            0x018F => Self::FillToRight,
            0x0190 => Self::FillToBottom,
            0x0191 => Self::FillRectLeft,
            0x0192 => Self::FillRectTop,
            0x0193 => Self::FillRectRight,
            0x0194 => Self::FillRectBottom,
            0x0195 => Self::FillDzType,
            0x0196 => Self::FillShadePreset,
            0x0197 => Self::FillShadeColors,
            0x0198 => Self::FillOriginX,
            0x0199 => Self::FillOriginY,
            0x019A => Self::FillShapeOriginX,
            0x019B => Self::FillShapeOriginY,
            0x019C => Self::FillShadeType,
            0x01BB => Self::Filled,
            0x01BC => Self::HitTestFill,
            0x01BD => Self::FillShape,
            0x01BE => Self::UseRect,
            0x01BF => Self::NoFillHitTest,
            0x01C0 => Self::LineColor,
            0x01C1 => Self::LineOpacity,
            0x01C2 => Self::LineBackColor,
            0x01C3 => Self::LineCrMod,
            0x01C4 => Self::LineType,
            0x01C5 => Self::LineFillBlip,
            0x01C6 => Self::LineFillBlipName,
            0x01C7 => Self::LineFillBlipFlags,
            0x01C8 => Self::LineFillWidth,
            0x01C9 => Self::LineFillHeight,
            0x01CA => Self::LineFillDzType,
            0x01CB => Self::LineWidth,
            0x01CC => Self::LineMiterLimit,
            0x01CD => Self::LineStyle,
            0x01CE => Self::LineDashing,
            0x01CF => Self::LineDashStyle,
            0x01D0 => Self::LineStartArrowhead,
            0x01D1 => Self::LineEndArrowhead,
            0x01D2 => Self::LineStartArrowWidth,
            0x01D3 => Self::LineStartArrowLength,
            0x01D4 => Self::LineEndArrowWidth,
            0x01D5 => Self::LineEndArrowLength,
            0x01D6 => Self::LineJoinStyle,
            0x01D7 => Self::LineEndCapStyle,
            0x01FB => Self::ArrowheadsOk,
            0x01FC => Self::AnyLine,
            0x01FD => Self::HitTestLine,
            0x01FE => Self::LineFillShape,
            0x01FF => Self::NoLineDrawDash,
            0x0200 => Self::ShadowType,
            0x0201 => Self::ShadowColor,
            0x0202 => Self::ShadowHighlight,
            0x0203 => Self::ShadowCrMod,
            0x0204 => Self::ShadowOpacity,
            0x0205 => Self::ShadowOffsetX,
            0x0206 => Self::ShadowOffsetY,
            0x0207 => Self::ShadowSecondOffsetX,
            0x0208 => Self::ShadowSecondOffsetY,
            0x0209 => Self::ShadowScaleXToX,
            0x020A => Self::ShadowScaleYToX,
            0x020B => Self::ShadowScaleXToY,
            0x020C => Self::ShadowScaleYToY,
            0x020D => Self::ShadowPerspectiveX,
            0x020E => Self::ShadowPerspectiveY,
            0x020F => Self::ShadowWeight,
            0x0210 => Self::ShadowOriginX,
            0x0211 => Self::ShadowOriginY,
            0x023E => Self::Shadow,
            0x023F => Self::ShadowObscured,
            0x0240 => Self::PerspectiveType,
            0x0241 => Self::PerspectiveOffsetX,
            0x0242 => Self::PerspectiveOffsetY,
            0x0243 => Self::PerspectiveScaleXToX,
            0x0244 => Self::PerspectiveScaleYToX,
            0x0245 => Self::PerspectiveScaleXToY,
            0x0246 => Self::PerspectiveScaleYToY,
            0x0247 => Self::PerspectivePerspectiveX,
            0x0248 => Self::PerspectivePerspectiveY,
            0x0249 => Self::PerspectiveWeight,
            0x024A => Self::PerspectiveOriginX,
            0x024B => Self::PerspectiveOriginY,
            0x027F => Self::PerspectiveOn,
            0x0280 => Self::ThreeDSpecularAmount,
            0x0281 => Self::ThreeDDiffuseAmount,
            0x0282 => Self::ThreeDShininess,
            0x0283 => Self::ThreeDEdgeThickness,
            0x0284 => Self::ThreeDExtrudeForward,
            0x0285 => Self::ThreeDExtrudeBackward,
            0x0287 => Self::ThreeDExtrusionColor,
            0x0288 => Self::ThreeDCrMod,
            0x0289 => Self::ThreeDExtrusionColorExt,
            0x02BC => Self::ThreeDEffect,
            0x02BD => Self::ThreeDMetallic,
            0x02BE => Self::ThreeDUseExtrusionColor,
            0x02BF => Self::ThreeDLightFace,
            0x02C0 => Self::ThreeDStyleYRotationAngle,
            0x02C1 => Self::ThreeDStyleXRotationAngle,
            0x02C2 => Self::ThreeDStyleRotationAxisX,
            0x02C3 => Self::ThreeDStyleRotationAxisY,
            0x02C4 => Self::ThreeDStyleRotationAxisZ,
            0x02C5 => Self::ThreeDStyleRotationAngle,
            0x02C6 => Self::ThreeDStyleRotationCenterX,
            0x02C7 => Self::ThreeDStyleRotationCenterY,
            0x02C8 => Self::ThreeDStyleRotationCenterZ,
            0x02C9 => Self::ThreeDStyleRenderMode,
            0x02CA => Self::ThreeDStyleTolerance,
            0x02CB => Self::ThreeDStyleXViewpoint,
            0x02CC => Self::ThreeDStyleYViewpoint,
            0x02CD => Self::ThreeDStyleZViewpoint,
            0x02CE => Self::ThreeDStyleOriginX,
            0x02CF => Self::ThreeDStyleOriginY,
            0x02D0 => Self::ThreeDStyleSkewAngle,
            0x02D1 => Self::ThreeDStyleSkewAmount,
            0x02D2 => Self::ThreeDStyleAmbientIntensity,
            0x02D3 => Self::ThreeDStyleKeyX,
            0x02D4 => Self::ThreeDStyleKeyY,
            0x02D5 => Self::ThreeDStyleKeyZ,
            0x02D6 => Self::ThreeDStyleKeyIntensity,
            0x02D7 => Self::ThreeDStyleFillX,
            0x02D8 => Self::ThreeDStyleFillY,
            0x02D9 => Self::ThreeDStyleFillZ,
            0x02DA => Self::ThreeDStyleFillIntensity,
            0x0301 => Self::ShapeMaster,
            0x0303 => Self::ShapeConnectorStyle,
            0x0304 => Self::ShapeBlackAndWhiteSettings,
            0x0305 => Self::ShapeWModePureBw,
            0x0306 => Self::ShapeWModeBw,
            0x033A => Self::ShapeOleIcon,
            0x033B => Self::ShapePreferRelativeResize,
            0x033C => Self::ShapeLockShapeType,
            0x033E => Self::ShapeDeleteAttachedObject,
            0x033F => Self::ShapeBackgroundShape,
            0x0340 => Self::CalloutType,
            0x0341 => Self::CalloutXYGap,
            0x0342 => Self::CalloutAngle,
            0x0343 => Self::CalloutDropType,
            0x0344 => Self::CalloutDrop,
            0x0345 => Self::CalloutLength,
            0x0380 => Self::GroupName,
            0x0381 => Self::GroupDescription,
            0x0382 => Self::Hyperlink,
            0x0500 => Self::DiagramType,
            0x0501 => Self::DiagramStyle,
            _ => Self::Unknown,
        }
    }
}

/// Escher shape property value.
#[derive(Debug, Clone)]
pub enum EscherPropertyValue<'data> {
    Simple(i32),
    Complex(&'data [u8]),
    Array(EscherArrayProperty<'data>),
}

/// Escher array property structure.
#[derive(Debug, Clone)]
pub struct EscherArrayProperty<'data> {
    data: &'data [u8],
}

impl<'data> EscherArrayProperty<'data> {
    #[inline]
    pub fn new(data: &'data [u8]) -> Option<Self> {
        if data.len() < 6 {
            return None;
        }
        Some(Self { data })
    }

    #[inline]
    pub fn element_count(&self) -> u16 {
        if self.data.len() < 2 {
            return 0;
        }
        u16::from_le_bytes([self.data[0], self.data[1]])
    }

    #[inline]
    pub fn element_count_in_memory(&self) -> u16 {
        if self.data.len() < 4 {
            return 0;
        }
        u16::from_le_bytes([self.data[2], self.data[3]])
    }

    #[inline]
    pub fn raw_element_size(&self) -> i16 {
        if self.data.len() < 6 {
            return 0;
        }
        i16::from_le_bytes([self.data[4], self.data[5]])
    }

    #[inline]
    pub fn element_size(&self) -> usize {
        let size = self.raw_element_size();
        if size < 0 {
            ((-size) >> 2) as usize
        } else {
            size as usize
        }
    }

    #[inline]
    pub fn get_element(&self, index: usize) -> Option<&'data [u8]> {
        let count = self.element_count() as usize;
        if index >= count {
            return None;
        }

        let elem_size = self.element_size();
        let start = 6 + index * elem_size;
        let end = start + elem_size;

        if end > self.data.len() {
            return None;
        }

        Some(&self.data[start..end])
    }

    pub fn elements(&self) -> impl Iterator<Item = &'data [u8]> {
        let count = self.element_count() as usize;
        let elem_size = self.element_size();
        let data = self.data;

        (0..count).filter_map(move |i| {
            let start = 6 + i * elem_size;
            let end = start + elem_size;
            if end <= data.len() {
                Some(&data[start..end])
            } else {
                None
            }
        })
    }

    #[inline]
    pub fn raw_data(&self) -> &'data [u8] {
        self.data
    }
}

/// Escher shape properties collection.
#[derive(Debug, Clone)]
pub struct EscherProperties<'data> {
    properties: HashMap<EscherPropertyId, EscherPropertyValue<'data>>,
}

#[derive(Debug, Clone, Copy)]
struct PropertyDescriptor {
    id: EscherPropertyId,
    id_raw: u16,
    value: i32,
}

impl<'data> EscherProperties<'data> {
    #[inline]
    pub fn new() -> Self {
        Self {
            properties: HashMap::new(),
        }
    }

    pub fn from_opt_record(opt: &EscherRecord<'data>) -> Self {
        let num_properties = opt.instance as usize;
        let mut properties = HashMap::with_capacity(num_properties);

        if opt.data.len() < 6 {
            return Self { properties };
        }

        let header_size = num_properties * 6;
        if header_size > opt.data.len() {
            return Self { properties };
        }

        let mut descriptors = Vec::with_capacity(num_properties);
        for i in 0..num_properties {
            let offset = i * 6;
            if offset + 6 > opt.data.len() {
                break;
            }

            let id_raw = u16::from_le_bytes([opt.data[offset], opt.data[offset + 1]]);
            let value = i32::from_le_bytes([
                opt.data[offset + 2],
                opt.data[offset + 3],
                opt.data[offset + 4],
                opt.data[offset + 5],
            ]);

            let id = EscherPropertyId::from(id_raw);
            descriptors.push(PropertyDescriptor { id, id_raw, value });
        }

        let mut complex_data_offset = header_size;

        for desc in descriptors {
            let is_complex = (desc.id_raw & IS_COMPLEX) != 0;

            let prop_value = if is_complex {
                let complex_len = desc.value as usize;
                let complex_end = complex_data_offset + complex_len;

                if complex_end > opt.data.len() {
                    complex_data_offset = complex_end;
                    continue;
                }

                let complex_data = &opt.data[complex_data_offset..complex_end];
                complex_data_offset = complex_end;

                if Self::is_array_property(complex_data) {
                    if let Some(array_prop) = EscherArrayProperty::new(complex_data) {
                        EscherPropertyValue::Array(array_prop)
                    } else {
                        EscherPropertyValue::Complex(complex_data)
                    }
                } else {
                    EscherPropertyValue::Complex(complex_data)
                }
            } else {
                EscherPropertyValue::Simple(desc.value)
            };

            properties.insert(desc.id, prop_value);
        }

        Self { properties }
    }

    fn is_array_property(data: &[u8]) -> bool {
        if data.len() < 6 {
            return false;
        }

        let num_elements = u16::from_le_bytes([data[0], data[1]]) as usize;
        let element_size_raw = i16::from_le_bytes([data[4], data[5]]);

        let element_size = if element_size_raw < 0 {
            ((-element_size_raw) >> 2) as usize
        } else {
            element_size_raw as usize
        };

        let expected_size_with_header = 6 + num_elements * element_size;
        let expected_size_without_header = num_elements * element_size;

        data.len() == expected_size_with_header || data.len() == expected_size_without_header
    }

    pub fn from_container(container: &EscherContainer<'data>) -> Self {
        if let Some(opt) = container.find_child(EscherRecordType::Opt) {
            Self::from_opt_record(&opt)
        } else {
            Self::new()
        }
    }

    #[inline]
    pub fn get(&self, id: EscherPropertyId) -> Option<&EscherPropertyValue<'data>> {
        self.properties.get(&id)
    }

    #[inline]
    pub fn get_int(&self, id: EscherPropertyId) -> Option<i32> {
        match self.properties.get(&id) {
            Some(EscherPropertyValue::Simple(v)) => Some(*v),
            _ => None,
        }
    }

    #[inline]
    pub fn get_color(&self, id: EscherPropertyId) -> Option<u32> {
        self.get_int(id).map(|v| v as u32)
    }

    #[inline]
    pub fn get_bool(&self, id: EscherPropertyId) -> Option<bool> {
        self.get_int(id).map(|v| v != 0)
    }

    #[inline]
    pub fn get_binary(&self, id: EscherPropertyId) -> Option<&'data [u8]> {
        match self.properties.get(&id) {
            Some(EscherPropertyValue::Complex(data)) => Some(data),
            _ => None,
        }
    }

    #[inline]
    pub fn get_array(&self, id: EscherPropertyId) -> Option<&EscherArrayProperty<'data>> {
        match self.properties.get(&id) {
            Some(EscherPropertyValue::Array(array)) => Some(array),
            _ => None,
        }
    }

    #[inline]
    pub fn has(&self, id: EscherPropertyId) -> bool {
        self.properties.contains_key(&id)
    }

    #[inline]
    pub fn len(&self) -> usize {
        self.properties.len()
    }

    #[inline]
    pub fn is_empty(&self) -> bool {
        self.properties.is_empty()
    }

    pub fn iter(&self) -> impl Iterator<Item = (&EscherPropertyId, &EscherPropertyValue<'data>)> {
        self.properties.iter()
    }

    #[inline]
    pub fn get_rgb(&self, id: EscherPropertyId) -> Option<(u8, u8, u8)> {
        self.get_color(id).map(|color| {
            let red = ((color >> 16) & 0xFF) as u8;
            let green = ((color >> 8) & 0xFF) as u8;
            let blue = (color & 0xFF) as u8;
            (red, green, blue)
        })
    }

    #[inline]
    pub fn get_rotation_degrees(&self, id: EscherPropertyId) -> Option<f32> {
        self.get_int(id)
            .map(|fixed_point| (fixed_point as f32) / 65536.0)
    }

    #[inline]
    pub fn get_opacity(&self, id: EscherPropertyId) -> Option<f32> {
        self.get_int(id).map(|fixed_point| {
            let opacity = (fixed_point as f32) / 65536.0;
            opacity.clamp(0.0, 1.0)
        })
    }

    #[inline]
    pub fn get_coord(&self, id: EscherPropertyId) -> Option<i32> {
        self.get_int(id)
    }

    #[inline]
    pub fn is_true(&self, id: EscherPropertyId) -> bool {
        self.get_bool(id).unwrap_or(false)
    }

    #[inline]
    pub fn get_line_width(&self) -> Option<i32> {
        self.get_int(EscherPropertyId::LineWidth)
    }

    #[inline]
    pub fn get_fill_color(&self) -> Option<(u8, u8, u8)> {
        self.get_rgb(EscherPropertyId::FillColor)
    }

    #[inline]
    pub fn get_line_color(&self) -> Option<(u8, u8, u8)> {
        self.get_rgb(EscherPropertyId::LineColor)
    }

    #[inline]
    pub fn is_filled(&self) -> bool {
        self.is_true(EscherPropertyId::Filled)
    }

    #[inline]
    pub fn has_line(&self) -> bool {
        self.is_true(EscherPropertyId::AnyLine)
    }

    #[inline]
    pub fn has_shadow(&self) -> bool {
        self.is_true(EscherPropertyId::Shadow)
    }

    #[inline]
    pub fn get_geometry_rect(&self) -> Option<(i32, i32, i32, i32)> {
        let left = self.get_coord(EscherPropertyId::GeomLeft)?;
        let top = self.get_coord(EscherPropertyId::GeomTop)?;
        let right = self.get_coord(EscherPropertyId::GeomRight)?;
        let bottom = self.get_coord(EscherPropertyId::GeomBottom)?;
        Some((left, top, right, bottom))
    }

    #[inline]
    pub fn get_text_margins(&self) -> Option<(i32, i32, i32, i32)> {
        let left = self.get_int(EscherPropertyId::TextLeft).unwrap_or(0);
        let top = self.get_int(EscherPropertyId::TextTop).unwrap_or(0);
        let right = self.get_int(EscherPropertyId::TextRight).unwrap_or(0);
        let bottom = self.get_int(EscherPropertyId::TextBottom).unwrap_or(0);
        Some((left, top, right, bottom))
    }

    #[inline]
    pub fn get_adjust(&self, id: EscherPropertyId) -> Option<i32> {
        self.get_int(id)
    }
}

impl<'data> Default for EscherProperties<'data> {
    fn default() -> Self {
        Self::new()
    }
}

/// Shape anchor (position and size).
#[derive(Debug, Clone, Copy)]
pub struct ShapeAnchor {
    pub left: i32,
    pub top: i32,
    pub right: i32,
    pub bottom: i32,
}

impl ShapeAnchor {
    #[inline]
    pub const fn new(left: i32, top: i32, right: i32, bottom: i32) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    #[inline]
    pub const fn width(&self) -> i32 {
        self.right - self.left
    }

    #[inline]
    pub const fn height(&self) -> i32 {
        self.bottom - self.top
    }

    pub fn from_child_anchor(anchor: &EscherRecord) -> Option<Self> {
        if anchor.data.len() < 16 {
            return None;
        }

        let left = i32::from_le_bytes([
            anchor.data[0],
            anchor.data[1],
            anchor.data[2],
            anchor.data[3],
        ]);
        let top = i32::from_le_bytes([
            anchor.data[4],
            anchor.data[5],
            anchor.data[6],
            anchor.data[7],
        ]);
        let right = i32::from_le_bytes([
            anchor.data[8],
            anchor.data[9],
            anchor.data[10],
            anchor.data[11],
        ]);
        let bottom = i32::from_le_bytes([
            anchor.data[12],
            anchor.data[13],
            anchor.data[14],
            anchor.data[15],
        ]);

        Some(Self::new(left, top, right, bottom))
    }

    pub fn from_client_anchor(anchor: &EscherRecord) -> Option<Self> {
        Self::from_child_anchor(anchor)
    }
}
