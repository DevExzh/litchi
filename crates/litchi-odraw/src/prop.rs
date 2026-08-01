//! OfficeArt shape-property parsing (`Opt` records).
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
//! - Wire-order retention plus indexed property lookup
//! - Pre-allocated capacity based on the checked property count

use crate::{Container, Error, Record, RecordKind, Result};
use std::collections::HashMap;

const IS_BLIP: u16 = 0x4000;
const IS_COMPLEX: u16 = 0x8000;
const PROPERTY_ID_MASK: u16 = 0x3FFF;

/// A property identifier not assigned by the OfficeArt specification known to
/// this crate.
///
/// The numeric value is private so an `UnknownId` can never contain flag bits
/// or alias one of the typed [`Id`] variants.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct UnknownId(u16);

impl UnknownId {
    /// Returns the exact 14-bit identifier read from the wire.
    pub const fn raw(self) -> u16 {
        self.0
    }
}

/// Comprehensive OfficeArt property identifier from [MS-ODRAW].
///
/// [MS-ODRAW]: https://learn.microsoft.com/openspecs/office_file_formats/ms-odraw/
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Id {
    Rotation,
    LockRotation,
    LockAspectRatio,
    LockPosition,
    LockAgainstSelect,
    LockCropping,
    LockVertices,
    LockText,
    LockAdjustHandles,
    LockAgainstGrouping,
    TextId,
    TextLeft,
    TextTop,
    TextRight,
    TextBottom,
    WrapText,
    UnusedText134,
    AnchorText,
    TextFlow,
    FontRotation,
    IdOfNextShape,
    TextDirection,
    SelectText,
    AutoTextMargin,
    UnusedTextBoolean189,
    FitShapeToText,
    TextBooleanProperties,
    GeoTextUnicode,
    GeoTextRtf,
    GeoTextAlignmentOnCurve,
    GeoTextDefaultPointSize,
    GeoTextSpacing,
    GeoTextFontFamilyName,
    GeoTextBoldFont,
    GeoTextItalicFont,
    GeoTextUnderlineFont,
    GeoTextShadowFont,
    GeoTextSmallCapsFont,
    GeoTextStrikethroughFont,
    BlipCropFromTop,
    BlipCropFromBottom,
    BlipCropFromLeft,
    BlipCropFromRight,
    BlipToDisplay,
    PictureFileName,
    BlipFlags,
    TransparentColor,
    PictureContrast,
    PictureBrightness,
    PictureGamma,
    PictureId,
    DoubleMod,
    PictureFillMod,
    PictureLine,
    PrintBlip,
    PrintBlipFilename,
    PrintFlags,
    NoHitTestPicture,
    PictureGray,
    PictureBilevel,
    PictureActive,
    GeomLeft,
    GeomTop,
    GeomRight,
    GeomBottom,
    ShapePath,
    Vertices,
    SegmentInfo,
    AdjustValue,
    Adjust2Value,
    Adjust3Value,
    Adjust4Value,
    Adjust5Value,
    Adjust6Value,
    Adjust7Value,
    Adjust8Value,
    Adjust9Value,
    Adjust10Value,
    ConnectionSites,
    ConnectionSitesDir,
    XLimo,
    YLimo,
    AdjustHandles,
    Guides,
    Inscribe,
    Cxk,
    Fragments,
    ShadowOk,
    ThreeDOk,
    LineOk,
    GeoTextOk,
    FillShadeShapeOk,
    FillOk,
    FillType,
    FillColor,
    FillOpacity,
    FillBackColor,
    FillBackOpacity,
    FillCrMod,
    FillBlip,
    FillBlipName,
    FillBlipFlags,
    FillWidth,
    FillHeight,
    FillAngle,
    FillFocus,
    FillToLeft,
    FillToTop,
    FillToRight,
    FillToBottom,
    FillRectLeft,
    FillRectTop,
    FillRectRight,
    FillRectBottom,
    FillDzType,
    FillShadePreset,
    FillShadeColors,
    FillOriginX,
    FillOriginY,
    FillShapeOriginX,
    FillShapeOriginY,
    FillShadeType,
    Filled,
    HitTestFill,
    FillShape,
    UseRect,
    NoFillHitTest,
    LineColor,
    LineOpacity,
    LineBackColor,
    LineCrMod,
    LineType,
    LineFillBlip,
    LineFillBlipName,
    LineFillBlipFlags,
    LineFillWidth,
    LineFillHeight,
    LineFillDzType,
    LineWidth,
    LineMiterLimit,
    LineStyle,
    LineDashing,
    LineDashStyle,
    LineStartArrowhead,
    LineEndArrowhead,
    LineStartArrowWidth,
    LineStartArrowLength,
    LineEndArrowWidth,
    LineEndArrowLength,
    LineJoinStyle,
    LineEndCapStyle,
    ArrowheadsOk,
    AnyLine,
    HitTestLine,
    LineFillShape,
    NoLineDrawDash,
    ShadowType,
    ShadowColor,
    ShadowHighlight,
    ShadowCrMod,
    ShadowOpacity,
    ShadowOffsetX,
    ShadowOffsetY,
    ShadowSecondOffsetX,
    ShadowSecondOffsetY,
    ShadowScaleXToX,
    ShadowScaleYToX,
    ShadowScaleXToY,
    ShadowScaleYToY,
    ShadowPerspectiveX,
    ShadowPerspectiveY,
    ShadowWeight,
    ShadowOriginX,
    ShadowOriginY,
    Shadow,
    ShadowObscured,
    PerspectiveType,
    PerspectiveOffsetX,
    PerspectiveOffsetY,
    PerspectiveScaleXToX,
    PerspectiveScaleYToX,
    PerspectiveScaleXToY,
    PerspectiveScaleYToY,
    PerspectivePerspectiveX,
    PerspectivePerspectiveY,
    PerspectiveWeight,
    PerspectiveOriginX,
    PerspectiveOriginY,
    PerspectiveOn,
    ThreeDSpecularAmount,
    ThreeDDiffuseAmount,
    ThreeDShininess,
    ThreeDEdgeThickness,
    ThreeDExtrudeForward,
    ThreeDExtrudeBackward,
    ThreeDExtrusionColor,
    ThreeDCrMod,
    ThreeDExtrusionColorExt,
    ThreeDEffect,
    ThreeDMetallic,
    ThreeDUseExtrusionColor,
    ThreeDLightFace,
    ThreeDStyleYRotationAngle,
    ThreeDStyleXRotationAngle,
    ThreeDStyleRotationAxisX,
    ThreeDStyleRotationAxisY,
    ThreeDStyleRotationAxisZ,
    ThreeDStyleRotationAngle,
    ThreeDStyleRotationCenterX,
    ThreeDStyleRotationCenterY,
    ThreeDStyleRotationCenterZ,
    ThreeDStyleRenderMode,
    ThreeDStyleTolerance,
    ThreeDStyleXViewpoint,
    ThreeDStyleYViewpoint,
    ThreeDStyleZViewpoint,
    ThreeDStyleOriginX,
    ThreeDStyleOriginY,
    ThreeDStyleSkewAngle,
    ThreeDStyleSkewAmount,
    ThreeDStyleAmbientIntensity,
    ThreeDStyleKeyX,
    ThreeDStyleKeyY,
    ThreeDStyleKeyZ,
    ThreeDStyleKeyIntensity,
    ThreeDStyleFillX,
    ThreeDStyleFillY,
    ThreeDStyleFillZ,
    ThreeDStyleFillIntensity,
    ShapeMaster,
    ShapeConnectorStyle,
    ShapeBlackAndWhiteSettings,
    ShapeWModePureBw,
    ShapeWModeBw,
    ShapeOleIcon,
    ShapePreferRelativeResize,
    ShapeLockShapeType,
    ShapeDeleteAttachedObject,
    ShapeBackgroundShape,
    CalloutType,
    CalloutXYGap,
    CalloutAngle,
    CalloutDropType,
    CalloutDrop,
    CalloutLength,
    GroupName,
    GroupDescription,
    Hyperlink,
    GroupTableProperties,
    GroupTableRowProperties,
    DiagramType,
    DiagramStyle,
    Unknown(UnknownId),
}

impl From<u16> for Id {
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
            0x0086 => Self::UnusedText134,
            0x0087 => Self::AnchorText,
            0x0088 => Self::TextFlow,
            0x0089 => Self::FontRotation,
            0x008A => Self::IdOfNextShape,
            0x008B => Self::TextDirection,
            0x00BB => Self::SelectText,
            0x00BC => Self::AutoTextMargin,
            0x00BD => Self::UnusedTextBoolean189,
            0x00BE => Self::FitShapeToText,
            0x00BF => Self::TextBooleanProperties,
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
            0x039F => Self::GroupTableProperties,
            0x03A0 => Self::GroupTableRowProperties,
            0x0500 => Self::DiagramType,
            0x0501 => Self::DiagramStyle,
            raw => Self::Unknown(UnknownId(raw)),
        }
    }
}

impl Id {
    /// Constructs an unassigned identifier after validating its invariant.
    ///
    /// Returns `None` when `raw` contains property flags or is already modeled
    /// by a typed variant.
    pub fn unknown(raw: u16) -> Option<Self> {
        if raw > PROPERTY_ID_MASK {
            return None;
        }
        match Self::from(raw) {
            id @ Self::Unknown(_) => Some(id),
            _ => None,
        }
    }

    /// Returns this property's exact wire identifier without flag bits.
    pub const fn raw(self) -> u16 {
        match self {
            Self::Rotation => 0x0004,
            Self::LockRotation => 0x0077,
            Self::LockAspectRatio => 0x0078,
            Self::LockPosition => 0x0079,
            Self::LockAgainstSelect => 0x007A,
            Self::LockCropping => 0x007B,
            Self::LockVertices => 0x007C,
            Self::LockText => 0x007D,
            Self::LockAdjustHandles => 0x007E,
            Self::LockAgainstGrouping => 0x007F,
            Self::TextId => 0x0080,
            Self::TextLeft => 0x0081,
            Self::TextTop => 0x0082,
            Self::TextRight => 0x0083,
            Self::TextBottom => 0x0084,
            Self::WrapText => 0x0085,
            Self::UnusedText134 => 0x0086,
            Self::AnchorText => 0x0087,
            Self::TextFlow => 0x0088,
            Self::FontRotation => 0x0089,
            Self::IdOfNextShape => 0x008A,
            Self::TextDirection => 0x008B,
            Self::SelectText => 0x00BB,
            Self::AutoTextMargin => 0x00BC,
            Self::UnusedTextBoolean189 => 0x00BD,
            Self::FitShapeToText => 0x00BE,
            Self::TextBooleanProperties => 0x00BF,
            Self::GeoTextUnicode => 0x00C0,
            Self::GeoTextRtf => 0x00C1,
            Self::GeoTextAlignmentOnCurve => 0x00C2,
            Self::GeoTextDefaultPointSize => 0x00C3,
            Self::GeoTextSpacing => 0x00C4,
            Self::GeoTextFontFamilyName => 0x00C5,
            Self::GeoTextBoldFont => 0x00FA,
            Self::GeoTextItalicFont => 0x00FB,
            Self::GeoTextUnderlineFont => 0x00FC,
            Self::GeoTextShadowFont => 0x00FD,
            Self::GeoTextSmallCapsFont => 0x00FE,
            Self::GeoTextStrikethroughFont => 0x00FF,
            Self::BlipCropFromTop => 0x0100,
            Self::BlipCropFromBottom => 0x0101,
            Self::BlipCropFromLeft => 0x0102,
            Self::BlipCropFromRight => 0x0103,
            Self::BlipToDisplay => 0x0104,
            Self::PictureFileName => 0x0105,
            Self::BlipFlags => 0x0106,
            Self::TransparentColor => 0x0107,
            Self::PictureContrast => 0x0108,
            Self::PictureBrightness => 0x0109,
            Self::PictureGamma => 0x010A,
            Self::PictureId => 0x010B,
            Self::DoubleMod => 0x010C,
            Self::PictureFillMod => 0x010D,
            Self::PictureLine => 0x010E,
            Self::PrintBlip => 0x010F,
            Self::PrintBlipFilename => 0x0110,
            Self::PrintFlags => 0x0111,
            Self::NoHitTestPicture => 0x013C,
            Self::PictureGray => 0x013D,
            Self::PictureBilevel => 0x013E,
            Self::PictureActive => 0x013F,
            Self::GeomLeft => 0x0140,
            Self::GeomTop => 0x0141,
            Self::GeomRight => 0x0142,
            Self::GeomBottom => 0x0143,
            Self::ShapePath => 0x0144,
            Self::Vertices => 0x0145,
            Self::SegmentInfo => 0x0146,
            Self::AdjustValue => 0x0147,
            Self::Adjust2Value => 0x0148,
            Self::Adjust3Value => 0x0149,
            Self::Adjust4Value => 0x014A,
            Self::Adjust5Value => 0x014B,
            Self::Adjust6Value => 0x014C,
            Self::Adjust7Value => 0x014D,
            Self::Adjust8Value => 0x014E,
            Self::Adjust9Value => 0x014F,
            Self::Adjust10Value => 0x0150,
            Self::ConnectionSites => 0x0151,
            Self::ConnectionSitesDir => 0x0152,
            Self::XLimo => 0x0153,
            Self::YLimo => 0x0154,
            Self::AdjustHandles => 0x0155,
            Self::Guides => 0x0156,
            Self::Inscribe => 0x0157,
            Self::Cxk => 0x0158,
            Self::Fragments => 0x0159,
            Self::ShadowOk => 0x017A,
            Self::ThreeDOk => 0x017B,
            Self::LineOk => 0x017C,
            Self::GeoTextOk => 0x017D,
            Self::FillShadeShapeOk => 0x017E,
            Self::FillOk => 0x017F,
            Self::FillType => 0x0180,
            Self::FillColor => 0x0181,
            Self::FillOpacity => 0x0182,
            Self::FillBackColor => 0x0183,
            Self::FillBackOpacity => 0x0184,
            Self::FillCrMod => 0x0185,
            Self::FillBlip => 0x0186,
            Self::FillBlipName => 0x0187,
            Self::FillBlipFlags => 0x0188,
            Self::FillWidth => 0x0189,
            Self::FillHeight => 0x018A,
            Self::FillAngle => 0x018B,
            Self::FillFocus => 0x018C,
            Self::FillToLeft => 0x018D,
            Self::FillToTop => 0x018E,
            Self::FillToRight => 0x018F,
            Self::FillToBottom => 0x0190,
            Self::FillRectLeft => 0x0191,
            Self::FillRectTop => 0x0192,
            Self::FillRectRight => 0x0193,
            Self::FillRectBottom => 0x0194,
            Self::FillDzType => 0x0195,
            Self::FillShadePreset => 0x0196,
            Self::FillShadeColors => 0x0197,
            Self::FillOriginX => 0x0198,
            Self::FillOriginY => 0x0199,
            Self::FillShapeOriginX => 0x019A,
            Self::FillShapeOriginY => 0x019B,
            Self::FillShadeType => 0x019C,
            Self::Filled => 0x01BB,
            Self::HitTestFill => 0x01BC,
            Self::FillShape => 0x01BD,
            Self::UseRect => 0x01BE,
            Self::NoFillHitTest => 0x01BF,
            Self::LineColor => 0x01C0,
            Self::LineOpacity => 0x01C1,
            Self::LineBackColor => 0x01C2,
            Self::LineCrMod => 0x01C3,
            Self::LineType => 0x01C4,
            Self::LineFillBlip => 0x01C5,
            Self::LineFillBlipName => 0x01C6,
            Self::LineFillBlipFlags => 0x01C7,
            Self::LineFillWidth => 0x01C8,
            Self::LineFillHeight => 0x01C9,
            Self::LineFillDzType => 0x01CA,
            Self::LineWidth => 0x01CB,
            Self::LineMiterLimit => 0x01CC,
            Self::LineStyle => 0x01CD,
            Self::LineDashing => 0x01CE,
            Self::LineDashStyle => 0x01CF,
            Self::LineStartArrowhead => 0x01D0,
            Self::LineEndArrowhead => 0x01D1,
            Self::LineStartArrowWidth => 0x01D2,
            Self::LineStartArrowLength => 0x01D3,
            Self::LineEndArrowWidth => 0x01D4,
            Self::LineEndArrowLength => 0x01D5,
            Self::LineJoinStyle => 0x01D6,
            Self::LineEndCapStyle => 0x01D7,
            Self::ArrowheadsOk => 0x01FB,
            Self::AnyLine => 0x01FC,
            Self::HitTestLine => 0x01FD,
            Self::LineFillShape => 0x01FE,
            Self::NoLineDrawDash => 0x01FF,
            Self::ShadowType => 0x0200,
            Self::ShadowColor => 0x0201,
            Self::ShadowHighlight => 0x0202,
            Self::ShadowCrMod => 0x0203,
            Self::ShadowOpacity => 0x0204,
            Self::ShadowOffsetX => 0x0205,
            Self::ShadowOffsetY => 0x0206,
            Self::ShadowSecondOffsetX => 0x0207,
            Self::ShadowSecondOffsetY => 0x0208,
            Self::ShadowScaleXToX => 0x0209,
            Self::ShadowScaleYToX => 0x020A,
            Self::ShadowScaleXToY => 0x020B,
            Self::ShadowScaleYToY => 0x020C,
            Self::ShadowPerspectiveX => 0x020D,
            Self::ShadowPerspectiveY => 0x020E,
            Self::ShadowWeight => 0x020F,
            Self::ShadowOriginX => 0x0210,
            Self::ShadowOriginY => 0x0211,
            Self::Shadow => 0x023E,
            Self::ShadowObscured => 0x023F,
            Self::PerspectiveType => 0x0240,
            Self::PerspectiveOffsetX => 0x0241,
            Self::PerspectiveOffsetY => 0x0242,
            Self::PerspectiveScaleXToX => 0x0243,
            Self::PerspectiveScaleYToX => 0x0244,
            Self::PerspectiveScaleXToY => 0x0245,
            Self::PerspectiveScaleYToY => 0x0246,
            Self::PerspectivePerspectiveX => 0x0247,
            Self::PerspectivePerspectiveY => 0x0248,
            Self::PerspectiveWeight => 0x0249,
            Self::PerspectiveOriginX => 0x024A,
            Self::PerspectiveOriginY => 0x024B,
            Self::PerspectiveOn => 0x027F,
            Self::ThreeDSpecularAmount => 0x0280,
            Self::ThreeDDiffuseAmount => 0x0281,
            Self::ThreeDShininess => 0x0282,
            Self::ThreeDEdgeThickness => 0x0283,
            Self::ThreeDExtrudeForward => 0x0284,
            Self::ThreeDExtrudeBackward => 0x0285,
            Self::ThreeDExtrusionColor => 0x0287,
            Self::ThreeDCrMod => 0x0288,
            Self::ThreeDExtrusionColorExt => 0x0289,
            Self::ThreeDEffect => 0x02BC,
            Self::ThreeDMetallic => 0x02BD,
            Self::ThreeDUseExtrusionColor => 0x02BE,
            Self::ThreeDLightFace => 0x02BF,
            Self::ThreeDStyleYRotationAngle => 0x02C0,
            Self::ThreeDStyleXRotationAngle => 0x02C1,
            Self::ThreeDStyleRotationAxisX => 0x02C2,
            Self::ThreeDStyleRotationAxisY => 0x02C3,
            Self::ThreeDStyleRotationAxisZ => 0x02C4,
            Self::ThreeDStyleRotationAngle => 0x02C5,
            Self::ThreeDStyleRotationCenterX => 0x02C6,
            Self::ThreeDStyleRotationCenterY => 0x02C7,
            Self::ThreeDStyleRotationCenterZ => 0x02C8,
            Self::ThreeDStyleRenderMode => 0x02C9,
            Self::ThreeDStyleTolerance => 0x02CA,
            Self::ThreeDStyleXViewpoint => 0x02CB,
            Self::ThreeDStyleYViewpoint => 0x02CC,
            Self::ThreeDStyleZViewpoint => 0x02CD,
            Self::ThreeDStyleOriginX => 0x02CE,
            Self::ThreeDStyleOriginY => 0x02CF,
            Self::ThreeDStyleSkewAngle => 0x02D0,
            Self::ThreeDStyleSkewAmount => 0x02D1,
            Self::ThreeDStyleAmbientIntensity => 0x02D2,
            Self::ThreeDStyleKeyX => 0x02D3,
            Self::ThreeDStyleKeyY => 0x02D4,
            Self::ThreeDStyleKeyZ => 0x02D5,
            Self::ThreeDStyleKeyIntensity => 0x02D6,
            Self::ThreeDStyleFillX => 0x02D7,
            Self::ThreeDStyleFillY => 0x02D8,
            Self::ThreeDStyleFillZ => 0x02D9,
            Self::ThreeDStyleFillIntensity => 0x02DA,
            Self::ShapeMaster => 0x0301,
            Self::ShapeConnectorStyle => 0x0303,
            Self::ShapeBlackAndWhiteSettings => 0x0304,
            Self::ShapeWModePureBw => 0x0305,
            Self::ShapeWModeBw => 0x0306,
            Self::ShapeOleIcon => 0x033A,
            Self::ShapePreferRelativeResize => 0x033B,
            Self::ShapeLockShapeType => 0x033C,
            Self::ShapeDeleteAttachedObject => 0x033E,
            Self::ShapeBackgroundShape => 0x033F,
            Self::CalloutType => 0x0340,
            Self::CalloutXYGap => 0x0341,
            Self::CalloutAngle => 0x0342,
            Self::CalloutDropType => 0x0343,
            Self::CalloutDrop => 0x0344,
            Self::CalloutLength => 0x0345,
            Self::GroupName => 0x0380,
            Self::GroupDescription => 0x0381,
            Self::Hyperlink => 0x0382,
            Self::GroupTableProperties => 0x039F,
            Self::GroupTableRowProperties => 0x03A0,
            Self::DiagramType => 0x0500,
            Self::DiagramStyle => 0x0501,
            Self::Unknown(raw) => raw.raw(),
        }
    }

    /// Returns whether this property is encoded as an `IMsoArray`.
    ///
    /// Classification is based on the property identifier, never on whether
    /// arbitrary complex bytes happen to resemble an array header.
    ///
    pub const fn is_array(self) -> bool {
        matches!(
            self.raw(),
            0x0145 // pVertices
                | 0x0146 // pSegmentInfo
                | 0x0151 // pConnectionSites
                | 0x0152 // pConnectionSitesDir
                | 0x0155 // pAdjustHandles
                | 0x0156 // pGuides
                | 0x0157 // pInscribe
                | 0x0197 // fillShadeColors
                | 0x01CF // lineDashStyle
                | 0x0383 // pWrapPolygonVertices
                | 0x03A0 // tableRowProperties
                | 0x0504 // pRelationTbl
                | 0x0508 // dgmConstrainBounds
                | 0x054F // lineLeftDashStyle
                | 0x058F // lineTopDashStyle
                | 0x05CF // lineRightDashStyle
                | 0x060F // lineBottomDashStyle
        )
    }
}

/// A lossless OfficeArt color reference.
///
/// Indirect palette, scheme, and system colors retain their exact bit pattern.
/// Call [`Self::rgb`] only when direct RGB semantics are required.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ColorRef(u32);

impl ColorRef {
    const FLAGS_MASK: u32 = 0xFF00_0000;

    /// Wraps an exact `OfficeArtCOLORREF` value.
    pub const fn from_raw(raw: u32) -> Self {
        Self(raw)
    }

    /// Returns the exact value read from the wire.
    pub const fn raw(self) -> u32 {
        self.0
    }

    /// Returns the high-byte flags without interpreting producer extensions.
    pub const fn flags(self) -> u8 {
        (self.0 >> 24) as u8
    }

    /// Returns an RGB triple only for a direct, unflagged color.
    ///
    /// `OfficeArtCOLORREF` stores red in the low byte, followed by green and
    /// blue. Flagged values are palette, scheme, or system references and are
    /// deliberately not flattened here.
    pub const fn rgb(self) -> Option<(u8, u8, u8)> {
        if self.0 & Self::FLAGS_MASK != 0 {
            return None;
        }
        Some((
            (self.0 & 0xFF) as u8,
            ((self.0 >> 8) & 0xFF) as u8,
            ((self.0 >> 16) & 0xFF) as u8,
        ))
    }
}

/// A decoded property value that continues to borrow complex bytes.
#[derive(Debug)]
pub enum Value<'data> {
    /// A four-byte scalar, retaining its signed wire representation.
    Simple(i32),
    /// Property-specific complex bytes.
    Complex(&'data [u8]),
    /// A validated `IMsoArray`.
    Array(Array<'data>),
}

/// A validated, zero-copy `IMsoArray` view.
#[derive(Debug, Clone, Copy)]
pub struct Array<'data> {
    data: &'data [u8],
}

impl<'data> Array<'data> {
    /// Validates an entire `IMsoArray`, including its exact payload extent.
    pub fn new(data: &'data [u8]) -> Result<Self> {
        let header = data.get(..6).ok_or(Error::MalformedProperties {
            reason: "array property is shorter than its six-byte header",
        })?;
        let count = u16::from_le_bytes([header[0], header[1]]);
        let allocated = u16::from_le_bytes([header[2], header[3]]);
        if allocated < count {
            return Err(Error::MalformedProperties {
                reason: "array allocation count is smaller than its element count",
            });
        }

        let raw_size = u16::from_le_bytes([header[4], header[5]]);
        let size = if raw_size == 0xFFF0 {
            4
        } else {
            usize::from(raw_size)
        };
        let payload_len =
            usize::from(count)
                .checked_mul(size)
                .ok_or(Error::ArithmeticOverflow {
                    context: "array-property payload length",
                })?;
        let expected = 6usize
            .checked_add(payload_len)
            .ok_or(Error::ArithmeticOverflow {
                context: "array-property extent",
            })?;
        if data.len() != expected {
            return Err(Error::MalformedProperties {
                reason: "array property does not have its exact declared extent",
            });
        }
        Ok(Self { data })
    }

    /// Returns the number of encoded elements.
    #[inline]
    pub fn element_count(&self) -> u16 {
        u16::from_le_bytes([self.data[0], self.data[1]])
    }

    /// Returns the maximum number of elements the producer allocated.
    #[inline]
    pub fn element_count_in_memory(&self) -> u16 {
        u16::from_le_bytes([self.data[2], self.data[3]])
    }

    /// Returns the exact unsigned `cbElem` field.
    #[inline]
    pub fn raw_element_size(&self) -> u16 {
        u16::from_le_bytes([self.data[4], self.data[5]])
    }

    /// Returns the number of encoded bytes per element.
    #[inline]
    pub fn element_size(&self) -> usize {
        match self.raw_element_size() {
            0xFFF0 => 4,
            size => usize::from(size),
        }
    }

    /// Borrows one encoded element.
    #[inline]
    pub fn get_element(&self, index: usize) -> Option<&'data [u8]> {
        if index >= usize::from(self.element_count()) {
            return None;
        }
        let size = self.element_size();
        let start = index.checked_mul(size)?.checked_add(6)?;
        let end = start.checked_add(size)?;
        self.data.get(start..end)
    }

    /// Iterates over every encoded element in order.
    pub fn elements(&self) -> impl Iterator<Item = &'data [u8]> {
        let count = usize::from(self.element_count());
        let array = *self;
        let mut index = 0;
        std::iter::from_fn(move || {
            if index == count {
                return None;
            }
            let element = array.get_element(index);
            index += 1;
            element
        })
    }

    /// Returns the complete encoded array, including its header.
    #[inline]
    pub fn raw_data(&self) -> &'data [u8] {
        self.data
    }
}

/// One ordered, lossless property-table descriptor and its decoded value.
#[derive(Debug)]
pub struct Prop<'data> {
    id: Id,
    raw_id: u16,
    blip: bool,
    complex: bool,
    raw_value: i32,
    value: Value<'data>,
}

impl<'data> Prop<'data> {
    /// Returns the typed identifier.
    pub const fn id(&self) -> Id {
        self.id
    }

    /// Returns the exact 14-bit identifier without flags.
    pub const fn raw_id(&self) -> u16 {
        self.raw_id
    }

    /// Reassembles the exact 16-bit identifier-and-flags field.
    pub const fn raw_opid(&self) -> u16 {
        self.raw_id
            | if self.blip { IS_BLIP } else { 0 }
            | if self.complex { IS_COMPLEX } else { 0 }
    }

    /// Returns whether `fBid` was set on the wire.
    pub const fn is_blip(&self) -> bool {
        self.blip
    }

    /// Returns whether `fComplex` was set on the wire.
    pub const fn is_complex(&self) -> bool {
        self.complex
    }

    /// Returns the exact signed four-byte `op` value.
    pub const fn raw_value(&self) -> i32 {
        self.raw_value
    }

    /// Borrows the decoded value.
    pub const fn value(&self) -> &Value<'data> {
        &self.value
    }
}

#[derive(Debug, Clone, Copy)]
struct Descriptor {
    id: Id,
    raw_id: u16,
    blip: bool,
    complex: bool,
    raw_value: i32,
}

/// An ordered OfficeArt shape-property collection.
#[derive(Debug)]
pub struct Props<'data> {
    properties: Vec<Prop<'data>>,
    by_id: HashMap<Id, usize>,
}

impl<'data> Props<'data> {
    /// Creates an empty property collection.
    #[inline]
    pub fn new() -> Self {
        Self {
            properties: Vec::new(),
            by_id: HashMap::new(),
        }
    }

    /// Parses an Opt-family record while borrowing all complex values.
    pub fn parse(opt: &Record<'data>) -> Result<Self> {
        if !matches!(
            opt.kind(),
            RecordKind::Opt | RecordKind::SecondaryOpt | RecordKind::TertiaryOpt
        ) {
            return Err(Error::MalformedProperties {
                reason: "record is not an Opt-family property table",
            });
        }
        if opt.version() != 3 {
            return Err(Error::MalformedProperties {
                reason: "Opt-family property table must have recVer 3",
            });
        }

        let count = usize::from(opt.instance());
        let data = opt.data();
        let header_size = count.checked_mul(6).ok_or(Error::ArithmeticOverflow {
            context: "property-header table length",
        })?;
        if header_size > data.len() {
            return Err(Error::MalformedProperties {
                reason: "property-header table exceeds recLen",
            });
        }

        let mut descriptors = Vec::with_capacity(count);
        let mut by_id = HashMap::with_capacity(count);
        for index in 0..count {
            let offset = index.checked_mul(6).ok_or(Error::ArithmeticOverflow {
                context: "property-header offset",
            })?;
            let end = offset.checked_add(6).ok_or(Error::ArithmeticOverflow {
                context: "property-header extent",
            })?;
            let header = data.get(offset..end).ok_or(Error::MalformedProperties {
                reason: "truncated property header",
            })?;
            let raw_opid = u16::from_le_bytes([header[0], header[1]]);
            let raw_id = raw_opid & PROPERTY_ID_MASK;
            let id = Id::from(raw_id);
            if by_id.insert(id, index).is_some() {
                return Err(Error::MalformedProperties {
                    reason: "duplicate property identifier",
                });
            }
            descriptors.push(Descriptor {
                id,
                raw_id,
                blip: raw_opid & IS_BLIP != 0,
                complex: raw_opid & IS_COMPLEX != 0,
                raw_value: i32::from_le_bytes([header[2], header[3], header[4], header[5]]),
            });
        }

        let mut properties = Vec::with_capacity(count);
        let mut complex_offset = header_size;
        for descriptor in descriptors {
            let value = if descriptor.complex {
                let complex_len = usize::try_from(descriptor.raw_value).map_err(|_| {
                    Error::MalformedProperties {
                        reason: "negative complex-property length",
                    }
                })?;
                let complex_end =
                    complex_offset
                        .checked_add(complex_len)
                        .ok_or(Error::ArithmeticOverflow {
                            context: "complex-property extent",
                        })?;
                let complex =
                    data.get(complex_offset..complex_end)
                        .ok_or(Error::MalformedProperties {
                            reason: "complex property exceeds recLen",
                        })?;
                complex_offset = complex_end;

                if descriptor.id.is_array() {
                    Value::Array(Array::new(complex)?)
                } else {
                    Value::Complex(complex)
                }
            } else {
                Value::Simple(descriptor.raw_value)
            };
            properties.push(Prop {
                id: descriptor.id,
                raw_id: descriptor.raw_id,
                blip: descriptor.blip,
                complex: descriptor.complex,
                raw_value: descriptor.raw_value,
                value,
            });
        }

        if complex_offset != data.len() {
            return Err(Error::MalformedProperties {
                reason: "unclaimed bytes follow the property table",
            });
        }
        Ok(Self { properties, by_id })
    }

    /// Parses the primary Opt child when present.
    pub fn from_container(container: &Container<'data>) -> Result<Self> {
        match container.find(RecordKind::Opt)? {
            Some(opt) => Self::parse(&opt),
            None => Ok(Self::new()),
        }
    }

    /// Returns the complete lossless descriptor for `id`.
    #[inline]
    pub fn prop(&self, id: Id) -> Option<&Prop<'data>> {
        self.by_id
            .get(&id)
            .and_then(|index| self.properties.get(*index))
    }

    /// Returns the decoded value for `id`.
    #[inline]
    pub fn get(&self, id: Id) -> Option<&Value<'data>> {
        self.prop(id).map(Prop::value)
    }

    /// Returns a simple signed value.
    #[inline]
    pub fn get_int(&self, id: Id) -> Option<i32> {
        match self.get(id) {
            Some(Value::Simple(v)) => Some(*v),
            _ => None,
        }
    }

    /// Returns a typed, lossless color reference.
    #[inline]
    pub fn get_color(&self, id: Id) -> Option<ColorRef> {
        self.get_int(id)
            .map(|value| ColorRef::from_raw(value as u32))
    }

    /// Resolves an explicitly encoded boolean without applying defaults.
    #[inline]
    pub fn get_bool(&self, id: Id) -> Option<bool> {
        let raw_id = id.raw();
        if let Some(terminal_id) = boolean_group_terminal(raw_id) {
            let terminal = Id::from(terminal_id);
            if let Some(value) = self.get_int(terminal) {
                let bit = u32::from(terminal_id - raw_id);
                let value_mask = 1u32 << bit;
                let use_mask = value_mask << 16;
                let value = value as u32;
                return (value & use_mask != 0).then_some(value & value_mask != 0);
            }
        }
        self.get_int(id).map(|value| value != 0)
    }

    /// Returns property-specific complex bytes.
    #[inline]
    pub fn get_binary(&self, id: Id) -> Option<&'data [u8]> {
        match self.get(id) {
            Some(Value::Complex(data)) => Some(data),
            _ => None,
        }
    }

    /// Returns a validated array property.
    #[inline]
    pub fn get_array(&self, id: Id) -> Option<&Array<'data>> {
        match self.get(id) {
            Some(Value::Array(array)) => Some(array),
            _ => None,
        }
    }

    /// Returns whether `id` is present.
    #[inline]
    pub fn has(&self, id: Id) -> bool {
        self.by_id.contains_key(&id)
    }

    /// Returns the number of descriptors.
    #[inline]
    pub fn len(&self) -> usize {
        self.properties.len()
    }

    /// Returns whether the collection has no descriptors.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.properties.is_empty()
    }

    /// Iterates over lossless descriptors in their original wire order.
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Prop<'data>> {
        self.properties.iter()
    }

    /// Returns direct RGB only; indirect color references return `None`.
    #[inline]
    pub fn get_rgb(&self, id: Id) -> Option<(u8, u8, u8)> {
        self.get_color(id).and_then(ColorRef::rgb)
    }

    #[inline]
    pub fn get_rotation_degrees(&self, id: Id) -> Option<f32> {
        self.get_int(id)
            .map(|fixed_point| (fixed_point as f32) / 65536.0)
    }

    #[inline]
    pub fn get_opacity(&self, id: Id) -> Option<f32> {
        self.get_int(id).map(|fixed_point| {
            let opacity = (fixed_point as f32) / 65536.0;
            opacity.clamp(0.0, 1.0)
        })
    }

    #[inline]
    pub fn get_coord(&self, id: Id) -> Option<i32> {
        self.get_int(id)
    }

    /// Returns whether a boolean is explicitly enabled, treating absence as
    /// false. Use [`Self::is_filled`] and [`Self::has_line`] for properties
    /// whose OfficeArt defaults are true.
    #[inline]
    pub fn is_true(&self, id: Id) -> bool {
        self.get_bool(id).unwrap_or(false)
    }

    #[inline]
    pub fn get_line_width(&self) -> Option<i32> {
        self.get_int(Id::LineWidth)
    }

    #[inline]
    pub fn get_fill_color(&self) -> Option<(u8, u8, u8)> {
        self.get_rgb(Id::FillColor)
    }

    #[inline]
    pub fn get_line_color(&self) -> Option<(u8, u8, u8)> {
        self.get_rgb(Id::LineColor)
    }

    /// Resolves the fill-enabled bit, whose specification default is `true`.
    #[inline]
    pub fn is_filled(&self) -> bool {
        self.get_bool(Id::Filled).unwrap_or(true)
    }

    /// Resolves the line-enabled bit, whose specification default is `true`.
    #[inline]
    pub fn has_line(&self) -> bool {
        self.get_bool(Id::AnyLine).unwrap_or(true)
    }

    #[inline]
    pub fn has_shadow(&self) -> bool {
        self.is_true(Id::Shadow)
    }

    #[inline]
    pub fn get_geometry_rect(&self) -> Option<(i32, i32, i32, i32)> {
        let left = self.get_coord(Id::GeomLeft)?;
        let top = self.get_coord(Id::GeomTop)?;
        let right = self.get_coord(Id::GeomRight)?;
        let bottom = self.get_coord(Id::GeomBottom)?;
        Some((left, top, right, bottom))
    }

    #[inline]
    pub fn get_text_margins(&self) -> Option<(i32, i32, i32, i32)> {
        const HORIZONTAL_DEFAULT: i32 = 0x0001_6530;
        const VERTICAL_DEFAULT: i32 = 0x0000_B298;
        let left = self.get_int(Id::TextLeft).unwrap_or(HORIZONTAL_DEFAULT);
        let top = self.get_int(Id::TextTop).unwrap_or(VERTICAL_DEFAULT);
        let right = self.get_int(Id::TextRight).unwrap_or(HORIZONTAL_DEFAULT);
        let bottom = self.get_int(Id::TextBottom).unwrap_or(VERTICAL_DEFAULT);
        Some((left, top, right, bottom))
    }

    #[inline]
    pub fn get_adjust(&self, id: Id) -> Option<i32> {
        self.get_int(id)
    }
}

fn boolean_group_terminal(id: u16) -> Option<u16> {
    match id {
        0x0077..=0x007F => Some(0x007F),
        0x00BB..=0x00BF => Some(0x00BF),
        0x00FA..=0x00FF => Some(0x00FF),
        0x013C..=0x013F => Some(0x013F),
        0x017A..=0x017F => Some(0x017F),
        0x01BB..=0x01BF => Some(0x01BF),
        0x01FB..=0x01FF => Some(0x01FF),
        0x023E..=0x023F => Some(0x023F),
        0x027F => Some(0x027F),
        0x02BC..=0x02BF => Some(0x02BF),
        0x033A..=0x033F => Some(0x033F),
        _ => None,
    }
}

impl<'data> Default for Props<'data> {
    fn default() -> Self {
        Self::new()
    }
}

/// Shape anchor (position and size).
#[derive(Debug, Clone, Copy)]
pub struct Anchor {
    pub left: i32,
    pub top: i32,
    pub right: i32,
    pub bottom: i32,
}

impl Anchor {
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
    pub const fn width(&self) -> Option<i32> {
        self.right.checked_sub(self.left)
    }

    #[inline]
    pub const fn height(&self) -> Option<i32> {
        self.bottom.checked_sub(self.top)
    }

    pub fn from_child_anchor(anchor: &Record) -> Option<Self> {
        if anchor.kind() != RecordKind::ChildAnchor || anchor.data().len() != 16 {
            return None;
        }

        let left = i32::from_le_bytes([
            anchor.data()[0],
            anchor.data()[1],
            anchor.data()[2],
            anchor.data()[3],
        ]);
        let top = i32::from_le_bytes([
            anchor.data()[4],
            anchor.data()[5],
            anchor.data()[6],
            anchor.data()[7],
        ]);
        let right = i32::from_le_bytes([
            anchor.data()[8],
            anchor.data()[9],
            anchor.data()[10],
            anchor.data()[11],
        ]);
        let bottom = i32::from_le_bytes([
            anchor.data()[12],
            anchor.data()[13],
            anchor.data()[14],
            anchor.data()[15],
        ]);

        Some(Self::new(left, top, right, bottom))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn opt_record(data: &[u8], instance: u16) -> Record<'_> {
        Record::from_parts(RecordKind::Opt, 3, instance, data).expect("valid fixture")
    }

    fn opt_record_with_version(data: &[u8], instance: u16, version: u8) -> Record<'_> {
        Record::from_parts(RecordKind::Opt, version, instance, data).expect("valid fixture")
    }

    fn push_property(data: &mut Vec<u8>, id: u16, value: i32) {
        data.extend_from_slice(&id.to_le_bytes());
        data.extend_from_slice(&value.to_le_bytes());
    }

    #[test]
    fn decodes_packed_boolean_property_groups() {
        let mut data = Vec::new();
        push_property(&mut data, 0x01BF, 0x0014_0010);
        push_property(&mut data, 0x01FF, 0x0008_0008);
        push_property(&mut data, 0x023F, 0x0002_0002);
        push_property(&mut data, 0x00FF, 0x0020_0020);
        push_property(&mut data, 0x033F, 0x0001_0001);
        push_property(&mut data, 0x00BF, 0x001A_0012);

        let properties = Props::parse(&opt_record(&data, 6)).expect("valid properties");

        assert_eq!(properties.get_bool(Id::Filled), Some(true));
        assert_eq!(properties.get_bool(Id::FillShape), Some(false));
        assert_eq!(properties.get_bool(Id::HitTestFill), None);
        assert!(properties.is_filled());
        assert_eq!(properties.get_bool(Id::AnyLine), Some(true));
        assert!(properties.has_line());
        assert_eq!(properties.get_bool(Id::Shadow), Some(true));
        assert_eq!(properties.get_bool(Id::ShadowObscured), None);
        assert!(properties.has_shadow());
        assert_eq!(properties.get_bool(Id::GeoTextBoldFont), Some(true));
        assert_eq!(properties.get_bool(Id::GeoTextUnderlineFont), None);
        assert_eq!(properties.get_bool(Id::ShapeBackgroundShape), Some(true));
        assert_eq!(properties.get_bool(Id::SelectText), Some(true));
        assert_eq!(properties.get_bool(Id::AutoTextMargin), Some(false));
        assert_eq!(properties.get_bool(Id::FitShapeToText), Some(true));
    }

    #[test]
    fn decodes_explicit_false_boolean_group_bits() {
        let mut data = Vec::new();
        push_property(&mut data, 0x01FF, 0x0008_0000);
        push_property(&mut data, 0x023F, 0x0002_0000);

        let properties = Props::parse(&opt_record(&data, 2)).expect("valid properties");

        assert_eq!(properties.get_bool(Id::AnyLine), Some(false));
        assert!(!properties.has_line());
        assert_eq!(properties.get_bool(Id::Shadow), Some(false));
        assert!(!properties.has_shadow());
    }

    #[test]
    fn text_margins_use_ms_odraw_defaults() {
        let properties = Props::new();
        assert_eq!(
            properties.get_text_margins(),
            Some((0x0001_6530, 0x0000_B298, 0x0001_6530, 0x0000_B298))
        );
    }

    #[test]
    fn fill_and_line_resolvers_apply_spec_defaults_without_hiding_absence() {
        let properties = Props::new();

        assert_eq!(properties.get_bool(Id::Filled), None);
        assert_eq!(properties.get_bool(Id::AnyLine), None);
        assert!(properties.is_filled());
        assert!(properties.has_line());
        assert!(!properties.has_shadow());
    }

    #[test]
    fn decodes_writer_fill_enabled_and_disabled_masks() {
        let mut enabled_data = Vec::new();
        push_property(&mut enabled_data, 0x01BF, 0x0015_0011);
        let enabled = Props::parse(&opt_record(&enabled_data, 1)).expect("valid properties");

        assert_eq!(enabled.get_bool(Id::Filled), Some(true));
        assert_eq!(enabled.get_bool(Id::FillShape), Some(false));
        assert_eq!(enabled.get_bool(Id::NoFillHitTest), Some(true));
        assert_eq!(enabled.get_bool(Id::HitTestFill), None);

        let mut disabled_data = Vec::new();
        push_property(&mut disabled_data, 0x01BF, 0x0010_0000);
        let disabled = Props::parse(&opt_record(&disabled_data, 1)).expect("valid properties");
        assert_eq!(disabled.get_bool(Id::Filled), Some(false));
    }

    #[test]
    fn accepts_direct_boolean_properties_from_lenient_producers() {
        let mut data = Vec::new();
        push_property(&mut data, Id::Filled.raw(), 1);

        let properties = Props::parse(&opt_record(&data, 1)).expect("valid properties");

        assert_eq!(properties.get_bool(Id::Filled), Some(true));
    }

    #[test]
    fn rejects_negative_complex_property_lengths_without_panicking() {
        let mut data = Vec::new();
        push_property(&mut data, IS_COMPLEX | Id::Vertices.raw(), -1);

        assert!(matches!(
            Props::parse(&opt_record(&data, 1)),
            Err(Error::MalformedProperties { .. })
        ));
    }

    #[test]
    fn preserves_distinct_unknown_property_ids() {
        let mut data = Vec::new();
        push_property(&mut data, 0x0600, 11);
        push_property(&mut data, 0x0601, 12);

        let properties = Props::parse(&opt_record(&data, 2)).expect("valid properties");
        let first = Id::unknown(0x0600).expect("unassigned identifier");
        let second = Id::unknown(0x0601).expect("unassigned identifier");
        assert_eq!(first.raw(), 0x0600);
        assert_eq!(properties.get_int(first), Some(11));
        assert_eq!(properties.get_int(second), Some(12));
        assert!(Id::unknown(IS_BLIP | 0x0600).is_none());
        assert!(Id::unknown(Id::FillColor.raw()).is_none());
    }

    #[test]
    fn preserves_order_flags_raw_ids_and_raw_values() {
        let mut data = Vec::new();
        push_property(&mut data, IS_BLIP | 0x0601, -7);
        push_property(&mut data, IS_BLIP | IS_COMPLEX | Id::GroupName.raw(), 3);
        data.extend_from_slice(&[0x41, 0x00, 0x00]);

        let properties = Props::parse(&opt_record(&data, 2)).expect("valid properties");
        let entries = properties.iter().collect::<Vec<_>>();
        assert_eq!(entries[0].raw_id(), 0x0601);
        assert_eq!(entries[0].raw_opid(), IS_BLIP | 0x0601);
        assert!(entries[0].is_blip());
        assert!(!entries[0].is_complex());
        assert_eq!(entries[0].raw_value(), -7);
        assert_eq!(entries[1].id(), Id::GroupName);
        assert_eq!(
            entries[1].raw_opid(),
            IS_BLIP | IS_COMPLEX | Id::GroupName.raw()
        );
        assert!(entries[1].is_blip());
        assert!(entries[1].is_complex());
        assert_eq!(entries[1].raw_value(), 3);
        assert_eq!(
            properties.get_binary(Id::GroupName),
            Some(&[0x41, 0, 0][..])
        );
    }

    #[test]
    fn rejects_duplicate_semantic_identifiers_even_when_flags_differ() {
        let mut data = Vec::new();
        push_property(&mut data, Id::FillColor.raw(), 1);
        push_property(&mut data, IS_BLIP | Id::FillColor.raw(), 2);

        assert!(matches!(
            Props::parse(&opt_record(&data, 2)),
            Err(Error::MalformedProperties {
                reason: "duplicate property identifier"
            })
        ));
    }

    #[test]
    fn requires_opt_family_version_three() {
        assert!(matches!(
            Props::parse(&opt_record_with_version(&[], 0, 2)),
            Err(Error::MalformedProperties {
                reason: "Opt-family property table must have recVer 3"
            })
        ));
    }

    #[test]
    fn classifies_arrays_by_property_id_instead_of_payload_shape() {
        let array_bytes = [0, 0, 0, 0, 4, 0];

        let mut scalar_complex = Vec::new();
        push_property(&mut scalar_complex, IS_COMPLEX | Id::GroupName.raw(), 6);
        scalar_complex.extend_from_slice(&array_bytes);
        let scalar = Props::parse(&opt_record(&scalar_complex, 1)).expect("valid complex value");
        assert!(matches!(scalar.get(Id::GroupName), Some(Value::Complex(_))));

        let mut typed_array = Vec::new();
        push_property(&mut typed_array, IS_COMPLEX | Id::Vertices.raw(), 6);
        typed_array.extend_from_slice(&array_bytes);
        let array = Props::parse(&opt_record(&typed_array, 1)).expect("valid array value");
        assert!(matches!(array.get(Id::Vertices), Some(Value::Array(_))));
    }

    #[test]
    fn validates_imsoarray_header_and_exact_extent() {
        let mut special = Vec::new();
        special.extend_from_slice(&2u16.to_le_bytes());
        special.extend_from_slice(&3u16.to_le_bytes());
        special.extend_from_slice(&0xFFF0u16.to_le_bytes());
        special.extend_from_slice(&[0; 8]);
        let special = Array::new(&special).expect("valid truncated-element array");
        assert_eq!(special.raw_element_size(), 0xFFF0);
        assert_eq!(special.element_size(), 4);
        assert_eq!(special.elements().count(), 2);

        let mut underallocated = Vec::new();
        underallocated.extend_from_slice(&2u16.to_le_bytes());
        underallocated.extend_from_slice(&1u16.to_le_bytes());
        underallocated.extend_from_slice(&4u16.to_le_bytes());
        underallocated.extend_from_slice(&[0; 8]);
        assert!(Array::new(&underallocated).is_err());

        let mut other_high_size = Vec::new();
        other_high_size.extend_from_slice(&1u16.to_le_bytes());
        other_high_size.extend_from_slice(&1u16.to_le_bytes());
        other_high_size.extend_from_slice(&0xFFF1u16.to_le_bytes());
        other_high_size.extend_from_slice(&[0; 4]);
        assert!(Array::new(&other_high_size).is_err());

        let mut trailing = Vec::new();
        trailing.extend_from_slice(&1u16.to_le_bytes());
        trailing.extend_from_slice(&1u16.to_le_bytes());
        trailing.extend_from_slice(&4u16.to_le_bytes());
        trailing.extend_from_slice(&[0; 5]);
        assert!(Array::new(&trailing).is_err());
    }

    #[test]
    fn color_ref_is_lossless_and_decodes_only_direct_rgb() {
        let direct = ColorRef::from_raw(0x0033_2211);
        assert_eq!(direct.raw(), 0x0033_2211);
        assert_eq!(direct.flags(), 0);
        assert_eq!(direct.rgb(), Some((0x11, 0x22, 0x33)));

        let scheme = ColorRef::from_raw(0x0800_0004);
        assert_eq!(scheme.raw(), 0x0800_0004);
        assert_eq!(scheme.flags(), 0x08);
        assert_eq!(scheme.rgb(), None);
    }
}
