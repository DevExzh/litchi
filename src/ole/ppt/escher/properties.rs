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

// Property ID flags (from MS-ODRAW)
const IS_BLIP: u16 = 0x4000; // Bit 14: is blip ID
const IS_COMPLEX: u16 = 0x8000; // Bit 15: is complex property
const PROPERTY_ID_MASK: u16 = 0x3FFF; // Lower 14 bits: property number

/// Escher shape property IDs (from MS-ODRAW and Apache POI).
///
/// This enum represents property IDs used in Office drawings.
/// The actual property number is stored in the lower 14 bits of the property ID.
///
/// Property groups are organized by their upper byte:
/// - 0x00xx: Transform properties
/// - 0x01xx: Protection and text properties  
/// - 0x01xx: Blip/picture properties
/// - 0x01xx: Geometry properties
/// - 0x01xx: Fill properties
/// - 0x01xx: Line properties
/// - 0x02xx: Shadow properties
/// - 0x02xx: Perspective properties
/// - 0x02xx: 3D properties
/// - 0x03xx: Shape properties
/// - 0x03xx: Callout properties
/// - 0x03xx: Group properties
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum EscherPropertyId {
    // Transform properties (0x0000 - 0x003F)
    /// Rotation angle (16.16 fixed point degrees)
    Rotation = 0x0004,

    // Protection properties (0x0040 - 0x007F)
    /// Lock rotation
    LockRotation = 0x0077,
    /// Lock aspect ratio
    LockAspectRatio = 0x0078,
    /// Lock position
    LockPosition = 0x0079,
    /// Lock against select
    LockAgainstSelect = 0x007A,
    /// Lock cropping
    LockCropping = 0x007B,
    /// Lock vertices
    LockVertices = 0x007C,
    /// Lock text
    LockText = 0x007D,
    /// Lock adjust handles
    LockAdjustHandles = 0x007E,
    /// Lock against grouping
    LockAgainstGrouping = 0x007F,

    // Text properties (0x0080 - 0x00BF)
    /// Text ID (reference to text)
    TextId = 0x0080,
    /// Text left margin
    TextLeft = 0x0081,
    /// Text top margin
    TextTop = 0x0082,
    /// Text right margin
    TextRight = 0x0083,
    /// Text bottom margin
    TextBottom = 0x0084,
    /// Wrap text
    WrapText = 0x0085,
    /// Scale text
    ScaleText = 0x0086,
    /// Anchor text
    AnchorText = 0x0087,
    /// Text flow
    TextFlow = 0x0088,
    /// Font rotation
    FontRotation = 0x0089,
    /// ID of next shape
    IdOfNextShape = 0x008A,
    /// Bidirectional text
    TextBidi = 0x008B,
    /// Single click selects
    SingleClickSelects = 0x00BB,
    /// Use host margins
    UseHostMargins = 0x00BC,
    /// Rotate text with shape
    RotateTextWithShape = 0x00BD,
    /// Size shape to fit text
    SizeShapeToFitText = 0x00BE,
    /// Size text to fit shape
    SizeTextToFitShape = 0x00BF,

    // GeoText properties (0x00C0 - 0x00FF)
    /// Unicode text
    GeoTextUnicode = 0x00C0,
    /// RTF text
    GeoTextRtf = 0x00C1,
    /// Alignment on curve
    GeoTextAlignmentOnCurve = 0x00C2,
    /// Default point size
    GeoTextDefaultPointSize = 0x00C3,
    /// Text spacing
    GeoTextSpacing = 0x00C4,
    /// Font family name
    GeoTextFontFamilyName = 0x00C5,
    /// Bold font
    GeoTextBoldFont = 0x00FA,
    /// Italic font
    GeoTextItalicFont = 0x00FB,
    /// Underline font
    GeoTextUnderlineFont = 0x00FC,
    /// Shadow font
    GeoTextShadowFont = 0x00FD,
    /// Small caps font
    GeoTextSmallCapsFont = 0x00FE,
    /// Strikethrough font
    GeoTextStrikethroughFont = 0x00FF,

    // Blip/Picture properties (0x0100 - 0x013F)
    /// Crop from top
    BlipCropFromTop = 0x0100,
    /// Crop from bottom
    BlipCropFromBottom = 0x0101,
    /// Crop from left
    BlipCropFromLeft = 0x0102,
    /// Crop from right
    BlipCropFromRight = 0x0103,
    /// Blip to display (reference)
    BlipToDisplay = 0x0104,
    /// Picture file name
    PictureFileName = 0x0105,
    /// Blip flags
    BlipFlags = 0x0106,
    /// Transparent color
    TransparentColor = 0x0107,
    /// Picture contrast
    PictureContrast = 0x0108,
    /// Picture brightness
    PictureBrightness = 0x0109,
    /// Gamma
    PictureGamma = 0x010A,
    /// Picture ID
    PictureId = 0x010B,
    /// Double mod
    DoubleMod = 0x010C,
    /// Picture fill mod
    PictureFillMod = 0x010D,
    /// Picture line
    PictureLine = 0x010E,
    /// Print blip
    PrintBlip = 0x010F,
    /// Print blip filename
    PrintBlipFilename = 0x0110,
    /// Print flags
    PrintFlags = 0x0111,
    /// No hit test picture
    NoHitTestPicture = 0x013C,
    /// Picture gray
    PictureGray = 0x013D,
    /// Picture bilevel
    PictureBilevel = 0x013E,
    /// Picture active
    PictureActive = 0x013F,

    // Geometry properties (0x0140 - 0x017F)
    /// Geometry left
    GeomLeft = 0x0140,
    /// Geometry top
    GeomTop = 0x0141,
    /// Geometry right
    GeomRight = 0x0142,
    /// Geometry bottom
    GeomBottom = 0x0143,
    /// Shape path (complex)
    ShapePath = 0x0144,
    /// Vertices (complex array)
    Vertices = 0x0145,
    /// Segment info (complex array)
    SegmentInfo = 0x0146,
    /// Adjust value
    AdjustValue = 0x0147,
    /// Adjust 2 value
    Adjust2Value = 0x0148,
    /// Adjust 3 value
    Adjust3Value = 0x0149,
    /// Adjust 4 value
    Adjust4Value = 0x014A,
    /// Adjust 5 value
    Adjust5Value = 0x014B,
    /// Adjust 6 value
    Adjust6Value = 0x014C,
    /// Adjust 7 value
    Adjust7Value = 0x014D,
    /// Adjust 8 value
    Adjust8Value = 0x014E,
    /// Adjust 9 value
    Adjust9Value = 0x014F,
    /// Adjust 10 value
    Adjust10Value = 0x0150,
    /// Connection sites
    ConnectionSites = 0x0151,
    /// Connection sites direction
    ConnectionSitesDir = 0x0152,
    /// X limit origin
    XLimo = 0x0153,
    /// Y limit origin
    YLimo = 0x0154,
    /// Adjust handles
    AdjustHandles = 0x0155,
    /// Guides
    Guides = 0x0156,
    /// Inscribe
    Inscribe = 0x0157,
    /// CXK
    Cxk = 0x0158,
    /// Fragments
    Fragments = 0x0159,
    /// Shadow OK
    ShadowOk = 0x017A,
    /// 3D OK
    ThreeDOk = 0x017B,
    /// Line OK
    LineOk = 0x017C,
    /// Geo text OK
    GeoTextOk = 0x017D,
    /// Fill shade shape OK
    FillShadeShapeOk = 0x017E,
    /// Fill OK
    FillOk = 0x017F,

    // Fill properties (0x0180 - 0x01BF)
    /// Fill type
    FillType = 0x0180,
    /// Fill color (RGB)
    FillColor = 0x0181,
    /// Fill opacity
    FillOpacity = 0x0182,
    /// Fill back color (RGB)
    FillBackColor = 0x0183,
    /// Back opacity
    FillBackOpacity = 0x0184,
    /// Color mod
    FillCrMod = 0x0185,
    /// Fill blip (picture/texture)
    FillBlip = 0x0186,
    /// Fill blip name
    FillBlipName = 0x0187,
    /// Fill blip flags
    FillBlipFlags = 0x0188,
    /// Fill width
    FillWidth = 0x0189,
    /// Fill height
    FillHeight = 0x018A,
    /// Fill angle
    FillAngle = 0x018B,
    /// Fill focus
    FillFocus = 0x018C,
    /// Fill to left
    FillToLeft = 0x018D,
    /// Fill to top
    FillToTop = 0x018E,
    /// Fill to right
    FillToRight = 0x018F,
    /// Fill to bottom
    FillToBottom = 0x0190,
    /// Fill rect left
    FillRectLeft = 0x0191,
    /// Fill rect top
    FillRectTop = 0x0192,
    /// Fill rect right
    FillRectRight = 0x0193,
    /// Fill rect bottom
    FillRectBottom = 0x0194,
    /// Fill DZ type
    FillDzType = 0x0195,
    /// Fill shade preset
    FillShadePreset = 0x0196,
    /// Fill shade colors (array)
    FillShadeColors = 0x0197,
    /// Fill origin X
    FillOriginX = 0x0198,
    /// Fill origin Y
    FillOriginY = 0x0199,
    /// Fill shape origin X
    FillShapeOriginX = 0x019A,
    /// Fill shape origin Y
    FillShapeOriginY = 0x019B,
    /// Fill shade type
    FillShadeType = 0x019C,
    /// Filled
    Filled = 0x01BB,
    /// Hit test fill
    HitTestFill = 0x01BC,
    /// Fill shape
    FillShape = 0x01BD,
    /// Use rect
    UseRect = 0x01BE,
    /// No fill hit test
    NoFillHitTest = 0x01BF,

    // Line properties (0x01C0 - 0x01FF)
    /// Line color (RGB)
    LineColor = 0x01C0,
    /// Line opacity
    LineOpacity = 0x01C1,
    /// Line back color
    LineBackColor = 0x01C2,
    /// Line color mod
    LineCrMod = 0x01C3,
    /// Line type
    LineType = 0x01C4,
    /// Line fill blip
    LineFillBlip = 0x01C5,
    /// Line fill blip name
    LineFillBlipName = 0x01C6,
    /// Line fill blip flags
    LineFillBlipFlags = 0x01C7,
    /// Line fill width
    LineFillWidth = 0x01C8,
    /// Line fill height
    LineFillHeight = 0x01C9,
    /// Line fill DZ type
    LineFillDzType = 0x01CA,
    /// Line width
    LineWidth = 0x01CB,
    /// Line miter limit
    LineMiterLimit = 0x01CC,
    /// Line style
    LineStyle = 0x01CD,
    /// Line dash style
    LineDashing = 0x01CE,
    /// Line dash style array
    LineDashStyle = 0x01CF,
    /// Line start arrow head
    LineStartArrowhead = 0x01D0,
    /// Line end arrow head
    LineEndArrowhead = 0x01D1,
    /// Line start arrow width
    LineStartArrowWidth = 0x01D2,
    /// Line start arrow length
    LineStartArrowLength = 0x01D3,
    /// Line end arrow width
    LineEndArrowWidth = 0x01D4,
    /// Line end arrow length
    LineEndArrowLength = 0x01D5,
    /// Line join style
    LineJoinStyle = 0x01D6,
    /// Line end cap style
    LineEndCapStyle = 0x01D7,
    /// Arrow heads OK
    ArrowheadsOk = 0x01FB,
    /// Any line
    AnyLine = 0x01FC,
    /// Hit test line
    HitTestLine = 0x01FD,
    /// Line fill shape
    LineFillShape = 0x01FE,
    /// No line draw dash
    NoLineDrawDash = 0x01FF,

    // Shadow properties (0x0200 - 0x023F)
    /// Shadow type
    ShadowType = 0x0200,
    /// Shadow color (RGB)
    ShadowColor = 0x0201,
    /// Shadow highlight
    ShadowHighlight = 0x0202,
    /// Shadow color mod
    ShadowCrMod = 0x0203,
    /// Shadow opacity
    ShadowOpacity = 0x0204,
    /// Shadow offset X
    ShadowOffsetX = 0x0205,
    /// Shadow offset Y
    ShadowOffsetY = 0x0206,
    /// Shadow second offset X
    ShadowSecondOffsetX = 0x0207,
    /// Shadow second offset Y
    ShadowSecondOffsetY = 0x0208,
    /// Shadow scale X to X
    ShadowScaleXToX = 0x0209,
    /// Shadow scale Y to X
    ShadowScaleYToX = 0x020A,
    /// Shadow scale X to Y
    ShadowScaleXToY = 0x020B,
    /// Shadow scale Y to Y
    ShadowScaleYToY = 0x020C,
    /// Shadow perspective X
    ShadowPerspectiveX = 0x020D,
    /// Shadow perspective Y
    ShadowPerspectiveY = 0x020E,
    /// Shadow weight
    ShadowWeight = 0x020F,
    /// Shadow origin X
    ShadowOriginX = 0x0210,
    /// Shadow origin Y
    ShadowOriginY = 0x0211,
    /// Shadow enabled
    Shadow = 0x023E,
    /// Shadow obscured
    ShadowObscured = 0x023F,

    // Perspective properties (0x0240 - 0x027F)
    /// Perspective type
    PerspectiveType = 0x0240,
    /// Perspective offset X
    PerspectiveOffsetX = 0x0241,
    /// Perspective offset Y
    PerspectiveOffsetY = 0x0242,
    /// Perspective scale X to X
    PerspectiveScaleXToX = 0x0243,
    /// Perspective scale Y to X
    PerspectiveScaleYToX = 0x0244,
    /// Perspective scale X to Y
    PerspectiveScaleXToY = 0x0245,
    /// Perspective scale Y to Y
    PerspectiveScaleYToY = 0x0246,
    /// Perspective X
    PerspectivePerspectiveX = 0x0247,
    /// Perspective Y
    PerspectivePerspectiveY = 0x0248,
    /// Perspective weight
    PerspectiveWeight = 0x0249,
    /// Perspective origin X
    PerspectiveOriginX = 0x024A,
    /// Perspective origin Y
    PerspectiveOriginY = 0x024B,
    /// Perspective on
    PerspectiveOn = 0x027F,

    // 3D properties (0x0280 - 0x02BF)
    /// Specular amount
    ThreeDSpecularAmount = 0x0280,
    /// Diffuse amount
    ThreeDDiffuseAmount = 0x0281,
    /// Shininess
    ThreeDShininess = 0x0282,
    /// Edge thickness
    ThreeDEdgeThickness = 0x0283,
    /// Extrude forward
    ThreeDExtrudeForward = 0x0284,
    /// Extrude backward
    ThreeDExtrudeBackward = 0x0285,
    /// Extrusion color
    ThreeDExtrusionColor = 0x0287,
    /// 3D color mod
    ThreeDCrMod = 0x0288,
    /// Extrusion color ext
    ThreeDExtrusionColorExt = 0x0289,
    /// 3D Effect
    ThreeDEffect = 0x02BC,
    /// Metallic
    ThreeDMetallic = 0x02BD,
    /// Use extrusion color
    ThreeDUseExtrusionColor = 0x02BE,
    /// Light face
    ThreeDLightFace = 0x02BF,

    // 3D Style properties (0x02C0 - 0x02FF)
    /// Y rotation angle
    ThreeDStyleYRotationAngle = 0x02C0,
    /// X rotation angle
    ThreeDStyleXRotationAngle = 0x02C1,
    /// Rotation axis X
    ThreeDStyleRotationAxisX = 0x02C2,
    /// Rotation axis Y
    ThreeDStyleRotationAxisY = 0x02C3,
    /// Rotation axis Z
    ThreeDStyleRotationAxisZ = 0x02C4,
    /// Rotation angle
    ThreeDStyleRotationAngle = 0x02C5,
    /// Rotation center X
    ThreeDStyleRotationCenterX = 0x02C6,
    /// Rotation center Y
    ThreeDStyleRotationCenterY = 0x02C7,
    /// Rotation center Z
    ThreeDStyleRotationCenterZ = 0x02C8,
    /// Render mode
    ThreeDStyleRenderMode = 0x02C9,
    /// Tolerance
    ThreeDStyleTolerance = 0x02CA,
    /// X viewpoint
    ThreeDStyleXViewpoint = 0x02CB,
    /// Y viewpoint
    ThreeDStyleYViewpoint = 0x02CC,
    /// Z viewpoint
    ThreeDStyleZViewpoint = 0x02CD,
    /// Origin X
    ThreeDStyleOriginX = 0x02CE,
    /// Origin Y
    ThreeDStyleOriginY = 0x02CF,
    /// Skew angle
    ThreeDStyleSkewAngle = 0x02D0,
    /// Skew amount
    ThreeDStyleSkewAmount = 0x02D1,
    /// Ambient intensity
    ThreeDStyleAmbientIntensity = 0x02D2,
    /// Key X
    ThreeDStyleKeyX = 0x02D3,
    /// Key Y
    ThreeDStyleKeyY = 0x02D4,
    /// Key Z
    ThreeDStyleKeyZ = 0x02D5,
    /// Key intensity
    ThreeDStyleKeyIntensity = 0x02D6,
    /// Fill X
    ThreeDStyleFillX = 0x02D7,
    /// Fill Y
    ThreeDStyleFillY = 0x02D8,
    /// Fill Z
    ThreeDStyleFillZ = 0x02D9,
    /// Fill intensity
    ThreeDStyleFillIntensity = 0x02DA,

    // Shape properties (0x0300 - 0x033F)
    /// Master shape
    ShapeMaster = 0x0301,
    /// Connector style
    ShapeConnectorStyle = 0x0303,
    /// Black and white settings
    ShapeBlackAndWhiteSettings = 0x0304,
    /// W mode pure BW
    ShapeWModePureBw = 0x0305,
    /// W mode BW
    ShapeWModeBw = 0x0306,
    /// OLE icon
    ShapeOleIcon = 0x033A,
    /// Prefer relative resize
    ShapePreferRelativeResize = 0x033B,
    /// Lock shape type
    ShapeLockShapeType = 0x033C,
    /// Delete attached object
    ShapeDeleteAttachedObject = 0x033E,
    /// Background shape
    ShapeBackgroundShape = 0x033F,

    // Callout properties (0x0340 - 0x037F)
    /// Callout type
    CalloutType = 0x0340,
    /// XY callout gap
    CalloutXYGap = 0x0341,
    /// Callout angle
    CalloutAngle = 0x0342,
    /// Callout drop type
    CalloutDropType = 0x0343,
    /// Callout drop
    CalloutDrop = 0x0344,
    /// Callout length
    CalloutLength = 0x0345,

    // Group properties (0x0380 - 0x03BF)
    /// Group name
    GroupName = 0x0380,
    /// Group description
    GroupDescription = 0x0381,
    /// Hyperlink
    Hyperlink = 0x0382,

    // Diagram properties (0x0500 - 0x057F)
    /// Diagram type
    DiagramType = 0x0500,
    /// Diagram style
    DiagramStyle = 0x0501,

    /// Unknown property
    Unknown = 0xFFFF,
}

impl From<u16> for EscherPropertyId {
    fn from(value: u16) -> Self {
        // Mask off the flags to get the property number
        let prop_num = value & PROPERTY_ID_MASK;

        // Comprehensive property ID mapping based on MS-ODRAW specification
        match prop_num {
            // Transform properties
            0x0004 => Self::Rotation,

            // Protection properties
            0x0077 => Self::LockRotation,
            0x0078 => Self::LockAspectRatio,
            0x0079 => Self::LockPosition,
            0x007A => Self::LockAgainstSelect,
            0x007B => Self::LockCropping,
            0x007C => Self::LockVertices,
            0x007D => Self::LockText,
            0x007E => Self::LockAdjustHandles,
            0x007F => Self::LockAgainstGrouping,

            // Text properties
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

            // GeoText properties
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

            // Blip/Picture properties
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

            // Geometry properties
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

            // Fill properties
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

            // Line properties
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

            // Shadow properties
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

            // Perspective properties
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

            // 3D properties
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

            // 3D Style properties
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

            // Shape properties
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

            // Callout properties
            0x0340 => Self::CalloutType,
            0x0341 => Self::CalloutXYGap,
            0x0342 => Self::CalloutAngle,
            0x0343 => Self::CalloutDropType,
            0x0344 => Self::CalloutDrop,
            0x0345 => Self::CalloutLength,

            // Group properties
            0x0380 => Self::GroupName,
            0x0381 => Self::GroupDescription,
            0x0382 => Self::Hyperlink,

            // Diagram properties
            0x0500 => Self::DiagramType,
            0x0501 => Self::DiagramStyle,

            // Unknown property
            _ => Self::Unknown,
        }
    }
}

/// Escher shape property value.
///
/// Properties can be simple (stored in 4 bytes) or complex (variable-length data).
/// Array properties are a special type of complex property with structured data.
#[derive(Debug, Clone)]
pub enum EscherPropertyValue<'data> {
    /// Simple property: 32-bit integer value
    Simple(i32),

    /// Complex property: binary data (zero-copy borrow)
    ///
    /// This is raw binary data stored in the complex part of the property.
    /// The data is borrowed from the original source for efficiency.
    Complex(&'data [u8]),

    /// Array property: structured array data
    ///
    /// Array properties have a 6-byte header followed by element data:
    /// - 2 bytes: number of elements in array
    /// - 2 bytes: number of elements in memory (reserved)
    /// - 2 bytes: size of each element (can be negative, see get_element_size)
    Array(EscherArrayProperty<'data>),
}

/// Escher array property structure.
///
/// Array properties are complex properties with a specific structure:
/// - 6-byte header (element count, reserved, element size)
/// - Variable-length element data
///
/// # Performance
///
/// - Zero-copy: Borrows data from source
/// - Lazy element access via iterator
/// - No intermediate allocations
#[derive(Debug, Clone)]
pub struct EscherArrayProperty<'data> {
    /// Raw array data including header
    data: &'data [u8],
}

impl<'data> EscherArrayProperty<'data> {
    /// Create array property from raw data.
    ///
    /// # Data Format
    ///
    /// - Bytes 0-1: Number of elements in array (unsigned 16-bit)
    /// - Bytes 2-3: Number of elements in memory / reserved (unsigned 16-bit)
    /// - Bytes 4-5: Size of each element (signed 16-bit, see get_element_size)
    /// - Bytes 6+: Element data
    #[inline]
    pub fn new(data: &'data [u8]) -> Option<Self> {
        if data.len() < 6 {
            return None;
        }
        Some(Self { data })
    }

    /// Get number of elements in array.
    #[inline]
    pub fn element_count(&self) -> u16 {
        if self.data.len() < 2 {
            return 0;
        }
        u16::from_le_bytes([self.data[0], self.data[1]])
    }

    /// Get number of elements in memory (reserved field).
    #[inline]
    pub fn element_count_in_memory(&self) -> u16 {
        if self.data.len() < 4 {
            return 0;
        }
        u16::from_le_bytes([self.data[2], self.data[3]])
    }

    /// Get raw element size value (can be negative).
    #[inline]
    pub fn raw_element_size(&self) -> i16 {
        if self.data.len() < 6 {
            return 0;
        }
        i16::from_le_bytes([self.data[4], self.data[5]])
    }

    /// Get actual element size in bytes.
    ///
    /// # Special Handling
    ///
    /// From MS-ODRAW: If the size is negative, the actual size is:
    /// `(-size) >> 2` (negate and right shift by 2)
    ///
    /// This weird encoding is used for some array properties.
    #[inline]
    pub fn element_size(&self) -> usize {
        let size = self.raw_element_size();
        if size < 0 {
            ((-size) >> 2) as usize
        } else {
            size as usize
        }
    }

    /// Get element at index (zero-copy).
    ///
    /// Returns None if index is out of bounds or element data is truncated.
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

    /// Iterate over all elements (zero-copy).
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

    /// Get raw array data (including header).
    #[inline]
    pub fn raw_data(&self) -> &'data [u8] {
        self.data
    }
}

/// Escher shape properties collection.
///
/// # Performance
///
/// - HashMap for O(1) property lookup
/// - Pre-allocated capacity based on property count
/// - Two-pass parsing for efficient complex data handling
/// - Zero-copy for complex and array properties
#[derive(Debug, Clone)]
pub struct EscherProperties<'data> {
    properties: HashMap<EscherPropertyId, EscherPropertyValue<'data>>,
}

/// Intermediate property descriptor used during two-pass parsing.
#[derive(Debug, Clone, Copy)]
struct PropertyDescriptor {
    id: EscherPropertyId,
    id_raw: u16,
    value: i32,
}

impl<'data> EscherProperties<'data> {
    /// Create empty properties collection.
    #[inline]
    pub fn new() -> Self {
        Self {
            properties: HashMap::new(),
        }
    }

    /// Parse properties from Escher Opt record using two-pass approach.
    ///
    /// # Algorithm (based on Apache POI)
    ///
    /// **Pass 1**: Parse all 6-byte property headers
    /// - Read property ID (2 bytes) with flags
    /// - Read value/length (4 bytes)
    /// - Create property descriptors
    ///
    /// **Pass 2**: Read complex data for complex properties
    /// - Complex data follows all headers sequentially
    /// - Use length from Pass 1 to read correct amount
    /// - Distinguish between complex and array properties
    ///
    /// # Performance
    ///
    /// - Pre-allocated HashMap capacity
    /// - Zero-copy for complex data (borrows from opt.data)
    /// - Efficient bit manipulation for flags
    /// - Single allocation for property descriptors
    pub fn from_opt_record(opt: &EscherRecord<'data>) -> Self {
        let num_properties = opt.instance as usize;
        let mut properties = HashMap::with_capacity(num_properties);

        if opt.data.len() < 6 {
            return Self { properties };
        }

        // Pass 1: Parse all property headers (6 bytes each)
        let header_size = num_properties * 6;
        if header_size > opt.data.len() {
            // Truncated data, parse what we can
            return Self { properties };
        }

        let mut descriptors = Vec::with_capacity(num_properties);
        for i in 0..num_properties {
            let offset = i * 6;
            if offset + 6 > opt.data.len() {
                break;
            }

            // Read property ID with flags (2 bytes)
            let id_raw = u16::from_le_bytes([opt.data[offset], opt.data[offset + 1]]);

            // Read value/length (4 bytes)
            let value = i32::from_le_bytes([
                opt.data[offset + 2],
                opt.data[offset + 3],
                opt.data[offset + 4],
                opt.data[offset + 5],
            ]);

            let id = EscherPropertyId::from(id_raw);

            descriptors.push(PropertyDescriptor { id, id_raw, value });
        }

        // Pass 2: Process properties and read complex data
        let mut complex_data_offset = header_size;

        for desc in descriptors {
            let is_complex = (desc.id_raw & IS_COMPLEX) != 0;
            let _is_blip = (desc.id_raw & IS_BLIP) != 0;

            let prop_value = if is_complex {
                // Complex property: value is the length of complex data
                let complex_len = desc.value as usize;
                let complex_end = complex_data_offset + complex_len;

                if complex_end > opt.data.len() {
                    // Truncated complex data, skip this property
                    // But still advance offset for next properties
                    complex_data_offset = complex_end;
                    continue;
                }

                let complex_data = &opt.data[complex_data_offset..complex_end];
                complex_data_offset = complex_end;

                // Try to detect array properties by checking if they have array structure
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
                // Simple property: value is the data itself
                EscherPropertyValue::Simple(desc.value)
            };

            properties.insert(desc.id, prop_value);
        }

        Self { properties }
    }

    /// Heuristic to detect if complex property is an array property.
    ///
    /// Array properties have a 6-byte header followed by array data.
    /// We check if:
    /// 1. Data is at least 6 bytes
    /// 2. Element count and size are reasonable
    /// 3. Total size matches array structure
    fn is_array_property(data: &[u8]) -> bool {
        if data.len() < 6 {
            return false;
        }

        let num_elements = u16::from_le_bytes([data[0], data[1]]) as usize;
        let element_size_raw = i16::from_le_bytes([data[4], data[5]]);

        // Compute actual element size
        let element_size = if element_size_raw < 0 {
            ((-element_size_raw) >> 2) as usize
        } else {
            element_size_raw as usize
        };

        // Check if the data size matches array structure
        // Some arrays don't include header in size calculation, so check both
        let expected_size_with_header = 6 + num_elements * element_size;
        let expected_size_without_header = num_elements * element_size;

        data.len() == expected_size_with_header || data.len() == expected_size_without_header
    }

    /// Parse properties from a container by finding Opt record.
    pub fn from_container(container: &EscherContainer<'data>) -> Self {
        if let Some(opt) = container.find_child(EscherRecordType::Opt) {
            Self::from_opt_record(&opt)
        } else {
            Self::new()
        }
    }

    /// Get property value by ID.
    #[inline]
    pub fn get(&self, id: EscherPropertyId) -> Option<&EscherPropertyValue<'data>> {
        self.properties.get(&id)
    }

    /// Get simple integer property value.
    ///
    /// # Example
    ///
    /// ```ignore
    /// if let Some(rotation) = props.get_int(EscherPropertyId::Rotation) {
    ///     println!("Rotation: {}", rotation);
    /// }
    /// ```
    #[inline]
    pub fn get_int(&self, id: EscherPropertyId) -> Option<i32> {
        match self.properties.get(&id) {
            Some(EscherPropertyValue::Simple(v)) => Some(*v),
            _ => None,
        }
    }

    /// Get color property value (RGB).
    ///
    /// Colors are stored as 32-bit values in BGR format (low byte = blue).
    #[inline]
    pub fn get_color(&self, id: EscherPropertyId) -> Option<u32> {
        self.get_int(id).map(|v| v as u32)
    }

    /// Get boolean property value.
    ///
    /// Interprets non-zero values as true, zero as false.
    #[inline]
    pub fn get_bool(&self, id: EscherPropertyId) -> Option<bool> {
        self.get_int(id).map(|v| v != 0)
    }

    /// Get complex binary property value (zero-copy).
    ///
    /// Returns borrowed slice from original data.
    #[inline]
    pub fn get_binary(&self, id: EscherPropertyId) -> Option<&'data [u8]> {
        match self.properties.get(&id) {
            Some(EscherPropertyValue::Complex(data)) => Some(data),
            _ => None,
        }
    }

    /// Get array property value.
    ///
    /// # Example
    ///
    /// ```ignore
    /// if let Some(vertices) = props.get_array(EscherPropertyId::Vertices) {
    ///     for vertex in vertices.elements() {
    ///         // Process vertex data
    ///     }
    /// }
    /// ```
    #[inline]
    pub fn get_array(&self, id: EscherPropertyId) -> Option<&EscherArrayProperty<'data>> {
        match self.properties.get(&id) {
            Some(EscherPropertyValue::Array(array)) => Some(array),
            _ => None,
        }
    }

    /// Check if property exists.
    #[inline]
    pub fn has(&self, id: EscherPropertyId) -> bool {
        self.properties.contains_key(&id)
    }

    /// Get number of properties.
    #[inline]
    pub fn len(&self) -> usize {
        self.properties.len()
    }

    /// Check if properties collection is empty.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.properties.is_empty()
    }

    /// Iterate over all properties.
    pub fn iter(&self) -> impl Iterator<Item = (&EscherPropertyId, &EscherPropertyValue<'data>)> {
        self.properties.iter()
    }

    // ===== Specialized Property Extractors =====
    // These methods provide type-safe, convenient access to common property types
    // following Apache POI's approach

    /// Get RGB color components from a color property.
    ///
    /// Colors in Escher are stored in BGR format (little-endian RGB):
    /// - Bits 0-7: Blue
    /// - Bits 8-15: Green
    /// - Bits 16-23: Red
    ///
    /// # Returns
    ///
    /// `(red, green, blue)` tuple with values 0-255
    ///
    /// # Example
    ///
    /// ```ignore
    /// if let Some((r, g, b)) = props.get_rgb(EscherPropertyId::FillColor) {
    ///     println!("Fill color: RGB({}, {}, {})", r, g, b);
    /// }
    /// ```
    #[inline]
    pub fn get_rgb(&self, id: EscherPropertyId) -> Option<(u8, u8, u8)> {
        self.get_color(id).map(|color| {
            let red = ((color >> 16) & 0xFF) as u8;
            let green = ((color >> 8) & 0xFF) as u8;
            let blue = (color & 0xFF) as u8;
            (red, green, blue)
        })
    }

    /// Get rotation angle in degrees from rotation property.
    ///
    /// Rotation values are stored as 16.16 fixed-point numbers.
    /// The upper 16 bits are the integer part, lower 16 bits are the fractional part.
    ///
    /// # Example
    ///
    /// ```ignore
    /// if let Some(angle) = props.get_rotation_degrees(EscherPropertyId::Rotation) {
    ///     println!("Rotation: {} degrees", angle);
    /// }
    /// ```
    #[inline]
    pub fn get_rotation_degrees(&self, id: EscherPropertyId) -> Option<f32> {
        self.get_int(id).map(|fixed_point| {
            // Convert 16.16 fixed-point to float degrees
            // Upper 16 bits = integer part, lower 16 bits = fractional part
            (fixed_point as f32) / 65536.0
        })
    }

    /// Get opacity value as percentage (0.0 - 1.0).
    ///
    /// Opacity is stored as a 16.16 fixed-point value where 65536 = 100%.
    ///
    /// # Example
    ///
    /// ```ignore
    /// if let Some(opacity) = props.get_opacity(EscherPropertyId::FillOpacity) {
    ///     println!("Fill opacity: {}%", opacity * 100.0);
    /// }
    /// ```
    #[inline]
    pub fn get_opacity(&self, id: EscherPropertyId) -> Option<f32> {
        self.get_int(id).map(|fixed_point| {
            // Convert 16.16 fixed-point to 0.0-1.0 range
            let opacity = (fixed_point as f32) / 65536.0;
            opacity.clamp(0.0, 1.0)
        })
    }

    /// Get coordinate value in master units.
    ///
    /// Coordinates in Escher are typically stored in master units (1/576 inch).
    /// This returns the raw coordinate value.
    ///
    /// # Example
    ///
    /// ```ignore
    /// if let Some(x) = props.get_coord(EscherPropertyId::GeomLeft) {
    ///     println!("Geometry left: {} master units", x);
    /// }
    /// ```
    #[inline]
    pub fn get_coord(&self, id: EscherPropertyId) -> Option<i32> {
        self.get_int(id)
    }

    /// Get adjust value for complex shapes.
    ///
    /// Adjust values control the geometry of complex shapes like arrows, stars, etc.
    /// They are typically stored as integers representing relative positions.
    ///
    /// # Example
    ///
    /// ```ignore
    /// if let Some(adj) = props.get_adjust(EscherPropertyId::AdjustValue) {
    ///     println!("Adjust value: {}", adj);
    /// }
    /// ```
    #[inline]
    pub fn get_adjust(&self, id: EscherPropertyId) -> Option<i32> {
        self.get_int(id)
    }

    /// Get string property from complex data.
    ///
    /// Text properties are stored as UTF-16LE complex data.
    ///
    /// # Example
    ///
    /// ```ignore
    /// if let Some(name) = props.get_string(EscherPropertyId::GroupName) {
    ///     println!("Group name: {}", name);
    /// }
    /// ```
    #[inline]
    pub fn get_string(&self, id: EscherPropertyId) -> Option<String> {
        self.get_binary(id).and_then(|data| {
            if data.is_empty() {
                return None;
            }

            // Try UTF-16LE decoding (most common for Office)
            if data.len() % 2 == 0 {
                let utf16_data: Vec<u16> = data
                    .chunks_exact(2)
                    .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                    .collect();

                String::from_utf16(&utf16_data).ok()
            } else {
                // Fall back to ASCII/Latin-1
                Some(String::from_utf8_lossy(data).into_owned())
            }
        })
    }

    /// Get blip (image) reference ID.
    ///
    /// Blip properties reference images stored in the BLIP store.
    ///
    /// # Example
    ///
    /// ```ignore
    /// if let Some(blip_id) = props.get_blip_id(EscherPropertyId::FillBlip) {
    ///     println!("Fill blip ID: {}", blip_id);
    /// }
    /// ```
    #[inline]
    pub fn get_blip_id(&self, id: EscherPropertyId) -> Option<u32> {
        self.get_int(id).map(|v| v as u32)
    }

    /// Check if a boolean property is true.
    ///
    /// Boolean properties are stored as integers where non-zero = true.
    ///
    /// # Example
    ///
    /// ```ignore
    /// if props.is_true(EscherPropertyId::Filled) {
    ///     println!("Shape is filled");
    /// }
    /// ```
    #[inline]
    pub fn is_true(&self, id: EscherPropertyId) -> bool {
        self.get_bool(id).unwrap_or(false)
    }

    /// Get line width in master units.
    ///
    /// Line widths are stored as integers in master units.
    /// The value represents 1/12700 of an inch in most cases.
    ///
    /// # Example
    ///
    /// ```ignore
    /// if let Some(width) = props.get_line_width() {
    ///     println!("Line width: {} EMUs", width);
    /// }
    /// ```
    #[inline]
    pub fn get_line_width(&self) -> Option<i32> {
        self.get_int(EscherPropertyId::LineWidth)
    }

    /// Get fill color as RGB components.
    ///
    /// Convenience method for accessing the fill color.
    ///
    /// # Returns
    ///
    /// `(red, green, blue)` tuple with values 0-255
    #[inline]
    pub fn get_fill_color(&self) -> Option<(u8, u8, u8)> {
        self.get_rgb(EscherPropertyId::FillColor)
    }

    /// Get line color as RGB components.
    ///
    /// Convenience method for accessing the line color.
    ///
    /// # Returns
    ///
    /// `(red, green, blue)` tuple with values 0-255
    #[inline]
    pub fn get_line_color(&self) -> Option<(u8, u8, u8)> {
        self.get_rgb(EscherPropertyId::LineColor)
    }

    /// Check if shape is filled.
    ///
    /// Convenience method for checking the Filled property.
    #[inline]
    pub fn is_filled(&self) -> bool {
        self.is_true(EscherPropertyId::Filled)
    }

    /// Check if shape has line.
    ///
    /// Convenience method based on the AnyLine property.
    #[inline]
    pub fn has_line(&self) -> bool {
        self.is_true(EscherPropertyId::AnyLine)
    }

    /// Check if shape has shadow.
    ///
    /// Convenience method for checking the Shadow property.
    #[inline]
    pub fn has_shadow(&self) -> bool {
        self.is_true(EscherPropertyId::Shadow)
    }

    /// Get geometry rectangle.
    ///
    /// Returns (left, top, right, bottom) geometry coordinates.
    ///
    /// # Example
    ///
    /// ```ignore
    /// if let Some((left, top, right, bottom)) = props.get_geometry_rect() {
    ///     let width = right - left;
    ///     let height = bottom - top;
    ///     println!("Geometry: {}x{} at ({}, {})", width, height, left, top);
    /// }
    /// ```
    #[inline]
    pub fn get_geometry_rect(&self) -> Option<(i32, i32, i32, i32)> {
        let left = self.get_coord(EscherPropertyId::GeomLeft)?;
        let top = self.get_coord(EscherPropertyId::GeomTop)?;
        let right = self.get_coord(EscherPropertyId::GeomRight)?;
        let bottom = self.get_coord(EscherPropertyId::GeomBottom)?;
        Some((left, top, right, bottom))
    }

    /// Get text margins.
    ///
    /// Returns (left, top, right, bottom) text margin values.
    ///
    /// # Example
    ///
    /// ```ignore
    /// if let Some((l, t, r, b)) = props.get_text_margins() {
    ///     println!("Text margins: L={}, T={}, R={}, B={}", l, t, r, b);
    /// }
    /// ```
    #[inline]
    pub fn get_text_margins(&self) -> Option<(i32, i32, i32, i32)> {
        let left = self.get_int(EscherPropertyId::TextLeft).unwrap_or(0);
        let top = self.get_int(EscherPropertyId::TextTop).unwrap_or(0);
        let right = self.get_int(EscherPropertyId::TextRight).unwrap_or(0);
        let bottom = self.get_int(EscherPropertyId::TextBottom).unwrap_or(0);
        Some((left, top, right, bottom))
    }
}

impl<'data> Default for EscherProperties<'data> {
    fn default() -> Self {
        Self::new()
    }
}

/// Shape anchor (position and size).
///
/// # Coordinates
///
/// - Coordinates are in master units (typically 1/576 inch)
/// - Origin is top-left corner
#[derive(Debug, Clone, Copy)]
pub struct ShapeAnchor {
    /// Left coordinate
    pub left: i32,
    /// Top coordinate
    pub top: i32,
    /// Right coordinate
    pub right: i32,
    /// Bottom coordinate
    pub bottom: i32,
}

impl ShapeAnchor {
    /// Create anchor from coordinates.
    #[inline]
    pub const fn new(left: i32, top: i32, right: i32, bottom: i32) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    /// Get width.
    #[inline]
    pub const fn width(&self) -> i32 {
        self.right - self.left
    }

    /// Get height.
    #[inline]
    pub const fn height(&self) -> i32 {
        self.bottom - self.top
    }

    /// Parse from ChildAnchor record.
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

    /// Parse from ClientAnchor record.
    pub fn from_client_anchor(anchor: &EscherRecord) -> Option<Self> {
        // ClientAnchor has same format as ChildAnchor
        Self::from_child_anchor(anchor)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_anchor_dimensions() {
        let anchor = ShapeAnchor::new(100, 200, 500, 600);
        assert_eq!(anchor.width(), 400);
        assert_eq!(anchor.height(), 400);
    }

    #[test]
    fn test_property_id_from_u16() {
        // Test transform properties
        assert_eq!(EscherPropertyId::from(0x0004), EscherPropertyId::Rotation);

        // Test protection properties
        assert_eq!(
            EscherPropertyId::from(0x0077),
            EscherPropertyId::LockRotation
        );
        assert_eq!(
            EscherPropertyId::from(0x0078),
            EscherPropertyId::LockAspectRatio
        );

        // Test text properties
        assert_eq!(EscherPropertyId::from(0x0080), EscherPropertyId::TextId);
        assert_eq!(EscherPropertyId::from(0x0081), EscherPropertyId::TextLeft);

        // Test geometry properties
        assert_eq!(EscherPropertyId::from(0x0140), EscherPropertyId::GeomLeft);
        assert_eq!(EscherPropertyId::from(0x0145), EscherPropertyId::Vertices);

        // Test fill properties
        assert_eq!(EscherPropertyId::from(0x0180), EscherPropertyId::FillType);
        assert_eq!(EscherPropertyId::from(0x0181), EscherPropertyId::FillColor);

        // Test line properties
        assert_eq!(EscherPropertyId::from(0x01C0), EscherPropertyId::LineColor);
        assert_eq!(EscherPropertyId::from(0x01CB), EscherPropertyId::LineWidth);

        // Test shadow properties
        assert_eq!(EscherPropertyId::from(0x0200), EscherPropertyId::ShadowType);
        assert_eq!(
            EscherPropertyId::from(0x0201),
            EscherPropertyId::ShadowColor
        );

        // Test 3D properties
        assert_eq!(
            EscherPropertyId::from(0x0280),
            EscherPropertyId::ThreeDSpecularAmount
        );

        // Test shape properties
        assert_eq!(
            EscherPropertyId::from(0x0301),
            EscherPropertyId::ShapeMaster
        );

        // Test group properties
        assert_eq!(EscherPropertyId::from(0x0380), EscherPropertyId::GroupName);

        // Test unknown property
        assert_eq!(EscherPropertyId::from(0xFFFF), EscherPropertyId::Unknown);
    }

    #[test]
    fn test_property_id_masking() {
        // Test that flags are properly masked off
        let with_complex_flag = 0x8140; // GeomLeft with IS_COMPLEX flag
        assert_eq!(
            EscherPropertyId::from(with_complex_flag),
            EscherPropertyId::GeomLeft
        );

        let with_blip_flag = 0x4186; // FillBlip with IS_BLIP flag  
        assert_eq!(
            EscherPropertyId::from(with_blip_flag),
            EscherPropertyId::FillBlip
        );

        let with_both_flags = 0xC181; // FillColor with both flags
        assert_eq!(
            EscherPropertyId::from(with_both_flags),
            EscherPropertyId::FillColor
        );
    }

    #[test]
    fn test_array_property_creation() {
        let data = [
            0x03, 0x00, // 3 elements
            0x03, 0x00, // 3 elements in memory
            0x08, 0x00, // 8 bytes per element
            // Element data would follow
            0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, // Element 1
            0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, // Element 2
            0x21, 0x22, 0x23, 0x24, 0x25, 0x26, 0x27, 0x28, // Element 3
        ];

        let array = EscherArrayProperty::new(&data).unwrap();
        assert_eq!(array.element_count(), 3);
        assert_eq!(array.element_count_in_memory(), 3);
        assert_eq!(array.element_size(), 8);

        // Test element access
        let elem1 = array.get_element(0).unwrap();
        assert_eq!(elem1, &[0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08]);

        let elem2 = array.get_element(1).unwrap();
        assert_eq!(elem2, &[0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18]);

        // Test iterator
        let elements: Vec<_> = array.elements().collect();
        assert_eq!(elements.len(), 3);
    }

    #[test]
    fn test_array_property_negative_element_size() {
        // Test negative element size (special encoding)
        let data = [
            0x02, 0x00, // 2 elements
            0x02, 0x00, // 2 elements in memory
            0xF0, 0xFF, // -16 as i16 -> actual size is (-(-16)) >> 2 = 4
            // Element data
            0x01, 0x02, 0x03, 0x04, // Element 1 (4 bytes)
            0x11, 0x12, 0x13, 0x14, // Element 2 (4 bytes)
        ];

        let array = EscherArrayProperty::new(&data).unwrap();
        assert_eq!(array.element_count(), 2);
        assert_eq!(array.raw_element_size(), -16);
        assert_eq!(array.element_size(), 4); // (-(-16)) >> 2 = 4

        let elem1 = array.get_element(0).unwrap();
        assert_eq!(elem1, &[0x01, 0x02, 0x03, 0x04]);
    }

    #[test]
    fn test_rgb_extraction() {
        let mut props = EscherProperties::new();

        // Mock a color property value
        // Color is stored in BGR format: 0x00RRGGBB -> 0x00BBGGRR in little-endian
        // For RGB(255, 128, 64) -> stored as 0x0040FF80 -> little-endian bytes
        let color_bgr = 0x0040FF80u32; // B=128, G=255, R=64 in BGR

        let mut test_props = std::collections::HashMap::new();
        test_props.insert(
            EscherPropertyId::FillColor,
            EscherPropertyValue::Simple(color_bgr as i32),
        );
        props.properties = test_props;

        let (r, g, b) = props.get_rgb(EscherPropertyId::FillColor).unwrap();
        assert_eq!(r, 64);
        assert_eq!(g, 255);
        assert_eq!(b, 128);
    }

    #[test]
    fn test_rotation_degrees_extraction() {
        let mut props = EscherProperties::new();

        // 90 degrees in 16.16 fixed-point format: 90 * 65536 = 5898240
        let rotation_fixed = 90 * 65536;

        let mut test_props = std::collections::HashMap::new();
        test_props.insert(
            EscherPropertyId::Rotation,
            EscherPropertyValue::Simple(rotation_fixed),
        );
        props.properties = test_props;

        let degrees = props
            .get_rotation_degrees(EscherPropertyId::Rotation)
            .unwrap();
        assert!((degrees - 90.0).abs() < 0.001);
    }

    #[test]
    fn test_opacity_extraction() {
        let mut props = EscherProperties::new();

        // 50% opacity in 16.16 fixed-point: 0.5 * 65536 = 32768
        let opacity_fixed = 32768;

        let mut test_props = std::collections::HashMap::new();
        test_props.insert(
            EscherPropertyId::FillOpacity,
            EscherPropertyValue::Simple(opacity_fixed),
        );
        props.properties = test_props;

        let opacity = props.get_opacity(EscherPropertyId::FillOpacity).unwrap();
        assert!((opacity - 0.5).abs() < 0.001);
    }

    #[test]
    fn test_boolean_properties() {
        let mut props = EscherProperties::new();

        let mut test_props = std::collections::HashMap::new();
        test_props.insert(EscherPropertyId::Filled, EscherPropertyValue::Simple(1));
        test_props.insert(EscherPropertyId::Shadow, EscherPropertyValue::Simple(0));
        props.properties = test_props;

        assert!(props.is_filled());
        assert!(!props.has_shadow());
        assert!(props.is_true(EscherPropertyId::Filled));
        assert!(!props.is_true(EscherPropertyId::Shadow));
    }

    #[test]
    fn test_geometry_rect_extraction() {
        let mut props = EscherProperties::new();

        let mut test_props = std::collections::HashMap::new();
        test_props.insert(EscherPropertyId::GeomLeft, EscherPropertyValue::Simple(100));
        test_props.insert(EscherPropertyId::GeomTop, EscherPropertyValue::Simple(200));
        test_props.insert(
            EscherPropertyId::GeomRight,
            EscherPropertyValue::Simple(500),
        );
        test_props.insert(
            EscherPropertyId::GeomBottom,
            EscherPropertyValue::Simple(600),
        );
        props.properties = test_props;

        let (left, top, right, bottom) = props.get_geometry_rect().unwrap();
        assert_eq!(left, 100);
        assert_eq!(top, 200);
        assert_eq!(right, 500);
        assert_eq!(bottom, 600);
    }

    #[test]
    fn test_text_margins_extraction() {
        let mut props = EscherProperties::new();

        let mut test_props = std::collections::HashMap::new();
        test_props.insert(EscherPropertyId::TextLeft, EscherPropertyValue::Simple(10));
        test_props.insert(EscherPropertyId::TextTop, EscherPropertyValue::Simple(20));
        test_props.insert(EscherPropertyId::TextRight, EscherPropertyValue::Simple(30));
        test_props.insert(
            EscherPropertyId::TextBottom,
            EscherPropertyValue::Simple(40),
        );
        props.properties = test_props;

        let (left, top, right, bottom) = props.get_text_margins().unwrap();
        assert_eq!(left, 10);
        assert_eq!(top, 20);
        assert_eq!(right, 30);
        assert_eq!(bottom, 40);
    }

    #[test]
    fn test_convenience_getters() {
        let mut props = EscherProperties::new();

        let mut test_props = std::collections::HashMap::new();
        test_props.insert(
            EscherPropertyId::FillColor,
            EscherPropertyValue::Simple(0x00FF8040), // BGR: B=64, G=128, R=255
        );
        test_props.insert(
            EscherPropertyId::LineColor,
            EscherPropertyValue::Simple(0x00408020), // BGR: B=32, G=128, R=64
        );
        test_props.insert(
            EscherPropertyId::LineWidth,
            EscherPropertyValue::Simple(12700),
        );
        test_props.insert(EscherPropertyId::AnyLine, EscherPropertyValue::Simple(1));
        props.properties = test_props;

        // Test fill color
        let (r, g, b) = props.get_fill_color().unwrap();
        assert_eq!((r, g, b), (255, 128, 64));

        // Test line color
        let (r, g, b) = props.get_line_color().unwrap();
        assert_eq!((r, g, b), (64, 128, 32));

        // Test line width
        assert_eq!(props.get_line_width().unwrap(), 12700);

        // Test has_line
        assert!(props.has_line());
    }

    #[test]
    fn test_string_property_utf16() {
        let mut props = EscherProperties::new();

        // Create UTF-16LE string "Test" = [0x54, 0x00, 0x65, 0x00, 0x73, 0x00, 0x74, 0x00]
        let utf16_bytes = vec![0x54, 0x00, 0x65, 0x00, 0x73, 0x00, 0x74, 0x00];

        let mut test_props = std::collections::HashMap::new();
        test_props.insert(
            EscherPropertyId::GroupName,
            EscherPropertyValue::Complex(&utf16_bytes),
        );
        props.properties = test_props;

        let name = props.get_string(EscherPropertyId::GroupName).unwrap();
        assert_eq!(name, "Test");
    }

    #[test]
    fn test_blip_id_extraction() {
        let mut props = EscherProperties::new();

        let mut test_props = std::collections::HashMap::new();
        test_props.insert(EscherPropertyId::FillBlip, EscherPropertyValue::Simple(42));
        props.properties = test_props;

        let blip_id = props.get_blip_id(EscherPropertyId::FillBlip).unwrap();
        assert_eq!(blip_id, 42);
    }
}
