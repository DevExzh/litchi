/// EMF Record Types and Common Structures
///
/// Comprehensive definitions based on [MS-EMF] specification and libUEMF
///
/// References:
/// - [MS-EMF]: Enhanced Metafile Format Specification
/// - libUEMF (version 0.2.8)
/// - LibreOffice emfreader.cxx
use zerocopy::{FromBytes, IntoBytes};

/// EMF Record Type enumeration
///
/// All 120+ EMF record types defined in [MS-EMF] specification
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u32)]
pub enum EmrType {
    // Control Records (1-14)
    Header = 1,
    PolyBezier = 2,
    Polygon = 3,
    Polyline = 4,
    PolyBezierTo = 5,
    PolyLineTo = 6,
    PolyPolyline = 7,
    PolyPolygon = 8,
    SetWindowExtEx = 9,
    SetWindowOrgEx = 10,
    SetViewportExtEx = 11,
    SetViewportOrgEx = 12,
    SetBrushOrgEx = 13,
    Eof = 14,

    // Drawing Records (15-68)
    SetPixelV = 15,
    SetMapperFlags = 16,
    SetMapMode = 17,
    SetBkMode = 18,
    SetPolyFillMode = 19,
    SetRop2 = 20,
    SetStretchBltMode = 21,
    SetTextAlign = 22,
    SetColorAdjustment = 23,
    SetTextColor = 24,
    SetBkColor = 25,
    OffsetClipRgn = 26,
    MoveToEx = 27,
    SetMetaRgn = 28,
    ExcludeClipRect = 29,
    IntersectClipRect = 30,
    ScaleViewportExtEx = 31,
    ScaleWindowExtEx = 32,
    SaveDc = 33,
    RestoreDc = 34,
    SetWorldTransform = 35,
    ModifyWorldTransform = 36,
    SelectObject = 37,
    CreatePen = 38,
    CreateBrushIndirect = 39,
    DeleteObject = 40,
    AngleArc = 41,
    Ellipse = 42,
    Rectangle = 43,
    RoundRect = 44,
    Arc = 45,
    Chord = 46,
    Pie = 47,
    SelectPalette = 48,
    CreatePalette = 49,
    SetPaletteEntries = 50,
    ResizePalette = 51,
    RealizePalette = 52,
    ExtFloodFill = 53,
    LineTo = 54,
    ArcTo = 55,
    PolyDraw = 56,
    SetArcDirection = 57,
    SetMiterLimit = 58,
    BeginPath = 59,
    EndPath = 60,
    CloseFigure = 61,
    FillPath = 62,
    StrokeAndFillPath = 63,
    StrokePath = 64,
    FlattenPath = 65,
    WidenPath = 66,
    SelectClipPath = 67,
    AbortPath = 68,

    // Comment and Region Records (70-75)
    Comment = 70,
    FillRgn = 71,
    FrameRgn = 72,
    InvertRgn = 73,
    PaintRgn = 74,
    ExtSelectClipRgn = 75,

    // Bitmap Records (76-81)
    BitBlt = 76,
    StretchBlt = 77,
    MaskBlt = 78,
    PlgBlt = 79,
    SetDIBitsToDevice = 80,
    StretchDIBits = 81,

    // Font and Text Records (82-84, 96-97, 108, 120)
    ExtCreateFontIndirectW = 82,
    ExtTextOutA = 83,
    ExtTextOutW = 84,
    PolyTextOutA = 96,
    PolyTextOutW = 97,
    SmallTextOut = 108,
    SetTextJustification = 120,

    // Polygon Records (85-92)
    PolyBezier16 = 85,
    Polygon16 = 86,
    Polyline16 = 87,
    PolyBezierTo16 = 88,
    PolyLineTo16 = 89,
    PolyPolyline16 = 90,
    PolyPolygon16 = 91,
    PolyDraw16 = 92,

    // Brush Records (93-95)
    CreateMonoBrush = 93,
    CreateDIBPatternBrushPt = 94,
    ExtCreatePen = 95,

    // Color Management Records (98-104)
    SetIcmMode = 98,
    CreateColorSpace = 99,
    SetColorSpace = 100,
    DeleteColorSpace = 101,
    GlsRecord = 102,
    GlsBoundedRecord = 103,
    PixelFormat = 104,

    // Advanced Records (105-120)
    DrawEscape = 105,
    ExtEscape = 106,
    StartDoc = 107,
    // SmallTextOut = 108 (defined above)
    ForceUfiMapping = 109,
    NamedEscape = 110,
    ColorCorrectPalette = 111,
    SetIcmProfileA = 112,
    SetIcmProfileW = 113,
    AlphaBlend = 114,
    SetLayout = 115,
    TransparentBlt = 116,
    // Reserved = 117,
    GradientFill = 118,
    SetLinkedUfis = 119,
    // SetTextJustification = 120 (defined above)
    ColorMatchToTargetW = 121,
    CreateColorSpaceW = 122,
}

impl EmrType {
    /// Convert from u32 value
    #[inline]
    pub const fn from_u32(value: u32) -> Option<Self> {
        match value {
            1 => Some(Self::Header),
            2 => Some(Self::PolyBezier),
            3 => Some(Self::Polygon),
            4 => Some(Self::Polyline),
            5 => Some(Self::PolyBezierTo),
            6 => Some(Self::PolyLineTo),
            7 => Some(Self::PolyPolyline),
            8 => Some(Self::PolyPolygon),
            9 => Some(Self::SetWindowExtEx),
            10 => Some(Self::SetWindowOrgEx),
            11 => Some(Self::SetViewportExtEx),
            12 => Some(Self::SetViewportOrgEx),
            13 => Some(Self::SetBrushOrgEx),
            14 => Some(Self::Eof),
            15 => Some(Self::SetPixelV),
            16 => Some(Self::SetMapperFlags),
            17 => Some(Self::SetMapMode),
            18 => Some(Self::SetBkMode),
            19 => Some(Self::SetPolyFillMode),
            20 => Some(Self::SetRop2),
            21 => Some(Self::SetStretchBltMode),
            22 => Some(Self::SetTextAlign),
            23 => Some(Self::SetColorAdjustment),
            24 => Some(Self::SetTextColor),
            25 => Some(Self::SetBkColor),
            26 => Some(Self::OffsetClipRgn),
            27 => Some(Self::MoveToEx),
            28 => Some(Self::SetMetaRgn),
            29 => Some(Self::ExcludeClipRect),
            30 => Some(Self::IntersectClipRect),
            31 => Some(Self::ScaleViewportExtEx),
            32 => Some(Self::ScaleWindowExtEx),
            33 => Some(Self::SaveDc),
            34 => Some(Self::RestoreDc),
            35 => Some(Self::SetWorldTransform),
            36 => Some(Self::ModifyWorldTransform),
            37 => Some(Self::SelectObject),
            38 => Some(Self::CreatePen),
            39 => Some(Self::CreateBrushIndirect),
            40 => Some(Self::DeleteObject),
            41 => Some(Self::AngleArc),
            42 => Some(Self::Ellipse),
            43 => Some(Self::Rectangle),
            44 => Some(Self::RoundRect),
            45 => Some(Self::Arc),
            46 => Some(Self::Chord),
            47 => Some(Self::Pie),
            48 => Some(Self::SelectPalette),
            49 => Some(Self::CreatePalette),
            50 => Some(Self::SetPaletteEntries),
            51 => Some(Self::ResizePalette),
            52 => Some(Self::RealizePalette),
            53 => Some(Self::ExtFloodFill),
            54 => Some(Self::LineTo),
            55 => Some(Self::ArcTo),
            56 => Some(Self::PolyDraw),
            57 => Some(Self::SetArcDirection),
            58 => Some(Self::SetMiterLimit),
            59 => Some(Self::BeginPath),
            60 => Some(Self::EndPath),
            61 => Some(Self::CloseFigure),
            62 => Some(Self::FillPath),
            63 => Some(Self::StrokeAndFillPath),
            64 => Some(Self::StrokePath),
            65 => Some(Self::FlattenPath),
            66 => Some(Self::WidenPath),
            67 => Some(Self::SelectClipPath),
            68 => Some(Self::AbortPath),
            70 => Some(Self::Comment),
            71 => Some(Self::FillRgn),
            72 => Some(Self::FrameRgn),
            73 => Some(Self::InvertRgn),
            74 => Some(Self::PaintRgn),
            75 => Some(Self::ExtSelectClipRgn),
            76 => Some(Self::BitBlt),
            77 => Some(Self::StretchBlt),
            78 => Some(Self::MaskBlt),
            79 => Some(Self::PlgBlt),
            80 => Some(Self::SetDIBitsToDevice),
            81 => Some(Self::StretchDIBits),
            82 => Some(Self::ExtCreateFontIndirectW),
            83 => Some(Self::ExtTextOutA),
            84 => Some(Self::ExtTextOutW),
            85 => Some(Self::PolyBezier16),
            86 => Some(Self::Polygon16),
            87 => Some(Self::Polyline16),
            88 => Some(Self::PolyBezierTo16),
            89 => Some(Self::PolyLineTo16),
            90 => Some(Self::PolyPolyline16),
            91 => Some(Self::PolyPolygon16),
            92 => Some(Self::PolyDraw16),
            93 => Some(Self::CreateMonoBrush),
            94 => Some(Self::CreateDIBPatternBrushPt),
            95 => Some(Self::ExtCreatePen),
            96 => Some(Self::PolyTextOutA),
            97 => Some(Self::PolyTextOutW),
            98 => Some(Self::SetIcmMode),
            99 => Some(Self::CreateColorSpace),
            100 => Some(Self::SetColorSpace),
            101 => Some(Self::DeleteColorSpace),
            102 => Some(Self::GlsRecord),
            103 => Some(Self::GlsBoundedRecord),
            104 => Some(Self::PixelFormat),
            105 => Some(Self::DrawEscape),
            106 => Some(Self::ExtEscape),
            107 => Some(Self::StartDoc),
            108 => Some(Self::SmallTextOut),
            109 => Some(Self::ForceUfiMapping),
            110 => Some(Self::NamedEscape),
            111 => Some(Self::ColorCorrectPalette),
            112 => Some(Self::SetIcmProfileA),
            113 => Some(Self::SetIcmProfileW),
            114 => Some(Self::AlphaBlend),
            115 => Some(Self::SetLayout),
            116 => Some(Self::TransparentBlt),
            118 => Some(Self::GradientFill),
            119 => Some(Self::SetLinkedUfis),
            120 => Some(Self::SetTextJustification),
            121 => Some(Self::ColorMatchToTargetW),
            122 => Some(Self::CreateColorSpaceW),
            _ => None,
        }
    }

    /// Get record type name for debugging
    pub const fn name(self) -> &'static str {
        match self {
            Self::Header => "EMR_HEADER",
            Self::PolyBezier => "EMR_POLYBEZIER",
            Self::Polygon => "EMR_POLYGON",
            Self::Polyline => "EMR_POLYLINE",
            Self::PolyBezierTo => "EMR_POLYBEZIERTO",
            Self::PolyLineTo => "EMR_POLYLINETO",
            Self::PolyPolyline => "EMR_POLYPOLYLINE",
            Self::PolyPolygon => "EMR_POLYPOLYGON",
            Self::SetWindowExtEx => "EMR_SETWINDOWEXTEX",
            Self::SetWindowOrgEx => "EMR_SETWINDOWORGEX",
            Self::SetViewportExtEx => "EMR_SETVIEWPORTEXTEX",
            Self::SetViewportOrgEx => "EMR_SETVIEWPORTORGEX",
            Self::SetBrushOrgEx => "EMR_SETBRUSHORGEX",
            Self::Eof => "EMR_EOF",
            Self::SetPixelV => "EMR_SETPIXELV",
            Self::SetMapperFlags => "EMR_SETMAPPERFLAGS",
            Self::SetMapMode => "EMR_SETMAPMODE",
            Self::SetBkMode => "EMR_SETBKMODE",
            Self::SetPolyFillMode => "EMR_SETPOLYFILLMODE",
            Self::SetRop2 => "EMR_SETROP2",
            Self::SetStretchBltMode => "EMR_SETSTRETCHBLTMODE",
            Self::SetTextAlign => "EMR_SETTEXTALIGN",
            Self::SetColorAdjustment => "EMR_SETCOLORADJUSTMENT",
            Self::SetTextColor => "EMR_SETTEXTCOLOR",
            Self::SetBkColor => "EMR_SETBKCOLOR",
            Self::OffsetClipRgn => "EMR_OFFSETCLIPRGN",
            Self::MoveToEx => "EMR_MOVETOEX",
            Self::SetMetaRgn => "EMR_SETMETARGN",
            Self::ExcludeClipRect => "EMR_EXCLUDECLIPRECT",
            Self::IntersectClipRect => "EMR_INTERSECTCLIPRECT",
            Self::ScaleViewportExtEx => "EMR_SCALEVIEWPORTEXTEX",
            Self::ScaleWindowExtEx => "EMR_SCALEWINDOWEXTEX",
            Self::SaveDc => "EMR_SAVEDC",
            Self::RestoreDc => "EMR_RESTOREDC",
            Self::SetWorldTransform => "EMR_SETWORLDTRANSFORM",
            Self::ModifyWorldTransform => "EMR_MODIFYWORLDTRANSFORM",
            Self::SelectObject => "EMR_SELECTOBJECT",
            Self::CreatePen => "EMR_CREATEPEN",
            Self::CreateBrushIndirect => "EMR_CREATEBRUSHINDIRECT",
            Self::DeleteObject => "EMR_DELETEOBJECT",
            Self::AngleArc => "EMR_ANGLEARC",
            Self::Ellipse => "EMR_ELLIPSE",
            Self::Rectangle => "EMR_RECTANGLE",
            Self::RoundRect => "EMR_ROUNDRECT",
            Self::Arc => "EMR_ARC",
            Self::Chord => "EMR_CHORD",
            Self::Pie => "EMR_PIE",
            Self::SelectPalette => "EMR_SELECTPALETTE",
            Self::CreatePalette => "EMR_CREATEPALETTE",
            Self::SetPaletteEntries => "EMR_SETPALETTEENTRIES",
            Self::ResizePalette => "EMR_RESIZEPALETTE",
            Self::RealizePalette => "EMR_REALIZEPALETTE",
            Self::ExtFloodFill => "EMR_EXTFLOODFILL",
            Self::LineTo => "EMR_LINETO",
            Self::ArcTo => "EMR_ARCTO",
            Self::PolyDraw => "EMR_POLYDRAW",
            Self::SetArcDirection => "EMR_SETARCDIRECTION",
            Self::SetMiterLimit => "EMR_SETMITERLIMIT",
            Self::BeginPath => "EMR_BEGINPATH",
            Self::EndPath => "EMR_ENDPATH",
            Self::CloseFigure => "EMR_CLOSEFIGURE",
            Self::FillPath => "EMR_FILLPATH",
            Self::StrokeAndFillPath => "EMR_STROKEANDFILLPATH",
            Self::StrokePath => "EMR_STROKEPATH",
            Self::FlattenPath => "EMR_FLATTENPATH",
            Self::WidenPath => "EMR_WIDENPATH",
            Self::SelectClipPath => "EMR_SELECTCLIPPATH",
            Self::AbortPath => "EMR_ABORTPATH",
            Self::Comment => "EMR_COMMENT",
            Self::FillRgn => "EMR_FILLRGN",
            Self::FrameRgn => "EMR_FRAMERGN",
            Self::InvertRgn => "EMR_INVERTRGN",
            Self::PaintRgn => "EMR_PAINTRGN",
            Self::ExtSelectClipRgn => "EMR_EXTSELECTCLIPRGN",
            Self::BitBlt => "EMR_BITBLT",
            Self::StretchBlt => "EMR_STRETCHBLT",
            Self::MaskBlt => "EMR_MASKBLT",
            Self::PlgBlt => "EMR_PLGBLT",
            Self::SetDIBitsToDevice => "EMR_SETDIBITSTODEVICE",
            Self::StretchDIBits => "EMR_STRETCHDIBITS",
            Self::ExtCreateFontIndirectW => "EMR_EXTCREATEFONTINDIRECTW",
            Self::ExtTextOutA => "EMR_EXTTEXTOUTA",
            Self::ExtTextOutW => "EMR_EXTTEXTOUTW",
            Self::PolyBezier16 => "EMR_POLYBEZIER16",
            Self::Polygon16 => "EMR_POLYGON16",
            Self::Polyline16 => "EMR_POLYLINE16",
            Self::PolyBezierTo16 => "EMR_POLYBEZIERTO16",
            Self::PolyLineTo16 => "EMR_POLYLINETO16",
            Self::PolyPolyline16 => "EMR_POLYPOLYLINE16",
            Self::PolyPolygon16 => "EMR_POLYPOLYGON16",
            Self::PolyDraw16 => "EMR_POLYDRAW16",
            Self::CreateMonoBrush => "EMR_CREATEMONOBRUSH",
            Self::CreateDIBPatternBrushPt => "EMR_CREATEDIBPATTERNBRUSHPT",
            Self::ExtCreatePen => "EMR_EXTCREATEPEN",
            Self::PolyTextOutA => "EMR_POLYTEXTOUTA",
            Self::PolyTextOutW => "EMR_POLYTEXTOUTW",
            Self::SetIcmMode => "EMR_SETICMMODE",
            Self::CreateColorSpace => "EMR_CREATECOLORSPACE",
            Self::SetColorSpace => "EMR_SETCOLORSPACE",
            Self::DeleteColorSpace => "EMR_DELETECOLORSPACE",
            Self::GlsRecord => "EMR_GLSRECORD",
            Self::GlsBoundedRecord => "EMR_GLSBOUNDEDRECORD",
            Self::PixelFormat => "EMR_PIXELFORMAT",
            Self::DrawEscape => "EMR_DRAWESCAPE",
            Self::ExtEscape => "EMR_EXTESCAPE",
            Self::StartDoc => "EMR_STARTDOC",
            Self::SmallTextOut => "EMR_SMALLTEXTOUT",
            Self::ForceUfiMapping => "EMR_FORCEUFIMAPPING",
            Self::NamedEscape => "EMR_NAMEDESCAPE",
            Self::ColorCorrectPalette => "EMR_COLORCORRECTPALETTE",
            Self::SetIcmProfileA => "EMR_SETICMPROFILEA",
            Self::SetIcmProfileW => "EMR_SETICMPROFILEW",
            Self::AlphaBlend => "EMR_ALPHABLEND",
            Self::SetLayout => "EMR_SETLAYOUT",
            Self::TransparentBlt => "EMR_TRANSPARENTBLT",
            Self::GradientFill => "EMR_GRADIENTFILL",
            Self::SetLinkedUfis => "EMR_SETLINKEDUFIS",
            Self::SetTextJustification => "EMR_SETTEXTJUSTIFICATION",
            Self::ColorMatchToTargetW => "EMR_COLORMATCHTOTARGETW",
            Self::CreateColorSpaceW => "EMR_CREATECOLORSPACEW",
        }
    }
}

// Common EMF structures using zerocopy for zero-copy parsing

/// EMF Point (POINTL) - 32-bit signed coordinates
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct PointL {
    pub x: i32,
    pub y: i32,
}

/// EMF Point (POINTS) - 16-bit signed coordinates
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct PointS {
    pub x: i16,
    pub y: i16,
}

/// EMF Rectangle (RECTL)
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct RectL {
    pub left: i32,
    pub top: i32,
    pub right: i32,
    pub bottom: i32,
}

/// EMF Size (SIZEL)
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct SizeL {
    pub cx: i32,
    pub cy: i32,
}

/// EMF Color (COLORREF) - 0x00bbggrr format
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct ColorRef {
    pub value: u32,
}

impl ColorRef {
    /// Create from RGB components
    #[inline]
    pub const fn from_rgb(r: u8, g: u8, b: u8) -> Self {
        Self {
            value: (r as u32) | ((g as u32) << 8) | ((b as u32) << 16),
        }
    }

    /// Extract red component
    #[inline]
    pub const fn r(self) -> u8 {
        (self.value & 0xFF) as u8
    }

    /// Extract green component
    #[inline]
    pub const fn g(self) -> u8 {
        ((self.value >> 8) & 0xFF) as u8
    }

    /// Extract blue component
    #[inline]
    pub const fn b(self) -> u8 {
        ((self.value >> 16) & 0xFF) as u8
    }

    /// Convert to SVG color string (compact format)
    pub fn to_svg_color(self) -> String {
        // Use shorthand hex notation when possible
        let r = self.r();
        let g = self.g();
        let b = self.b();

        if r == g && g == b {
            // Grayscale - check for common values
            match r {
                0 => return "black".to_string(),
                255 => return "white".to_string(),
                _ => {},
            }
        }

        // Check for standard named colors (minimal output)
        match (r, g, b) {
            (255, 0, 0) => "red".to_string(),
            (0, 255, 0) => "lime".to_string(),
            (0, 0, 255) => "blue".to_string(),
            _ => {
                // Use 3-digit hex when possible
                if r & 0x0F == (r >> 4) && g & 0x0F == (g >> 4) && b & 0x0F == (b >> 4) {
                    format!("#{:x}{:x}{:x}", r >> 4, g >> 4, b >> 4)
                } else {
                    format!("#{:02x}{:02x}{:02x}", r, g, b)
                }
            },
        }
    }
}

/// EMF World Transform (XFORM)
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct XForm {
    pub m11: f32, // Horizontal scaling
    pub m12: f32, // Horizontal shearing
    pub m21: f32, // Vertical shearing
    pub m22: f32, // Vertical scaling
    pub dx: f32,  // Horizontal translation
    pub dy: f32,  // Vertical translation
}

impl Default for XForm {
    fn default() -> Self {
        Self {
            m11: 1.0,
            m12: 0.0,
            m21: 0.0,
            m22: 1.0,
            dx: 0.0,
            dy: 0.0,
        }
    }
}

impl XForm {
    /// Check if this is the identity transform
    #[inline]
    pub fn is_identity(&self) -> bool {
        self.m11 == 1.0
            && self.m12 == 0.0
            && self.m21 == 0.0
            && self.m22 == 1.0
            && self.dx == 0.0
            && self.dy == 0.0
    }

    /// Convert to SVG transform string (minimal format like SVGO)
    pub fn to_svg_transform(&self) -> Option<String> {
        if self.is_identity() {
            return None;
        }

        // Check for simple transforms
        const EPSILON: f32 = 0.0001;

        // Pure translation
        if (self.m11 - 1.0).abs() < EPSILON
            && self.m12.abs() < EPSILON
            && self.m21.abs() < EPSILON
            && (self.m22 - 1.0).abs() < EPSILON
        {
            if self.dy.abs() < EPSILON {
                return Some(format!("translate({})", self.dx));
            }
            return Some(format!("translate({} {})", self.dx, self.dy));
        }

        // Pure scaling
        if self.dx.abs() < EPSILON
            && self.dy.abs() < EPSILON
            && self.m12.abs() < EPSILON
            && self.m21.abs() < EPSILON
        {
            if (self.m11 - self.m22).abs() < EPSILON {
                return Some(format!("scale({})", self.m11));
            }
            return Some(format!("scale({} {})", self.m11, self.m22));
        }

        // General matrix
        Some(format!(
            "matrix({} {} {} {} {} {})",
            self.m11, self.m12, self.m21, self.m22, self.dx, self.dy
        ))
    }

    /// Transform a point
    #[inline]
    pub fn transform_point(&self, x: f64, y: f64) -> (f64, f64) {
        let x = x as f32;
        let y = y as f32;
        (
            (self.m11 * x + self.m21 * y + self.dx) as f64,
            (self.m12 * x + self.m22 * y + self.dy) as f64,
        )
    }

    /// Multiply two transforms
    pub fn multiply(&self, other: &XForm) -> XForm {
        XForm {
            m11: self.m11 * other.m11 + self.m12 * other.m21,
            m12: self.m11 * other.m12 + self.m12 * other.m22,
            m21: self.m21 * other.m11 + self.m22 * other.m21,
            m22: self.m21 * other.m12 + self.m22 * other.m22,
            dx: self.m11 * other.dx + self.m21 * other.dy + self.dx,
            dy: self.m12 * other.dx + self.m22 * other.dy + self.dy,
        }
    }
}

/// EMF Record Header (common to all records)
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrHeader {
    pub record_type: u32,
    pub record_size: u32,
}

/// Stock object indices (used with SelectObject)
pub mod stock_objects {
    pub const WHITE_BRUSH: u32 = 0x80000000;
    pub const LTGRAY_BRUSH: u32 = 0x80000001;
    pub const GRAY_BRUSH: u32 = 0x80000002;
    pub const DKGRAY_BRUSH: u32 = 0x80000003;
    pub const BLACK_BRUSH: u32 = 0x80000004;
    pub const NULL_BRUSH: u32 = 0x80000005;
    pub const WHITE_PEN: u32 = 0x80000006;
    pub const BLACK_PEN: u32 = 0x80000007;
    pub const NULL_PEN: u32 = 0x80000008;
    pub const OEM_FIXED_FONT: u32 = 0x8000000A;
    pub const ANSI_FIXED_FONT: u32 = 0x8000000B;
    pub const ANSI_VAR_FONT: u32 = 0x8000000C;
    pub const SYSTEM_FONT: u32 = 0x8000000D;
    pub const DEVICE_DEFAULT_FONT: u32 = 0x8000000E;
    pub const DEFAULT_PALETTE: u32 = 0x8000000F;
    pub const SYSTEM_FIXED_FONT: u32 = 0x80000010;
    pub const DEFAULT_GUI_FONT: u32 = 0x80000011;
    pub const DC_BRUSH: u32 = 0x80000012;
    pub const DC_PEN: u32 = 0x80000013;

    /// Check if value is a stock object
    #[inline]
    pub const fn is_stock_object(value: u32) -> bool {
        (value & 0x80000000) != 0
    }
}
