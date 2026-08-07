//! Decoding of the payloads carried by `EmfPlusObject` records.
//!
//! The wire format has a surprising number of length-prefixed, recursive and
//! optional structures.  This module intentionally owns decoded values: an
//! object table can consequently outlive the source comment buffer without
//! retaining an entire metafile in memory.

#![allow(
    clippy::missing_errors_doc,
    clippy::module_name_repetitions,
    clippy::struct_excessive_bools,
    clippy::too_many_lines,
    reason = "The EMF+ names are specification terms and the wire decoder is necessarily centralised."
)]

use std::convert::TryFrom;

use litchi_core::error::{Error, Result};

use super::types::ObjectType;

const BRUSH_PATH: u32 = 0x01;
const BRUSH_TRANSFORM: u32 = 0x02;
const BRUSH_PRESET_COLORS: u32 = 0x04;
const BRUSH_BLEND_H: u32 = 0x08;
const BRUSH_BLEND_V: u32 = 0x10;
const BRUSH_FOCUS_SCALES: u32 = 0x40;
const PEN_TRANSFORM: u32 = 0x01;
const PEN_START_CAP: u32 = 0x02;
const PEN_END_CAP: u32 = 0x04;
const PEN_JOIN: u32 = 0x08;
const PEN_MITER: u32 = 0x10;
const PEN_STYLE: u32 = 0x20;
const PEN_DASH_CAP: u32 = 0x40;
const PEN_DASH_OFFSET: u32 = 0x80;
const PEN_DASH: u32 = 0x100;
const PEN_NON_CENTER: u32 = 0x200;
const PEN_COMPOUND: u32 = 0x400;
const PEN_CUSTOM_START: u32 = 0x800;
const PEN_CUSTOM_END: u32 = 0x1000;

/// Resource limits for decoding one assembled EMF+ object.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct DecodeLimits {
    pub max_bytes: usize,
    pub max_points: usize,
    pub max_recursion: usize,
}

impl DecodeLimits {
    pub fn validate(self) -> Result<Self> {
        if self.max_bytes == 0 || self.max_points == 0 || self.max_recursion == 0 {
            return Err(error("EMF+ decode limits must be greater than zero"));
        }
        Ok(self)
    }
}

impl Default for DecodeLimits {
    fn default() -> Self {
        Self {
            max_bytes: 64 * 1024 * 1024,
            max_points: 1_000_000,
            max_recursion: 64,
        }
    }
}

/// An EMF+ ARGB value, stored in the byte order used by the specification.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
#[repr(transparent)]
pub struct Argb(pub u32);

impl Argb {
    #[must_use]
    pub const fn alpha(self) -> u8 {
        self.0.to_be_bytes()[0]
    }
    #[must_use]
    pub const fn red(self) -> u8 {
        self.0.to_be_bytes()[1]
    }
    #[must_use]
    pub const fn green(self) -> u8 {
        self.0.to_be_bytes()[2]
    }
    #[must_use]
    pub const fn blue(self) -> u8 {
        self.0.to_be_bytes()[3]
    }
    #[must_use]
    pub const fn rgba(self) -> [u8; 4] {
        [self.red(), self.green(), self.blue(), self.alpha()]
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Point {
    pub x: f32,
    pub y: f32,
}
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Rect {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
}
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Transform {
    pub m11: f32,
    pub m12: f32,
    pub m21: f32,
    pub m22: f32,
    pub dx: f32,
    pub dy: f32,
}

/// Information retained when a legal extension or deliberately opaque format is encountered.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DecodeDiagnostic {
    pub message: String,
    pub bytes: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq)]
pub enum GraphicsObject {
    Brush(Brush),
    CustomLineCap(CustomLineCap),
    Font(Font),
    Image(Image),
    ImageAttributes(ImageAttributes),
    Path(Path),
    Pen(Pen),
    Region(Region),
    StringFormat(StringFormat),
}

#[derive(Clone, Debug, PartialEq)]
pub struct Brush {
    pub version: u32,
    pub kind: BrushKind,
    pub diagnostics: Vec<DecodeDiagnostic>,
}
#[derive(Clone, Debug, PartialEq)]
pub enum BrushKind {
    Solid {
        color: Argb,
    },
    Hatch {
        style: u32,
        foreground: Argb,
        background: Argb,
    },
    Texture(TextureBrush),
    LinearGradient(LinearGradientBrush),
    PathGradient(PathGradientBrush),
    Unsupported {
        type_code: u32,
        data: Vec<u8>,
    },
}
#[derive(Clone, Debug, PartialEq)]
pub struct TextureBrush {
    pub flags: u32,
    pub wrap_mode: i32,
    pub transform: Option<Transform>,
    pub image: Option<Box<Image>>,
}
#[derive(Clone, Debug, PartialEq)]
pub struct LinearGradientBrush {
    pub flags: u32,
    pub wrap_mode: i32,
    pub rect: Rect,
    pub start: Argb,
    pub end: Argb,
    pub transform: Option<Transform>,
    pub blends: Vec<Blend>,
}
#[derive(Clone, Debug, PartialEq)]
pub struct PathGradientBrush {
    pub flags: u32,
    pub wrap_mode: i32,
    pub center_color: Argb,
    pub center: Point,
    pub surrounding: Vec<Argb>,
    pub boundary: GradientBoundary,
    pub transform: Option<Transform>,
    pub blends: Vec<Blend>,
    pub focus_scales: Option<(f32, f32)>,
}
#[derive(Clone, Debug, PartialEq)]
pub enum GradientBoundary {
    Path(Box<Path>),
    Points(Vec<Point>),
}
#[derive(Clone, Debug, PartialEq)]
pub enum Blend {
    Colors(Vec<(f32, Argb)>),
    Factors(Vec<(f32, f32)>),
}

#[derive(Clone, Debug, PartialEq)]
pub struct CustomLineCap {
    pub version: u32,
    pub kind: CustomLineCapKind,
    pub diagnostics: Vec<DecodeDiagnostic>,
}
#[derive(Clone, Debug, PartialEq)]
pub enum CustomLineCapKind {
    Default {
        flags: u32,
        base_cap: u32,
        base_inset: f32,
        stroke_start_cap: u32,
        stroke_end_cap: u32,
        stroke_join: u32,
        miter_limit: f32,
        width_scale: f32,
        fill_path: Option<Path>,
        outline_path: Option<Path>,
    },
    AdjustableArrow {
        width: f32,
        height: f32,
        middle_inset: f32,
        filled: bool,
        start_cap: u32,
        end_cap: u32,
        line_join: u32,
        miter_limit: f32,
        width_scale: f32,
    },
    Unsupported {
        type_code: i32,
        data: Vec<u8>,
    },
}

#[derive(Clone, Debug, PartialEq)]
pub struct Font {
    pub version: u32,
    pub em_size: f32,
    pub unit: u32,
    pub style: i32,
    pub family: String,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Image {
    pub version: u32,
    pub kind: ImageKind,
    pub diagnostics: Vec<DecodeDiagnostic>,
}
#[derive(Clone, Debug, PartialEq)]
pub enum ImageKind {
    Bitmap(Bitmap),
    Compressed(Vec<u8>),
    Metafile { type_code: u32, data: Vec<u8> },
    Unsupported { type_code: u32, data: Vec<u8> },
}
#[derive(Clone, Debug, PartialEq)]
pub struct Bitmap {
    pub width: i32,
    pub height: i32,
    pub stride: i32,
    pub pixel_format: u32,
    pub pixels: Vec<u8>,
    pub compressed: bool,
}

#[derive(Clone, Debug, PartialEq)]
pub struct ImageAttributes {
    pub version: u32,
    pub wrap_mode: u32,
    pub clamp_color: Argb,
    pub clamp: i32,
    pub diagnostics: Vec<DecodeDiagnostic>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Path {
    pub version: u32,
    pub flags: u32,
    pub points: Vec<Point>,
    pub point_types: Vec<u8>,
}
impl Path {
    /// Renderer-ready segments inferred from EMF+ point type values.
    #[must_use]
    pub fn segments(&self) -> Vec<PathSegment> {
        self.points
            .iter()
            .copied()
            .zip(self.point_types.iter().copied())
            .map(|(point, point_type)| PathSegment {
                point,
                kind: point_type & 0x07,
                flags: point_type & !0x07,
            })
            .collect()
    }
}
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PathSegment {
    pub point: Point,
    pub kind: u8,
    pub flags: u8,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Pen {
    pub version: u32,
    pub flags: u32,
    pub unit: u32,
    pub width: f32,
    pub transform: Option<Transform>,
    pub start_cap: Option<i32>,
    pub end_cap: Option<i32>,
    pub join: Option<i32>,
    pub miter_limit: Option<f32>,
    pub line_style: Option<i32>,
    pub dashed_cap: Option<i32>,
    pub dash_offset: Option<f32>,
    pub dashes: Vec<f32>,
    pub alignment: Option<i32>,
    pub compound: Vec<f32>,
    pub brush: Box<Brush>,
    pub diagnostics: Vec<DecodeDiagnostic>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Region {
    pub version: u32,
    pub root: RegionNode,
    pub diagnostics: Vec<DecodeDiagnostic>,
}
#[derive(Clone, Debug, PartialEq)]
pub enum RegionNode {
    And(Box<RegionNode>, Box<RegionNode>),
    Or(Box<RegionNode>, Box<RegionNode>),
    Xor(Box<RegionNode>, Box<RegionNode>),
    Exclude(Box<RegionNode>, Box<RegionNode>),
    Complement(Box<RegionNode>, Box<RegionNode>),
    Rect(Rect),
    Path(Path),
    Empty,
    Infinite,
}

#[derive(Clone, Debug, PartialEq)]
pub struct StringFormat {
    pub version: u32,
    pub flags: u32,
    pub language: u16,
    pub alignment: u32,
    pub line_alignment: u32,
    pub digit_substitution: u32,
    pub digit_language: u16,
    pub first_tab_offset: f32,
    pub hotkey_prefix: i32,
    pub leading_margin: f32,
    pub trailing_margin: f32,
    pub tracking: f32,
    pub trimming: u32,
    pub tab_stops: Vec<f32>,
    pub ranges: Vec<CharacterRange>,
}
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct CharacterRange {
    pub first: i32,
    pub length: i32,
}

/// Decode a fully reassembled `EmfPlusObject` payload.
pub fn decode_object(
    object_type: ObjectType,
    data: &[u8],
    limits: DecodeLimits,
) -> Result<GraphicsObject> {
    limits.validate()?;
    if data.len() > limits.max_bytes {
        return Err(error("EMF+ object exceeds configured byte limit"));
    }
    let mut cursor = Cursor::new(data, limits);
    let object = match object_type {
        ObjectType::Brush => GraphicsObject::Brush(decode_brush(&mut cursor)?),
        ObjectType::CustomLineCap => GraphicsObject::CustomLineCap(decode_cap(&mut cursor)?),
        ObjectType::Font => GraphicsObject::Font(decode_font(&mut cursor)?),
        ObjectType::Image => GraphicsObject::Image(decode_image(&mut cursor)?),
        ObjectType::ImageAttributes => {
            GraphicsObject::ImageAttributes(decode_attributes(&mut cursor)?)
        },
        ObjectType::Path => GraphicsObject::Path(decode_path(&mut cursor)?),
        ObjectType::Pen => GraphicsObject::Pen(decode_pen(&mut cursor)?),
        ObjectType::Region => GraphicsObject::Region(decode_region(&mut cursor)?),
        ObjectType::StringFormat => {
            GraphicsObject::StringFormat(decode_string_format(&mut cursor)?)
        },
        ObjectType::Invalid => return Err(error("ObjectTypeInvalid is not a graphics object")),
    };
    cursor.padding()?;
    Ok(object)
}

fn decode_brush(c: &mut Cursor<'_>) -> Result<Brush> {
    let version = c.u32()?;
    let type_code = c.u32()?;
    let kind = match type_code {
        0 => BrushKind::Solid {
            color: Argb(c.u32()?),
        },
        1 => BrushKind::Hatch {
            style: c.u32()?,
            foreground: Argb(c.u32()?),
            background: Argb(c.u32()?),
        },
        2 => BrushKind::Texture(decode_texture(c)?),
        3 => BrushKind::PathGradient(decode_path_gradient(c)?),
        4 => BrushKind::LinearGradient(decode_linear_gradient(c)?),
        _ => BrushKind::Unsupported {
            type_code,
            data: c.rest(),
        },
    };
    let diagnostics = match &kind {
        BrushKind::Unsupported { data, .. } => {
            vec![diagnostic("unsupported EMF+ brush type", data.clone())]
        },
        BrushKind::Solid { .. }
        | BrushKind::Hatch { .. }
        | BrushKind::Texture(_)
        | BrushKind::LinearGradient(_)
        | BrushKind::PathGradient(_) => Vec::new(),
    };
    Ok(Brush {
        version,
        kind,
        diagnostics,
    })
}

fn decode_texture(c: &mut Cursor<'_>) -> Result<TextureBrush> {
    let flags = c.u32()?;
    let wrap_mode = c.i32()?;
    let transform = if flags & BRUSH_TRANSFORM != 0 {
        Some(c.transform()?)
    } else {
        None
    };
    let image = if c.remaining() == 0 {
        None
    } else {
        Some(Box::new(decode_image(c)?))
    };
    Ok(TextureBrush {
        flags,
        wrap_mode,
        transform,
        image,
    })
}

fn decode_linear_gradient(c: &mut Cursor<'_>) -> Result<LinearGradientBrush> {
    let flags = c.u32()?;
    let wrap_mode = c.i32()?;
    let rect = c.rect()?;
    let start = Argb(c.u32()?);
    let end = Argb(c.u32()?);
    let _reserved = c.u64()?;
    let transform = if flags & BRUSH_TRANSFORM != 0 {
        Some(c.transform()?)
    } else {
        None
    };
    let blends = decode_blends(c, flags, true)?;
    Ok(LinearGradientBrush {
        flags,
        wrap_mode,
        rect,
        start,
        end,
        transform,
        blends,
    })
}

fn decode_path_gradient(c: &mut Cursor<'_>) -> Result<PathGradientBrush> {
    let flags = c.u32()?;
    let wrap_mode = c.i32()?;
    let center_color = Argb(c.u32()?);
    let center = c.point()?;
    let raw_count = c.u32()?;
    let count = c.count(raw_count, "surrounding colors")?;
    let mut surrounding = Vec::new();
    c.reserve(&mut surrounding, count, "surrounding colors")?;
    for _ in 0..count {
        surrounding.push(Argb(c.u32()?));
    }
    let boundary = if flags & BRUSH_PATH != 0 {
        GradientBoundary::Path(Box::new(decode_sized_path(c, "gradient boundary path")?))
    } else {
        let raw_count = c.u32()?;
        let count = c.count(raw_count, "gradient boundary points")?;
        GradientBoundary::Points(c.points_f32(count)?)
    };
    let transform = if flags & BRUSH_TRANSFORM != 0 {
        Some(c.transform()?)
    } else {
        None
    };
    let blends = decode_blends(c, flags, false)?;
    let focus_scales = if flags & BRUSH_FOCUS_SCALES != 0 {
        Some((c.f32()?, c.f32()?))
    } else {
        None
    };
    Ok(PathGradientBrush {
        flags,
        wrap_mode,
        center_color,
        center,
        surrounding,
        boundary,
        transform,
        blends,
        focus_scales,
    })
}

fn decode_blends(c: &mut Cursor<'_>, flags: u32, linear: bool) -> Result<Vec<Blend>> {
    if flags & BRUSH_PRESET_COLORS != 0 && flags & (BRUSH_BLEND_H | BRUSH_BLEND_V) != 0 {
        return Err(error(
            "EMF+ gradient cannot combine preset colors and blend factors",
        ));
    }
    if !linear && flags & BRUSH_BLEND_V != 0 {
        return Err(error(
            "vertical blend factors are only valid for a linear gradient",
        ));
    }
    let mut blends = Vec::new();
    if flags & BRUSH_PRESET_COLORS != 0 {
        blends.push(Blend::Colors(c.blend_colors()?));
    }
    if flags & BRUSH_BLEND_H != 0 {
        blends.push(Blend::Factors(c.blend_factors()?));
    }
    if linear && flags & BRUSH_BLEND_V != 0 {
        blends.push(Blend::Factors(c.blend_factors()?));
    }
    Ok(blends)
}

fn decode_cap(c: &mut Cursor<'_>) -> Result<CustomLineCap> {
    let version = c.u32()?;
    let type_code = c.i32()?;
    let kind = match type_code {
        0 => {
            let flags = c.u32()?;
            let base_cap = c.u32()?;
            let base_inset = c.f32()?;
            let stroke_start_cap = c.u32()?;
            let stroke_end_cap = c.u32()?;
            let stroke_join = c.u32()?;
            let miter_limit = c.f32()?;
            let width_scale = c.f32()?;
            c.skip(16)?;
            let fill_path = if flags & 1 != 0 {
                Some(decode_sized_path(c, "custom cap fill path")?)
            } else {
                None
            };
            let outline_path = if flags & 2 != 0 {
                Some(decode_sized_path(c, "custom cap outline path")?)
            } else {
                None
            };
            CustomLineCapKind::Default {
                flags,
                base_cap,
                base_inset,
                stroke_start_cap,
                stroke_end_cap,
                stroke_join,
                miter_limit,
                width_scale,
                fill_path,
                outline_path,
            }
        },
        1 => {
            let width = c.f32()?;
            let height = c.f32()?;
            let middle_inset = c.f32()?;
            let filled = c.u32()? != 0;
            let start_cap = c.u32()?;
            let end_cap = c.u32()?;
            let line_join = c.u32()?;
            let miter_limit = c.f32()?;
            let width_scale = c.f32()?;
            c.skip(16)?;
            CustomLineCapKind::AdjustableArrow {
                width,
                height,
                middle_inset,
                filled,
                start_cap,
                end_cap,
                line_join,
                miter_limit,
                width_scale,
            }
        },
        _ => CustomLineCapKind::Unsupported {
            type_code,
            data: c.rest(),
        },
    };
    let diagnostics = match &kind {
        CustomLineCapKind::Unsupported { data, .. } => {
            vec![diagnostic("unsupported custom line cap type", data.clone())]
        },
        CustomLineCapKind::Default { .. } | CustomLineCapKind::AdjustableArrow { .. } => Vec::new(),
    };
    Ok(CustomLineCap {
        version,
        kind,
        diagnostics,
    })
}

fn decode_font(c: &mut Cursor<'_>) -> Result<Font> {
    let version = c.u32()?;
    let em_size = c.f32()?;
    let unit = c.u32()?;
    let style = c.i32()?;
    let _reserved = c.u32()?;
    let raw_length = c.u32()?;
    let length = c.count(raw_length, "font family characters")?;
    let units = c.take(
        length
            .checked_mul(2)
            .ok_or_else(|| error("font family length overflow"))?,
    )?;
    let utf16: Vec<u16> = units
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect();
    let family = String::from_utf16(&utf16).map_err(|_| error("invalid UTF-16 font family"))?;
    Ok(Font {
        version,
        em_size,
        unit,
        style,
        family,
    })
}

fn decode_image(c: &mut Cursor<'_>) -> Result<Image> {
    let version = c.u32()?;
    let type_code = c.u32()?;
    let kind = match type_code {
        1 => decode_bitmap(c)?,
        2 => decode_metafile(c)?,
        _ => ImageKind::Unsupported {
            type_code,
            data: c.rest(),
        },
    };
    let diagnostics = match &kind {
        ImageKind::Metafile { data, .. } => vec![diagnostic(
            "EMF+ metafile image retained but not rendered",
            data.clone(),
        )],
        ImageKind::Compressed(data) => vec![diagnostic(
            "compressed EMF+ bitmap retained for an image decoder",
            data.clone(),
        )],
        ImageKind::Unsupported { data, .. } => {
            vec![diagnostic("unsupported EMF+ image type", data.clone())]
        },
        ImageKind::Bitmap(_) => Vec::new(),
    };
    Ok(Image {
        version,
        kind,
        diagnostics,
    })
}

fn decode_metafile(c: &mut Cursor<'_>) -> Result<ImageKind> {
    let type_code = c.u32()?;
    let size = usize::try_from(c.u32()?).map_err(|_| error("metafile size does not fit usize"))?;
    if size > c.limits.max_bytes {
        return Err(error("metafile image exceeds configured byte limit"));
    }
    Ok(ImageKind::Metafile {
        type_code,
        data: c.take(size)?.to_vec(),
    })
}

fn decode_bitmap(c: &mut Cursor<'_>) -> Result<ImageKind> {
    let width = c.i32()?;
    let height = c.i32()?;
    let stride = c.i32()?;
    let pixel_format = c.u32()?;
    let data_type = c.u32()?;
    if width < 0 || height < 0 || stride < 0 || stride % 4 != 0 {
        return Err(error("invalid EMF+ bitmap dimensions or stride"));
    }
    let pixels = c.rest();
    Ok(if data_type == 0 {
        ImageKind::Bitmap(Bitmap {
            width,
            height,
            stride,
            pixel_format,
            pixels,
            compressed: false,
        })
    } else if data_type == 1 {
        ImageKind::Compressed(pixels)
    } else {
        ImageKind::Unsupported {
            type_code: data_type,
            data: pixels,
        }
    })
}

fn decode_attributes(c: &mut Cursor<'_>) -> Result<ImageAttributes> {
    let version = c.u32()?;
    let _reserved1 = c.u32()?;
    let wrap_mode = c.u32()?;
    let clamp_color = Argb(c.u32()?);
    let clamp = c.i32()?;
    let _reserved2 = c.u32()?;
    let extra = c.rest();
    let diagnostics = if extra.is_empty() {
        Vec::new()
    } else {
        vec![diagnostic("unparsed ImageAttributes extension", extra)]
    };
    Ok(ImageAttributes {
        version,
        wrap_mode,
        clamp_color,
        clamp,
        diagnostics,
    })
}

fn decode_path(c: &mut Cursor<'_>) -> Result<Path> {
    let version = c.u32()?;
    let raw_count = c.u32()?;
    let count = c.count(raw_count, "path points")?;
    let flags = c.u32()?;
    let relative = flags & 0x800 != 0;
    let rle = flags & 0x1000 != 0;
    let compressed = flags & 0x4000 != 0;
    let points = if relative {
        c.points_relative(count)?
    } else if compressed {
        c.points_i16(count)?
    } else {
        c.points_f32(count)?
    };
    let point_types = if rle {
        c.path_types_rle(count)?
    } else {
        c.take(count)?.to_vec()
    };
    Ok(Path {
        version,
        flags,
        points,
        point_types,
    })
}

fn decode_sized_path(c: &mut Cursor<'_>, label: &'static str) -> Result<Path> {
    let raw_size = c.i32()?;
    if raw_size < 0 {
        return Err(error(format!("negative {label} size")));
    }
    let size =
        usize::try_from(raw_size).map_err(|_| error(format!("{label} size does not fit usize")))?;
    if size > c.limits.max_bytes {
        return Err(error(format!("{label} exceeds configured byte limit")));
    }
    let data = c.take(size)?;
    let mut inner = Cursor::new(data, c.limits);
    let path = decode_path(&mut inner)?;
    inner.padding()?;
    Ok(path)
}

fn decode_pen(c: &mut Cursor<'_>) -> Result<Pen> {
    let version = c.u32()?;
    let type_code = c.u32()?;
    if type_code != 0 {
        return Err(error("EMF+ pen Type must be zero"));
    }
    let flags = c.u32()?;
    let unit = c.u32()?;
    let width = c.f32()?;
    let transform = optional_transform(c, flags)?;
    let start_cap = optional_i32(c, flags, PEN_START_CAP)?;
    let end_cap = optional_i32(c, flags, PEN_END_CAP)?;
    let join = optional_i32(c, flags, PEN_JOIN)?;
    let miter_limit = optional_f32(c, flags, PEN_MITER)?;
    let line_style = optional_i32(c, flags, PEN_STYLE)?;
    let dashed_cap = optional_i32(c, flags, PEN_DASH_CAP)?;
    let dash_offset = optional_f32(c, flags, PEN_DASH_OFFSET)?;
    let dashes = if flags & PEN_DASH != 0 {
        c.f32_array("dash values")?
    } else {
        Vec::new()
    };
    let alignment = optional_i32(c, flags, PEN_NON_CENTER)?;
    let compound = if flags & PEN_COMPOUND != 0 {
        c.f32_array("compound line values")?
    } else {
        Vec::new()
    };
    let mut diagnostics = Vec::new();
    if flags & PEN_CUSTOM_START != 0 {
        diagnostics.push(diagnostic(
            "custom pen start-cap data retained as unsupported",
            c.sized_data()?,
        ));
    }
    if flags & PEN_CUSTOM_END != 0 {
        diagnostics.push(diagnostic(
            "custom pen end-cap data retained as unsupported",
            c.sized_data()?,
        ));
    }
    let brush = Box::new(decode_brush(c)?);
    Ok(Pen {
        version,
        flags,
        unit,
        width,
        transform,
        start_cap,
        end_cap,
        join,
        miter_limit,
        line_style,
        dashed_cap,
        dash_offset,
        dashes,
        alignment,
        compound,
        brush,
        diagnostics,
    })
}
fn optional_transform(c: &mut Cursor<'_>, flags: u32) -> Result<Option<Transform>> {
    if flags & PEN_TRANSFORM != 0 {
        Ok(Some(c.transform()?))
    } else {
        Ok(None)
    }
}
fn optional_i32(c: &mut Cursor<'_>, flags: u32, bit: u32) -> Result<Option<i32>> {
    if flags & bit != 0 {
        Ok(Some(c.i32()?))
    } else {
        Ok(None)
    }
}
fn optional_f32(c: &mut Cursor<'_>, flags: u32, bit: u32) -> Result<Option<f32>> {
    if flags & bit != 0 {
        Ok(Some(c.f32()?))
    } else {
        Ok(None)
    }
}

fn decode_region(c: &mut Cursor<'_>) -> Result<Region> {
    let version = c.u32()?;
    let raw_child_count = c.u32()?;
    let child_count = c.count(raw_child_count, "region nodes")?;
    let nodes = child_count
        .checked_add(1)
        .ok_or_else(|| error("region node count overflow"))?;
    if nodes > c.limits.max_points {
        return Err(error("region node count exceeds configured limit"));
    }
    let mut decoded_nodes = 0usize;
    let root = decode_region_node(c, 0, &mut decoded_nodes)?;
    if decoded_nodes != nodes {
        return Err(error("EMF+ region node count does not match its tree"));
    }
    let diagnostics = if c.remaining() == 0 {
        Vec::new()
    } else {
        vec![diagnostic("extra region node data", c.rest())]
    };
    Ok(Region {
        version,
        root,
        diagnostics,
    })
}
fn decode_region_node(
    c: &mut Cursor<'_>,
    depth: usize,
    decoded_nodes: &mut usize,
) -> Result<RegionNode> {
    if depth >= c.limits.max_recursion {
        return Err(error("EMF+ region nesting exceeds configured limit"));
    }
    *decoded_nodes = decoded_nodes
        .checked_add(1)
        .ok_or_else(|| error("EMF+ region node count overflow"))?;
    match c.u32()? {
        1 => binary_region(c, depth, decoded_nodes, RegionNode::And),
        2 => binary_region(c, depth, decoded_nodes, RegionNode::Or),
        3 => binary_region(c, depth, decoded_nodes, RegionNode::Xor),
        4 => binary_region(c, depth, decoded_nodes, RegionNode::Exclude),
        5 => binary_region(c, depth, decoded_nodes, RegionNode::Complement),
        0x1000_0000 => Ok(RegionNode::Rect(c.rect()?)),
        0x1000_0001 => Ok(RegionNode::Path(decode_path(c)?)),
        0x1000_0002 => Ok(RegionNode::Empty),
        0x1000_0003 => Ok(RegionNode::Infinite),
        value => Err(error(format!(
            "unknown EMF+ region node type 0x{value:08X}"
        ))),
    }
}
fn binary_region(
    c: &mut Cursor<'_>,
    depth: usize,
    decoded_nodes: &mut usize,
    combine: fn(Box<RegionNode>, Box<RegionNode>) -> RegionNode,
) -> Result<RegionNode> {
    Ok(combine(
        Box::new(decode_region_node(c, depth + 1, decoded_nodes)?),
        Box::new(decode_region_node(c, depth + 1, decoded_nodes)?),
    ))
}

fn decode_string_format(c: &mut Cursor<'_>) -> Result<StringFormat> {
    let version = c.u32()?;
    let flags = c.u32()?;
    let language = c.u16()?;
    c.skip(2)?;
    let alignment = c.u32()?;
    let line_alignment = c.u32()?;
    let digit_substitution = c.u32()?;
    let digit_language = c.u16()?;
    c.skip(2)?;
    let first_tab_offset = c.f32()?;
    let hotkey_prefix = c.i32()?;
    let leading_margin = c.f32()?;
    let trailing_margin = c.f32()?;
    let tracking = c.f32()?;
    let trimming = c.u32()?;
    let raw_tab_count = c.i32()?;
    let tab_count = c.count_i32(raw_tab_count, "tab stops")?;
    let raw_range_count = c.i32()?;
    let range_count = c.count_i32(raw_range_count, "character ranges")?;
    let mut tab_stops = Vec::new();
    c.reserve(&mut tab_stops, tab_count, "tab stops")?;
    for _ in 0..tab_count {
        tab_stops.push(c.f32()?);
    }
    let mut ranges = Vec::new();
    c.reserve(&mut ranges, range_count, "character ranges")?;
    for _ in 0..range_count {
        ranges.push(CharacterRange {
            first: c.i32()?,
            length: c.i32()?,
        });
    }
    Ok(StringFormat {
        version,
        flags,
        language,
        alignment,
        line_alignment,
        digit_substitution,
        digit_language,
        first_tab_offset,
        hotkey_prefix,
        leading_margin,
        trailing_margin,
        tracking,
        trimming,
        tab_stops,
        ranges,
    })
}

struct Cursor<'a> {
    bytes: &'a [u8],
    offset: usize,
    limits: DecodeLimits,
}
impl<'a> Cursor<'a> {
    fn new(bytes: &'a [u8], limits: DecodeLimits) -> Self {
        Self {
            bytes,
            offset: 0,
            limits,
        }
    }
    fn remaining(&self) -> usize {
        self.bytes.len().saturating_sub(self.offset)
    }
    fn take(&mut self, count: usize) -> Result<&'a [u8]> {
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| error("EMF+ offset overflow"))?;
        let value = self
            .bytes
            .get(self.offset..end)
            .ok_or_else(|| error("truncated EMF+ object"))?;
        self.offset = end;
        Ok(value)
    }
    fn rest(&mut self) -> Vec<u8> {
        let result = self.bytes[self.offset..].to_vec();
        self.offset = self.bytes.len();
        result
    }
    fn skip(&mut self, count: usize) -> Result<()> {
        let _ = self.take(count)?;
        Ok(())
    }
    fn u16(&mut self) -> Result<u16> {
        let b = self.take(2)?;
        Ok(u16::from_le_bytes([b[0], b[1]]))
    }
    fn u32(&mut self) -> Result<u32> {
        let b = self.take(4)?;
        Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
    }
    fn u64(&mut self) -> Result<u64> {
        let b = self.take(8)?;
        Ok(u64::from_le_bytes([
            b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7],
        ]))
    }
    fn i32(&mut self) -> Result<i32> {
        Ok(self.u32()?.cast_signed())
    }
    fn f32(&mut self) -> Result<f32> {
        Ok(f32::from_bits(self.u32()?))
    }
    fn point(&mut self) -> Result<Point> {
        Ok(Point {
            x: self.f32()?,
            y: self.f32()?,
        })
    }
    fn rect(&mut self) -> Result<Rect> {
        Ok(Rect {
            x: self.f32()?,
            y: self.f32()?,
            width: self.f32()?,
            height: self.f32()?,
        })
    }
    fn transform(&mut self) -> Result<Transform> {
        Ok(Transform {
            m11: self.f32()?,
            m12: self.f32()?,
            m21: self.f32()?,
            m22: self.f32()?,
            dx: self.f32()?,
            dy: self.f32()?,
        })
    }
    fn count(&self, value: u32, label: &'static str) -> Result<usize> {
        let count = usize::try_from(value)
            .map_err(|_| error(format!("{label} count does not fit usize")))?;
        if count > self.limits.max_points {
            return Err(error(format!("{label} count exceeds configured limit")));
        }
        Ok(count)
    }
    fn count_i32(&self, value: i32, label: &'static str) -> Result<usize> {
        if value < 0 {
            return Err(error(format!("negative {label} count")));
        }
        self.count(
            u32::try_from(value).map_err(|_| error("count conversion failed"))?,
            label,
        )
    }
    #[allow(
        clippy::unused_self,
        reason = "kept as a cursor operation for uniform bounded allocation calls"
    )]
    fn reserve<T>(&self, value: &mut Vec<T>, count: usize, label: &'static str) -> Result<()> {
        value
            .try_reserve(count)
            .map_err(|source| Error::Allocation {
                resource: label,
                source,
            })
    }
    fn points_f32(&mut self, count: usize) -> Result<Vec<Point>> {
        let mut result = Vec::new();
        self.reserve(&mut result, count, "path points")?;
        for _ in 0..count {
            result.push(self.point()?);
        }
        Ok(result)
    }
    fn points_i16(&mut self, count: usize) -> Result<Vec<Point>> {
        let mut result = Vec::new();
        self.reserve(&mut result, count, "path points")?;
        for _ in 0..count {
            result.push(Point {
                x: f32::from(i16::from_le_bytes(
                    self.take(2)?
                        .try_into()
                        .map_err(|_| error("point conversion failed"))?,
                )),
                y: f32::from(i16::from_le_bytes(
                    self.take(2)?
                        .try_into()
                        .map_err(|_| error("point conversion failed"))?,
                )),
            });
        }
        Ok(result)
    }
    fn points_relative(&mut self, count: usize) -> Result<Vec<Point>> {
        let mut result = Vec::new();
        self.reserve(&mut result, count, "path points")?;
        let mut last = Point { x: 0.0, y: 0.0 };
        for _ in 0..count {
            let x = self.integer15()?;
            let y = self.integer15()?;
            last = Point {
                x: last.x + f32::from(x),
                y: last.y + f32::from(y),
            };
            result.push(last);
        }
        Ok(result)
    }
    fn integer15(&mut self) -> Result<i16> {
        let first = self.take(1)?[0];
        if first & 0x80 == 0 {
            Ok(i16::from(first.cast_signed()) << 1 >> 1)
        } else {
            let second = self.take(1)?[0];
            Ok((i16::from(first & 0x7f) << 8 | i16::from(second)) << 1 >> 1)
        }
    }
    fn path_types_rle(&mut self, count: usize) -> Result<Vec<u8>> {
        let mut result = Vec::new();
        self.reserve(&mut result, count, "path point types")?;
        while result.len() < count {
            let run = self.take(1)?[0];
            let run_count = usize::from((run >> 1) & 0x3f);
            if run_count == 0 || run_count > count - result.len() {
                return Err(error("invalid EMF+ path point RLE run"));
            }
            let typ = self.take(1)?[0];
            result.extend(std::iter::repeat_n(typ, run_count));
        }
        Ok(result)
    }
    fn f32_array(&mut self, label: &'static str) -> Result<Vec<f32>> {
        let raw_count = self.u32()?;
        let count = self.count(raw_count, label)?;
        let mut result = Vec::new();
        self.reserve(&mut result, count, label)?;
        for _ in 0..count {
            result.push(self.f32()?);
        }
        Ok(result)
    }
    fn sized_data(&mut self) -> Result<Vec<u8>> {
        let size =
            usize::try_from(self.u32()?).map_err(|_| error("EMF+ data size does not fit usize"))?;
        if size > self.limits.max_bytes {
            return Err(error("EMF+ embedded data exceeds configured byte limit"));
        }
        Ok(self.take(size)?.to_vec())
    }
    fn blend_factors(&mut self) -> Result<Vec<(f32, f32)>> {
        let raw_count = self.u32()?;
        let count = self.count(raw_count, "blend factors")?;
        let mut positions = Vec::new();
        self.reserve(&mut positions, count, "blend factors")?;
        for _ in 0..count {
            positions.push(self.f32()?);
        }
        let mut result = Vec::new();
        self.reserve(&mut result, count, "blend factors")?;
        for position in positions {
            result.push((position, self.f32()?));
        }
        Ok(result)
    }
    fn blend_colors(&mut self) -> Result<Vec<(f32, Argb)>> {
        let raw_count = self.u32()?;
        let count = self.count(raw_count, "blend colors")?;
        let mut positions = Vec::new();
        self.reserve(&mut positions, count, "blend colors")?;
        for _ in 0..count {
            positions.push(self.f32()?);
        }
        let mut result = Vec::new();
        self.reserve(&mut result, count, "blend colors")?;
        for position in positions {
            result.push((position, Argb(self.u32()?)));
        }
        Ok(result)
    }
    fn padding(&mut self) -> Result<()> {
        if self.remaining() > 3 {
            return Err(error("unexpected trailing EMF+ object bytes"));
        }
        self.skip(self.remaining())
    }
}

fn diagnostic(message: &str, bytes: Vec<u8>) -> DecodeDiagnostic {
    DecodeDiagnostic {
        message: message.to_owned(),
        bytes,
    }
}
fn error(message: impl Into<String>) -> Error {
    Error::ParseError(message.into())
}

#[cfg(test)]
mod tests {
    use super::{Argb, DecodeLimits, GraphicsObject, ObjectType, decode_object};
    fn words(values: &[u32]) -> Vec<u8> {
        values
            .iter()
            .flat_map(|value| value.to_le_bytes())
            .collect()
    }
    #[test]
    fn solid_brush_decodes_argb() {
        let object = decode_object(
            ObjectType::Brush,
            &words(&[0, 0, 0x8040_2010]),
            DecodeLimits::default(),
        )
        .expect("valid brush");
        match object {
            GraphicsObject::Brush(brush) => match brush.kind {
                super::BrushKind::Solid { color } => {
                    assert_eq!(color.rgba(), [0x40, 0x20, 0x10, 0x80])
                },
                _ => panic!("solid"),
            },
            _ => panic!("brush"),
        }
    }
    #[test]
    fn truncation_is_an_error() {
        assert!(decode_object(ObjectType::Path, &[0; 5], DecodeLimits::default()).is_err());
    }
    #[test]
    fn point_limit_is_enforced() {
        let data = words(&[0, 2, 0]);
        let limits = DecodeLimits {
            max_points: 1,
            ..DecodeLimits::default()
        };
        assert!(decode_object(ObjectType::Path, &data, limits).is_err());
    }
    #[test]
    fn argb_accessors_are_stable() {
        let color = Argb(0x1122_3344);
        assert_eq!(
            (color.alpha(), color.red(), color.green(), color.blue()),
            (0x11, 0x22, 0x33, 0x44)
        );
    }
}
