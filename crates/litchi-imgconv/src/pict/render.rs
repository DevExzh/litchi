//! Bounded PICT playback to SVG.
//!
//! Playback follows QuickDraw's picture semantics from *Inside Macintosh:
//! Imaging With QuickDraw* (chapter 3, "QuickDraw Drawing", chapter 4,
//! "Color QuickDraw", chapter 7, "Pictures", and Appendix A, "Picture
//! Opcodes"): shapes use the pen and fill patterns with the port colours,
//! text uses the current font state, bitmaps are copied from their source
//! rectangle into their destination rectangle, and each `ClipRgn` replaces
//! the clipping region for everything that follows.

use std::collections::{HashMap, HashSet};
use std::fmt::Write as _;

use base64::Engine as _;
use image::{DynamicImage, ImageFormat, RgbaImage};
use litchi_core::error::{Error, Result};

use super::bitmap::{self, DecodeError, DecodedPixels, PixPatKind, PixelLimits};
use super::parser::{PictParser, PictVersion};
use super::quicktime::{self, Decoded};
use super::region::{self, Region, RegionError};
use super::shape::{self, Bounds};
use super::types::{
    PictRect, mac_roman_to_string, read_fixed, read_i16, read_u16, read_u32, symbol_to_string,
};
use crate::codec::ConversionDiagnostic;
use crate::svg_utils::{write_num, write_xml_escaped};

type Rgb = [u8; 3];

const BLACK: Rgb = [0, 0, 0];
const WHITE: Rgb = [255, 255, 255];
const SOLID: [u8; 8] = [0xFF; 8];
const CLEAR: [u8; 8] = [0; 8];

/// SVG produced by playing back one picture.
#[derive(Debug, Clone)]
pub(crate) struct RenderedPicture {
    pub(crate) svg: String,
    /// Natural raster width: the picture size in points, enlarged to the
    /// resolution of its finest image.
    pub(crate) width: u32,
    /// Natural raster height.
    pub(crate) height: u32,
    /// Whether every painting record is a bitmap or QuickTime image.
    pub(crate) bitmap_only: bool,
    pub(crate) diagnostics: Vec<ConversionDiagnostic>,
}

/// Plays back a parsed picture.
pub(crate) fn render(parser: &PictParser, limits: &crate::Limits) -> Result<RenderedPicture> {
    let geometry = Geometry::new(parser)?;
    let mut renderer = Renderer::new(parser, limits, geometry);
    for warning in parser.warnings() {
        renderer.note(warning.code, warning.message.clone());
    }
    renderer.play()?;
    renderer.finish()
}

/// Picture coordinate space and output size.
#[derive(Debug, Clone, Copy)]
struct Geometry {
    /// Drawing-space rectangle shown by the picture.
    view: Bounds,
    /// SVG document size in points (1/72 inch).
    document: (f64, f64),
}

impl Geometry {
    /// An extended version 2 picture draws in the native-resolution source
    /// rectangle of its header, whose size in points follows from the
    /// resolution; other pictures draw in their frame, which is in points.
    fn new(parser: &PictParser) -> Result<Self> {
        let frame = parser.frame();
        let native = parser.header_op().filter(|header| {
            header.version == -2
                && header.source_rect.is_nonempty()
                && header.h_res > 0.0
                && header.v_res > 0.0
        });
        if let Some(header) = native {
            let view = to_bounds(header.source_rect);
            let document = (
                view.width() * 72.0 / header.h_res,
                view.height() * 72.0 / header.v_res,
            );
            return Ok(Self { view, document });
        }
        let view_rect = (frame.is_nonempty().then_some(frame))
            .or_else(|| drawing_bounds(parser))
            .ok_or_else(|| Error::ParseError("PICT picture frame is empty".into()))?;
        let view = to_bounds(view_rect);
        Ok(Self {
            view,
            document: (view.width(), view.height()),
        })
    }

    /// Points per drawing unit, horizontally and vertically.
    fn points_per_unit(&self) -> (f64, f64) {
        (
            self.document.0 / self.view.width(),
            self.document.1 / self.view.height(),
        )
    }
}

fn to_bounds(rect: PictRect) -> Bounds {
    Bounds {
        left: f64::from(rect.left),
        top: f64::from(rect.top),
        right: f64::from(rect.right),
        bottom: f64::from(rect.bottom),
    }
}

/// Union of the rectangles of drawing records, for pictures whose frame is
/// empty.
fn drawing_bounds(parser: &PictParser) -> Option<PictRect> {
    let mut union: Option<PictRect> = None;
    for record in &parser.records {
        let rect = match record.opcode {
            0x0030..=0x0034 | 0x0040..=0x0044 | 0x0050..=0x0054 | 0x0060..=0x0064 => {
                PictRect::read(&record.data, 0)
            },
            0x0090 | 0x0091 | 0x0098..=0x009B => {
                bitmap::parse_bits_record(&record.data, record.opcode)
                    .map(|(bits, _)| bits.dst_rect)
            },
            _ => None,
        };
        let Some(rect) = rect.filter(PictRect::is_nonempty) else {
            continue;
        };
        union = Some(match union {
            None => rect,
            Some(current) => PictRect {
                top: current.top.min(rect.top),
                left: current.left.min(rect.left),
                bottom: current.bottom.max(rect.bottom),
                right: current.right.max(rect.right),
            },
        });
    }
    union
}

/// A pen, fill or background pattern.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum Pattern {
    /// An 8x8 one-bit pattern drawn in the port colours.
    Bits([u8; 8]),
    /// A decoded full-colour `PixPat`, by index into the renderer's table.
    Color(usize),
    /// A dither pattern standing for an RGB colour.
    Dither(Rgb),
}

/// How a pattern or source combines with the destination.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Transfer {
    Copy,
    Or,
    Xor,
    Bic,
}

/// A transfer mode split into its Boolean operation and source inversion.
/// Arithmetic and highlight modes are drawn as copies and reported.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Mode {
    transfer: Transfer,
    invert: bool,
    approximated: bool,
}

impl Mode {
    fn from_code(code: u16) -> Self {
        // `ditherCopy` (64) only allows dithering; it does not change the
        // operation.
        let code = code & !64;
        let boolean = |transfer| Self {
            transfer,
            invert: code & 4 != 0,
            approximated: false,
        };
        match code {
            0..=15 => boolean(match code & 3 {
                0 => Transfer::Copy,
                1 => Transfer::Or,
                2 => Transfer::Xor,
                _ => Transfer::Bic,
            }),
            // `transparent` leaves background-coloured pixels undrawn, which
            // for patterns and text is the `or` operation.
            36 => Self {
                transfer: Transfer::Or,
                invert: false,
                approximated: false,
            },
            _ => Self {
                transfer: Transfer::Copy,
                invert: false,
                approximated: true,
            },
        }
    }
}

/// How an image combines with what is already drawn.
#[derive(Debug, Clone, Copy, PartialEq)]
enum Composite {
    Normal,
    Blend(&'static str),
    Opacity(f64),
}

#[derive(Debug, Clone, PartialEq)]
enum Paint {
    None,
    Solid(Rgb),
    Url(String),
    /// A colour composited with the destination by an SVG blend mode.
    Blend(Rgb, &'static str),
}

#[derive(Debug, Clone)]
struct TextState {
    font: u16,
    face: u8,
    size: u16,
    mode: u16,
    ratio: (f64, f64),
    space_extra: f64,
    char_extra: f64,
    location: (i32, i32),
}

#[derive(Debug, Clone)]
struct State {
    foreground: Rgb,
    background: Rgb,
    pen_pattern: Pattern,
    fill_pattern: Pattern,
    background_pattern: Pattern,
    pen_size: (i32, i32),
    pen_mode: u16,
    pen_location: (i32, i32),
    oval_size: (i32, i32),
    origin: (i32, i32),
    text: TextState,
    last_rect: PictRect,
    last_round_rect: PictRect,
    last_oval: PictRect,
    last_arc: PictRect,
    last_polygon: Vec<(i32, i32)>,
    last_region: Option<Region>,
}

impl State {
    fn new(view: Bounds) -> Self {
        let start = (view.left as i32, view.top as i32);
        Self {
            foreground: BLACK,
            background: WHITE,
            pen_pattern: Pattern::Bits(SOLID),
            fill_pattern: Pattern::Bits(SOLID),
            background_pattern: Pattern::Bits(CLEAR),
            pen_size: (1, 1),
            pen_mode: 8,
            pen_location: start,
            oval_size: (0, 0),
            origin: (0, 0),
            text: TextState {
                font: 0,
                face: 0,
                size: 0,
                mode: 1,
                ratio: (1.0, 1.0),
                space_extra: 0.0,
                char_extra: 0.0,
                location: start,
            },
            last_rect: PictRect::default(),
            last_round_rect: PictRect::default(),
            last_oval: PictRect::default(),
            last_arc: PictRect::default(),
            last_polygon: Vec::new(),
            last_region: None,
        }
    }
}

/// The drawing method encoded in the low three bits of a shape opcode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Verb {
    Frame,
    Paint,
    Erase,
    Invert,
    Fill,
}

impl Verb {
    fn from_opcode(opcode: u16) -> Option<Self> {
        match opcode & 7 {
            0 => Some(Self::Frame),
            1 => Some(Self::Paint),
            2 => Some(Self::Erase),
            3 => Some(Self::Invert),
            4 => Some(Self::Fill),
            _ => None,
        }
    }
}

enum Shape {
    Rect(PictRect),
    RoundRect(PictRect),
    Oval(PictRect),
    Arc(PictRect, i16, i16),
    Polygon(Vec<(i32, i32)>),
    Region(Region),
}

/// Where and how an image is drawn.
#[derive(Debug, Clone, Copy)]
struct Placement {
    /// The image rectangle, before `matrix`.
    bounds: Bounds,
    /// A transform from image to picture coordinates.
    matrix: Option<[f64; 6]>,
    /// Whether enlarged pixels stay square, as QuickDraw copies them.
    pixelated: bool,
    composite: Composite,
}

/// A decoded bitmap waiting to be emitted. QuickDraw applications store
/// large images as horizontal bands; consecutive bands are joined so that
/// scaled output shows no seams between them.
struct PendingImage {
    image: RgbaImage,
    placed: Bounds,
    composite: Composite,
    clip: Option<String>,
}

impl PendingImage {
    fn can_append(&self, band: &PendingImage, limits: &crate::Limits) -> bool {
        let (width, height) = self.image.dimensions();
        let (band_width, band_height) = band.image.dimensions();
        let row_height = self.placed.height() / f64::from(height);
        let band_row_height = band.placed.height() / f64::from(band_height);
        let rows = u64::from(height) + u64::from(band_height);
        self.composite == Composite::Normal
            && band.composite == Composite::Normal
            && self.clip == band.clip
            && width == band_width
            && self.placed.left == band.placed.left
            && self.placed.right == band.placed.right
            && (self.placed.bottom - band.placed.top).abs() < 1e-9
            // Writers round each band's destination to whole units, so
            // band scales differ slightly; joining them moves rows by less
            // than one unit, where separate images would show seams.
            && (row_height - band_row_height).abs() <= 0.02 * row_height
            && rows <= u64::from(limits.max_height)
            && rows * u64::from(width) <= limits.max_pixels
    }

    fn append(&mut self, band: PendingImage) {
        let (width, height) = self.image.dimensions();
        let rows = height + band.image.height();
        let mut pixels = std::mem::take(&mut self.image).into_raw();
        pixels.extend_from_slice(band.image.as_raw());
        self.image = RgbaImage::from_raw(width, rows, pixels).unwrap_or_default();
        self.placed.bottom = band.placed.bottom;
    }
}

/// Most distinct diagnostics kept per picture.
const MAX_DIAGNOSTICS: usize = 64;
/// Most font names remembered from `FontName` records.
const MAX_FONT_NAMES: usize = 256;
/// Largest region, in pixels, framed by eroding its pixel mask.
const MAX_FRAME_REGION_PIXELS: u64 = 4 * 1024 * 1024;

struct Renderer<'a> {
    parser: &'a PictParser,
    limits: &'a crate::Limits,
    geometry: Geometry,
    state: State,
    body: String,
    defs: String,
    next_id: usize,
    pattern_ids: HashMap<([u8; 8], Option<Rgb>, Option<Rgb>), String>,
    color_patterns: Vec<Option<String>>,
    font_names: HashMap<u16, String>,
    /// The current clipping path, applied to every drawn element.
    clip: Option<String>,
    /// Whether an element blends with the destination, which then needs
    /// QuickDraw's white port background behind the picture.
    uses_blend: bool,
    /// Consecutive bitmap bands waiting to be emitted as one image.
    pending: Option<PendingImage>,
    /// Bounds of everything drawn in a colour other than white; outside
    /// them the destination is still the white port background.
    painted: Option<Bounds>,
    /// The `OpColor` used by arithmetic transfer modes, once set.
    op_color: Option<Rgb>,
    /// Index of the record after the one being played.
    next_record: usize,
    elements: usize,
    element_limit_hit: bool,
    diagnostics: Vec<ConversionDiagnostic>,
    noted: HashSet<&'static str>,
    painted_bitmap: bool,
    painted_vector: bool,
    /// Largest source-to-destination scale of a drawn bitmap.
    bitmap_scale: f64,
}

impl<'a> Renderer<'a> {
    fn new(parser: &'a PictParser, limits: &'a crate::Limits, geometry: Geometry) -> Self {
        Self {
            parser,
            limits,
            geometry,
            state: State::new(geometry.view),
            body: String::new(),
            defs: String::new(),
            next_id: 0,
            pattern_ids: HashMap::new(),
            color_patterns: Vec::new(),
            font_names: HashMap::new(),
            clip: None,
            uses_blend: false,
            pending: None,
            painted: None,
            op_color: None,
            next_record: 0,
            elements: 0,
            element_limit_hit: false,
            diagnostics: Vec::new(),
            noted: HashSet::new(),
            painted_bitmap: false,
            painted_vector: false,
            bitmap_scale: 1.0,
        }
    }

    fn note(&mut self, code: &'static str, message: impl Into<String>) {
        if self.noted.len() < MAX_DIAGNOSTICS && self.noted.insert(code) {
            self.diagnostics.push(ConversionDiagnostic {
                code,
                message: message.into(),
            });
        }
    }

    fn new_id(&mut self, kind: &str) -> String {
        let id = format!("pict-{kind}{}", self.next_id);
        self.next_id += 1;
        id
    }

    fn check_output(&self) -> Result<()> {
        let size = self.body.len().saturating_add(self.defs.len());
        if size > self.limits.max_output_bytes {
            return Err(Error::ParseError(format!(
                "PICT SVG output exceeds limit {} bytes",
                self.limits.max_output_bytes
            )));
        }
        Ok(())
    }

    /// Reserves one SVG element; returns `false` once the element limit is
    /// reached, after reporting it.
    fn reserve_element(&mut self) -> bool {
        if self.elements >= self.limits.max_svg_elements {
            if !self.element_limit_hit {
                self.element_limit_hit = true;
                let limit = self.limits.max_svg_elements;
                self.note(
                    "pict-svg-element-limit",
                    format!("picture needs more than {limit} SVG elements; the rest is not drawn"),
                );
            }
            return false;
        }
        self.elements += 1;
        true
    }

    fn map(&self, x: i32, y: i32) -> (f64, f64) {
        (
            f64::from(x) - f64::from(self.state.origin.0),
            f64::from(y) - f64::from(self.state.origin.1),
        )
    }

    fn map_rect(&self, rect: PictRect) -> Bounds {
        let (left, top) = self.map(i32::from(rect.left), i32::from(rect.top));
        let (right, bottom) = self.map(i32::from(rect.right), i32::from(rect.bottom));
        Bounds {
            left,
            top,
            right,
            bottom,
        }
    }

    fn play(&mut self) -> Result<()> {
        let records = &self.parser.records;
        let mut index = 0usize;
        while index < records.len() {
            let record = &records[index];
            index += 1;
            if self.element_limit_hit {
                break;
            }
            if !matches!(record.opcode, 0x0090 | 0x0091 | 0x0098..=0x009B) {
                self.flush_pending()?;
            }
            self.next_record = index;
            let drew_quicktime = self.play_record(record.opcode, &record.data)?;
            if drew_quicktime
                && self.parser.header.version == PictVersion::V2
                && let Some(fallback) = records.get(index).and_then(fallback_length)
            {
                index = self.skip_fallback(index, fallback);
            }
            self.check_output()?;
        }
        self.flush_pending()
    }

    /// Returns the index of the first record that starts at least `bytes`
    /// after the end of the marker record `marker`.
    fn skip_fallback(&self, marker: usize, bytes: usize) -> usize {
        let records = &self.parser.records;
        let Some(start) = self.parser.record_offset(marker) else {
            return marker + 1;
        };
        // The marker is a `PnSize` record: a two-byte opcode and four bytes.
        let end = start + 6 + bytes;
        let mut index = marker + 1;
        while index < records.len()
            && self
                .parser
                .record_offset(index)
                .is_some_and(|offset| offset < end)
        {
            index += 1;
        }
        index
    }

    /// Plays one record. Returns whether it drew a QuickTime image.
    fn play_record(&mut self, opcode: u16, data: &[u8]) -> Result<bool> {
        match opcode {
            0x0000 | 0x0011 | 0x0C00 | 0x00A0 | 0x00A1 => {},
            0x0001 => self.set_clip(data),
            0x0002 => self.state.background_pattern = Pattern::Bits(pattern_bits(data)),
            0x0003 => self.state.text.font = read_u16(data, 0).unwrap_or(0),
            0x0004 => self.state.text.face = data.first().copied().unwrap_or(0),
            0x0005 => self.state.text.mode = read_u16(data, 0).unwrap_or(1),
            0x0006 => self.state.text.space_extra = read_fixed(data, 0).unwrap_or(0.0),
            0x0007 => {
                let height = read_i16(data, 0).unwrap_or(1);
                let width = read_i16(data, 2).unwrap_or(1);
                self.state.pen_size = (i32::from(width), i32::from(height));
            },
            0x0008 => self.state.pen_mode = read_u16(data, 0).unwrap_or(8),
            0x0009 => self.state.pen_pattern = Pattern::Bits(pattern_bits(data)),
            0x000A => self.state.fill_pattern = Pattern::Bits(pattern_bits(data)),
            0x000B => {
                let height = read_i16(data, 0).unwrap_or(0);
                let width = read_i16(data, 2).unwrap_or(0);
                self.state.oval_size = (i32::from(width), i32::from(height));
            },
            0x000C => {
                // `Origin` stores the horizontal offset first. After it, a
                // coordinate `c` names the place `c - offset` named before;
                // the pen and text locations keep their numeric values, as
                // `SetOrigin` leaves them in local coordinates.
                let dh = i32::from(read_i16(data, 0).unwrap_or(0));
                let dv = i32::from(read_i16(data, 2).unwrap_or(0));
                self.state.origin.0 = self.state.origin.0.saturating_add(dh);
                self.state.origin.1 = self.state.origin.1.saturating_add(dv);
            },
            0x000D => self.state.text.size = read_u16(data, 0).unwrap_or(0),
            0x000E => {
                let color = read_u32(data, 0).unwrap_or(33);
                self.state.foreground = self.classic_color(color);
            },
            0x000F => {
                let color = read_u32(data, 0).unwrap_or(30);
                self.state.background = self.classic_color(color);
            },
            0x0010 => self.set_text_ratio(data),
            0x0012..=0x0014 => self.set_pixpat(opcode, data),
            0x001F => self.op_color = read_rgb(data),
            0x0015 | 0x0017..=0x0019 | 0x001C..=0x001E | 0x002D | 0x002E => {},
            0x0016 => {
                // chExtra is a 4.12 fixed-point value.
                self.state.text.char_extra = f64::from(read_i16(data, 0).unwrap_or(0)) / 4096.0;
            },
            0x001A => self.state.foreground = read_rgb(data).unwrap_or(BLACK),
            0x001B => self.state.background = read_rgb(data).unwrap_or(WHITE),
            0x0020..=0x0023 => self.draw_line(opcode, data),
            0x0028..=0x002B => self.draw_text(opcode, data),
            0x002C => self.set_font_name(data),
            0x0030..=0x0034 => {
                if let Some(rect) = PictRect::read(data, 0) {
                    self.state.last_rect = rect;
                    self.draw_shape(opcode, Shape::Rect(rect));
                }
            },
            0x0038..=0x003C => self.draw_shape(opcode, Shape::Rect(self.state.last_rect)),
            0x0040..=0x0044 => {
                if let Some(rect) = PictRect::read(data, 0) {
                    self.state.last_round_rect = rect;
                    self.draw_shape(opcode, Shape::RoundRect(rect));
                }
            },
            0x0048..=0x004C => {
                self.draw_shape(opcode, Shape::RoundRect(self.state.last_round_rect));
            },
            0x0050..=0x0054 => {
                if let Some(rect) = PictRect::read(data, 0) {
                    self.state.last_oval = rect;
                    self.draw_shape(opcode, Shape::Oval(rect));
                }
            },
            0x0058..=0x005C => self.draw_shape(opcode, Shape::Oval(self.state.last_oval)),
            0x0060..=0x0064 => {
                if let Some(rect) = PictRect::read(data, 0) {
                    self.state.last_arc = rect;
                    let start = read_i16(data, 8).unwrap_or(0);
                    let sweep = read_i16(data, 10).unwrap_or(0);
                    self.draw_shape(opcode, Shape::Arc(rect, start, sweep));
                }
            },
            0x0068..=0x006C => {
                let start = read_i16(data, 0).unwrap_or(0);
                let sweep = read_i16(data, 2).unwrap_or(0);
                self.draw_shape(opcode, Shape::Arc(self.state.last_arc, start, sweep));
            },
            0x0070..=0x0074 => {
                let points = read_polygon(data);
                if points.len() > self.limits.max_path_points {
                    self.note("pict-path-limit", "a polygon exceeds the path point limit");
                    return Ok(false);
                }
                self.state.last_polygon = points.clone();
                self.draw_shape(opcode, Shape::Polygon(points));
            },
            0x0078..=0x007C => {
                let points = self.state.last_polygon.clone();
                self.draw_shape(opcode, Shape::Polygon(points));
            },
            0x0080..=0x0084 => {
                if let Some(region) = self.read_region(data) {
                    self.state.last_region = Some(region.clone());
                    self.draw_shape(opcode, Shape::Region(region));
                }
            },
            0x0088..=0x008C => {
                if let Some(region) = self.state.last_region.clone() {
                    self.draw_shape(opcode, Shape::Region(region));
                }
            },
            0x0090 | 0x0091 | 0x0098..=0x009B => self.draw_bits(opcode, data)?,
            0x8200 => return self.draw_quicktime(data),
            0x8201 => self.note(
                "pict-quicktime-matte",
                "an uncompressed QuickTime matte is not applied to the image that follows",
            ),
            _ => self.note(
                "noncanonical-pict-reserved-opcode",
                format!("reserved opcode 0x{opcode:04X} is skipped as QuickDraw does"),
            ),
        }
        Ok(false)
    }

    /// Maps the eight classic QuickDraw colour constants to the RGB values
    /// Color QuickDraw uses for them.
    fn classic_color(&mut self, value: u32) -> Rgb {
        match value {
            30 => WHITE,
            33 => BLACK,
            69 => [0xFC, 0xF3, 0x05],
            137 => [0xF2, 0x08, 0x84],
            205 => [0xDD, 0x08, 0x06],
            273 => [0x02, 0xAB, 0xEA],
            341 => [0x00, 0x80, 0x11],
            409 => [0x00, 0x00, 0xD4],
            _ => {
                self.note(
                    "pict-classic-color",
                    format!("classic colour value {value} is not a QuickDraw colour constant; black is used"),
                );
                BLACK
            },
        }
    }

    fn set_text_ratio(&mut self, data: &[u8]) {
        let ratio = |offset: usize| {
            let numerator = read_i16(data, offset).map_or(1.0, f64::from);
            let denominator = read_i16(data, offset + 4).map_or(1.0, f64::from);
            if numerator > 0.0 && denominator > 0.0 {
                numerator / denominator
            } else {
                1.0
            }
        };
        // `numer` and `denom` are points: vertical, then horizontal.
        self.state.text.ratio = (ratio(2), ratio(0));
    }

    /// The name of a font number: from a `FontName` record, or the classic
    /// Macintosh font of that number.
    fn font_name(&self, id: u16) -> String {
        match self.font_names.get(&id) {
            Some(name) => name.clone(),
            None => classic_font_name(id).to_string(),
        }
    }

    fn set_font_name(&mut self, data: &[u8]) {
        let Some(id) = read_u16(data, 2) else {
            return;
        };
        let Some(&length) = data.get(4) else {
            return;
        };
        let Some(name) = data.get(5..5 + usize::from(length)) else {
            return;
        };
        if self.font_names.len() < MAX_FONT_NAMES || self.font_names.contains_key(&id) {
            self.font_names.insert(id, mac_roman_to_string(name));
        }
    }

    fn set_pixpat(&mut self, opcode: u16, data: &[u8]) {
        let Some((pixpat, _)) = bitmap::parse_pixpat(data) else {
            return;
        };
        let pattern = match pixpat.kind {
            PixPatKind::Bits => Pattern::Bits(pixpat.pat1),
            PixPatKind::Dither(rgb) => Pattern::Dither(rgb),
            PixPatKind::Color { .. } if self.color_patterns.len() >= self.limits.max_objects => {
                self.note(
                    "pict-pixpat-skipped",
                    "colour patterns beyond the object limit are drawn with their monochrome form",
                );
                Pattern::Bits(pixpat.pat1)
            },
            PixPatKind::Color { pixmap, pixels } => {
                match bitmap::decode_pixpat(&pixmap, pixels, self.pixel_limits()) {
                    Ok(decoded) => {
                        let id = self.color_pattern_def(&decoded);
                        self.color_patterns.push(id);
                        Pattern::Color(self.color_patterns.len() - 1)
                    },
                    Err(error) => {
                        self.note(
                            "pict-pixpat-skipped",
                            format!(
                                "a colour pattern was drawn with its monochrome form: {}",
                                decode_error_text(&error)
                            ),
                        );
                        Pattern::Bits(pixpat.pat1)
                    },
                }
            },
        };
        match opcode {
            0x0012 => self.state.background_pattern = pattern,
            0x0013 => self.state.pen_pattern = pattern,
            _ => self.state.fill_pattern = pattern,
        }
    }

    fn color_pattern_def(&mut self, decoded: &DecodedPixels) -> Option<String> {
        let png = encode_png(
            decoded.width,
            decoded.height,
            decoded.rgba.clone(),
            self.limits,
        )?;
        let id = self.new_id("pixpat");
        let _ = write!(
            self.defs,
            r#"<pattern id="{id}" width="{w}" height="{h}" patternUnits="userSpaceOnUse"><image width="{w}" height="{h}" href="data:image/png;base64,"#,
            w = decoded.width,
            h = decoded.height
        );
        base64::engine::general_purpose::STANDARD.encode_string(&png, &mut self.defs);
        self.defs.push_str(r#""/></pattern>"#);
        Some(id)
    }

    fn pixel_limits(&self) -> PixelLimits {
        PixelLimits {
            max_width: self.limits.max_width,
            max_height: self.limits.max_height,
            max_pixels: self.limits.max_pixels,
        }
    }

    fn read_region(&mut self, data: &[u8]) -> Option<Region> {
        let max_rects = self.limits.max_path_points / 4;
        match region::parse_region(data, max_rects) {
            Ok(region) => Some(region),
            Err(RegionError::Truncated) => {
                self.note(
                    "pict-region-malformed",
                    "a region's scan-line data is truncated and is not drawn",
                );
                None
            },
            Err(RegionError::TooComplex) => {
                self.note(
                    "pict-path-limit",
                    "a region exceeds the path point limit and is not drawn",
                );
                None
            },
        }
    }

    fn set_clip(&mut self, data: &[u8]) {
        let Some(region) = self.read_region(data) else {
            return;
        };
        let view = self.geometry.view;
        let covers_view = region.is_rectangular() && {
            let bounds = self.map_rect(region.bounds);
            bounds.left <= view.left
                && bounds.top <= view.top
                && bounds.right >= view.right
                && bounds.bottom >= view.bottom
        };
        if covers_view {
            self.clip = None;
            return;
        }
        let id = self.new_id("clip");
        let mut path = String::new();
        let origin = self.state.origin;
        region.write_path(&mut path, |x, y| {
            (x - f64::from(origin.0), y - f64::from(origin.1))
        });
        let _ = write!(
            self.defs,
            r#"<clipPath id="{id}"><path d="{path}"/></clipPath>"#
        );
        self.clip = Some(id);
    }

    /// Writes the current clipping path attribute.
    fn write_clip(&mut self) {
        if let Some(id) = &self.clip {
            let _ = write!(self.body, r#" clip-path="url(#{id})""#);
        }
    }

    /// Writes the attribute that composites an element with the
    /// destination.
    fn write_composite(&mut self, composite: Composite) {
        match composite {
            Composite::Normal => {},
            Composite::Blend(blend) => {
                self.uses_blend = true;
                let _ = write!(self.body, r#" style="mix-blend-mode:{blend}""#);
            },
            Composite::Opacity(opacity) => {
                self.body.push_str(r#" opacity=""#);
                write_num(&mut self.body, opacity);
                self.body.push('"');
            },
        }
    }

    /// Records drawing in a colour other than white over `bounds`.
    fn mark_painted(&mut self, bounds: Bounds) {
        self.painted = Some(match self.painted {
            None => bounds,
            Some(current) => Bounds {
                left: current.left.min(bounds.left),
                top: current.top.min(bounds.top),
                right: current.right.max(bounds.right),
                bottom: current.bottom.max(bounds.bottom),
            },
        });
    }

    /// Whether nothing but white has been drawn under `bounds`.
    fn untouched(&self, bounds: Bounds) -> bool {
        self.painted.is_none_or(|painted| {
            painted.right <= bounds.left
                || bounds.right <= painted.left
                || painted.bottom <= bounds.top
                || bounds.bottom <= painted.top
        })
    }

    /// Resolves a pattern drawn in `mode` to an SVG paint.
    ///
    /// One-bit patterns draw their set bits in the foreground colour and
    /// their clear bits in the background colour (`patCopy`), only the set
    /// bits (`patOr`), or only the set bits in the background colour
    /// (`patBic`). Colour patterns are exact only in copy mode.
    fn pattern_paint(&mut self, pattern: Pattern, mode: u16) -> Paint {
        let mode = Mode::from_code(mode);
        if mode.approximated {
            self.note(
                "pict-transfer-mode",
                "an arithmetic or highlight transfer mode is drawn as a copy",
            );
        }
        match pattern {
            Pattern::Bits(bits) => {
                let bits = if mode.invert {
                    bits.map(|byte| !byte)
                } else {
                    bits
                };
                let (one, zero) = match mode.transfer {
                    Transfer::Copy => (Some(self.state.foreground), Some(self.state.background)),
                    Transfer::Or => (Some(self.state.foreground), None),
                    Transfer::Bic => (Some(self.state.background), None),
                    Transfer::Xor => {
                        self.note_xor();
                        if bits == SOLID {
                            return Paint::Blend(WHITE, "difference");
                        }
                        (Some(BLACK), None)
                    },
                };
                self.bits_paint(bits, one, zero)
            },
            Pattern::Dither(rgb) => match mode.transfer {
                Transfer::Copy if !mode.invert => Paint::Solid(rgb),
                _ => {
                    self.note(
                        "pict-transfer-mode",
                        "a colour pattern in a non-copy mode is drawn as a copy",
                    );
                    Paint::Solid(rgb)
                },
            },
            Pattern::Color(index) => {
                if mode.transfer != Transfer::Copy || mode.invert {
                    self.note(
                        "pict-transfer-mode",
                        "a colour pixel pattern in a non-copy mode is drawn as a copy",
                    );
                }
                match self.color_patterns.get(index).cloned().flatten() {
                    Some(id) => Paint::Url(id),
                    None => Paint::Solid(self.state.foreground),
                }
            },
        }
    }

    fn note_xor(&mut self) {
        self.note(
            "pict-xor-mode",
            "an exclusive-or or inverting operation is drawn as a difference over a white destination",
        );
    }

    fn bits_paint(&mut self, bits: [u8; 8], one: Option<Rgb>, zero: Option<Rgb>) -> Paint {
        let solid = |color: Option<Rgb>| color.map_or(Paint::None, Paint::Solid);
        if bits == SOLID || one == zero {
            return if bits == CLEAR {
                solid(zero)
            } else {
                solid(one)
            };
        }
        if bits == CLEAR {
            return solid(zero);
        }
        let key = (bits, one, zero);
        if let Some(id) = self.pattern_ids.get(&key) {
            return Paint::Url(id.clone());
        }
        let id = self.new_id("pat");
        let _ = write!(
            self.defs,
            r#"<pattern id="{id}" width="8" height="8" patternUnits="userSpaceOnUse">"#
        );
        if let Some(zero) = zero {
            let _ = write!(self.defs, r#"<path d="M0 0H8V8H0Z" fill="{}"/>"#, hex(zero));
        }
        if let Some(one) = one {
            let mut path = String::new();
            for (row, &byte) in bits.iter().enumerate() {
                let mut column = 0usize;
                while column < 8 {
                    if byte & (0x80 >> column) == 0 {
                        column += 1;
                        continue;
                    }
                    let start = column;
                    while column < 8 && byte & (0x80 >> column) != 0 {
                        column += 1;
                    }
                    let _ = write!(path, "M{start} {row}H{column}V{}H{start}Z", row + 1);
                }
            }
            let _ = write!(self.defs, r#"<path d="{path}" fill="{}"/>"#, hex(one));
        }
        self.defs.push_str("</pattern>");
        self.pattern_ids.insert(key, id.clone());
        Paint::Url(id)
    }

    fn emit_path(&mut self, path: &str, paint: Paint, even_odd: bool, bounds: Bounds) {
        let (fill, composite) = match &paint {
            Paint::None => return,
            Paint::Solid(color) => (hex(*color), Composite::Normal),
            Paint::Url(id) => (format!("url(#{id})"), Composite::Normal),
            Paint::Blend(color, blend) => (hex(*color), Composite::Blend(blend)),
        };
        if path.is_empty() || !self.reserve_element() {
            return;
        }
        if paint != Paint::Solid(WHITE) {
            self.mark_painted(bounds);
        }
        let _ = write!(self.body, r#"<path d="{path}" fill="{fill}""#);
        if even_odd {
            self.body.push_str(r#" fill-rule="evenodd""#);
        }
        self.write_clip();
        self.write_composite(composite);
        self.body.push_str("/>");
    }

    /// The pen size in drawing units. `DrawPicture` scales the pen with
    /// `ScalePt`, which keeps a nonzero pen at least one pixel wide; the
    /// pen is therefore at least one point, the picture's natural pixel.
    fn pen(&self) -> Option<(f64, f64)> {
        let (width, height) = self.state.pen_size;
        let (x_points, y_points) = self.geometry.points_per_unit();
        (width > 0 && height > 0).then(|| {
            (
                f64::from(width).max(1.0 / x_points),
                f64::from(height).max(1.0 / y_points),
            )
        })
    }

    fn draw_line(&mut self, opcode: u16, data: &[u8]) {
        let point = |offset: usize| -> Option<(i32, i32)> {
            Some((
                i32::from(read_i16(data, offset + 2)?),
                i32::from(read_i16(data, offset)?),
            ))
        };
        let delta = |offset: usize| -> Option<(i32, i32)> {
            Some((
                i32::from(*data.get(offset)? as i8),
                i32::from(*data.get(offset + 1)? as i8),
            ))
        };
        let segment = match opcode {
            0x0020 => point(0).zip(point(4)),
            0x0021 => point(0).map(|to| (self.state.pen_location, to)),
            0x0022 => point(0)
                .zip(delta(4))
                .map(|(from, (dh, dv))| (from, (from.0 + dh, from.1 + dv))),
            _ => delta(0).map(|(dh, dv)| {
                let from = self.state.pen_location;
                (from, (from.0 + dh, from.1 + dv))
            }),
        };
        let Some((from, to)) = segment else {
            return;
        };
        self.state.pen_location = to;
        self.painted_vector = true;
        let Some(pen) = self.pen() else {
            return;
        };
        if self.hidden_by_pen_mode() {
            return;
        }
        let mut path = String::new();
        let (start, end) = (self.map(from.0, from.1), self.map(to.0, to.1));
        shape::pen_line(&mut path, start, end, pen);
        let area = points_bounds(&[start, end], pen);
        let paint = self.pattern_paint(self.state.pen_pattern, self.state.pen_mode);
        self.emit_path(&path, paint, false, area);
    }

    /// Pen mode 23 is the PostScript-only mode applications use for drawing
    /// that only PostScript printers show; QuickDraw draws nothing with it.
    fn hidden_by_pen_mode(&self) -> bool {
        self.state.pen_mode == 23
    }

    fn draw_shape(&mut self, opcode: u16, shape: Shape) {
        let Some(verb) = Verb::from_opcode(opcode) else {
            return;
        };
        self.painted_vector = true;
        if matches!(verb, Verb::Frame | Verb::Paint) && self.hidden_by_pen_mode() {
            return;
        }
        let pen = self.pen();
        if verb == Verb::Frame && pen.is_none() {
            return;
        }
        let pen = pen.unwrap_or((1.0, 1.0));
        let mut path = String::new();
        let mut even_odd = false;
        let area = match &shape {
            Shape::Rect(rect)
            | Shape::RoundRect(rect)
            | Shape::Oval(rect)
            | Shape::Arc(rect, ..) => self.map_rect(*rect),
            Shape::Polygon(points) => {
                let mapped: Vec<(f64, f64)> = points.iter().map(|&(x, y)| self.map(x, y)).collect();
                points_bounds(&mapped, pen)
            },
            Shape::Region(region) => self.map_rect(region.bounds),
        };
        match &shape {
            Shape::Rect(rect)
            | Shape::RoundRect(rect)
            | Shape::Oval(rect)
            | Shape::Arc(rect, ..) => {
                if !rect.is_nonempty() {
                    return;
                }
                let bounds = self.map_rect(*rect);
                let oval = (
                    f64::from(self.state.oval_size.0),
                    f64::from(self.state.oval_size.1),
                );
                match (&shape, verb) {
                    (Shape::Rect(_), Verb::Frame) => {
                        shape::frame_rect(&mut path, bounds, pen);
                        even_odd = true;
                    },
                    (Shape::Rect(_), _) => shape::rect(&mut path, bounds),
                    (Shape::RoundRect(_), Verb::Frame) => {
                        shape::frame_round_rect(&mut path, bounds, oval, pen);
                        even_odd = true;
                    },
                    (Shape::RoundRect(_), _) => {
                        shape::round_rect(&mut path, bounds, oval.0, oval.1)
                    },
                    (Shape::Oval(_), Verb::Frame) => {
                        shape::frame_oval(&mut path, bounds, pen);
                        even_odd = true;
                    },
                    (Shape::Oval(_), _) => shape::oval(&mut path, bounds),
                    (Shape::Arc(_, start, sweep), Verb::Frame) => {
                        shape::arc_band(
                            &mut path,
                            bounds,
                            f64::from(*start),
                            f64::from(*sweep),
                            pen,
                        );
                        even_odd = true;
                    },
                    (Shape::Arc(_, start, sweep), _) => {
                        shape::wedge(&mut path, bounds, f64::from(*start), f64::from(*sweep));
                    },
                    _ => {},
                }
            },
            Shape::Polygon(points) => {
                if points.len() < 2 {
                    return;
                }
                let mapped: Vec<(f64, f64)> = points.iter().map(|&(x, y)| self.map(x, y)).collect();
                if verb == Verb::Frame {
                    for pair in mapped.windows(2) {
                        shape::pen_line(&mut path, pair[0], pair[1], pen);
                    }
                } else {
                    shape::polygon(&mut path, &mapped);
                    even_odd = true;
                }
            },
            Shape::Region(region) => {
                let origin = self.state.origin;
                let map = |x: f64, y: f64| (x - f64::from(origin.0), y - f64::from(origin.1));
                if verb == Verb::Frame {
                    match frame_region(region, self.state.pen_size) {
                        Some(frame) => frame.write_path(&mut path, map),
                        None => {
                            self.note(
                                "pict-frame-region",
                                "a region frame is too large to compute and is not drawn",
                            );
                            return;
                        },
                    }
                } else {
                    region.write_path(&mut path, map);
                }
            },
        }
        let paint = match verb {
            Verb::Frame | Verb::Paint => {
                self.pattern_paint(self.state.pen_pattern, self.state.pen_mode)
            },
            Verb::Erase => self.pattern_paint(self.state.background_pattern, 8),
            Verb::Fill => self.pattern_paint(self.state.fill_pattern, 8),
            Verb::Invert => {
                self.note_xor();
                Paint::Blend(WHITE, "difference")
            },
        };
        self.emit_path(&path, paint, even_odd, area);
    }

    fn draw_text(&mut self, opcode: u16, data: &[u8]) {
        let location = &mut self.state.text.location;
        let text_start = match opcode {
            0x0028 => {
                let (Some(v), Some(h)) = (read_i16(data, 0), read_i16(data, 2)) else {
                    return;
                };
                *location = (i32::from(h), i32::from(v));
                4
            },
            0x0029 => {
                location.0 += i32::from(data.first().copied().unwrap_or(0));
                1
            },
            0x002A => {
                location.1 += i32::from(data.first().copied().unwrap_or(0));
                1
            },
            _ => {
                location.0 += i32::from(data.first().copied().unwrap_or(0));
                location.1 += i32::from(data.get(1).copied().unwrap_or(0));
                2
            },
        };
        let Some(&count) = data.get(text_start) else {
            return;
        };
        let Some(bytes) = data.get(text_start + 1..text_start + 1 + usize::from(count)) else {
            return;
        };
        self.painted_vector = true;
        let symbol = self
            .font_name(self.state.text.font)
            .eq_ignore_ascii_case("symbol");
        let decoded = if symbol {
            symbol_to_string(bytes)
        } else {
            mac_roman_to_string(bytes)
        };
        let text: String = decoded
            .chars()
            .map(|character| {
                if character.is_control() {
                    ' '
                } else {
                    character
                }
            })
            .collect();
        let text = text.trim_end();
        if text.is_empty() {
            return;
        }
        let size = if self.state.text.size == 0 {
            12
        } else {
            self.state.text.size
        };
        let (x, y) = self.map(self.state.text.location.0, self.state.text.location.1);
        let (sx, sy) = self.state.text.ratio;
        // Glyph metrics are unknown; the box covers the text in any common
        // font.
        let area = {
            let size = f64::from(size);
            let length = text.chars().count() as f64;
            Bounds {
                left: x - size,
                top: y - 1.5 * size * sy,
                right: x + (length + 1.0) * size * sx,
                bottom: y + size * sy,
            }
        };
        let fill = match Mode::from_code(self.state.text.mode) {
            mode if mode.approximated => {
                self.note(
                    "pict-text-mode",
                    "text in an arithmetic or highlight mode is drawn in the foreground colour",
                );
                self.state.foreground
            },
            Mode {
                transfer: Transfer::Or,
                invert: false,
                ..
            } => self.state.foreground,
            Mode {
                transfer: Transfer::Bic,
                invert: false,
                ..
            } => self.state.background,
            // `srcCopy` also paints the background colour around the glyphs,
            // which leaves an untouched white destination unchanged.
            Mode {
                transfer: Transfer::Copy,
                invert: false,
                ..
            } if self.state.background == WHITE && self.untouched(area) => self.state.foreground,
            _ => {
                self.note("pict-text-mode", "text in a copy, inverting or exclusive-or mode is drawn without its background");
                self.state.foreground
            },
        };
        if !self.reserve_element() {
            return;
        }
        let state = &self.state.text;
        // Symbol-font text is decoded to Unicode, so it needs a font with
        // Greek and mathematical coverage rather than a Symbol font.
        let family = if symbol {
            SYMBOL_FAMILY.to_string()
        } else {
            font_family_for_name(&self.font_name(state.font))
        };
        let mut element = String::with_capacity(text.len() + 128);
        // A uniform text ratio scales the font size; an uneven one needs a
        // transform, and a clip path must then be applied by a group so
        // that the transform does not move it.
        let uneven = (sx - sy).abs() > 1e-9;
        let scale = if uneven { 1.0 } else { sx };
        let clip_group = uneven.then(|| self.clip.clone()).flatten();
        if let Some(id) = &clip_group {
            let _ = write!(element, r#"<g clip-path="url(#{id})">"#);
        }
        element.push_str("<text");
        if uneven {
            element.push_str(r#" transform="matrix("#);
            for (index, value) in [sx, 0.0, 0.0, sy, x, y].into_iter().enumerate() {
                if index > 0 {
                    element.push(' ');
                }
                write_num(&mut element, value);
            }
            element.push_str(r#")""#);
        } else {
            element.push_str(r#" x=""#);
            write_num(&mut element, x);
            element.push_str(r#"" y=""#);
            write_num(&mut element, y);
            element.push('"');
        }
        element.push_str(r#" font-size=""#);
        write_num(&mut element, f64::from(size) * scale);
        element.push_str(r#"" font-family=""#);
        write_xml_escaped(&mut element, &family);
        element.push('"');
        let face = state.face;
        if face & 0x01 != 0 {
            element.push_str(r#" font-weight="bold""#);
        }
        if face & 0x02 != 0 {
            element.push_str(r#" font-style="italic""#);
        }
        if face & 0x04 != 0 {
            element.push_str(r#" text-decoration="underline""#);
        }
        let mut letter_spacing = state.char_extra;
        if face & 0x20 != 0 {
            letter_spacing -= 1.0;
        }
        if face & 0x40 != 0 {
            letter_spacing += 1.0;
        }
        if letter_spacing != 0.0 {
            element.push_str(r#" letter-spacing=""#);
            write_num(&mut element, letter_spacing * scale);
            element.push('"');
        }
        if state.space_extra != 0.0 {
            element.push_str(r#" word-spacing=""#);
            write_num(&mut element, state.space_extra * scale);
            element.push('"');
        }
        if face & 0x18 != 0 {
            // Outline and shadow faces draw hollow glyphs.
            let _ = write!(
                element,
                r#" fill="none" stroke="{}" stroke-width="1""#,
                hex(fill)
            );
            if face & 0x10 != 0 {
                self.note(
                    "pict-text-shadow",
                    "shadowed text is drawn as outlined text",
                );
            }
        } else {
            let _ = write!(element, r#" fill="{}""#, hex(fill));
        }
        if let (false, Some(id)) = (uneven, &self.clip) {
            let _ = write!(element, r#" clip-path="url(#{id})""#);
        }
        element.push_str(r#" xml:space="preserve">"#);
        if fill != WHITE {
            self.mark_painted(area);
        }
        write_xml_escaped(&mut element, text);
        element.push_str("</text>");
        if clip_group.is_some() {
            element.push_str("</g>");
        }
        self.body.push_str(&element);
    }

    fn draw_bits(&mut self, opcode: u16, data: &[u8]) -> Result<()> {
        self.painted_bitmap = true;
        let Some((record, _)) = bitmap::parse_bits_record(data, opcode) else {
            self.flush_pending()?;
            self.note("pict-bitmap-skipped", "a bitmap record could not be read");
            return Ok(());
        };
        let mut decoded = match record.decode(self.pixel_limits()) {
            Ok(decoded) => decoded,
            Err(error) => {
                self.flush_pending()?;
                self.note(
                    "pict-bitmap-skipped",
                    format!("a bitmap is not drawn: {}", decode_error_text(&error)),
                );
                return Ok(());
            },
        };
        if decoded.short_rows {
            self.note(
                "pict-bitmap-short-rows",
                "packed bitmap rows end early; the missing pixels are drawn as zero",
            );
        }
        if decoded.overlong_rows {
            self.note("noncanonical-pict-bitmap-row-padding", "packed bitmap rows hold more data than the row width; QuickDraw ignores the excess");
        }
        // Crop the source rectangle to the pixel map and shrink the
        // destination in proportion.
        let bounds = record.pixmap.bounds;
        let src = record.src_rect;
        let dst = self.map_rect(record.dst_rect);
        let left = src.left.max(bounds.left);
        let top = src.top.max(bounds.top);
        let right = src.right.min(bounds.right);
        let bottom = src.bottom.min(bounds.bottom);
        if !src.is_nonempty() || dst.is_empty() || right <= left || bottom <= top {
            return Ok(());
        }
        let scale_x = dst.width() / f64::from(src.width());
        let scale_y = dst.height() / f64::from(src.height());
        let placed = Bounds {
            left: dst.left + f64::from(i32::from(left) - i32::from(src.left)) * scale_x,
            top: dst.top + f64::from(i32::from(top) - i32::from(src.top)) * scale_y,
            right: dst.left + f64::from(i32::from(right) - i32::from(src.left)) * scale_x,
            bottom: dst.top + f64::from(i32::from(bottom) - i32::from(src.top)) * scale_y,
        };
        let composite = self.apply_transfer(&mut decoded, &record.pixmap, record.mode, placed);
        let crop = (
            u32::try_from(i32::from(left) - i32::from(bounds.left)).unwrap_or(0),
            u32::try_from(i32::from(top) - i32::from(bounds.top)).unwrap_or(0),
            u32::try_from(i32::from(right) - i32::from(left)).unwrap_or(0),
            u32::try_from(i32::from(bottom) - i32::from(top)).unwrap_or(0),
        );
        let Some(image) = RgbaImage::from_raw(decoded.width, decoded.height, decoded.rgba) else {
            return Ok(());
        };
        let image = if crop == (0, 0, decoded.width, decoded.height) {
            image
        } else {
            image::imageops::crop_imm(&image, crop.0, crop.1, crop.2, crop.3).to_image()
        };
        self.note_bitmap_scale(f64::from(crop.2), f64::from(crop.3), placed);
        let mask = match record.mask_region {
            Some(mask) => match self.read_region(mask) {
                Some(region) if !region.is_rectangular() => Some(region),
                Some(_) => None,
                None => return self.flush_pending(),
            },
            None => None,
        };
        self.mark_painted(placed);
        let band = PendingImage {
            image,
            placed,
            composite,
            clip: self.clip.clone(),
        };
        if mask.is_none()
            && let Some(pending) = self.pending.as_mut()
            && pending.can_append(&band, self.limits)
        {
            pending.append(band);
            return Ok(());
        }
        self.flush_pending()?;
        if mask.is_some() {
            return self.emit_band(band, mask.as_ref());
        }
        self.pending = Some(band);
        Ok(())
    }

    /// Records the pixels per point of a placed image.
    fn note_bitmap_scale(&mut self, width: f64, height: f64, placed: Bounds) {
        let (x_points, y_points) = self.geometry.points_per_unit();
        let horizontal = width / (placed.width() * x_points);
        let vertical = height / (placed.height() * y_points);
        if horizontal.is_finite() && vertical.is_finite() {
            self.bitmap_scale = self.bitmap_scale.max(horizontal).max(vertical);
        }
    }

    fn flush_pending(&mut self) -> Result<()> {
        match self.pending.take() {
            Some(band) => self.emit_band(band, None),
            None => Ok(()),
        }
    }

    fn emit_band(&mut self, band: PendingImage, mask: Option<&Region>) -> Result<()> {
        let (width, height) = band.image.dimensions();
        let pixelated = band.placed.width() >= 2.0 * f64::from(width)
            || band.placed.height() >= 2.0 * f64::from(height);
        let Some(png) = encode_png(width, height, band.image.into_raw(), self.limits) else {
            self.note(
                "pict-bitmap-skipped",
                "a bitmap could not be encoded within the output limit",
            );
            return Ok(());
        };
        let clip = std::mem::replace(&mut self.clip, band.clip);
        let placement = Placement {
            bounds: band.placed,
            matrix: None,
            pixelated,
            composite: band.composite,
        };
        let result = self.emit_image("image/png", &png, &placement, mask);
        self.clip = clip;
        result
    }

    /// Applies a CopyBits transfer mode to decoded pixels and returns how
    /// they composite with the destination.
    ///
    /// One-bit sources are drawn in the port colours: set bits in the
    /// foreground and clear bits in the background colour (`srcCopy`), only
    /// set bits (`srcOr`), or set bits in the background colour (`srcBic`).
    /// Colour sources follow Color QuickDraw's colour rules: `srcCopy`
    /// colourises with the foreground and background colours; `srcOr` draws
    /// black source pixels in the foreground colour and `srcBic` draws them
    /// in the background colour, leaving white ones undrawn, which the
    /// darkening and lightening blends reproduce. Of the arithmetic modes,
    /// `blend` mixes by the `OpColor` weight, `addPin` and `subPin` add or
    /// subtract with the `OpColor` as limit, `adMax` and `adMin` keep the
    /// lighter or darker colour and `transparent` leaves source pixels of the
    /// background colour undrawn (*Inside Macintosh: Imaging With QuickDraw*,
    /// "Boolean Transfer Modes With Color Pixels" and "Arithmetic Transfer
    /// Modes").
    fn apply_transfer(
        &mut self,
        decoded: &mut DecodedPixels,
        pixmap: &bitmap::PixMap,
        code: u16,
        placed: Bounds,
    ) -> Composite {
        let code = code & !64;
        let foreground = self.state.foreground;
        let background = self.state.background;
        let mono = decoded
            .mono
            .take()
            .filter(|_| pixmap.is_bitmap || pixmap.palette.get(..2) == Some(&[WHITE, BLACK][..]));
        if code > 15 {
            if let Some(mono) = &mono {
                colorize_mono(
                    &mut decoded.rgba,
                    mono,
                    Some(foreground),
                    Some(background),
                    false,
                );
            }
            return self.arithmetic_transfer(decoded, code, placed);
        }
        let mode = Mode::from_code(code);
        if let Some(mono) = &mono {
            let (one, zero, composite) = match mode.transfer {
                Transfer::Copy => (Some(foreground), Some(background), Composite::Normal),
                Transfer::Or => (Some(foreground), None, Composite::Normal),
                Transfer::Bic => (Some(background), None, Composite::Normal),
                Transfer::Xor => {
                    self.note_xor();
                    (Some(WHITE), None, Composite::Blend("difference"))
                },
            };
            colorize_mono(&mut decoded.rgba, mono, one, zero, mode.invert);
            return composite;
        }
        if mode.invert {
            invert_rgb(&mut decoded.rgba);
        }
        let two_tone = is_two_tone(&decoded.rgba);
        match mode.transfer {
            Transfer::Copy => {
                if foreground != BLACK || background != WHITE {
                    for pixel in decoded.rgba.chunks_exact_mut(4) {
                        for ((component, fg), bk) in
                            pixel[..3].iter_mut().zip(foreground).zip(background)
                        {
                            let source = i32::from(*component);
                            let value =
                                i32::from(fg) + (i32::from(bk) - i32::from(fg)) * source / 255;
                            *component = u8::try_from(value.clamp(0, 255)).unwrap_or(0);
                        }
                    }
                }
                Composite::Normal
            },
            Transfer::Or => {
                if foreground != BLACK {
                    self.note(
                        "pict-colorize",
                        "a colour bitmap is not colourised by a non-black foreground",
                    );
                }
                if !two_tone && !self.untouched(placed) {
                    self.note_overlap();
                }
                Composite::Blend("multiply")
            },
            Transfer::Bic => {
                if background != WHITE {
                    self.note(
                        "pict-colorize",
                        "a colour bitmap is not colourised by a non-white background",
                    );
                }
                // A bit-clear pass followed by a `transparent` pass of the
                // same image paints over everything the first pass changed.
                if !two_tone && !self.untouched(placed) && !self.next_is_transparent_copy(placed) {
                    self.note_overlap();
                }
                invert_rgb(&mut decoded.rgba);
                Composite::Blend("screen")
            },
            Transfer::Xor => {
                self.note_xor();
                Composite::Blend("difference")
            },
        }
    }

    fn arithmetic_transfer(
        &mut self,
        decoded: &mut DecodedPixels,
        code: u16,
        placed: Bounds,
    ) -> Composite {
        let two_tone = is_two_tone(&decoded.rgba);
        let exact_over_white = two_tone || self.untouched(placed);
        let approximate = |renderer: &mut Self| {
            renderer.note(
                "pict-transfer-mode",
                format!("a bitmap in arithmetic transfer mode {code} is approximated"),
            );
        };
        match code {
            32 => match self.op_color {
                Some([red, green, blue]) if red == green && green == blue => {
                    Composite::Opacity(f64::from(red) / 255.0)
                },
                _ => {
                    approximate(self);
                    Composite::Opacity(0.5)
                },
            },
            33 => {
                if self.op_color != Some(WHITE) || !exact_over_white {
                    approximate(self);
                }
                Composite::Blend("lighten")
            },
            35 => {
                if self.op_color != Some(BLACK) || !exact_over_white {
                    approximate(self);
                }
                invert_rgb(&mut decoded.rgba);
                Composite::Blend("multiply")
            },
            36 => {
                let background = self.state.background;
                for pixel in decoded.rgba.chunks_exact_mut(4) {
                    if pixel[..3] == background {
                        pixel.copy_from_slice(&[0, 0, 0, 0]);
                    }
                }
                Composite::Normal
            },
            37 => Composite::Blend("lighten"),
            39 => Composite::Blend("darken"),
            _ => {
                approximate(self);
                Composite::Normal
            },
        }
    }

    fn note_overlap(&mut self) {
        self.note(
            "pict-transfer-mode",
            "a colour bitmap in an or or bit-clear mode over earlier drawing is composited by darkening or lightening",
        );
    }

    /// Whether the next record copies a bitmap to `placed` in the
    /// `transparent` mode.
    fn next_is_transparent_copy(&self, placed: Bounds) -> bool {
        let Some(record) = self.parser.records.get(self.next_record) else {
            return false;
        };
        if !matches!(record.opcode, 0x0090 | 0x0091 | 0x0098..=0x009B) {
            return false;
        }
        bitmap::parse_bits_record(&record.data, record.opcode).is_some_and(|(next, _)| {
            next.mode & !64 == 36 && self.map_rect(next.dst_rect) == placed
        })
    }

    fn emit_image(
        &mut self,
        mime: &str,
        bytes: &[u8],
        placement: &Placement,
        mask: Option<&Region>,
    ) -> Result<()> {
        let Placement {
            bounds: placed,
            matrix,
            pixelated,
            composite,
        } = *placement;
        let encoded_len = bytes.len().div_ceil(3).saturating_mul(4);
        if self
            .body
            .len()
            .saturating_add(self.defs.len())
            .saturating_add(encoded_len)
            > self.limits.max_output_bytes
        {
            return Err(Error::ParseError(format!(
                "PICT SVG output exceeds limit {} bytes",
                self.limits.max_output_bytes
            )));
        }
        if !self.reserve_element() {
            return Ok(());
        }
        let mut clip = self.clip.clone();
        if let Some(mask) = mask {
            let id = self.new_id("mask");
            let mut path = String::new();
            let origin = self.state.origin;
            mask.write_path(&mut path, |x, y| {
                (x - f64::from(origin.0), y - f64::from(origin.1))
            });
            let _ = write!(self.defs, r#"<clipPath id="{id}""#);
            if let Some(outer) = &clip {
                let _ = write!(self.defs, r#" clip-path="url(#{outer})""#);
            }
            let _ = write!(self.defs, r#"><path d="{path}"/></clipPath>"#);
            clip = Some(id);
        }
        if let (Some(_), Some(id)) = (matrix, &clip) {
            // The clip is in picture space; the image's own transform must
            // not move it.
            let _ = write!(self.body, r#"<g clip-path="url(#{id})">"#);
        }
        self.body.push_str("<image");
        if let Some(matrix) = matrix {
            self.body.push_str(r#" transform="matrix("#);
            for (index, value) in matrix.into_iter().enumerate() {
                if index > 0 {
                    self.body.push(' ');
                }
                let _ = write!(self.body, "{}", round6(value));
            }
            self.body.push_str(r#")""#);
        }
        for (name, value) in [
            ("x", placed.left),
            ("y", placed.top),
            ("width", placed.width()),
            ("height", placed.height()),
        ] {
            let _ = write!(self.body, r#" {name}=""#);
            write_num(&mut self.body, value);
            self.body.push('"');
        }
        self.body.push_str(r#" preserveAspectRatio="none""#);
        if pixelated {
            self.body.push_str(r#" image-rendering="optimizeSpeed""#);
        }
        if let (None, Some(id)) = (matrix, &clip) {
            let _ = write!(self.body, r#" clip-path="url(#{id})""#);
        }
        self.write_composite(composite);
        let _ = write!(self.body, r#" href="data:{mime};base64,"#);
        base64::engine::general_purpose::STANDARD.encode_string(bytes, &mut self.body);
        self.body.push_str(r#""/>"#);
        if matrix.is_some() && clip.is_some() {
            self.body.push_str("</g>");
        }
        Ok(())
    }

    fn draw_quicktime(&mut self, data: &[u8]) -> Result<bool> {
        self.painted_bitmap = true;
        let Some(image) = quicktime::parse_compressed(data) else {
            self.note(
                "pict-quicktime-malformed",
                "a QuickTime image record could not be read",
            );
            return Ok(false);
        };
        let codec = image.description.codec_name();
        let embedded = match quicktime::decode(&image, self.limits) {
            Decoded::Image(embedded) => embedded,
            Decoded::Unsupported => {
                self.note(
                    "pict-quicktime-codec",
                    format!("QuickTime codec '{codec}' is not decoded; the picture's fallback drawing is shown"),
                );
                return Ok(false);
            },
            Decoded::Failed(reason) => {
                self.note(
                    "pict-quicktime-undecodable",
                    format!("QuickTime '{codec}' image could not be decoded ({reason}); the picture's fallback drawing is shown"),
                );
                return Ok(false);
            },
        };
        for diagnostic in &embedded.diagnostics {
            self.note(diagnostic.code, diagnostic.message.clone());
        }
        if image.has_matte {
            self.note(
                "pict-quicktime-matte",
                "a QuickTime image matte is not applied",
            );
        }
        // Copy, optionally dithered, or the QuickTime graphics modes that
        // composite by the image's own straight alpha.
        if !matches!(image.mode & !0x40, 0 | 0x100 | 0x103) {
            self.note(
                "pict-transfer-mode",
                "a QuickTime image in a non-copy mode is drawn as a copy",
            );
        }
        let width = if image.description.width > 0 {
            f64::from(image.description.width)
        } else {
            f64::from(embedded.width)
        };
        let height = if image.description.height > 0 {
            f64::from(image.description.height)
        } else {
            f64::from(embedded.height)
        };
        let [a, b, c, d, tx, ty] = image.matrix;
        let origin = (
            f64::from(self.state.origin.0),
            f64::from(self.state.origin.1),
        );
        let matrix = [a, b, c, d, tx - origin.0, ty - origin.1];
        if let Some(destination) = quicktime::axis_aligned_destination(&image) {
            let placed = self.map_rect(destination);
            if !placed.is_empty() {
                self.note_bitmap_scale(
                    f64::from(embedded.width),
                    f64::from(embedded.height),
                    placed,
                );
            }
        }
        let mask = match image.mask_region {
            Some(mask) => match self.read_region(mask) {
                Some(region) if !region.is_rectangular() => Some(region),
                Some(_) => None,
                None => return Ok(true),
            },
            None => None,
        };
        let placed = Bounds {
            left: 0.0,
            top: 0.0,
            right: width,
            bottom: height,
        };
        let placement = Placement {
            bounds: placed,
            matrix: Some(matrix),
            pixelated: false,
            composite: Composite::Normal,
        };
        self.emit_image(embedded.mime, &embedded.bytes, &placement, mask.as_ref())?;
        let corners = [(0.0, 0.0), (width, 0.0), (0.0, height), (width, height)]
            .map(|(x, y)| (a * x + c * y + matrix[4], b * x + d * y + matrix[5]));
        self.mark_painted(points_bounds(&corners, (0.0, 0.0)));
        Ok(true)
    }

    fn finish(mut self) -> Result<RenderedPicture> {
        let view = self.geometry.view;
        let (document_width, document_height) = self.geometry.document;
        let bitmap_only = self.painted_bitmap && !self.painted_vector;
        // The natural raster size is the document size in points, enlarged
        // to the resolution of the finest image so that bitmaps keep their
        // pixels.
        let scale = self.bitmap_scale.clamp(1.0, 64.0);
        let natural = (document_width * scale, document_height * scale);
        let (width, height) = fit_dimensions(natural, self.limits);
        let mut svg = String::with_capacity(self.body.len() + self.defs.len() + 256);
        svg.push_str(r#"<svg xmlns="http://www.w3.org/2000/svg" width=""#);
        write_num(&mut svg, document_width.max(1.0));
        svg.push_str(r#"" height=""#);
        write_num(&mut svg, document_height.max(1.0));
        svg.push_str(r#"" viewBox=""#);
        for (index, value) in [view.left, view.top, view.width(), view.height()]
            .into_iter()
            .enumerate()
        {
            if index > 0 {
                svg.push(' ');
            }
            write_num(&mut svg, value);
        }
        svg.push_str(r#"" preserveAspectRatio="none">"#);
        if !self.defs.is_empty() {
            svg.push_str("<defs>");
            svg.push_str(&self.defs);
            svg.push_str("</defs>");
        }
        if self.uses_blend {
            // Blending needs a destination: QuickDraw ports start white.
            svg.push_str("<rect");
            for (name, value) in [
                ("x", view.left),
                ("y", view.top),
                ("width", view.width()),
                ("height", view.height()),
            ] {
                let _ = write!(svg, r#" {name}=""#);
                write_num(&mut svg, value);
                svg.push('"');
            }
            svg.push_str(r##" fill="#ffffff"/>"##);
        }
        svg.push_str(&self.body);
        svg.push_str("</svg>");
        if svg.len() > self.limits.max_output_bytes {
            return Err(Error::ParseError(format!(
                "PICT SVG output exceeds limit {} bytes",
                self.limits.max_output_bytes
            )));
        }
        Ok(RenderedPicture {
            svg,
            width,
            height,
            bitmap_only,
            diagnostics: std::mem::take(&mut self.diagnostics),
        })
    }
}

/// QuickTime follows a compressed image with drawing for readers without
/// QuickTime, announced by a `PnSize` record whose vertical size is `0xAE`
/// and whose horizontal size is the length in bytes of the fallback records
/// that follow it. Returns that length for a marker record.
fn fallback_length(record: &super::parser::PictRecord) -> Option<usize> {
    if record.opcode != 0x0007 || read_u16(&record.data, 0) != Some(0x00AE) {
        return None;
    }
    read_u16(&record.data, 2).map(usize::from)
}

fn round6(value: f64) -> f64 {
    (value * 1_000_000.0).round() / 1_000_000.0
}

/// Scales `size` down, preserving its aspect ratio, until it fits the
/// caller's dimension and pixel limits.
fn fit_dimensions(size: (f64, f64), limits: &crate::Limits) -> (u32, u32) {
    let width = size.0.max(1.0);
    let height = size.1.max(1.0);
    let mut scale = 1.0f64
        .min(f64::from(limits.max_width) / width)
        .min(f64::from(limits.max_height) / height);
    let pixels = width * height * scale * scale;
    if pixels > limits.max_pixels as f64 {
        scale *= (limits.max_pixels as f64 / pixels).sqrt();
    }
    let fit = |value: f64, maximum: u32, round: fn(f64) -> f64| {
        (round(value * scale).max(1.0) as u32).min(maximum.max(1))
    };
    let rounded = (
        fit(width, limits.max_width, f64::round),
        fit(height, limits.max_height, f64::round),
    );
    if u64::from(rounded.0) * u64::from(rounded.1) <= limits.max_pixels {
        return rounded;
    }
    let floored = (
        fit(width, limits.max_width, f64::floor),
        fit(height, limits.max_height, f64::floor),
    );
    let width = floored.0.min(
        u32::try_from(limits.max_pixels / u64::from(floored.1))
            .unwrap_or(u32::MAX)
            .max(1),
    );
    (width, floored.1)
}

/// Draws one-bit pixels: set bits (clear bits when `invert`) in `one`,
/// the others in `zero`; `None` leaves pixels transparent.
fn colorize_mono(
    rgba: &mut [u8],
    mono: &[bool],
    one: Option<Rgb>,
    zero: Option<Rgb>,
    invert: bool,
) {
    for (pixel, &bit) in rgba.chunks_exact_mut(4).zip(mono) {
        match if bit != invert { one } else { zero } {
            Some(color) => pixel.copy_from_slice(&[color[0], color[1], color[2], 255]),
            None => pixel.copy_from_slice(&[0, 0, 0, 0]),
        }
    }
}

fn invert_rgb(rgba: &mut [u8]) {
    for pixel in rgba.chunks_exact_mut(4) {
        for component in &mut pixel[..3] {
            *component = !*component;
        }
    }
}

/// Whether every opaque pixel is black or white.
fn is_two_tone(rgba: &[u8]) -> bool {
    rgba.chunks_exact(4)
        .all(|pixel| pixel[3] == 0 || pixel[..3] == BLACK || pixel[..3] == WHITE)
}

/// Bounds of points swept by a `pen.0` by `pen.1` pen.
fn points_bounds(points: &[(f64, f64)], pen: (f64, f64)) -> Bounds {
    let mut bounds = Bounds {
        left: f64::INFINITY,
        top: f64::INFINITY,
        right: f64::NEG_INFINITY,
        bottom: f64::NEG_INFINITY,
    };
    for &(x, y) in points {
        bounds.left = bounds.left.min(x);
        bounds.top = bounds.top.min(y);
        bounds.right = bounds.right.max(x + pen.0);
        bounds.bottom = bounds.bottom.max(y + pen.1);
    }
    bounds
}

fn pattern_bits(data: &[u8]) -> [u8; 8] {
    data.get(..8)
        .and_then(|bytes| bytes.try_into().ok())
        .unwrap_or(SOLID)
}

fn read_rgb(data: &[u8]) -> Option<Rgb> {
    Some([
        (read_u16(data, 0)? >> 8) as u8,
        (read_u16(data, 2)? >> 8) as u8,
        (read_u16(data, 4)? >> 8) as u8,
    ])
}

/// Reads polygon points: `polySize`, `polyBBox`, then `(v, h)` points.
fn read_polygon(data: &[u8]) -> Vec<(i32, i32)> {
    let size = read_u16(data, 0).map_or(0, usize::from).min(data.len());
    (10..size.saturating_sub(3))
        .step_by(4)
        .filter_map(|offset| {
            Some((
                i32::from(read_i16(data, offset + 2)?),
                i32::from(read_i16(data, offset)?),
            ))
        })
        .collect()
}

/// `FrameRgn`: the region minus the region inset by the pen, computed on the
/// region's pixel mask.
fn frame_region(region: &Region, pen: (i32, i32)) -> Option<Region> {
    let (pen_width, pen_height) = pen;
    if pen_width <= 0 || pen_height <= 0 || region.rects.is_empty() {
        return Some(Region {
            bounds: region.bounds,
            rects: Vec::new(),
        });
    }
    let left = region.rects.iter().map(|rect| rect[0]).min()?;
    let top = region.rects.iter().map(|rect| rect[1]).min()?;
    let right = region.rects.iter().map(|rect| rect[2]).max()?;
    let bottom = region.rects.iter().map(|rect| rect[3]).max()?;
    let width = usize::try_from(right - left).ok()?;
    let height = usize::try_from(bottom - top).ok()?;
    if (width as u64) * (height as u64) > MAX_FRAME_REGION_PIXELS {
        return None;
    }
    let mut mask = vec![false; width * height];
    for &[x0, y0, x1, y1] in &region.rects {
        for y in (y0 - top) as usize..(y1 - top) as usize {
            mask[y * width + (x0 - left) as usize..y * width + (x1 - left) as usize].fill(true);
        }
    }
    let inside = |x: isize, y: isize, mask: &[bool]| {
        x >= 0
            && y >= 0
            && (x as usize) < width
            && (y as usize) < height
            && mask[y as usize * width + x as usize]
    };
    let (dx, dy) = (pen_width as isize, pen_height as isize);
    let mut horizontal = vec![false; width * height];
    for y in 0..height as isize {
        for x in 0..width as isize {
            horizontal[y as usize * width + x as usize] =
                (-dx..=dx).all(|offset| inside(x + offset, y, &mask));
        }
    }
    let mut rects = Vec::new();
    for y in 0..height {
        let mut x = 0usize;
        while x < width {
            let edge = |x: usize| {
                mask[y * width + x]
                    && !(-dy..=dy)
                        .all(|offset| inside(x as isize, y as isize + offset, &horizontal))
            };
            if !edge(x) {
                x += 1;
                continue;
            }
            let start = x;
            while x < width && edge(x) {
                x += 1;
            }
            let row = top + y as i32;
            rects.push([left + start as i32, row, left + x as i32, row + 1]);
        }
    }
    Some(Region {
        bounds: region.bounds,
        rects,
    })
}

fn encode_png(width: u32, height: u32, rgba: Vec<u8>, limits: &crate::Limits) -> Option<Vec<u8>> {
    let image = RgbaImage::from_raw(width, height, rgba)?;
    crate::codec::encode(
        &DynamicImage::ImageRgba8(image),
        ImageFormat::Png,
        limits.max_output_bytes,
    )
    .ok()
}

fn decode_error_text(error: &DecodeError) -> String {
    match error {
        DecodeError::Unsupported(reason) => format!("unsupported pixel format ({reason})"),
        DecodeError::TooLarge(reason) => format!("exceeds the pixel limits ({reason})"),
    }
}

fn hex(color: Rgb) -> String {
    format!("#{:02x}{:02x}{:02x}", color[0], color[1], color[2])
}

/// Font families for text decoded from the Symbol font encoding.
const SYMBOL_FAMILY: &str = "'DejaVu Serif', 'Cambria Math', serif";

/// Names of the classic Macintosh font numbers.
fn classic_font_name(id: u16) -> &'static str {
    match id {
        0 => "Chicago",
        1 | 3 => "Geneva",
        2 => "New York",
        4 => "Monaco",
        5 => "Venice",
        6 => "London",
        7 => "Athens",
        8 => "San Francisco",
        9 => "Toronto",
        11 => "Cairo",
        12 => "Los Angeles",
        13 => "Zapf Dingbats",
        14 => "Bookman",
        15 => "Helvetica Narrow",
        16 => "Palatino",
        18 => "Zapf Chancery",
        20 => "Times",
        21 => "Helvetica",
        22 => "Courier",
        23 => "Symbol",
        33 => "Avant Garde",
        34 => "New Century Schlbk",
        _ => "",
    }
}

/// A CSS font-family list: the picture's font followed by a generic family
/// chosen from well-known Macintosh and PostScript font names.
fn font_family_for_name(name: &str) -> String {
    let lower = name.to_ascii_lowercase();
    let generic = if ["courier", "monaco", "mono", "consol"]
        .iter()
        .any(|key| lower.contains(key))
    {
        "monospace"
    } else if [
        "times",
        "new york",
        "palatino",
        "bookman",
        "schoolbook",
        "schlbk",
        "garamond",
        "georgia",
        "roman",
        "serif",
        "minion",
        "caslon",
        "baskerville",
        "century",
        "symbol",
    ]
    .iter()
    .any(|key| lower.contains(key))
        && !lower.contains("sans")
    {
        "serif"
    } else {
        "sans-serif"
    };
    let name = name.trim();
    if name.is_empty() {
        return generic.to_string();
    }
    let quoted: String = name
        .chars()
        .filter(|character| !matches!(character, '\'' | '"' | ',' | ';'))
        .collect();
    format!("'{quoted}', {generic}")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pict::bitmap::tests::push_rect;
    use crate::pict::parser::tests::version2_picture;
    use crate::pict::quicktime::tests::compressed_record;

    fn op(records: &mut Vec<u8>, opcode: u16, data: &[u8]) {
        records.extend_from_slice(&opcode.to_be_bytes());
        records.extend_from_slice(data);
        if data.len() % 2 == 1 {
            records.push(0);
        }
    }

    fn rect(rect: [i16; 4]) -> Vec<u8> {
        let mut data = Vec::new();
        push_rect(&mut data, rect);
        data
    }

    fn play(records: &[u8]) -> RenderedPicture {
        play_frame([0, 0, 100, 100], records)
    }

    fn play_frame(frame: [i16; 4], records: &[u8]) -> RenderedPicture {
        let mut records = records.to_vec();
        records.extend_from_slice(&0x00FFu16.to_be_bytes());
        let data = version2_picture(frame, &records);
        let parser = PictParser::new(&data, &crate::Limits::default()).unwrap();
        render(&parser, &crate::Limits::default()).unwrap()
    }

    fn codes(rendered: &RenderedPicture) -> Vec<&'static str> {
        rendered
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.code)
            .collect()
    }

    #[test]
    fn shapes_and_lines_become_filled_paths() {
        let mut records = Vec::new();
        op(&mut records, 0x0031, &rect([10, 10, 40, 60])); // paintRect
        op(&mut records, 0x0007, &[0, 2, 0, 2]); // PnSize 2x2
        op(&mut records, 0x0020, &[0, 0, 0, 0, 0, 0, 0, 50]); // Line (0,0)-(50,0)
        op(&mut records, 0x0030, &rect([50, 0, 60, 10])); // frameRect
        let rendered = play(&records);
        let svg = &rendered.svg;
        assert!(
            svg.contains(r##"<path d="M10 10H60V40H10Z" fill="#000000"/>"##),
            "{svg}"
        );
        assert!(
            svg.contains(r##"<path d="M0 0L52 0L52 2L0 2Z" fill="#000000"/>"##),
            "{svg}"
        );
        assert!(
            svg.contains(r##"d="M0 50H10V60H0ZM2 52H8V58H2Z" fill="#000000" fill-rule="evenodd""##),
            "{svg}"
        );
        assert!(rendered.diagnostics.is_empty());
        assert!(!rendered.bitmap_only);
        assert_eq!((rendered.width, rendered.height), (100, 100));
    }

    #[test]
    fn colours_and_one_bit_patterns_follow_the_port_state() {
        let mut records = Vec::new();
        op(&mut records, 0x001A, &[0xFF, 0xFF, 0, 0, 0, 0]); // RGBFgCol red
        op(
            &mut records,
            0x000A,
            &[0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55],
        ); // FillPat
        op(&mut records, 0x0034, &rect([0, 0, 10, 10])); // fillRect
        op(&mut records, 0x000F, &409u32.to_be_bytes()); // BkColor blue
        op(&mut records, 0x0032, &rect([20, 20, 30, 30])); // eraseRect
        let svg = play(&records).svg;
        assert!(svg.contains(r##"<pattern id="pict-pat0" width="8" height="8" patternUnits="userSpaceOnUse"><path d="M0 0H8V8H0Z" fill="#ffffff"/><path d="M0 0H1V1H0Z"##), "{svg}");
        assert!(svg.contains(r##"fill="#ff0000"/></pattern>"##), "{svg}");
        assert!(
            svg.contains(r#"<path d="M0 0H10V10H0Z" fill="url(#pict-pat0)"/>"#),
            "{svg}"
        );
        assert!(
            svg.contains(r##"<path d="M20 20H30V30H20Z" fill="#0000d4"/>"##),
            "{svg}"
        );
    }

    #[test]
    fn text_uses_font_state_and_mac_roman() {
        let mut records = Vec::new();
        op(
            &mut records,
            0x002C,
            &[
                0, 12, 0, 7, 9, b'H', b'e', b'l', b'v', b'e', b't', b'i', b'c', b'a',
            ],
        );
        op(&mut records, 0x0003, &7u16.to_be_bytes()); // TxFont
        op(&mut records, 0x0004, &[0x03]); // bold italic
        op(&mut records, 0x000D, &14u16.to_be_bytes()); // TxSize
        op(&mut records, 0x0028, &[0, 30, 0, 20, 3, b'A', 0xC9, b'B']); // LongText at (20, 30)
        op(&mut records, 0x0029, &[10, 1, b'C']); // DHText
        op(&mut records, 0x0003, &23u16.to_be_bytes()); // Symbol
        op(&mut records, 0x002A, &[12, 1, b'p']); // DVText
        let svg = play(&records).svg;
        assert!(svg.contains(r##"<text x="20" y="30" font-size="14" font-family="&apos;Helvetica&apos;, sans-serif" font-weight="bold" font-style="italic" fill="#000000" xml:space="preserve">A…B</text>"##), "{svg}");
        assert!(
            svg.contains(r#"<text x="30" y="30" font-size="14""#),
            "{svg}"
        );
        assert!(svg.contains(">π</text>"), "{svg}");
        assert!(svg.contains(r#"<text x="30" y="42""#), "{svg}");
    }

    #[test]
    fn text_ratios_scale_glyphs_without_moving_the_clip() {
        let mut records = Vec::new();
        op(&mut records, 0x0001, &[0, 10, 0, 0, 0, 0, 0, 50, 0, 50]); // Clip
        // TxRatio: numerator (v 2, h 2), denominator (v 1, h 1).
        op(&mut records, 0x0010, &[0, 2, 0, 2, 0, 1, 0, 1]);
        op(&mut records, 0x0028, &[0, 30, 0, 20, 1, b'A']);
        // An uneven ratio: horizontal 3, vertical 2.
        op(&mut records, 0x0010, &[0, 2, 0, 3, 0, 1, 0, 1]);
        op(&mut records, 0x0029, &[5, 1, b'B']);
        let svg = play(&records).svg;
        assert!(
            svg.contains(r#"<text x="20" y="30" font-size="24""#),
            "{svg}"
        );
        assert!(svg.contains(r#"<g clip-path="url(#pict-clip0)"><text transform="matrix(3 0 0 2 25 30)" font-size="12""#), "{svg}");
        assert!(svg.contains(">B</text></g>"), "{svg}");
    }

    #[test]
    fn clip_regions_apply_to_following_elements() {
        let mut records = Vec::new();
        op(&mut records, 0x0001, &[0, 10, 0, 0, 0, 0, 0, 50, 0, 50]); // Clip
        op(&mut records, 0x0031, &rect([0, 0, 100, 100]));
        op(&mut records, 0x0001, &[0, 10, 0, 0, 0, 0, 0, 100, 0, 100]); // whole frame
        op(&mut records, 0x0031, &rect([0, 0, 10, 10]));
        let svg = play(&records).svg;
        assert!(
            svg.contains(r#"<clipPath id="pict-clip0"><path d="M0 0H50V50H0Z"/></clipPath>"#),
            "{svg}"
        );
        assert!(
            svg.contains(
                r##"<path d="M0 0H100V100H0Z" fill="#000000" clip-path="url(#pict-clip0)"/>"##
            ),
            "{svg}"
        );
        assert!(
            svg.contains(r##"<path d="M0 0H10V10H0Z" fill="#000000"/>"##),
            "{svg}"
        );
    }

    #[test]
    fn origin_offsets_later_coordinates() {
        let mut records = Vec::new();
        op(&mut records, 0x000C, &[0, 5, 0xFF, 0xF6]); // Origin dh 5, dv -10
        op(&mut records, 0x0031, &rect([0, 10, 10, 20]));
        let svg = play(&records).svg;
        assert!(svg.contains(r#"d="M5 10H15V20H5Z""#), "{svg}");
    }

    #[test]
    fn src_or_one_bit_bitmaps_leave_clear_bits_transparent() {
        let mut bits = Vec::new();
        bits.extend_from_slice(&2u16.to_be_bytes()); // BitMap rowBytes
        for value in [[0i16, 0, 1, 2], [0, 0, 1, 2], [0, 0, 10, 20]] {
            push_rect(&mut bits, value);
        }
        bits.extend_from_slice(&1u16.to_be_bytes()); // srcOr
        bits.extend_from_slice(&[0b1000_0000, 0]);
        let mut records = Vec::new();
        op(&mut records, 0x0090, &bits);
        let rendered = play(&records);
        assert!(rendered.bitmap_only);
        assert!(rendered.diagnostics.is_empty());
        let start = rendered.svg.find("base64,").unwrap() + 7;
        let end = start + rendered.svg[start..].find('"').unwrap();
        let png = base64::engine::general_purpose::STANDARD
            .decode(&rendered.svg[start..end])
            .unwrap();
        let image = image::load_from_memory(&png).unwrap().to_rgba8();
        assert_eq!(image.get_pixel(0, 0).0, [0, 0, 0, 255]);
        assert_eq!(image.get_pixel(1, 0).0[3], 0);
        assert!(rendered.svg.contains(r#"<image x="0" y="0" width="20" height="10" preserveAspectRatio="none" image-rendering="optimizeSpeed""#));
        // The natural raster keeps one pixel per bitmap pixel.
        assert_eq!((rendered.width, rendered.height), (100, 100));
    }

    #[test]
    fn exclusive_or_and_invert_are_reported_approximations() {
        let mut records = Vec::new();
        op(&mut records, 0x0033, &rect([0, 0, 10, 10])); // invertRect
        let rendered = play(&records);
        assert_eq!(codes(&rendered), ["pict-xor-mode"]);
        assert!(
            rendered
                .svg
                .contains(r#"style="mix-blend-mode:difference""#)
        );
        assert!(
            rendered
                .svg
                .contains(r##"<rect x="0" y="0" width="100" height="100" fill="#ffffff"/>"##)
        );
    }

    #[test]
    fn decoded_quicktime_images_skip_their_fallback_drawing() {
        let mut jpeg = std::io::Cursor::new(Vec::new());
        DynamicImage::ImageRgb8(image::RgbImage::new(4, 2))
            .write_to(&mut jpeg, ImageFormat::Jpeg)
            .unwrap();
        let mut records = Vec::new();
        op(
            &mut records,
            0x8200,
            &compressed_record(b"jpeg", 4, 2, 24, (3, 5), &jpeg.into_inner()),
        );
        let mut fallback = Vec::new();
        op(&mut fallback, 0x0007, &[0, 1, 0, 1]);
        op(&mut fallback, 0x0028, &[0, 9, 0, 9, 3, b'Q', b'T', b'!']);
        op(&mut records, 0x0007, &[0, 0xAE, 0, fallback.len() as u8]);
        records.extend_from_slice(&fallback);
        op(&mut records, 0x0031, &rect([50, 50, 60, 60]));
        let rendered = play(&records);
        let svg = &rendered.svg;
        assert!(svg.contains(r#"<image transform="matrix(1 0 0 1 3 5)" x="0" y="0" width="4" height="2" preserveAspectRatio="none" href="data:image/jpeg;base64,"#), "{svg}");
        assert!(!svg.contains("<text"), "{svg}");
        assert!(svg.contains(r#"d="M50 50H60V60H50Z""#), "{svg}");
        assert!(rendered.diagnostics.is_empty());
    }

    #[test]
    fn undecoded_quicktime_codecs_draw_the_fallback_and_report() {
        let mut records = Vec::new();
        op(
            &mut records,
            0x8200,
            &compressed_record(b"cvid", 4, 2, 24, (0, 0), &[0; 8]),
        );
        op(&mut records, 0x0007, &[0, 0xAE, 0, 18]);
        op(&mut records, 0x0007, &[0, 1, 0, 1]);
        op(&mut records, 0x0028, &[0, 9, 0, 9, 3, b'Q', b'T', b'!']);
        let rendered = play(&records);
        assert_eq!(codes(&rendered), ["pict-quicktime-codec"]);
        assert!(rendered.svg.contains(">QT!</text>"));
    }

    #[test]
    fn native_resolution_sets_document_size_and_minimum_pen() {
        // A 300 dpi picture: 300 drawing units are one inch (72 points).
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_be_bytes());
        push_rect(&mut data, [0, 0, 72, 144]);
        data.extend_from_slice(&[0x00, 0x11, 0x02, 0xFF, 0x0C, 0x00, 0xFF, 0xFE, 0, 0]);
        data.extend_from_slice(&(300i32 << 16).to_be_bytes());
        data.extend_from_slice(&(300i32 << 16).to_be_bytes());
        push_rect(&mut data, [0, 0, 300, 600]);
        data.extend_from_slice(&[0; 4]);
        op(&mut data, 0x0020, &[0, 0, 0, 0, 0, 0, 1, 0x2C]); // Line (0,0)-(300,0)
        data.extend_from_slice(&0x00FFu16.to_be_bytes());
        let parser = PictParser::new(&data, &crate::Limits::default()).unwrap();
        let rendered = render(&parser, &crate::Limits::default()).unwrap();
        assert!(rendered.svg.starts_with(r#"<svg xmlns="http://www.w3.org/2000/svg" width="144" height="72" viewBox="0 0 600 300""#), "{}", rendered.svg);
        // The one-unit pen is drawn one point (300/72 units) wide.
        assert!(
            rendered
                .svg
                .contains(r#"d="M0 0L304.17 0L304.17 4.17L0 4.17Z""#),
            "{}",
            rendered.svg
        );
        assert_eq!((rendered.width, rendered.height), (144, 72));
    }

    #[test]
    fn element_limit_stops_drawing_with_a_diagnostic() {
        let mut records = Vec::new();
        for _ in 0..5 {
            op(&mut records, 0x0031, &rect([0, 0, 10, 10]));
        }
        records.extend_from_slice(&0x00FFu16.to_be_bytes());
        let data = version2_picture([0, 0, 100, 100], &records);
        let limits = crate::Limits {
            max_svg_elements: 3,
            ..crate::Limits::default()
        };
        let parser = PictParser::new(&data, &limits).unwrap();
        let rendered = render(&parser, &limits).unwrap();
        assert_eq!(rendered.svg.matches("<path").count(), 3);
        assert_eq!(codes(&rendered), ["pict-svg-element-limit"]);
    }
}
