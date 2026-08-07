//! Bounded semantic playback for framed EMF+ records.
//!
//! This layer intentionally does not interpret graphics object payloads.  It
//! assembles them into a 64-slot table and emits object references in drawing
//! commands, leaving brush, pen, path, image, and font decoding to a renderer.

use litchi_core::error::{Error, Result};

use super::{EmfPlusRecord, ObjectId, ObjectType, ParserLimits, RecordType};

const SOLID_COLOR_FLAG: u16 = 0x8000;
const COMPRESSED_FLAG: u16 = 0x4000;
const RELATIVE_FLAG: u16 = 0x0800;
const POST_MULTIPLY_FLAG: u16 = 0x2000;

/// A two-dimensional point in EMF+ world coordinates.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct Point {
    pub x: f32,
    pub y: f32,
}

/// A rectangle in EMF+ world coordinates.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct Rect {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
}

/// An affine matrix using the EMF+/GDI+ layout `[m11, m12, m21, m22, dx, dy]`.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Matrix {
    pub m11: f32,
    pub m12: f32,
    pub m21: f32,
    pub m22: f32,
    pub dx: f32,
    pub dy: f32,
}

impl Matrix {
    pub const IDENTITY: Self = Self {
        m11: 1.0,
        m12: 0.0,
        m21: 0.0,
        m22: 1.0,
        dx: 0.0,
        dy: 0.0,
    };

    #[must_use]
    pub fn multiply(self, rhs: Self) -> Self {
        Self {
            m11: self.m11 * rhs.m11 + self.m12 * rhs.m21,
            m12: self.m11 * rhs.m12 + self.m12 * rhs.m22,
            m21: self.m21 * rhs.m11 + self.m22 * rhs.m21,
            m22: self.m21 * rhs.m12 + self.m22 * rhs.m22,
            dx: self.dx * rhs.m11 + self.dy * rhs.m21 + rhs.dx,
            dy: self.dx * rhs.m12 + self.dy * rhs.m22 + rhs.dy,
        }
    }

    #[must_use]
    pub fn transform_point(self, point: Point) -> Point {
        Point {
            x: point.x * self.m11 + point.y * self.m21 + self.dx,
            y: point.x * self.m12 + point.y * self.m22 + self.dy,
        }
    }
}

/// Unit used by a page transform.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Unit {
    World,
    Display,
    Pixel,
    Point,
    Inch,
    Document,
    Millimeter,
}

impl Unit {
    fn parse(value: u8) -> Option<Self> {
        Some(match value {
            0 => Self::World,
            1 => Self::Display,
            2 => Self::Pixel,
            3 => Self::Point,
            4 => Self::Inch,
            5 => Self::Document,
            6 => Self::Millimeter,
            _ => return None,
        })
    }
}

/// EMF+ page-space conversion parameters.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PageTransform {
    pub unit: Unit,
    pub scale: f32,
}
impl Default for PageTransform {
    fn default() -> Self {
        Self {
            unit: Unit::Display,
            scale: 1.0,
        }
    }
}

/// A region-combination operation.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum CombineMode {
    Replace,
    Intersect,
    Union,
    Xor,
    Exclude,
    Complement,
}
impl CombineMode {
    fn parse(value: u8) -> Option<Self> {
        Some(match value {
            0 => Self::Replace,
            1 => Self::Intersect,
            2 => Self::Union,
            3 => Self::Xor,
            4 => Self::Exclude,
            5 => Self::Complement,
            _ => return None,
        })
    }
}

/// A clip expression.  Path and region values refer to the object table.
#[derive(Clone, Debug, PartialEq)]
pub enum Clip {
    Infinite,
    Rect { mode: CombineMode, rect: Rect },
    Path { mode: CombineMode, path: ObjectId },
    Region { mode: CombineMode, region: ObjectId },
    Offset { clip: Box<Self>, dx: f32, dy: f32 },
}

/// Source-over or source-copy compositing.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum CompositingMode {
    SourceOver,
    SourceCopy,
}
/// Rendering modes are stored as their valid EMF+ numeric values.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct RenderingSettings {
    pub smoothing_mode: u8,
    pub text_rendering_hint: u8,
    pub text_contrast: u8,
    pub interpolation_mode: u8,
    pub pixel_offset_mode: u8,
    pub compositing_mode: CompositingMode,
    pub compositing_quality: u8,
    pub rendering_origin: (i32, i32),
}
impl Default for RenderingSettings {
    fn default() -> Self {
        Self {
            smoothing_mode: 0,
            text_rendering_hint: 0,
            text_contrast: 4,
            interpolation_mode: 3,
            pixel_offset_mode: 0,
            compositing_mode: CompositingMode::SourceOver,
            compositing_quality: 1,
            rendering_origin: (0, 0),
        }
    }
}

/// Complete graphics state captured on every output command.
#[derive(Clone, Debug, PartialEq)]
pub struct GraphicsState {
    pub world_transform: Matrix,
    pub page_transform: PageTransform,
    pub clip: Clip,
    pub rendering: RenderingSettings,
}
impl Default for GraphicsState {
    fn default() -> Self {
        Self {
            world_transform: Matrix::IDENTITY,
            page_transform: PageTransform::default(),
            clip: Clip::Infinite,
            rendering: RenderingSettings::default(),
        }
    }
}

/// A brush reference, either an object-table entry or a solid ARGB color.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Brush {
    Object(ObjectId),
    Solid(u32),
}

/// A fully assembled object definition.  Its bytes deliberately remain opaque.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ObjectBlob {
    pub id: ObjectId,
    pub object_type: ObjectType,
    pub bytes: Vec<u8>,
}

/// Structured draw operation independent of object decoding.
#[derive(Clone, Debug, PartialEq)]
pub enum DrawCommandKind {
    Clear {
        color: u32,
    },
    FillRects {
        brush: Brush,
        rects: Vec<Rect>,
    },
    DrawRects {
        pen: ObjectId,
        rects: Vec<Rect>,
    },
    FillPolygon {
        brush: Brush,
        points: Vec<Point>,
        fill_mode: u8,
    },
    DrawLines {
        pen: ObjectId,
        points: Vec<Point>,
    },
    FillEllipse {
        brush: Brush,
        rect: Rect,
    },
    DrawEllipse {
        pen: ObjectId,
        rect: Rect,
    },
    FillPie {
        brush: Brush,
        rect: Rect,
        start_angle: f32,
        sweep_angle: f32,
    },
    DrawPie {
        pen: ObjectId,
        rect: Rect,
        start_angle: f32,
        sweep_angle: f32,
    },
    DrawArc {
        pen: ObjectId,
        rect: Rect,
        start_angle: f32,
        sweep_angle: f32,
    },
    FillRegion {
        brush: Brush,
        region: ObjectId,
    },
    FillPath {
        brush: Brush,
        path: ObjectId,
    },
    DrawPath {
        pen: ObjectId,
        path: ObjectId,
    },
    FillClosedCurve {
        brush: Brush,
        points: Vec<Point>,
        tension: f32,
        fill_mode: u8,
    },
    DrawClosedCurve {
        pen: ObjectId,
        points: Vec<Point>,
        tension: f32,
    },
    DrawCurve {
        pen: ObjectId,
        points: Vec<Point>,
        tension: f32,
        offset: u32,
        segments: u32,
    },
    DrawBeziers {
        pen: ObjectId,
        points: Vec<Point>,
    },
    DrawImage {
        image: ObjectId,
        attributes: Option<ObjectId>,
        dest: Rect,
        src: Rect,
        src_unit: Unit,
    },
    DrawImagePoints {
        image: ObjectId,
        attributes: Option<ObjectId>,
        points: Vec<Point>,
        src: Rect,
        src_unit: Unit,
    },
    DrawString {
        brush: Brush,
        font: ObjectId,
        format: Option<ObjectId>,
        text: String,
        layout: Rect,
    },
    DrawDriverString {
        brush: Brush,
        font: ObjectId,
        options: u32,
        glyphs: Vec<u16>,
        positions: Vec<Point>,
        transform: Option<Matrix>,
    },
    StrokeFillPath {
        pen: ObjectId,
        brush: Brush,
        path: ObjectId,
    },
}

/// One renderer-facing operation with the graphics state active at its record.
#[derive(Clone, Debug, PartialEq)]
pub struct DrawCommand {
    pub offset: usize,
    pub state: GraphicsState,
    pub kind: DrawCommandKind,
}

/// Diagnostic importance.  Playback continues after warning/error diagnostics.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum PlaybackDiagnosticSeverity {
    Info,
    Warning,
    Error,
}
/// A non-silent semantic problem, tied to the input record's byte offset.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct PlaybackDiagnostic {
    pub severity: PlaybackDiagnosticSeverity,
    pub code: &'static str,
    pub message: String,
    pub record_offset: usize,
}

/// Independent resource ceilings for semantic playback.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct PlaybackLimits {
    pub max_points: usize,
    pub max_output: usize,
    pub max_depth: usize,
    pub max_objects: usize,
    pub max_bytes: usize,
}
impl PlaybackLimits {
    pub fn validate(self) -> Result<Self> {
        if self.max_points == 0
            || self.max_output == 0
            || self.max_depth == 0
            || self.max_objects == 0
            || self.max_objects > 64
            || self.max_bytes == 0
        {
            return Err(Error::ParseError(
                "EMF+ playback limits must be non-zero; max_objects is 1..=64".into(),
            ));
        }
        Ok(self)
    }
}
impl Default for PlaybackLimits {
    fn default() -> Self {
        Self {
            max_points: 1_000_000,
            max_output: 1_000_000,
            max_depth: 1024,
            max_objects: 64,
            max_bytes: 64 * 1024 * 1024,
        }
    }
}

#[derive(Clone, Debug)]
struct SavedState {
    index: u32,
    state: GraphicsState,
}
#[derive(Clone, Debug)]
struct PendingObject {
    id: ObjectId,
    object_type: ObjectType,
    total: usize,
    bytes: Vec<u8>,
}

/// Stateful EMF+ playback engine.  Feed it records in stream order.
#[derive(Debug)]
pub struct PlaybackEngine {
    limits: PlaybackLimits,
    parser_limits: ParserLimits,
    state: GraphicsState,
    saves: Vec<SavedState>,
    containers: Vec<SavedState>,
    objects: [Option<ObjectBlob>; 64],
    pending: Option<PendingObject>,
    diagnostics: Vec<PlaybackDiagnostic>,
    output_count: usize,
    object_bytes: usize,
    ended: bool,
}

impl PlaybackEngine {
    pub fn new(limits: PlaybackLimits) -> Result<Self> {
        let limits = limits.validate()?;
        let parser_limits = ParserLimits {
            max_bytes: limits.max_bytes,
            max_records: usize::MAX,
            max_object_slots: limits.max_objects,
        }
        .validate()?;
        Ok(Self {
            limits,
            parser_limits,
            state: GraphicsState::default(),
            saves: Vec::new(),
            containers: Vec::new(),
            objects: std::array::from_fn(|_| None),
            pending: None,
            diagnostics: Vec::new(),
            output_count: 0,
            object_bytes: 0,
            ended: false,
        })
    }

    #[must_use]
    pub fn state(&self) -> &GraphicsState {
        &self.state
    }
    #[must_use]
    pub fn diagnostics(&self) -> &[PlaybackDiagnostic] {
        &self.diagnostics
    }
    #[must_use]
    pub fn object(&self, id: ObjectId) -> Option<&ObjectBlob> {
        self.objects[usize::from(id.get())].as_ref()
    }
    #[must_use]
    pub fn is_finished(&self) -> bool {
        self.ended
    }

    /// Consume one framed record and return zero or more renderer-facing draws.
    pub fn push(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        if self.ended {
            return self.fail(
                record.offset,
                "record_after_eof",
                "record appeared after EndOfFile",
            );
        }
        if self.pending.is_some() && record.header.record_type != RecordType::Object {
            self.diag(
                PlaybackDiagnosticSeverity::Error,
                "object_interrupted",
                "continued object was interrupted",
                record.offset,
            );
            self.pending = None;
        }
        let result = match record.header.record_type {
            RecordType::Header => self.header(record),
            RecordType::EndOfFile => self.eof(record),
            RecordType::Object => self.object_record(record),
            RecordType::Save => self.save(record),
            RecordType::Restore => self.restore(record),
            RecordType::BeginContainer => self.begin_container(record, true),
            RecordType::BeginContainerNoParams => self.begin_container(record, false),
            RecordType::EndContainer => self.end_container(record),
            RecordType::SetWorldTransform => self.set_matrix(record),
            RecordType::ResetWorldTransform => {
                self.require(record, 0)?;
                self.state.world_transform = Matrix::IDENTITY;
                Ok(Vec::new())
            },
            RecordType::MultiplyWorldTransform => self.multiply_matrix(record),
            RecordType::TranslateWorldTransform => self.translate(record),
            RecordType::ScaleWorldTransform => self.scale(record),
            RecordType::RotateWorldTransform => self.rotate(record),
            RecordType::SetPageTransform => self.page_transform(record),
            RecordType::ResetClip => {
                self.require(record, 0)?;
                self.state.clip = Clip::Infinite;
                Ok(Vec::new())
            },
            RecordType::SetClipRect => self.clip_rect(record),
            RecordType::SetClipPath => self.clip_object(record, true),
            RecordType::SetClipRegion => self.clip_object(record, false),
            RecordType::OffsetClip => self.offset_clip(record),
            RecordType::SetRenderingOrigin
            | RecordType::SetAntiAliasMode
            | RecordType::SetTextRenderingHint
            | RecordType::SetTextContrast
            | RecordType::SetInterpolationMode
            | RecordType::SetPixelOffsetMode
            | RecordType::SetCompositingMode
            | RecordType::SetCompositingQuality => self.rendering(record),
            RecordType::Clear
            | RecordType::FillRects
            | RecordType::DrawRects
            | RecordType::FillPolygon
            | RecordType::DrawLines
            | RecordType::FillEllipse
            | RecordType::DrawEllipse
            | RecordType::FillPie
            | RecordType::DrawPie
            | RecordType::DrawArc
            | RecordType::FillRegion
            | RecordType::FillPath
            | RecordType::DrawPath
            | RecordType::FillClosedCurve
            | RecordType::DrawClosedCurve
            | RecordType::DrawCurve
            | RecordType::DrawBeziers
            | RecordType::DrawImage
            | RecordType::DrawImagePoints
            | RecordType::DrawString
            | RecordType::DrawDriverString
            | RecordType::StrokeFillPath => self.draw(record),
            RecordType::MultiFormatStart
            | RecordType::MultiFormatSection
            | RecordType::MultiFormatEnd => {
                self.diag(
                    PlaybackDiagnosticSeverity::Error,
                    "reserved_record",
                    "reserved MultiFormat record encountered",
                    record.offset,
                );
                Ok(Vec::new())
            },
            RecordType::GetDc => {
                self.diag(
                    PlaybackDiagnosticSeverity::Warning,
                    "get_dc",
                    "GetDc delegates drawing to the underlying EMF device context",
                    record.offset,
                );
                Ok(Vec::new())
            },
            RecordType::Comment => {
                self.diag(
                    PlaybackDiagnosticSeverity::Info,
                    "comment",
                    "EMF+ comment record ignored",
                    record.offset,
                );
                Ok(Vec::new())
            },
            RecordType::SerializableObject | RecordType::SetTsGraphics | RecordType::SetTsClip => {
                self.diag(
                    PlaybackDiagnosticSeverity::Warning,
                    "unsupported_control",
                    format!("{:?} is not supported", record.header.record_type),
                    record.offset,
                );
                Ok(Vec::new())
            },
        };
        if let Err(error) = &result {
            self.diag(
                PlaybackDiagnosticSeverity::Error,
                "semantic_decode_failed",
                error.to_string(),
                record.offset,
            );
        }
        result
    }

    /// Ensure no incomplete continuation remains at stream end.
    pub fn finish(&mut self) -> Result<()> {
        if self.pending.is_some() {
            return Err(Error::ParseError(
                "unterminated EMF+ object continuation".into(),
            ));
        }
        if !self.ended {
            return Err(Error::ParseError(
                "EMF+ playback stream lacks EndOfFile".into(),
            ));
        }
        Ok(())
    }

    fn header(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        self.require(record, 16)?;
        Ok(Vec::new())
    }
    fn eof(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        self.require(record, 0)?;
        self.ended = true;
        Ok(Vec::new())
    }
    fn save(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        let index = u32_at(record.data, 0)?;
        self.require(record, 4)?;
        self.push_saved_to(true, index, record.offset)?;
        Ok(Vec::new())
    }
    fn restore(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        let index = u32_at(record.data, 0)?;
        self.require(record, 4)?;
        self.restore_saved_from(true, index, record.offset);
        Ok(Vec::new())
    }
    fn begin_container(
        &mut self,
        record: EmfPlusRecord<'_>,
        parameters: bool,
    ) -> Result<Vec<DrawCommand>> {
        let index = if parameters {
            self.require(record, 36)?;
            let dest = rect_at(record.data, 0)?;
            let src = rect_at(record.data, 16)?;
            let index = u32_at(record.data, 32)?;
            self.push_saved_to(false, index, record.offset)?;
            self.state.world_transform = self.state.world_transform.multiply(rect_map(dest, src));
            return Ok(Vec::new());
        } else {
            self.require(record, 4)?;
            u32_at(record.data, 0)?
        };
        self.push_saved_to(false, index, record.offset)?;
        Ok(Vec::new())
    }
    fn end_container(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        let index = u32_at(record.data, 0)?;
        self.require(record, 4)?;
        self.restore_saved_from(false, index, record.offset);
        Ok(Vec::new())
    }
    fn set_matrix(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        self.require(record, 24)?;
        self.state.world_transform = matrix_at(record.data, 0)?;
        Ok(Vec::new())
    }
    fn multiply_matrix(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        self.require(record, 24)?;
        self.apply_matrix(matrix_at(record.data, 0)?, record.header.flags.raw());
        Ok(Vec::new())
    }
    fn translate(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        self.require(record, 8)?;
        self.apply_matrix(
            Matrix {
                dx: f32_at(record.data, 0)?,
                dy: f32_at(record.data, 4)?,
                ..Matrix::IDENTITY
            },
            record.header.flags.raw(),
        );
        Ok(Vec::new())
    }
    fn scale(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        self.require(record, 8)?;
        self.apply_matrix(
            Matrix {
                m11: f32_at(record.data, 0)?,
                m22: f32_at(record.data, 4)?,
                ..Matrix::IDENTITY
            },
            record.header.flags.raw(),
        );
        Ok(Vec::new())
    }
    fn rotate(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        self.require(record, 4)?;
        let radians = f32_at(record.data, 0)?.to_radians();
        self.apply_matrix(
            Matrix {
                m11: radians.cos(),
                m12: radians.sin(),
                m21: -radians.sin(),
                m22: radians.cos(),
                dx: 0.0,
                dy: 0.0,
            },
            record.header.flags.raw(),
        );
        Ok(Vec::new())
    }
    fn page_transform(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        self.require(record, 4)?;
        let unit = Unit::parse(record.header.flags.raw().to_le_bytes()[0])
            .ok_or_else(|| Error::ParseError("invalid EMF+ page unit".into()))?;
        self.state.page_transform = PageTransform {
            unit,
            scale: f32_at(record.data, 0)?,
        };
        Ok(Vec::new())
    }
    fn clip_rect(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        self.require(record, 16)?;
        let mode = combine(record.header.flags.raw(), record.offset)?;
        self.state.clip = Clip::Rect {
            mode,
            rect: rect_at(record.data, 0)?,
        };
        Ok(Vec::new())
    }
    fn clip_object(&mut self, record: EmfPlusRecord<'_>, path: bool) -> Result<Vec<DrawCommand>> {
        self.require(record, 0)?;
        let mode = combine(record.header.flags.raw(), record.offset)?;
        let id = self.object_id(record)?;
        self.state.clip = if path {
            Clip::Path { mode, path: id }
        } else {
            Clip::Region { mode, region: id }
        };
        Ok(Vec::new())
    }
    fn offset_clip(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        self.require(record, 8)?;
        self.state.clip = Clip::Offset {
            clip: Box::new(self.state.clip.clone()),
            dx: f32_at(record.data, 0)?,
            dy: f32_at(record.data, 4)?,
        };
        Ok(Vec::new())
    }
    fn apply_matrix(&mut self, matrix: Matrix, flags: u16) {
        self.state.world_transform = if flags & POST_MULTIPLY_FLAG != 0 {
            self.state.world_transform.multiply(matrix)
        } else {
            matrix.multiply(self.state.world_transform)
        };
    }

    fn rendering(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        match record.header.record_type {
            RecordType::SetRenderingOrigin => {
                self.require(record, 8)?;
                self.state.rendering.rendering_origin =
                    (i32_at(record.data, 0)?, i32_at(record.data, 4)?);
            },
            RecordType::SetTextContrast => {
                self.require(record, 0)?;
                self.state.rendering.text_contrast = record.header.flags.raw().to_le_bytes()[0];
            },
            RecordType::SetAntiAliasMode => {
                self.require(record, 0)?;
                self.state.rendering.smoothing_mode = record.header.flags.raw().to_le_bytes()[0];
            },
            RecordType::SetTextRenderingHint => {
                self.require(record, 0)?;
                self.state.rendering.text_rendering_hint =
                    record.header.flags.raw().to_le_bytes()[0];
            },
            RecordType::SetInterpolationMode => {
                self.require(record, 0)?;
                self.state.rendering.interpolation_mode =
                    record.header.flags.raw().to_le_bytes()[0];
            },
            RecordType::SetPixelOffsetMode => {
                self.require(record, 0)?;
                self.state.rendering.pixel_offset_mode = record.header.flags.raw().to_le_bytes()[0];
            },
            RecordType::SetCompositingQuality => {
                self.require(record, 0)?;
                self.state.rendering.compositing_quality =
                    record.header.flags.raw().to_le_bytes()[0];
            },
            RecordType::SetCompositingMode => {
                self.require(record, 0)?;
                self.state.rendering.compositing_mode =
                    match record.header.flags.raw().to_le_bytes()[0] {
                        0 => CompositingMode::SourceOver,
                        1 => CompositingMode::SourceCopy,
                        _ => {
                            return self.fail(
                                record.offset,
                                "invalid_compositing_mode",
                                "invalid compositing mode",
                            )?;
                        },
                    };
            },
            RecordType::Header
            | RecordType::EndOfFile
            | RecordType::Comment
            | RecordType::GetDc
            | RecordType::MultiFormatStart
            | RecordType::MultiFormatSection
            | RecordType::MultiFormatEnd
            | RecordType::Object
            | RecordType::Clear
            | RecordType::FillRects
            | RecordType::DrawRects
            | RecordType::FillPolygon
            | RecordType::DrawLines
            | RecordType::FillEllipse
            | RecordType::DrawEllipse
            | RecordType::FillPie
            | RecordType::DrawPie
            | RecordType::DrawArc
            | RecordType::FillRegion
            | RecordType::FillPath
            | RecordType::DrawPath
            | RecordType::FillClosedCurve
            | RecordType::DrawClosedCurve
            | RecordType::DrawCurve
            | RecordType::DrawBeziers
            | RecordType::DrawImage
            | RecordType::DrawImagePoints
            | RecordType::DrawString
            | RecordType::Save
            | RecordType::Restore
            | RecordType::BeginContainer
            | RecordType::BeginContainerNoParams
            | RecordType::EndContainer
            | RecordType::SetWorldTransform
            | RecordType::ResetWorldTransform
            | RecordType::MultiplyWorldTransform
            | RecordType::TranslateWorldTransform
            | RecordType::ScaleWorldTransform
            | RecordType::RotateWorldTransform
            | RecordType::SetPageTransform
            | RecordType::ResetClip
            | RecordType::SetClipRect
            | RecordType::SetClipPath
            | RecordType::SetClipRegion
            | RecordType::OffsetClip
            | RecordType::DrawDriverString
            | RecordType::StrokeFillPath
            | RecordType::SerializableObject
            | RecordType::SetTsGraphics
            | RecordType::SetTsClip => {
                return self.fail(
                    record.offset,
                    "invalid_rendering_record",
                    "invalid rendering record",
                );
            },
        }
        if record.header.record_type != RecordType::SetCompositingMode {
            self.diag(
                PlaybackDiagnosticSeverity::Warning,
                "rendering_property_not_represented",
                format!(
                    "{:?} is retained in playback state but has no exact SVG representation",
                    record.header.record_type
                ),
                record.offset,
            );
        }
        Ok(Vec::new())
    }

    fn object_record(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        let fragment = record.object_fragment(self.parser_limits)?;
        let fragment_len = fragment.data.len();
        if let Some(pending) = self.pending.as_mut() {
            if pending.id != fragment.flags.object_id
                || pending.object_type != fragment.flags.object_type
            {
                return self.fail(
                    record.offset,
                    "object_continuation_mismatch",
                    "object continuation has a different id or type",
                );
            }
            if fragment
                .total_object_size
                .is_some_and(|value| value as usize != pending.total)
            {
                return self.fail(
                    record.offset,
                    "object_size_mismatch",
                    "object continuation TotalObjectSize changed",
                );
            }
            append_limited(
                &mut pending.bytes,
                fragment.data,
                pending.total,
                "EMF+ object bytes",
            )?;
            if pending.bytes.len() == pending.total {
                let pending = self.pending.take().ok_or_else(|| {
                    Error::ParseError("missing EMF+ pending object continuation".into())
                })?;
                self.store_object(pending, record.offset)?;
            } else if !fragment.flags.continued {
                return self.fail(
                    record.offset,
                    "object_size_mismatch",
                    "object continuation ended before TotalObjectSize",
                );
            }
        } else if fragment.flags.continued {
            let total = usize::try_from(fragment.total_object_size.unwrap_or(0))
                .map_err(|_| Error::ParseError("object size does not fit usize".into()))?;
            if total < fragment_len || total > self.limits.max_bytes {
                return self.fail(
                    record.offset,
                    "invalid_object_size",
                    "continued object size is invalid or exceeds byte limit",
                );
            }
            let mut bytes = Vec::new();
            bytes
                .try_reserve(fragment_len)
                .map_err(|e| Error::Allocation {
                    resource: "EMF+ object bytes",
                    source: e,
                })?;
            bytes.extend_from_slice(fragment.data);
            let pending = PendingObject {
                id: fragment.flags.object_id,
                object_type: fragment.flags.object_type,
                total,
                bytes,
            };
            if pending.bytes.len() == pending.total {
                self.store_object(pending, record.offset)?;
            } else {
                self.pending = Some(pending);
            }
        } else {
            self.store_object(
                PendingObject {
                    id: fragment.flags.object_id,
                    object_type: fragment.flags.object_type,
                    total: fragment_len,
                    bytes: copy_bytes(fragment.data, "EMF+ object bytes")?,
                },
                record.offset,
            )?;
        }
        Ok(Vec::new())
    }
    fn store_object(&mut self, object: PendingObject, offset: usize) -> Result<()> {
        let new_total = self
            .object_bytes
            .checked_add(object.bytes.len())
            .ok_or_else(|| Error::ParseError("EMF+ object byte count overflow".into()))?;
        if new_total > self.limits.max_bytes {
            let message = "assembled object bytes exceed playback byte limit";
            self.diag(
                PlaybackDiagnosticSeverity::Error,
                "object_bytes_limit",
                message,
                offset,
            );
            return Err(Error::ParseError(message.into()));
        }
        self.object_bytes = new_total;
        self.objects[usize::from(object.id.get())] = Some(ObjectBlob {
            id: object.id,
            object_type: object.object_type,
            bytes: object.bytes,
        });
        Ok(())
    }

    fn draw(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        let flags = record.header.flags.raw();
        let data = record.data;
        let kind = match record.header.record_type {
            RecordType::Clear => {
                self.require(record, 4)?;
                DrawCommandKind::Clear {
                    color: u32_at(data, 0)?,
                }
            },
            RecordType::FillRects => DrawCommandKind::FillRects {
                brush: self.brush(record, 0)?,
                rects: self.rects(record, 4)?,
            },
            RecordType::DrawRects => DrawCommandKind::DrawRects {
                pen: self.object_id(record)?,
                rects: self.rects(record, 0)?,
            },
            RecordType::FillPolygon => DrawCommandKind::FillPolygon {
                brush: self.brush(record, 0)?,
                points: self.points(record, 4)?,
                fill_mode: u8::from(((flags >> 8) & 1) != 0),
            },
            RecordType::DrawLines => DrawCommandKind::DrawLines {
                pen: self.object_id(record)?,
                points: self.points(record, 0)?,
            },
            RecordType::FillEllipse => DrawCommandKind::FillEllipse {
                brush: self.brush(record, 0)?,
                rect: rect_at(data, 4)?,
            },
            RecordType::DrawEllipse => {
                self.require(record, 16)?;
                DrawCommandKind::DrawEllipse {
                    pen: self.object_id(record)?,
                    rect: rect_at(data, 0)?,
                }
            },
            RecordType::FillPie => {
                self.require(record, 28)?;
                DrawCommandKind::FillPie {
                    brush: self.brush(record, 0)?,
                    start_angle: f32_at(data, 4)?,
                    sweep_angle: f32_at(data, 8)?,
                    rect: rect_at(data, 12)?,
                }
            },
            RecordType::DrawPie | RecordType::DrawArc => {
                self.require(record, 24)?;
                let pen = self.object_id(record)?;
                let start_angle = f32_at(data, 0)?;
                let sweep_angle = f32_at(data, 4)?;
                let rect = rect_at(data, 8)?;
                if record.header.record_type == RecordType::DrawPie {
                    DrawCommandKind::DrawPie {
                        pen,
                        rect,
                        start_angle,
                        sweep_angle,
                    }
                } else {
                    DrawCommandKind::DrawArc {
                        pen,
                        rect,
                        start_angle,
                        sweep_angle,
                    }
                }
            },
            RecordType::FillRegion => {
                self.require(record, 4)?;
                DrawCommandKind::FillRegion {
                    brush: self.brush(record, 0)?,
                    region: self.object_id(record)?,
                }
            },
            RecordType::FillPath => {
                self.require(record, 4)?;
                DrawCommandKind::FillPath {
                    brush: self.brush(record, 0)?,
                    path: self.object_id(record)?,
                }
            },
            RecordType::DrawPath => {
                self.require(record, 4)?;
                DrawCommandKind::DrawPath {
                    pen: self.object_id_from(data, 0)?,
                    path: self.object_id(record)?,
                }
            },
            RecordType::FillClosedCurve => {
                let points = self.points(record, 8)?;
                DrawCommandKind::FillClosedCurve {
                    brush: self.brush(record, 0)?,
                    tension: f32_at(data, 4)?,
                    points,
                    fill_mode: u8::from(((flags >> 8) & 1) != 0),
                }
            },
            RecordType::DrawClosedCurve => {
                let points = self.points(record, 4)?;
                DrawCommandKind::DrawClosedCurve {
                    pen: self.object_id(record)?,
                    tension: f32_at(data, 0)?,
                    points,
                }
            },
            RecordType::DrawCurve => {
                let points = self.points(record, 12)?;
                DrawCommandKind::DrawCurve {
                    pen: self.object_id(record)?,
                    tension: f32_at(data, 0)?,
                    offset: u32_at(data, 4)?,
                    segments: u32_at(data, 8)?,
                    points,
                }
            },
            RecordType::DrawBeziers => DrawCommandKind::DrawBeziers {
                pen: self.object_id(record)?,
                points: self.points(record, 0)?,
            },
            RecordType::DrawImage | RecordType::DrawImagePoints => return self.draw_image(record),
            RecordType::DrawString => return self.draw_string(record),
            RecordType::DrawDriverString => return self.draw_driver_string(record),
            RecordType::StrokeFillPath => {
                self.diag(
                    PlaybackDiagnosticSeverity::Warning,
                    "unsupported_stroke_fill_path",
                    "StrokeFillPath requires renderer-specific interpretation",
                    record.offset,
                );
                return Ok(Vec::new());
            },
            RecordType::Header
            | RecordType::EndOfFile
            | RecordType::Comment
            | RecordType::GetDc
            | RecordType::MultiFormatStart
            | RecordType::MultiFormatSection
            | RecordType::MultiFormatEnd
            | RecordType::Object
            | RecordType::SetRenderingOrigin
            | RecordType::SetAntiAliasMode
            | RecordType::SetTextRenderingHint
            | RecordType::SetTextContrast
            | RecordType::SetInterpolationMode
            | RecordType::SetPixelOffsetMode
            | RecordType::SetCompositingMode
            | RecordType::SetCompositingQuality
            | RecordType::Save
            | RecordType::Restore
            | RecordType::BeginContainer
            | RecordType::BeginContainerNoParams
            | RecordType::EndContainer
            | RecordType::SetWorldTransform
            | RecordType::ResetWorldTransform
            | RecordType::MultiplyWorldTransform
            | RecordType::TranslateWorldTransform
            | RecordType::ScaleWorldTransform
            | RecordType::RotateWorldTransform
            | RecordType::SetPageTransform
            | RecordType::ResetClip
            | RecordType::SetClipRect
            | RecordType::SetClipPath
            | RecordType::SetClipRegion
            | RecordType::OffsetClip
            | RecordType::SerializableObject
            | RecordType::SetTsGraphics
            | RecordType::SetTsClip => {
                return self.fail(record.offset, "invalid_draw_record", "invalid draw record");
            },
        };
        self.emit(record.offset, kind)
    }

    fn draw_image(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        let d = record.data;
        self.require_at(
            record,
            0,
            if record.header.record_type == RecordType::DrawImage {
                40
            } else {
                28
            },
        )?;
        let image = self.object_id(record)?;
        let attributes = object_opt(self, u32_at(d, 0)?)?;
        let src_unit_raw = u8::try_from(u32_at(d, 4)?)
            .map_err(|_| Error::ParseError("image source unit exceeds u8".into()))?;
        let src_unit = Unit::parse(src_unit_raw)
            .ok_or_else(|| Error::ParseError("invalid image source unit".into()))?;
        let src = rect_at(d, 8)?;
        let kind = if record.header.record_type == RecordType::DrawImage {
            DrawCommandKind::DrawImage {
                image,
                attributes,
                src,
                src_unit,
                dest: rect_at(d, 24)?,
            }
        } else {
            let points = points_from(
                d,
                28,
                self.limits.max_points,
                record.offset,
                record.header.flags.raw(),
            )?;
            DrawCommandKind::DrawImagePoints {
                image,
                attributes,
                src,
                src_unit,
                points,
            }
        };
        self.emit(record.offset, kind)
    }
    fn draw_string(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        self.require_at(record, 0, 28)?;
        let d = record.data;
        let length = usize::try_from(u32_at(d, 8)?)
            .map_err(|_| Error::ParseError("string length does not fit usize".into()))?;
        let bytes = length
            .checked_mul(2)
            .ok_or_else(|| Error::ParseError("string length overflow".into()))?;
        self.require_at(record, 28, bytes)?;
        let mut chars = Vec::new();
        chars.try_reserve(length).map_err(|e| Error::Allocation {
            resource: "EMF+ string",
            source: e,
        })?;
        for chunk in d[28..28 + bytes].chunks_exact(2) {
            chars.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
        let text = String::from_utf16_lossy(&chars);
        self.emit(
            record.offset,
            DrawCommandKind::DrawString {
                brush: self.brush(record, 0)?,
                format: object_opt(self, u32_at(d, 4)?)?,
                font: self.object_id(record)?,
                layout: rect_at(d, 12)?,
                text,
            },
        )
    }
    fn draw_driver_string(&mut self, record: EmfPlusRecord<'_>) -> Result<Vec<DrawCommand>> {
        self.require_at(record, 0, 16)?;
        let d = record.data;
        let count = checked_count(u32_at(d, 12)?, self.limits.max_points, record.offset)?;
        let glyph_bytes = count
            .checked_mul(2)
            .ok_or_else(|| Error::ParseError("glyph byte count overflow".into()))?;
        let positions_start = 16usize
            .checked_add(glyph_bytes)
            .ok_or_else(|| Error::ParseError("driver string offset overflow".into()))?;
        self.require_at(
            record,
            positions_start,
            count
                .checked_mul(8)
                .ok_or_else(|| Error::ParseError("driver string position count overflow".into()))?,
        )?;
        let mut glyphs = Vec::new();
        glyphs.try_reserve(count).map_err(|e| Error::Allocation {
            resource: "EMF+ glyphs",
            source: e,
        })?;
        for part in d[16..positions_start].chunks_exact(2) {
            glyphs.push(u16::from_le_bytes([part[0], part[1]]));
        }
        let positions = fixed_points(d, positions_start, count)?;
        let positions_bytes = count.checked_mul(8).ok_or_else(|| {
            Error::ParseError("driver string position byte count overflow".into())
        })?;
        let matrix_start = positions_start
            .checked_add(positions_bytes)
            .ok_or_else(|| Error::ParseError("driver string matrix offset overflow".into()))?;
        let transform = if u32_at(d, 8)? == 0 {
            None
        } else {
            self.require_at(record, matrix_start, 24)?;
            Some(matrix_at(d, matrix_start)?)
        };
        self.emit(
            record.offset,
            DrawCommandKind::DrawDriverString {
                brush: self.brush(record, 0)?,
                font: self.object_id(record)?,
                options: u32_at(d, 4)?,
                glyphs,
                positions,
                transform,
            },
        )
    }

    fn emit(&mut self, offset: usize, kind: DrawCommandKind) -> Result<Vec<DrawCommand>> {
        self.output_count = self
            .output_count
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("EMF+ output count overflow".into()))?;
        if self.output_count > self.limits.max_output {
            return self.fail(offset, "output_limit", "playback output limit exceeded");
        }
        let mut output = Vec::new();
        output.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "EMF+ draw commands",
            source,
        })?;
        output.push(DrawCommand {
            offset,
            state: self.state.clone(),
            kind,
        });
        Ok(output)
    }
    fn rects(&self, record: EmfPlusRecord<'_>, start: usize) -> Result<Vec<Rect>> {
        self.require_at(record, start, 4)?;
        let count = checked_count(
            u32_at(record.data, start)?,
            self.limits.max_points,
            record.offset,
        )?;
        let point_start = start
            .checked_add(4)
            .ok_or_else(|| Error::ParseError("rectangle offset overflow".into()))?;
        let compressed = record.header.flags.raw() & COMPRESSED_FLAG != 0;
        let item_size = if compressed { 8 } else { 16 };
        let bytes = count
            .checked_mul(item_size)
            .ok_or_else(|| Error::ParseError("rectangle byte count overflow".into()))?;
        self.require_at(record, point_start, bytes)?;
        let mut result = Vec::new();
        result.try_reserve(count).map_err(|e| Error::Allocation {
            resource: "EMF+ rectangles",
            source: e,
        })?;
        for i in 0..count {
            let at = point_start + i * item_size;
            result.push(if compressed {
                rect_s_at(record.data, at)?
            } else {
                rect_at(record.data, at)?
            });
        }
        Ok(result)
    }
    fn points(&self, record: EmfPlusRecord<'_>, start: usize) -> Result<Vec<Point>> {
        self.require_at(record, start, 4)?;
        points_from(
            record.data,
            start + 4,
            self.limits.max_points,
            record.offset,
            record.header.flags.raw(),
        )
    }
    fn brush(&self, record: EmfPlusRecord<'_>, start: usize) -> Result<Brush> {
        self.brush_from(record.data, start, record.header.flags.raw())
    }
    fn brush_from(&self, data: &[u8], start: usize, flags: u16) -> Result<Brush> {
        if flags & SOLID_COLOR_FLAG != 0 {
            Ok(Brush::Solid(u32_at(data, start)?))
        } else {
            Ok(Brush::Object(self.object_id_from(data, start)?))
        }
    }
    fn object_id(&self, record: EmfPlusRecord<'_>) -> Result<ObjectId> {
        record.header.flags.object_id(self.parser_limits)
    }
    fn object_id_from(&self, data: &[u8], start: usize) -> Result<ObjectId> {
        let value = u32_at(data, start)?;
        let id = u8::try_from(value)
            .map_err(|_| Error::ParseError("EMF+ object id exceeds u8".into()))?;
        ObjectId::new(id, self.limits.max_objects)
    }
    fn push_saved_to(&mut self, save: bool, index: u32, offset: usize) -> Result<()> {
        let at_depth_limit = if save {
            self.saves.len()
        } else {
            self.containers.len()
        } >= self.limits.max_depth;
        if at_depth_limit {
            let message = "graphics state stack depth limit exceeded";
            self.diag(
                PlaybackDiagnosticSeverity::Error,
                "state_depth_limit",
                message,
                offset,
            );
            return Err(Error::ParseError(message.into()));
        }
        let stack = if save {
            &mut self.saves
        } else {
            &mut self.containers
        };
        stack.try_reserve(1).map_err(|e| Error::Allocation {
            resource: "EMF+ graphics state stack",
            source: e,
        })?;
        stack.push(SavedState {
            index,
            state: self.state.clone(),
        });
        Ok(())
    }
    fn restore_saved_from(&mut self, save: bool, index: u32, offset: usize) {
        let stack = if save {
            &mut self.saves
        } else {
            &mut self.containers
        };
        if let Some(position) = stack.iter().rposition(|saved| saved.index == index) {
            let saved = stack.remove(position);
            self.state = saved.state;
            stack.truncate(position);
        } else {
            self.diag(
                PlaybackDiagnosticSeverity::Error,
                "unknown_state_index",
                format!("no graphics state exists for index {index}"),
                offset,
            );
        }
    }
    #[allow(
        clippy::unused_self,
        reason = "kept as an engine invariant helper at every call site"
    )]
    fn require(&self, record: EmfPlusRecord<'_>, length: usize) -> Result<()> {
        if record.data.len() != length {
            return Err(Error::ParseError(format!(
                "EMF+ {:?} at offset {} has DataSize {}, expected {length}",
                record.header.record_type,
                record.offset,
                record.data.len()
            )));
        }
        Ok(())
    }
    #[allow(
        clippy::unused_self,
        reason = "kept as an engine invariant helper at every call site"
    )]
    fn require_at(&self, record: EmfPlusRecord<'_>, start: usize, length: usize) -> Result<()> {
        let end = start
            .checked_add(length)
            .ok_or_else(|| Error::ParseError("EMF+ data range overflow".into()))?;
        if end > record.data.len() {
            return Err(Error::ParseError(format!(
                "truncated EMF+ {:?} body at offset {}",
                record.header.record_type, record.offset
            )));
        }
        Ok(())
    }
    fn diag(
        &mut self,
        severity: PlaybackDiagnosticSeverity,
        code: &'static str,
        message: impl Into<String>,
        record_offset: usize,
    ) {
        self.diagnostics.push(PlaybackDiagnostic {
            severity,
            code,
            message: message.into(),
            record_offset,
        });
    }
    fn fail<T>(
        &mut self,
        offset: usize,
        code: &'static str,
        message: impl Into<String>,
    ) -> Result<T> {
        let message = message.into();
        self.diag(
            PlaybackDiagnosticSeverity::Error,
            code,
            message.clone(),
            offset,
        );
        Err(Error::ParseError(message))
    }
}

fn combine(flags: u16, _offset: usize) -> Result<CombineMode> {
    CombineMode::parse(flags.to_le_bytes()[1] & 0x0f)
        .ok_or_else(|| Error::ParseError("invalid EMF+ combine mode".into()))
}
fn object_opt(engine: &PlaybackEngine, value: u32) -> Result<Option<ObjectId>> {
    if value == u32::MAX {
        Ok(None)
    } else {
        let id = u8::try_from(value)
            .map_err(|_| Error::ParseError("EMF+ object id exceeds u8".into()))?;
        Ok(Some(ObjectId::new(id, engine.limits.max_objects)?))
    }
}
fn checked_count(value: u32, max: usize, offset: usize) -> Result<usize> {
    let count = usize::try_from(value)
        .map_err(|_| Error::ParseError("EMF+ count does not fit usize".into()))?;
    if count > max {
        return Err(Error::ParseError(format!(
            "EMF+ point count {count} exceeds limit {max} at offset {offset}"
        )));
    }
    Ok(count)
}
fn append_limited(
    target: &mut Vec<u8>,
    source: &[u8],
    total: usize,
    resource: &'static str,
) -> Result<()> {
    let end = target
        .len()
        .checked_add(source.len())
        .ok_or_else(|| Error::ParseError("EMF+ object length overflow".into()))?;
    if end > total {
        return Err(Error::ParseError(
            "EMF+ object continuation exceeds TotalObjectSize".into(),
        ));
    }
    target
        .try_reserve(source.len())
        .map_err(|e| Error::Allocation {
            resource,
            source: e,
        })?;
    target.extend_from_slice(source);
    Ok(())
}
fn copy_bytes(source: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve(source.len())
        .map_err(|error| Error::Allocation {
            resource,
            source: error,
        })?;
    output.extend_from_slice(source);
    Ok(output)
}
fn u32_at(data: &[u8], at: usize) -> Result<u32> {
    let bytes = data
        .get(
            at..at
                .checked_add(4)
                .ok_or_else(|| Error::ParseError("EMF+ offset overflow".into()))?,
        )
        .ok_or_else(|| Error::ParseError("truncated EMF+ u32".into()))?;
    Ok(u32::from_le_bytes(bytes.try_into().map_err(|_| {
        Error::ParseError("invalid EMF+ u32 range".into())
    })?))
}
fn i32_at(data: &[u8], at: usize) -> Result<i32> {
    Ok(i32::from_le_bytes(u32_at(data, at)?.to_le_bytes()))
}
fn f32_at(data: &[u8], at: usize) -> Result<f32> {
    Ok(f32::from_bits(u32_at(data, at)?))
}
fn rect_at(data: &[u8], at: usize) -> Result<Rect> {
    Ok(Rect {
        x: f32_at(data, at)?,
        y: f32_at(data, at + 4)?,
        width: f32_at(data, at + 8)?,
        height: f32_at(data, at + 12)?,
    })
}
fn i16_at(data: &[u8], at: usize) -> Result<i16> {
    let end = at
        .checked_add(2)
        .ok_or_else(|| Error::ParseError("EMF+ offset overflow".into()))?;
    let bytes = data
        .get(at..end)
        .ok_or_else(|| Error::ParseError("truncated EMF+ i16".into()))?;
    Ok(i16::from_le_bytes([bytes[0], bytes[1]]))
}
fn rect_s_at(data: &[u8], at: usize) -> Result<Rect> {
    Ok(Rect {
        x: f32::from(i16_at(data, at)?),
        y: f32::from(i16_at(data, at + 2)?),
        width: f32::from(i16_at(data, at + 4)?),
        height: f32::from(i16_at(data, at + 6)?),
    })
}
fn matrix_at(data: &[u8], at: usize) -> Result<Matrix> {
    Ok(Matrix {
        m11: f32_at(data, at)?,
        m12: f32_at(data, at + 4)?,
        m21: f32_at(data, at + 8)?,
        m22: f32_at(data, at + 12)?,
        dx: f32_at(data, at + 16)?,
        dy: f32_at(data, at + 20)?,
    })
}
fn fixed_points(data: &[u8], at: usize, count: usize) -> Result<Vec<Point>> {
    let bytes = count
        .checked_mul(8)
        .ok_or_else(|| Error::ParseError("point byte count overflow".into()))?;
    if at
        .checked_add(bytes)
        .ok_or_else(|| Error::ParseError("point range overflow".into()))?
        > data.len()
    {
        return Err(Error::ParseError("truncated EMF+ point array".into()));
    }
    let mut result = Vec::new();
    result.try_reserve(count).map_err(|e| Error::Allocation {
        resource: "EMF+ points",
        source: e,
    })?;
    for i in 0..count {
        result.push(Point {
            x: f32_at(data, at + i * 8)?,
            y: f32_at(data, at + i * 8 + 4)?,
        });
    }
    Ok(result)
}
fn short_points(data: &[u8], at: usize, count: usize) -> Result<Vec<Point>> {
    let bytes = count
        .checked_mul(4)
        .ok_or_else(|| Error::ParseError("short point byte count overflow".into()))?;
    let end = at
        .checked_add(bytes)
        .ok_or_else(|| Error::ParseError("short point range overflow".into()))?;
    if end > data.len() {
        return Err(Error::ParseError("truncated EMF+ PointS array".into()));
    }
    let mut result = Vec::new();
    result
        .try_reserve(count)
        .map_err(|source| Error::Allocation {
            resource: "EMF+ points",
            source,
        })?;
    for index in 0..count {
        let point = at + index * 4;
        result.push(Point {
            x: f32::from(i16_at(data, point)?),
            y: f32::from(i16_at(data, point + 2)?),
        });
    }
    Ok(result)
}
fn relative_points(data: &[u8], at: usize, count: usize) -> Result<Vec<Point>> {
    let mut cursor = at;
    let mut current = Point::default();
    let mut result = Vec::new();
    result
        .try_reserve(count)
        .map_err(|source| Error::Allocation {
            resource: "EMF+ points",
            source,
        })?;
    for _ in 0..count {
        current.x += f32::from(integer_r(data, &mut cursor)?);
        current.y += f32::from(integer_r(data, &mut cursor)?);
        result.push(current);
    }
    Ok(result)
}
fn integer_r(data: &[u8], cursor: &mut usize) -> Result<i16> {
    let first = *data
        .get(*cursor)
        .ok_or_else(|| Error::ParseError("truncated EMF+ PointR coordinate".into()))?;
    if first & 1 == 0 {
        *cursor = cursor
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("EMF+ PointR offset overflow".into()))?;
        return Ok(i16::from(first.cast_signed() >> 1));
    }
    let end = cursor
        .checked_add(2)
        .ok_or_else(|| Error::ParseError("EMF+ PointR offset overflow".into()))?;
    let pair = data
        .get(*cursor..end)
        .ok_or_else(|| Error::ParseError("truncated EMF+ PointR coordinate".into()))?;
    *cursor = end;
    Ok((i16::from_le_bytes([pair[0], pair[1]])) >> 1)
}
fn points_from(
    data: &[u8],
    at: usize,
    max: usize,
    offset: usize,
    flags: u16,
) -> Result<Vec<Point>> {
    let count = checked_count(u32_at(data, at - 4)?, max, offset)?;
    if flags & RELATIVE_FLAG != 0 {
        return relative_points(data, at, count);
    }
    if flags & COMPRESSED_FLAG != 0 {
        return short_points(data, at, count);
    }
    fixed_points(data, at, count)
}
fn rect_map(dest: Rect, src: Rect) -> Matrix {
    let sx = if dest.width == 0.0 {
        1.0
    } else {
        src.width / dest.width
    };
    let sy = if dest.height == 0.0 {
        1.0
    } else {
        src.height / dest.height
    };
    Matrix {
        m11: sx,
        m12: 0.0,
        m21: 0.0,
        m22: sy,
        dx: src.x - dest.x * sx,
        dy: src.y - dest.y * sy,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::emfplus::{EmfPlusRecordIter, ParserLimits};

    fn record(kind: RecordType, flags: u16, data: &[u8]) -> Vec<u8> {
        let size = 12 + data.len();
        let mut out = Vec::new();
        out.extend_from_slice(&kind.raw().to_le_bytes());
        out.extend_from_slice(&flags.to_le_bytes());
        out.extend_from_slice(&(size as u32).to_le_bytes());
        out.extend_from_slice(&(data.len() as u32).to_le_bytes());
        out.extend_from_slice(data);
        out
    }
    fn one(bytes: &[u8]) -> EmfPlusRecord<'_> {
        EmfPlusRecordIter::new(bytes, ParserLimits::default())
            .unwrap()
            .next()
            .unwrap()
            .unwrap()
    }
    #[test]
    fn matrix_and_clear_are_emitted() {
        let mut engine = PlaybackEngine::new(PlaybackLimits::default()).unwrap();
        let mut matrix = Vec::new();
        for value in [2.0_f32, 0.0, 0.0, 3.0, 4.0, 5.0] {
            matrix.extend_from_slice(&value.to_le_bytes());
        }
        engine
            .push(one(&record(RecordType::SetWorldTransform, 0, &matrix)))
            .unwrap();
        let commands = engine
            .push(one(&record(
                RecordType::Clear,
                0,
                &0xff00_1122u32.to_le_bytes(),
            )))
            .unwrap();
        assert_eq!(commands.len(), 1);
        assert_eq!(commands[0].state.world_transform.dx, 4.);
        assert!(matches!(
            commands[0].kind,
            DrawCommandKind::Clear { color: 0xff00_1122 }
        ));
    }
    #[test]
    fn object_continuation_is_assembled() {
        let mut engine = PlaybackEngine::new(PlaybackLimits::default()).unwrap();
        let mut first = 8u32.to_le_bytes().to_vec();
        first.extend_from_slice(&[1, 2, 3, 4]);
        engine
            .push(one(&record(RecordType::Object, 0x8103, &first)))
            .unwrap();
        engine
            .push(one(&record(RecordType::Object, 0x0103, &[5, 6, 7, 8])))
            .unwrap();
        let id = ObjectId::new(3, 64).unwrap();
        assert_eq!(
            engine.object(id).unwrap().bytes,
            vec![1, 2, 3, 4, 5, 6, 7, 8]
        );
    }
    #[test]
    fn bad_container_is_diagnostic() {
        let mut engine = PlaybackEngine::new(PlaybackLimits::default()).unwrap();
        engine
            .push(one(&record(
                RecordType::EndContainer,
                0,
                &7u32.to_le_bytes(),
            )))
            .unwrap();
        assert_eq!(engine.diagnostics()[0].code, "unknown_state_index");
    }
}
