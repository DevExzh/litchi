//! Ordered EMF+ comment playback into safe SVG.
//!
//! The stateful renderer is intended to sit beside classic EMF playback.  It
//! reports the EMF+/classic multiplexing state after every comment without
//! accepting raw SVG or executing driver payloads.

use litchi_core::error::{Error, Result};

use super::objects::{
    self, Argb, BrushKind, DecodeLimits, GraphicsObject, ImageKind, Path as ObjectPath, RegionNode,
};
use super::playback::{
    Brush, Clip, ClipSpace, CombineMode, CompositingMode, DrawCommand, DrawCommandKind,
    GraphicsState, Matrix, PlaybackEngine, PlaybackLimits, Point, Rect, Unit,
};
use super::svg::{
    SvgBuilder, SvgColor, SvgCompositingMode, SvgDocument, SvgGradientStop, SvgId, SvgImage,
    SvgImageMime, SvgImageSource, SvgLineCap, SvgLineJoin, SvgLinearGradient, SvgPaint, SvgPath,
    SvgPathCommand, SvgPoint, SvgRect, SvgSpread, SvgStroke, SvgStyle, SvgText, SvgTransform,
};
use super::{
    EmfPlusRecord, EmfPlusRecordIter, EmfPlusStreamValidator, ObjectId, ParserLimits, RecordType,
    try_extract_emfplus_comment_body,
};

/// All independent ceilings used by parsing, object decoding, playback, and SVG emission.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct RendererLimits {
    pub parser: ParserLimits,
    pub playback: PlaybackLimits,
    pub objects: DecodeLimits,
    pub svg: super::svg::SvgLimits,
}

/// The rendering track advertised by `EmfPlusHeader.Flags.D`.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MetafileKind {
    EmfPlusOnly,
    Dual,
}

/// Multiplexing state for an outer classic-EMF dispatcher.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct MuxState {
    /// Header kind, once the first EMF+ record has been seen.
    pub kind: Option<MetafileKind>,
    /// Whether following ordinary EMF records should be played.
    ///
    /// This is always false for Dual streams because this renderer selects the
    /// complete EMF+ track.  In EMF+-only streams it becomes true after `GetDC`
    /// and becomes false when the next EMF+ record arrives.
    pub classic_emf_enabled: bool,
    /// Number of `GetDC` boundaries observed.
    pub get_dc_count: usize,
}

/// Result of offering one classic `EMR_COMMENT` body to the renderer.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CommentUpdate {
    pub was_emfplus: bool,
    pub records: usize,
    pub mux: MuxState,
}

/// A non-fatal, non-silent rendering degradation.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct RendererDiagnostic {
    pub code: &'static str,
    pub message: String,
    pub record_offset: Option<usize>,
}

/// Completed EMF+ SVG and semantic/mux metadata.
#[derive(Clone, Debug, PartialEq)]
pub struct RenderOutput {
    pub document: SvgDocument,
    pub diagnostics: Vec<RendererDiagnostic>,
    pub mux: MuxState,
}

impl RenderOutput {
    /// Complete standalone SVG source.
    #[must_use]
    pub fn svg(&self) -> &str {
        self.document.as_str()
    }

    /// Safe definitions/body fragment for insertion by a classic EMF composer.
    #[must_use]
    pub fn fragment(&self) -> &super::svg::SvgFragment {
        self.document.fragment()
    }
}

/// Stateful ordered-comment renderer.
#[derive(Debug)]
pub struct EmfPlusSvgRenderer {
    limits: RendererLimits,
    validator: EmfPlusStreamValidator,
    playback: PlaybackEngine,
    svg: SvgBuilder,
    diagnostics: Vec<RendererDiagnostic>,
    mux: MuxState,
    width: f64,
    height: f64,
    view_x: f64,
    view_y: f64,
    reference: ReferenceDevice,
    /// The last clip expression and the SVG clip defined for it, so that
    /// consecutive draws under one clip share one definition.
    clip_cache: Option<(Clip, Option<SvgId>)>,
}

/// The reference device described by the `EmfPlusHeader` record.
#[derive(Clone, Copy, Debug, PartialEq)]
struct ReferenceDevice {
    dpi_x: f64,
    dpi_y: f64,
    /// Whether the metafile was recorded for a video display rather than a
    /// printer, which decides the size of `UnitTypeDisplay`.
    video_display: bool,
}

impl Default for ReferenceDevice {
    fn default() -> Self {
        Self {
            dpi_x: 96.0,
            dpi_y: 96.0,
            video_display: true,
        }
    }
}

impl EmfPlusSvgRenderer {
    pub fn new(width: f64, height: f64, limits: RendererLimits) -> Result<Self> {
        limits.parser.validate()?;
        limits.playback.validate()?;
        limits.objects.validate()?;
        limits.svg.validate()?;
        Ok(Self {
            limits,
            validator: EmfPlusStreamValidator::new(limits.parser)?,
            playback: PlaybackEngine::new(limits.playback)?,
            svg: SvgBuilder::new(width, height, limits.svg)?,
            diagnostics: Vec::new(),
            mux: MuxState::default(),
            width,
            height,
            view_x: 0.0,
            view_y: 0.0,
            reference: ReferenceDevice::default(),
            clip_cache: None,
        })
    }

    /// Place the top-left corner of the output at the device point `(x, y)`.
    ///
    /// EMF+ device space is the device space of the enclosing EMF, so this is
    /// the corner of the classic picture's bounds.
    #[must_use]
    pub fn with_view_origin(mut self, x: f64, y: f64) -> Self {
        self.view_x = x;
        self.view_y = y;
        self.svg = self.svg.with_view_origin(x, y);
        self
    }

    #[must_use]
    pub const fn mux_state(&self) -> MuxState {
        self.mux
    }

    /// Offer an outer `EMR_COMMENT` body (`DataSize`, identifier, payload, padding).
    /// Valid non-EMF+ comments are ignored and return `was_emfplus = false`.
    pub fn push_comment_body(&mut self, body: &[u8]) -> Result<CommentUpdate> {
        let Some(payload) = try_extract_emfplus_comment_body(body, self.limits.parser)? else {
            return Ok(CommentUpdate {
                was_emfplus: false,
                records: 0,
                mux: self.mux,
            });
        };
        let records = self.push_payload(payload)?;
        Ok(CommentUpdate {
            was_emfplus: true,
            records,
            mux: self.mux,
        })
    }

    /// Feed one already-extracted EMF+ payload in its original stream order.
    pub fn push_payload(&mut self, payload: &[u8]) -> Result<usize> {
        let mut count = 0usize;
        for framed in EmfPlusRecordIter::new(payload, self.limits.parser)? {
            let record = framed?;
            self.validator.push(record)?;
            self.update_mux(record);
            let commands = self.playback.push(record)?;
            for command in &commands {
                self.render_command(command)?;
            }
            count = count
                .checked_add(1)
                .ok_or_else(|| Error::ParseError("EMF+ comment record count overflow".into()))?;
        }
        Ok(count)
    }

    /// Complete the logical stream. Header, EOF, and object continuations are required.
    pub fn finish(mut self) -> Result<RenderOutput> {
        self.validator.finish()?;
        self.playback.finish()?;
        self.diagnostics
            .extend(
                self.playback
                    .diagnostics()
                    .iter()
                    .map(|diagnostic| RendererDiagnostic {
                        code: diagnostic.code,
                        message: diagnostic.message.clone(),
                        record_offset: Some(diagnostic.record_offset),
                    }),
            );
        let document = self.svg.finish()?;
        self.diagnostics
            .extend(document.fragment().diagnostics().iter().map(|diagnostic| {
                let (code, message) = match diagnostic.feature {
                    super::svg::SvgUnsupportedFeature::RawMarkup => (
                        "raw_svg_rejected",
                        "raw SVG markup was rejected by the safe SVG emitter",
                    ),
                    super::svg::SvgUnsupportedFeature::ExternalImage => (
                        "external_image_rejected",
                        "an external image reference was rejected by the safe SVG emitter",
                    ),
                    super::svg::SvgUnsupportedFeature::SourceCopyCompositing => (
                        "source_copy_approximated",
                        "source-copy compositing has no faithful ordinary SVG equivalent",
                    ),
                };
                RendererDiagnostic {
                    code,
                    message: message.to_owned(),
                    record_offset: None,
                }
            }));
        Ok(RenderOutput {
            document,
            diagnostics: self.diagnostics,
            mux: self.mux,
        })
    }

    fn update_mux(&mut self, record: EmfPlusRecord<'_>) {
        // Any new EMF+ record closes the GetDC interval before that record.
        self.mux.classic_emf_enabled = false;
        if record.header.record_type == RecordType::Header {
            let kind = if record.header.flags.raw() & 1 != 0 {
                MetafileKind::Dual
            } else {
                MetafileKind::EmfPlusOnly
            };
            self.mux.kind = Some(kind);
            // Version, EmfPlusFlags, LogicalDpiX, LogicalDpiY.
            let field = |at: usize| {
                record
                    .data
                    .get(at..at + 4)
                    .map(|bytes| u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
            };
            if let (Some(flags), Some(dpi_x), Some(dpi_y)) = (field(4), field(8), field(12)) {
                self.reference.video_display = flags & 1 != 0;
                if dpi_x != 0 && dpi_y != 0 {
                    self.reference.dpi_x = f64::from(dpi_x);
                    self.reference.dpi_y = f64::from(dpi_y);
                }
            }
        } else if record.header.record_type == RecordType::GetDc {
            self.mux.get_dc_count = self.mux.get_dc_count.saturating_add(1);
            self.mux.classic_emf_enabled = self.mux.kind == Some(MetafileKind::EmfPlusOnly);
        }
    }

    fn render_command(&mut self, command: &DrawCommand) -> Result<()> {
        let transform = Some(self.device_transform(&command.state, command.offset));
        let clip = self.clip_id(&command.state.clip, command.offset)?;
        match &command.kind {
            DrawCommandKind::Clear { color } => {
                if Argb(*color).alpha() != 0xff {
                    // Clear replaces the pixels instead of blending over them.
                    self.diag(
                        "translucent_clear",
                        "clearing with a translucent color is drawn by blending",
                        command.offset,
                    );
                }
                let style =
                    self.fill_style(Brush::Solid(*color), &command.state, command.offset)?;
                self.svg.rect(
                    SvgRect {
                        x: self.view_x,
                        y: self.view_y,
                        width: self.width,
                        height: self.height,
                    },
                    &style,
                    None,
                    clip.as_ref(),
                )?;
            },
            DrawCommandKind::FillRects { brush, rects } => {
                let style = self.fill_style(*brush, &command.state, command.offset)?;
                for rect in rects {
                    self.svg
                        .rect(svg_rect(*rect), &style, transform, clip.as_ref())?;
                }
            },
            DrawCommandKind::DrawRects { pen, rects } => {
                let style = self.stroke_style(*pen, &command.state, command.offset)?;
                for rect in rects {
                    self.svg
                        .rect(svg_rect(*rect), &style, transform, clip.as_ref())?;
                }
            },
            DrawCommandKind::FillPolygon {
                brush,
                points,
                fill_mode,
            } => {
                let path = polygon_path(points, true);
                let mut style = self.fill_style(*brush, &command.state, command.offset)?;
                style.even_odd_fill = *fill_mode == 0;
                self.svg.path(&path, &style, transform, clip.as_ref())?;
            },
            DrawCommandKind::DrawLines { pen, points } => {
                let path = polygon_path(points, false);
                let style = self.stroke_style(*pen, &command.state, command.offset)?;
                self.svg.path(&path, &style, transform, clip.as_ref())?;
            },
            DrawCommandKind::FillEllipse { brush, rect } => {
                let style = self.fill_style(*brush, &command.state, command.offset)?;
                self.svg
                    .ellipse(svg_rect(*rect), &style, transform, clip.as_ref())?;
            },
            DrawCommandKind::DrawEllipse { pen, rect } => {
                let style = self.stroke_style(*pen, &command.state, command.offset)?;
                self.svg
                    .ellipse(svg_rect(*rect), &style, transform, clip.as_ref())?;
            },
            DrawCommandKind::FillPie {
                brush,
                rect,
                start_angle,
                sweep_angle,
            } => {
                let style = self.fill_style(*brush, &command.state, command.offset)?;
                self.svg.path(
                    &arc_path(*rect, *start_angle, *sweep_angle, true),
                    &style,
                    transform,
                    clip.as_ref(),
                )?;
            },
            DrawCommandKind::DrawPie {
                pen,
                rect,
                start_angle,
                sweep_angle,
            } => {
                let style = self.stroke_style(*pen, &command.state, command.offset)?;
                self.svg.path(
                    &arc_path(*rect, *start_angle, *sweep_angle, true),
                    &style,
                    transform,
                    clip.as_ref(),
                )?;
            },
            DrawCommandKind::DrawArc {
                pen,
                rect,
                start_angle,
                sweep_angle,
            } => {
                let style = self.stroke_style(*pen, &command.state, command.offset)?;
                self.svg.path(
                    &arc_path(*rect, *start_angle, *sweep_angle, false),
                    &style,
                    transform,
                    clip.as_ref(),
                )?;
            },
            DrawCommandKind::FillRegion { brush, region } => {
                let style = self.fill_style(*brush, &command.state, command.offset)?;
                self.render_region(*region, &style, transform, clip.as_ref(), command.offset)?;
            },
            DrawCommandKind::FillPath { brush, path } => {
                let (path, even_odd) = self.object_path(*path)?;
                let mut style = self.fill_style(*brush, &command.state, command.offset)?;
                style.even_odd_fill = even_odd;
                self.svg.path(&path, &style, transform, clip.as_ref())?;
            },
            DrawCommandKind::DrawPath { pen, path } => {
                let (path, _) = self.object_path(*path)?;
                let style = self.stroke_style(*pen, &command.state, command.offset)?;
                self.svg.path(&path, &style, transform, clip.as_ref())?;
            },
            DrawCommandKind::FillClosedCurve {
                brush,
                points,
                fill_mode,
                ..
            } => {
                self.diag(
                    "curve_approximation",
                    "cardinal curve rendered as a closed polygon",
                    command.offset,
                );
                let mut style = self.fill_style(*brush, &command.state, command.offset)?;
                style.even_odd_fill = *fill_mode == 0;
                self.svg.path(
                    &polygon_path(points, true),
                    &style,
                    transform,
                    clip.as_ref(),
                )?;
            },
            DrawCommandKind::DrawClosedCurve { pen, points, .. } => {
                self.diag(
                    "curve_approximation",
                    "cardinal curve rendered as a closed polyline",
                    command.offset,
                );
                let style = self.stroke_style(*pen, &command.state, command.offset)?;
                self.svg.path(
                    &polygon_path(points, true),
                    &style,
                    transform,
                    clip.as_ref(),
                )?;
            },
            DrawCommandKind::DrawCurve { pen, points, .. } => {
                self.diag(
                    "curve_approximation",
                    "cardinal curve rendered as a polyline",
                    command.offset,
                );
                let style = self.stroke_style(*pen, &command.state, command.offset)?;
                self.svg.path(
                    &polygon_path(points, false),
                    &style,
                    transform,
                    clip.as_ref(),
                )?;
            },
            DrawCommandKind::DrawBeziers { pen, points } => {
                let style = self.stroke_style(*pen, &command.state, command.offset)?;
                self.svg
                    .path(&bezier_path(points), &style, transform, clip.as_ref())?;
            },
            DrawCommandKind::DrawImage {
                image,
                src,
                src_unit,
                dest,
                ..
            } => {
                self.render_image(
                    *image,
                    ImageSource {
                        rect: *src,
                        unit: *src_unit,
                    },
                    ImageFrame::rect(*dest),
                    transform,
                    clip.as_ref(),
                    command.offset,
                )?;
            },
            DrawCommandKind::DrawImagePoints {
                image,
                src,
                src_unit,
                points,
                ..
            } => {
                let frame = parallelogram(points);
                if frame.is_none() {
                    // Only a parallelogram can be mapped exactly; other point
                    // counts keep their bounding box.
                    self.diag(
                        "image_parallelogram",
                        "image destination points are not a parallelogram; its bounding box was used",
                        command.offset,
                    );
                }
                let Some(frame) = frame.or_else(|| bounds(points).map(ImageFrame::rect)) else {
                    return Err(Error::ParseError(
                        "DrawImagePoints has no destination".into(),
                    ));
                };
                self.render_image(
                    *image,
                    ImageSource {
                        rect: *src,
                        unit: *src_unit,
                    },
                    frame,
                    transform,
                    clip.as_ref(),
                    command.offset,
                )?;
            },
            DrawCommandKind::DrawString {
                brush,
                font,
                text,
                layout,
                ..
            } => {
                self.render_text(
                    *brush,
                    *font,
                    text.clone(),
                    *layout,
                    &command.state,
                    clip.as_ref(),
                    command.offset,
                )?;
            },
            DrawCommandKind::DrawDriverString {
                brush,
                font,
                options,
                glyphs,
                positions,
                transform: matrix,
            } => {
                self.render_glyphs(
                    *brush,
                    *font,
                    DriverString {
                        options: *options,
                        glyphs,
                        positions,
                        matrix: *matrix,
                    },
                    &command.state,
                    clip.as_ref(),
                    command.offset,
                )?;
            },
            DrawCommandKind::StrokeFillPath { .. } => {
                self.diag(
                    "underspecified_stroke_fill_path",
                    "record has no normative body layout and was not guessed",
                    command.offset,
                );
            },
        }
        Ok(())
    }

    fn fill_style(
        &mut self,
        brush: Brush,
        state: &GraphicsState,
        offset: usize,
    ) -> Result<SvgStyle> {
        Ok(SvgStyle {
            fill: self.brush_paint(brush, offset)?,
            stroke: None,
            opacity: 1.0,
            compositing: compositing(state.rendering.compositing_mode),
            even_odd_fill: false,
        })
    }

    fn stroke_style(
        &mut self,
        id: ObjectId,
        state: &GraphicsState,
        offset: usize,
    ) -> Result<SvgStyle> {
        let object = self.decoded(id)?;
        let GraphicsObject::Pen(pen) = object else {
            return Err(Error::ParseError(
                "EMF+ pen reference has the wrong object type".into(),
            ));
        };
        let paint = self.object_brush_paint(&pen.brush.kind, offset)?;
        if !pen.diagnostics.is_empty() {
            self.diag(
                "pen_custom_cap",
                "custom pen line caps were not drawn",
                offset,
            );
        }
        if pen.transform.is_some() {
            self.diag("pen_transform", "pen transform was not applied", offset);
        }
        if pen.alignment.unwrap_or(PEN_ALIGNMENT_CENTER) != PEN_ALIGNMENT_CENTER {
            self.diag("pen_alignment", "pen alignment drawn centered", offset);
        }
        if !pen.compound.is_empty() {
            self.diag("pen_compound", "compound pen drawn as one line", offset);
        }
        if pen.unit != PEN_UNIT_WORLD {
            self.diag("pen_unit", "pen width unit used as world units", offset);
        }
        let start_cap = pen.start_cap.unwrap_or(LINE_CAP_FLAT);
        let end_cap = pen.end_cap.unwrap_or(LINE_CAP_FLAT);
        if start_cap != end_cap {
            self.diag(
                "pen_line_caps",
                "different start and end caps drawn with the start cap",
                offset,
            );
        }
        let line_cap = match start_cap {
            LINE_CAP_FLAT => SvgLineCap::Butt,
            LINE_CAP_SQUARE => SvgLineCap::Square,
            LINE_CAP_ROUND => SvgLineCap::Round,
            _ => {
                self.diag("pen_line_cap", "triangle or anchor cap drawn flat", offset);
                SvgLineCap::Butt
            },
        };
        let line_join = match pen.join.unwrap_or(LINE_JOIN_MITER) {
            LINE_JOIN_MITER => SvgLineJoin::Miter,
            LINE_JOIN_BEVEL => SvgLineJoin::Bevel,
            LINE_JOIN_ROUND => SvgLineJoin::Round,
            _ => {
                self.diag(
                    "pen_line_join",
                    "clipped miter join drawn as a miter",
                    offset,
                );
                SvgLineJoin::Miter
            },
        };
        let mut width = f64::from(pen.width.abs());
        if width == 0.0 {
            // GDI+ draws a zero-width pen one device pixel wide.
            let transform = self.device_transform(state, offset);
            let scale = (transform.a * transform.d - transform.b * transform.c)
                .abs()
                .sqrt();
            if scale > 0.0 {
                width = 1.0 / scale;
            }
        }
        // Dash lengths and the dash offset are multiples of the pen width.
        let pattern: &[f32] = match pen.line_style {
            _ if !pen.dashes.is_empty() => &pen.dashes,
            Some(LINE_STYLE_DASH) => &[3.0, 1.0],
            Some(LINE_STYLE_DOT) => &[1.0, 1.0],
            Some(LINE_STYLE_DASH_DOT) => &[3.0, 1.0, 1.0, 1.0],
            Some(LINE_STYLE_DASH_DOT_DOT) => &[3.0, 1.0, 1.0, 1.0, 1.0, 1.0],
            _ => &[],
        };
        let dashes: Vec<f64> = pattern
            .iter()
            .map(|value| f64::from(value.abs()) * width)
            .collect();
        if !dashes.is_empty() {
            let dash_cap = match pen.dashed_cap.unwrap_or(DASHED_LINE_CAP_FLAT) {
                DASHED_LINE_CAP_FLAT => Some(SvgLineCap::Butt),
                DASHED_LINE_CAP_ROUND => Some(SvgLineCap::Round),
                _ => None,
            };
            if dash_cap != Some(line_cap) || start_cap != end_cap {
                self.diag(
                    "pen_dash_cap",
                    "dash caps that differ from the line caps drawn as the line caps",
                    offset,
                );
            }
        }
        Ok(SvgStyle {
            fill: SvgPaint::None,
            even_odd_fill: false,
            stroke: Some(SvgStroke {
                paint,
                width,
                line_cap,
                line_join,
                dashes,
                dash_offset: f64::from(pen.dash_offset.unwrap_or(0.0)) * width,
                miter_limit: f64::from(pen.miter_limit.unwrap_or(DEFAULT_MITER_LIMIT)),
            }),
            opacity: 1.0,
            compositing: compositing(state.rendering.compositing_mode),
        })
    }

    fn brush_paint(&mut self, brush: Brush, offset: usize) -> Result<SvgPaint> {
        match brush {
            Brush::Solid(color) => Ok(SvgPaint::Solid(svg_color(Argb(color)))),
            Brush::Object(id) => {
                let object = self.decoded(id)?;
                let GraphicsObject::Brush(brush) = object else {
                    return Err(Error::ParseError(
                        "EMF+ brush reference has the wrong object type".into(),
                    ));
                };
                self.object_brush_paint(&brush.kind, offset)
            },
        }
    }

    fn object_brush_paint(&mut self, brush: &BrushKind, offset: usize) -> Result<SvgPaint> {
        match brush {
            BrushKind::Solid { color } => Ok(SvgPaint::Solid(svg_color(*color))),
            BrushKind::Hatch { foreground, .. } => {
                self.diag(
                    "hatch_approximation",
                    "hatch brush represented by its foreground color",
                    offset,
                );
                Ok(SvgPaint::Solid(svg_color(*foreground)))
            },
            BrushKind::LinearGradient(gradient) => {
                if !gradient.blends.is_empty() {
                    self.diag(
                        "gradient_blend",
                        "gradient blend factors or preset colors drawn as two colors",
                        offset,
                    );
                }
                if gradient.flags & BRUSH_GAMMA_CORRECTED != 0 {
                    self.diag(
                        "gradient_gamma",
                        "gamma-corrected gradient drawn without gamma correction",
                        offset,
                    );
                }
                // The colors run across the rectangle from its left edge to
                // its right edge; other directions are part of the transform.
                let start = SvgPoint {
                    x: f64::from(gradient.rect.x),
                    y: f64::from(gradient.rect.y),
                };
                let end = SvgPoint {
                    x: f64::from(gradient.rect.x + gradient.rect.width),
                    y: f64::from(gradient.rect.y),
                };
                // WrapMode: Tile, TileFlipX, TileFlipY, TileFlipXY, Clamp. A
                // vertical flip does not change a horizontal gradient.
                let spread = match gradient.wrap_mode {
                    0 | 2 => SvgSpread::Repeat,
                    1 | 3 => SvgSpread::Reflect,
                    _ => SvgSpread::Pad,
                };
                let id = self.svg.define_linear_gradient(&SvgLinearGradient {
                    start,
                    end,
                    stops: vec![
                        SvgGradientStop {
                            offset: 0.0,
                            color: svg_color(gradient.start),
                        },
                        SvgGradientStop {
                            offset: 1.0,
                            color: svg_color(gradient.end),
                        },
                    ],
                    transform: gradient.transform.map(object_transform),
                    spread,
                })?;
                Ok(SvgPaint::Reference(id))
            },
            BrushKind::Texture(_) | BrushKind::PathGradient(_) | BrushKind::Unsupported { .. } => {
                self.diag(
                    "unsupported_brush",
                    "brush cannot be represented by the safe SVG subset",
                    offset,
                );
                Ok(SvgPaint::None)
            },
        }
    }

    fn decoded(&self, id: ObjectId) -> Result<GraphicsObject> {
        let blob = self
            .playback
            .object(id)
            .ok_or_else(|| Error::ParseError(format!("EMF+ object {} is undefined", id.get())))?;
        objects::decode_object(blob.object_type, &blob.bytes, self.limits.objects)
    }

    /// A path object and whether it fills with the alternate (even-odd) rule.
    fn object_path(&self, id: ObjectId) -> Result<(SvgPath, bool)> {
        let GraphicsObject::Path(path) = self.decoded(id)? else {
            return Err(Error::ParseError(
                "EMF+ path reference has the wrong object type".into(),
            ));
        };
        Ok((svg_object_path(&path), path_is_alternate(&path)))
    }

    /// The SVG clip of a clip expression, defined once per run of draws that
    /// share it.
    fn clip_id(&mut self, clip: &Clip, offset: usize) -> Result<Option<SvgId>> {
        if let Some((cached, id)) = &self.clip_cache
            && cached == clip
        {
            return Ok(id.clone());
        }
        let id = self.define_clip(clip, None, offset)?;
        self.clip_cache = Some((clip.clone(), id.clone()));
        Ok(id)
    }

    /// Define `clip` in device coordinates, intersected with `outer`.
    fn define_clip(
        &mut self,
        clip: &Clip,
        outer: Option<&SvgId>,
        offset: usize,
    ) -> Result<Option<SvgId>> {
        match clip {
            Clip::Infinite => Ok(outer.cloned()),
            Clip::Rect { mode, rect, space } => {
                self.combine_diagnostic(*mode, offset);
                let transform = self.space_transform(*space, offset);
                Ok(Some(self.svg.define_clip_rect(
                    svg_rect(*rect),
                    Some(transform),
                    outer,
                )?))
            },
            Clip::Path { mode, path, space } => {
                self.combine_diagnostic(*mode, offset);
                let GraphicsObject::Path(path) =
                    objects::decode_object(path.object_type, &path.bytes, self.limits.objects)?
                else {
                    return Err(Error::ParseError(
                        "EMF+ clip path has the wrong object type".into(),
                    ));
                };
                let transform = self.space_transform(*space, offset);
                Ok(Some(self.svg.define_clip_path(
                    &svg_object_path(&path),
                    Some(transform),
                    outer,
                    path_is_alternate(&path),
                )?))
            },
            Clip::Region {
                mode,
                region,
                space,
            } => {
                self.combine_diagnostic(*mode, offset);
                let GraphicsObject::Region(region) =
                    objects::decode_object(region.object_type, &region.bytes, self.limits.objects)?
                else {
                    return Err(Error::ParseError(
                        "EMF+ clip region has the wrong object type".into(),
                    ));
                };
                let transform = self.space_transform(*space, offset);
                match region.root {
                    RegionNode::Rect(rect) => Ok(Some(self.svg.define_clip_rect(
                        object_rect(rect),
                        Some(transform),
                        outer,
                    )?)),
                    RegionNode::Path(path) => Ok(Some(self.svg.define_clip_path(
                        &svg_object_path(&path),
                        Some(transform),
                        outer,
                        path_is_alternate(&path),
                    )?)),
                    RegionNode::Infinite => Ok(outer.cloned()),
                    // An empty clip hides everything drawn under it.
                    RegionNode::Empty => Ok(Some(self.svg.define_clip_rect(
                        SvgRect {
                            x: 0.0,
                            y: 0.0,
                            width: 0.0,
                            height: 0.0,
                        },
                        None,
                        outer,
                    )?)),
                    RegionNode::And(_, _)
                    | RegionNode::Or(_, _)
                    | RegionNode::Xor(_, _)
                    | RegionNode::Exclude(_, _)
                    | RegionNode::Complement(_, _) => {
                        self.diag(
                            "complex_region_clip",
                            "complex Boolean region clip was not applied",
                            offset,
                        );
                        Ok(outer.cloned())
                    },
                }
            },
            Clip::Intersect(members) => {
                let mut current = outer.cloned();
                for member in members {
                    current = self.define_clip(member, current.as_ref(), offset)?;
                }
                Ok(current)
            },
            Clip::Offset { clip, .. } => {
                self.diag(
                    "offset_clip",
                    "offset clip represented without its offset",
                    offset,
                );
                self.define_clip(clip, outer, offset)
            },
        }
    }

    fn combine_diagnostic(&mut self, mode: CombineMode, offset: usize) {
        if !matches!(mode, CombineMode::Replace | CombineMode::Intersect) {
            self.diag(
                "clip_combine_approximation",
                "SVG clip cannot represent this Boolean combine mode",
                offset,
            );
        }
    }

    /// The SVG transform from world coordinates to device coordinates.
    fn device_transform(&mut self, state: &GraphicsState, offset: usize) -> SvgTransform {
        self.space_transform(
            ClipSpace {
                world_transform: state.world_transform,
                page_transform: state.page_transform,
            },
            offset,
        )
    }

    /// The SVG transform of a world transform followed by a page transform.
    fn space_transform(&mut self, space: ClipSpace, offset: usize) -> SvgTransform {
        let (unit_x, unit_y) = self.unit_scale(space.page_transform.unit, offset);
        let scale = f64::from(space.page_transform.scale);
        let (sx, sy) = (unit_x * scale, unit_y * scale);
        let world = svg_transform(space.world_transform);
        SvgTransform {
            a: world.a * sx,
            b: world.b * sy,
            c: world.c * sx,
            d: world.d * sy,
            e: world.e * sx,
            f: world.f * sy,
        }
    }

    /// Device pixels per unit, horizontally and vertically.
    fn unit_scale(&mut self, unit: Unit, offset: usize) -> (f64, f64) {
        let ReferenceDevice {
            dpi_x,
            dpi_y,
            video_display,
        } = self.reference;
        match unit {
            Unit::Pixel => (1.0, 1.0),
            // A display unit is a pixel of a video display and 1/100 inch on a
            // printer.
            Unit::Display if video_display => (1.0, 1.0),
            Unit::Display => (dpi_x / 100.0, dpi_y / 100.0),
            Unit::Point => (dpi_x / 72.0, dpi_y / 72.0),
            Unit::Inch => (dpi_x, dpi_y),
            Unit::Document => (dpi_x / 300.0, dpi_y / 300.0),
            Unit::Millimeter => (dpi_x / 25.4, dpi_y / 25.4),
            Unit::World => {
                self.diag(
                    "world_page_unit",
                    "the world unit is not a page unit; one pixel per unit was used",
                    offset,
                );
                (1.0, 1.0)
            },
        }
    }

    fn render_region(
        &mut self,
        id: ObjectId,
        style: &SvgStyle,
        transform: Option<SvgTransform>,
        clip: Option<&SvgId>,
        offset: usize,
    ) -> Result<()> {
        let GraphicsObject::Region(region) = self.decoded(id)? else {
            return Err(Error::ParseError(
                "EMF+ region reference has the wrong object type".into(),
            ));
        };
        match region.root {
            RegionNode::Rect(rect) => self.svg.rect(object_rect(rect), style, transform, clip),
            RegionNode::Path(path) => {
                let style = SvgStyle {
                    even_odd_fill: path_is_alternate(&path),
                    ..style.clone()
                };
                self.svg
                    .path(&svg_object_path(&path), &style, transform, clip)
            },
            RegionNode::Empty => Ok(()),
            RegionNode::Infinite => {
                self.diag(
                    "infinite_region",
                    "infinite region fill bounded to SVG canvas is unavailable",
                    offset,
                );
                Ok(())
            },
            RegionNode::And(_, _)
            | RegionNode::Or(_, _)
            | RegionNode::Xor(_, _)
            | RegionNode::Exclude(_, _)
            | RegionNode::Complement(_, _) => {
                self.diag(
                    "complex_region",
                    "Boolean region fill was not representable",
                    offset,
                );
                Ok(())
            },
        }
    }

    fn render_image(
        &mut self,
        id: ObjectId,
        src: ImageSource,
        frame: ImageFrame,
        transform: Option<SvgTransform>,
        clip: Option<&SvgId>,
        offset: usize,
    ) -> Result<()> {
        let GraphicsObject::Image(image) = self.decoded(id)? else {
            return Err(Error::ParseError(
                "EMF+ image reference has the wrong object type".into(),
            ));
        };
        let bytes = match image.kind {
            ImageKind::Compressed(bytes) => bytes,
            ImageKind::Bitmap(bitmap) => match self.bitmap_png(&bitmap) {
                Ok(Some(bytes)) => bytes,
                Ok(None) => {
                    self.diag(
                        "unsupported_image",
                        "bitmap pixel format is not supported",
                        offset,
                    );
                    return Ok(());
                },
                Err(error) => {
                    self.diag("invalid_bitmap", &error.to_string(), offset);
                    return Ok(());
                },
            },
            ImageKind::Metafile { .. } | ImageKind::Unsupported { .. } => {
                self.diag(
                    "unsupported_image",
                    "nested metafile image is not drawn",
                    offset,
                );
                return Ok(());
            },
        };
        let Some(mime) = image_mime(&bytes) else {
            self.diag(
                "unknown_image_encoding",
                "compressed image signature is unsupported",
                offset,
            );
            return Ok(());
        };
        let Some((width, height)) = image_dimensions(mime, &bytes) else {
            self.diag(
                "image_size",
                "image size is unknown, so its source rectangle was not applied",
                offset,
            );
            return Ok(());
        };
        if src.unit != Unit::Pixel {
            self.diag(
                "image_source_unit",
                "image source rectangle unit used as pixels",
                offset,
            );
        }
        let device = transform.unwrap_or(SvgTransform::identity());
        let src = src.rect;
        // The source rectangle selects the part of the image that fills the
        // destination frame. Placing the whole image so that the source
        // rectangle covers the frame, clipped to the frame, keeps that part.
        let whole = src.x == 0.0
            && src.y == 0.0
            && f64::from(src.width) == width
            && f64::from(src.height) == height;
        let (frame, clip) = if whole || src.width <= 0.0 || src.height <= 0.0 {
            (frame, clip.cloned())
        } else {
            let crop = self
                .svg
                .define_clip_path(&frame.path(), Some(device), clip, false)?;
            (frame.crop(src, width, height), Some(crop))
        };
        self.svg.image(
            &SvgImage {
                rect: SvgRect {
                    x: 0.0,
                    y: 0.0,
                    width: 1.0,
                    height: 1.0,
                },
                source: SvgImageSource::Embedded { mime, bytes },
                transform: Some(frame.transform(device)),
                opacity: 1.0,
            },
            clip.as_ref(),
        )
    }

    /// Encode an uncompressed EMF+ bitmap as PNG.
    ///
    /// Scan lines are stored top-down with GDI+'s little-endian pixel layout;
    /// indexed formats start with an `EmfPlusPalette`. Returns `None` for a
    /// pixel format without a decoder.
    fn bitmap_png(&self, bitmap: &objects::Bitmap) -> Result<Option<Vec<u8>>> {
        let invalid = |message: &str| Error::ParseError(format!("EMF+ bitmap {message}"));
        let width = u32::try_from(bitmap.width).map_err(|_| invalid("width is negative"))?;
        let height = u32::try_from(bitmap.height).map_err(|_| invalid("height is negative"))?;
        let stride = usize::try_from(bitmap.stride).map_err(|_| invalid("stride is negative"))?;
        let bits = match bitmap.pixel_format {
            PIXEL_FORMAT_1BPP_INDEXED => 1,
            PIXEL_FORMAT_4BPP_INDEXED => 4,
            PIXEL_FORMAT_8BPP_INDEXED => 8,
            PIXEL_FORMAT_16BPP_RGB555 | PIXEL_FORMAT_16BPP_RGB565 => 16,
            PIXEL_FORMAT_24BPP_RGB => 24,
            PIXEL_FORMAT_32BPP_RGB | PIXEL_FORMAT_32BPP_ARGB | PIXEL_FORMAT_32BPP_PARGB => 32,
            _ => return Ok(None),
        };
        let pixels = u64::from(width) * u64::from(height);
        if width == 0 || height == 0 || pixels > self.limits.objects.max_bytes as u64 {
            return Err(invalid("size is empty or exceeds the object byte limit"));
        }
        let row_bytes =
            (usize::try_from(width).map_err(|_| invalid("width overflows"))? * bits).div_ceil(8);
        if stride < row_bytes {
            return Err(invalid("stride is shorter than a scan line"));
        }
        let mut data = bitmap.pixels.as_slice();
        let mut palette = Vec::new();
        if bitmap.pixel_format & PIXEL_FORMAT_INDEXED != 0 {
            let word = |data: &[u8], at: usize| {
                data.get(at..at + 4)
                    .map(|bytes| u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
            };
            let count = word(data, 4).ok_or_else(|| invalid("palette is truncated"))?;
            let count = usize::try_from(count)
                .ok()
                .filter(|count| *count <= 256)
                .ok_or_else(|| invalid("palette has more than 256 entries"))?;
            for index in 0..count {
                palette.push(
                    word(data, 8 + index * 4).ok_or_else(|| invalid("palette is truncated"))?,
                );
            }
            data = &data[8 + count * 4..];
        }
        let height_usize = usize::try_from(height).map_err(|_| invalid("height overflows"))?;
        let needed = stride
            .checked_mul(height_usize - 1)
            .and_then(|value| value.checked_add(row_bytes))
            .ok_or_else(|| invalid("pixel data size overflows"))?;
        if data.len() < needed {
            return Err(invalid("pixel data is truncated"));
        }
        let mut image = image::RgbaImage::new(width, height);
        for (y, row) in image.rows_mut().enumerate() {
            let line = &data[y * stride..y * stride + row_bytes];
            for (x, pixel) in row.enumerate() {
                let argb = match bits {
                    1 | 4 | 8 => {
                        let bit = x * bits;
                        let byte = line[bit / 8];
                        let index = (byte >> (8 - bits - bit % 8)) & ((1_u16 << bits) - 1) as u8;
                        palette
                            .get(usize::from(index))
                            .copied()
                            .unwrap_or(0xff00_0000)
                    },
                    16 => {
                        let value = u32::from(u16::from_le_bytes([line[x * 2], line[x * 2 + 1]]));
                        let (red, green, blue) = if bitmap.pixel_format == PIXEL_FORMAT_16BPP_RGB565
                        {
                            ((value >> 11) & 0x1f, (value >> 5) & 0x3f, value & 0x1f)
                        } else {
                            ((value >> 10) & 0x1f, (value >> 5) & 0x1f, value & 0x1f)
                        };
                        let green_max = if bitmap.pixel_format == PIXEL_FORMAT_16BPP_RGB565 {
                            63
                        } else {
                            31
                        };
                        0xff00_0000
                            | ((red * 255 / 31) << 16)
                            | ((green * 255 / green_max) << 8)
                            | (blue * 255 / 31)
                    },
                    24 => {
                        let at = x * 3;
                        0xff00_0000
                            | (u32::from(line[at + 2]) << 16)
                            | (u32::from(line[at + 1]) << 8)
                            | u32::from(line[at])
                    },
                    _ => {
                        let at = x * 4;
                        let value = u32::from_le_bytes([
                            line[at],
                            line[at + 1],
                            line[at + 2],
                            line[at + 3],
                        ]);
                        match bitmap.pixel_format {
                            PIXEL_FORMAT_32BPP_RGB => value | 0xff00_0000,
                            PIXEL_FORMAT_32BPP_PARGB => unpremultiply(value),
                            _ => value,
                        }
                    },
                };
                let color = Argb(argb);
                *pixel = image::Rgba([color.red(), color.green(), color.blue(), color.alpha()]);
            }
        }
        let mut png = std::io::Cursor::new(Vec::new());
        image::DynamicImage::ImageRgba8(image)
            .write_to(&mut png, image::ImageFormat::Png)
            .map_err(|error| invalid(&format!("PNG encoding failed: {error}")))?;
        Ok(Some(png.into_inner()))
    }

    /// Draw an `EmfPlusDrawString` text.
    ///
    /// The layout rectangle and string format decide line breaks and
    /// alignment, which need font metrics; the text is drawn as one line from
    /// the top-left corner of the layout rectangle and reported.
    fn render_text(
        &mut self,
        brush: Brush,
        font: ObjectId,
        value: String,
        layout: Rect,
        state: &GraphicsState,
        clip: Option<&SvgId>,
        offset: usize,
    ) -> Result<()> {
        let GraphicsObject::Font(font) = self.decoded(font)? else {
            return Err(Error::ParseError(
                "EMF+ font reference has the wrong object type".into(),
            ));
        };
        if value.is_empty() {
            return Ok(());
        }
        self.diag(
            "draw_string_layout",
            "string layout and format drawn as one line at the layout rectangle",
            offset,
        );
        let font_size = self.font_size(&font, state, offset);
        let style = self.fill_style(brush, state, offset)?;
        let transform = self.device_transform(state, offset);
        self.svg.text(
            &SvgText {
                value,
                origin: SvgPoint {
                    x: f64::from(layout.x),
                    y: f64::from(layout.y) + font_size,
                },
                positions: Vec::new(),
                font_family: font.family.clone(),
                font_size,
                bold: font.style & FONT_STYLE_BOLD != 0,
                italic: font.style & FONT_STYLE_ITALIC != 0,
                underline: font.style & FONT_STYLE_UNDERLINE != 0,
                strikeout: font.style & FONT_STYLE_STRIKEOUT != 0,
                style,
                transform: Some(transform),
            },
            clip,
        )
    }

    /// Draw an `EmfPlusDrawDriverString` text with its glyphs at their
    /// recorded baseline positions.
    fn render_glyphs(
        &mut self,
        brush: Brush,
        font: ObjectId,
        text: DriverString<'_>,
        state: &GraphicsState,
        clip: Option<&SvgId>,
        offset: usize,
    ) -> Result<()> {
        let GraphicsObject::Font(font) = self.decoded(font)? else {
            return Err(Error::ParseError(
                "EMF+ font reference has the wrong object type".into(),
            ));
        };
        if text.options & DRIVER_STRING_CMAP_LOOKUP == 0 {
            self.diag(
                "driver_string_glyph_indices",
                "glyph indices need the font's glyph table and were not drawn",
                offset,
            );
            return Ok(());
        }
        if text.options & DRIVER_STRING_VERTICAL != 0 {
            self.diag(
                "driver_string_vertical",
                "vertical driver string drawn horizontally",
                offset,
            );
        }
        // With realized advances only the first position is recorded and the
        // font's advances place the other glyphs, as SVG text does.
        let realized = text.options & DRIVER_STRING_REALIZED_ADVANCE != 0;
        let mut value = String::new();
        let mut positions = Vec::new();
        let mut index = 0usize;
        while index < text.glyphs.len() {
            let (glyph, step) = glyph_text(text.glyphs, index);
            if !realized {
                let position = text.positions.get(index).copied().unwrap_or_default();
                positions.extend(glyph.chars().map(|_| svg_point(position)));
            }
            value.push_str(&glyph);
            index += step;
        }
        if value.is_empty() {
            return Ok(());
        }
        let origin = text.positions.first().copied().unwrap_or_default();
        let font_size = self.font_size(&font, state, offset);
        let style = self.fill_style(brush, state, offset)?;
        // The record matrix applies to the glyphs and their positions before
        // the world transform.
        let world = text.matrix.map_or(state.world_transform, |matrix| {
            matrix.multiply(state.world_transform)
        });
        let transform = self.space_transform(
            ClipSpace {
                world_transform: world,
                page_transform: state.page_transform,
            },
            offset,
        );
        self.svg.text(
            &SvgText {
                value,
                origin: svg_point(origin),
                positions,
                font_family: font.family.clone(),
                font_size,
                bold: font.style & FONT_STYLE_BOLD != 0,
                italic: font.style & FONT_STYLE_ITALIC != 0,
                underline: font.style & FONT_STYLE_UNDERLINE != 0,
                strikeout: font.style & FONT_STYLE_STRIKEOUT != 0,
                style,
                transform: Some(transform),
            },
            clip,
        )
    }

    /// The em size of `font` in world units.
    ///
    /// A font measured in a physical unit keeps its device size under the page
    /// transform, so the page scale is divided out.
    fn font_size(&mut self, font: &objects::Font, state: &GraphicsState, offset: usize) -> f64 {
        let em = f64::from(font.em_size.abs());
        let unit = u8::try_from(font.unit).ok().and_then(Unit::parse);
        match unit {
            Some(Unit::World) => em,
            Some(unit) => {
                let (font_scale, _) = self.unit_scale(unit, offset);
                let (page_unit, _) = self.unit_scale(state.page_transform.unit, offset);
                let page_scale = page_unit * f64::from(state.page_transform.scale);
                if page_scale == 0.0 {
                    em
                } else {
                    em * font_scale / page_scale
                }
            },
            None => {
                self.diag(
                    "font_unit",
                    "font size unit is invalid; the size was used as world units",
                    offset,
                );
                em
            },
        }
    }

    fn diag(&mut self, code: &'static str, message: &str, offset: usize) {
        if self.diagnostics.len() < self.limits.svg.max_diagnostics {
            self.diagnostics.push(RendererDiagnostic {
                code,
                message: message.to_owned(),
                record_offset: Some(offset),
            });
        }
    }
}

/// Render an ordered sequence of outer `EMR_COMMENT` bodies into standalone SVG.
pub fn render_emfplus_comments_to_svg<'a, I>(
    width: f64,
    height: f64,
    comments: I,
    limits: RendererLimits,
) -> Result<RenderOutput>
where
    I: IntoIterator<Item = &'a [u8]>,
{
    let mut renderer = EmfPlusSvgRenderer::new(width, height, limits)?;
    for comment in comments {
        renderer.push_comment_body(comment)?;
    }
    renderer.finish()
}

/// Render already-extracted ordered EMF+ payloads.
pub fn render_emfplus_payloads_to_svg<'a, I>(
    width: f64,
    height: f64,
    payloads: I,
    limits: RendererLimits,
) -> Result<RenderOutput>
where
    I: IntoIterator<Item = &'a [u8]>,
{
    let mut renderer = EmfPlusSvgRenderer::new(width, height, limits)?;
    for payload in payloads {
        renderer.push_payload(payload)?;
    }
    renderer.finish()
}

/// `PixelFormat` values ([MS-EMFPLUS] 2.1.1.24) with a bitmap decoder.
const PIXEL_FORMAT_INDEXED: u32 = 0x0001_0000;
const PIXEL_FORMAT_1BPP_INDEXED: u32 = 0x0003_0101;
const PIXEL_FORMAT_4BPP_INDEXED: u32 = 0x0003_0402;
const PIXEL_FORMAT_8BPP_INDEXED: u32 = 0x0003_0803;
const PIXEL_FORMAT_16BPP_RGB555: u32 = 0x0002_1005;
const PIXEL_FORMAT_16BPP_RGB565: u32 = 0x0002_1006;
const PIXEL_FORMAT_24BPP_RGB: u32 = 0x0002_1808;
const PIXEL_FORMAT_32BPP_RGB: u32 = 0x0002_2009;
const PIXEL_FORMAT_32BPP_ARGB: u32 = 0x0026_200a;
const PIXEL_FORMAT_32BPP_PARGB: u32 = 0x000e_200b;

/// Undo the premultiplication of a premultiplied ARGB pixel.
fn unpremultiply(value: u32) -> u32 {
    let alpha = value >> 24;
    if alpha == 0 {
        return 0;
    }
    let channel = |shift: u32| (((value >> shift) & 0xff) * 255 / alpha).min(255) << shift;
    (alpha << 24) | channel(16) | channel(8) | channel(0)
}

/// `BrushDataIsGammaCorrected` ([MS-EMFPLUS] 2.1.2.1).
const BRUSH_GAMMA_CORRECTED: u32 = 0x80;

/// `LineCapType` values ([MS-EMFPLUS] 2.1.1.17).
const LINE_CAP_FLAT: i32 = 0;
const LINE_CAP_SQUARE: i32 = 1;
const LINE_CAP_ROUND: i32 = 2;

/// `LineJoinType` values ([MS-EMFPLUS] 2.1.1.18).
const LINE_JOIN_MITER: i32 = 0;
const LINE_JOIN_BEVEL: i32 = 1;
const LINE_JOIN_ROUND: i32 = 2;

/// `LineStyle` values ([MS-EMFPLUS] 2.1.1.19) with a predefined pattern.
const LINE_STYLE_DASH: i32 = 1;
const LINE_STYLE_DOT: i32 = 2;
const LINE_STYLE_DASH_DOT: i32 = 3;
const LINE_STYLE_DASH_DOT_DOT: i32 = 4;

/// `DashedLineCapType` values ([MS-EMFPLUS] 2.1.1.10).
const DASHED_LINE_CAP_FLAT: i32 = 0;
const DASHED_LINE_CAP_ROUND: i32 = 2;

/// `PenAlignmentCenter` ([MS-EMFPLUS] 2.1.1.23).
const PEN_ALIGNMENT_CENTER: i32 = 0;

/// `UnitTypeWorld` ([MS-EMFPLUS] 2.1.1.32).
const PEN_UNIT_WORLD: u32 = 0;

/// The miter limit of a GDI+ pen that does not record one.
const DEFAULT_MITER_LIMIT: f32 = 10.0;

/// The `PathPointFlags` bit GDI+ sets on a path that fills with the winding
/// rule; the classic fallback records of such a path select `WINDING`.
const PATH_WINDING_FILL: u32 = 0x2000;

/// Whether a path object fills with the alternate (even-odd) rule, the GDI+
/// default.
fn path_is_alternate(path: &ObjectPath) -> bool {
    path.flags & PATH_WINDING_FILL == 0
}

/// `FontStyle` flags of an EMF+ font ([MS-EMFPLUS] 2.1.2.4).
const FONT_STYLE_BOLD: i32 = 0x1;
const FONT_STYLE_ITALIC: i32 = 0x2;
const FONT_STYLE_UNDERLINE: i32 = 0x4;
const FONT_STYLE_STRIKEOUT: i32 = 0x8;

/// `DriverStringOptions` flags ([MS-EMFPLUS] 2.1.2.3).
const DRIVER_STRING_CMAP_LOOKUP: u32 = 0x1;
const DRIVER_STRING_VERTICAL: u32 = 0x2;
const DRIVER_STRING_REALIZED_ADVANCE: u32 = 0x4;

/// The text content of one `EmfPlusDrawDriverString` record.
struct DriverString<'a> {
    options: u32,
    glyphs: &'a [u16],
    positions: &'a [Point],
    matrix: Option<Matrix>,
}

/// The source rectangle of an image drawing record.
struct ImageSource {
    rect: Rect,
    unit: Unit,
}

/// The parallelogram an image is drawn into.
///
/// `origin` is the upper-left corner and `u`/`v` are the upper-right and
/// lower-left edges, so the drawing area is `origin + u*s + v*t` for
/// `s, t` in `0..=1`.
#[derive(Clone, Copy, Debug, PartialEq)]
struct ImageFrame {
    origin: Point,
    u: Point,
    v: Point,
}

impl ImageFrame {
    /// The frame that fills an axis-aligned rectangle.
    fn rect(rect: Rect) -> Self {
        // A negative width or height mirrors the image, so the rectangle is
        // kept as recorded.
        Self {
            origin: Point {
                x: rect.x,
                y: rect.y,
            },
            u: Point {
                x: rect.width,
                y: 0.0,
            },
            v: Point {
                x: 0.0,
                y: rect.height,
            },
        }
    }

    /// The frame that maps a source rectangle of a `width` x `height` image.
    fn crop(self, src: Rect, width: f64, height: f64) -> Self {
        let sx = f64::from(src.x) / f64::from(src.width);
        let sy = f64::from(src.y) / f64::from(src.height);
        let su = width / f64::from(src.width);
        let sv = height / f64::from(src.height);
        let origin = Point {
            x: self.origin.x - (self.u.x as f64 * sx) as f32 - (self.v.x as f64 * sy) as f32,
            y: self.origin.y - (self.u.y as f64 * sx) as f32 - (self.v.y as f64 * sy) as f32,
        };
        Self {
            origin,
            u: Point {
                x: (f64::from(self.u.x) * su) as f32,
                y: (f64::from(self.u.y) * su) as f32,
            },
            v: Point {
                x: (f64::from(self.v.x) * sv) as f32,
                y: (f64::from(self.v.y) * sv) as f32,
            },
        }
    }

    /// The frame outline, used as the clip of a cropped image.
    fn path(self) -> SvgPath {
        let corner = |su: f32, sv: f32| SvgPoint {
            x: f64::from(self.origin.x)
                + f64::from(self.u.x) * f64::from(su)
                + f64::from(self.v.x) * f64::from(sv),
            y: f64::from(self.origin.y)
                + f64::from(self.u.y) * f64::from(su)
                + f64::from(self.v.y) * f64::from(sv),
        };
        let mut path = SvgPath::new();
        path.push(SvgPathCommand::MoveTo(corner(0.0, 0.0)));
        path.push(SvgPathCommand::LineTo(corner(1.0, 0.0)));
        path.push(SvgPathCommand::LineTo(corner(1.0, 1.0)));
        path.push(SvgPathCommand::LineTo(corner(0.0, 1.0)));
        path.push(SvgPathCommand::Close);
        path
    }

    /// The element transform of a unit image inside this frame, after `outer`.
    fn transform(self, outer: SvgTransform) -> SvgTransform {
        let inner = SvgTransform {
            a: f64::from(self.u.x),
            b: f64::from(self.u.y),
            c: f64::from(self.v.x),
            d: f64::from(self.v.y),
            e: f64::from(self.origin.x),
            f: f64::from(self.origin.y),
        };
        SvgTransform {
            a: outer.a * inner.a + outer.c * inner.b,
            b: outer.b * inner.a + outer.d * inner.b,
            c: outer.a * inner.c + outer.c * inner.d,
            d: outer.b * inner.c + outer.d * inner.d,
            e: outer.a * inner.e + outer.c * inner.f + outer.e,
            f: outer.b * inner.e + outer.d * inner.f + outer.f,
        }
    }
}

/// The parallelogram defined by the first three destination points.
fn parallelogram(points: &[Point]) -> Option<ImageFrame> {
    let [first, second, third, ..] = points else {
        return None;
    };
    let u = Point {
        x: second.x - first.x,
        y: second.y - first.y,
    };
    let v = Point {
        x: third.x - first.x,
        y: third.y - first.y,
    };
    if (u.x == 0.0 && u.y == 0.0) || (v.x == 0.0 && v.y == 0.0) {
        return None;
    }
    Some(ImageFrame {
        origin: *first,
        u,
        v,
    })
}

/// The pixel size of an embedded image whose header size can be read without
/// decoding it.
fn image_dimensions(mime: SvgImageMime, bytes: &[u8]) -> Option<(f64, f64)> {
    let be = |at: usize| -> Option<u32> {
        Some(u32::from_be_bytes(bytes.get(at..at + 4)?.try_into().ok()?))
    };
    match mime {
        SvgImageMime::Png => {
            if bytes.get(12..16)? != b"IHDR" {
                return None;
            }
            Some((f64::from(be(16)?), f64::from(be(20)?)))
        },
        SvgImageMime::Jpeg => jpeg_dimensions(bytes),
        SvgImageMime::Gif => {
            let width = u16::from_le_bytes(bytes.get(6..8)?.try_into().ok()?);
            let height = u16::from_le_bytes(bytes.get(8..10)?.try_into().ok()?);
            Some((f64::from(width), f64::from(height)))
        },
        SvgImageMime::Webp => None,
    }
}

/// The frame size of a JPEG's start-of-frame marker.
fn jpeg_dimensions(bytes: &[u8]) -> Option<(f64, f64)> {
    let mut at = 2usize;
    while at + 9 < bytes.len() {
        if bytes[at] != 0xFF {
            at += 1;
            continue;
        }
        let marker = bytes[at + 1];
        if marker == 0xFF {
            at += 1;
            continue;
        }
        if (0xD0..=0xD9).contains(&marker) || marker == 0x01 {
            at += 2;
            continue;
        }
        let length = usize::from(u16::from_be_bytes([bytes[at + 2], bytes[at + 3]]));
        if (0xC0..=0xCF).contains(&marker) && marker != 0xC4 && marker != 0xC8 && marker != 0xCC {
            let height = u16::from_be_bytes([bytes[at + 5], bytes[at + 6]]);
            let width = u16::from_be_bytes([bytes[at + 7], bytes[at + 8]]);
            return Some((f64::from(width), f64::from(height)));
        }
        at = at
            .checked_add(2)
            .and_then(|value| value.checked_add(length))?;
    }
    None
}

/// The text of one driver-string glyph entry, and the code units it covers.
///
/// A surrogate pair spans two entries; its position is the one recorded for
/// the first entry.
fn glyph_text(glyphs: &[u16], index: usize) -> (String, usize) {
    let first = glyphs[index];
    let second = glyphs.get(index + 1).copied();
    if (0xd800..0xdc00).contains(&first)
        && let Some(second) = second.filter(|unit| (0xdc00..0xe000).contains(unit))
        && let Some(Ok(value)) = char::decode_utf16([first, second]).next()
    {
        return (value.to_string(), 2);
    }
    (String::from_utf16_lossy(&glyphs[index..=index]), 1)
}

fn compositing(mode: CompositingMode) -> SvgCompositingMode {
    match mode {
        CompositingMode::SourceOver => SvgCompositingMode::SourceOver,
        CompositingMode::SourceCopy => SvgCompositingMode::SourceCopy,
    }
}

fn svg_color(color: Argb) -> SvgColor {
    SvgColor::rgba(color.red(), color.green(), color.blue(), color.alpha())
}

fn svg_point(point: Point) -> SvgPoint {
    SvgPoint {
        x: f64::from(point.x),
        y: f64::from(point.y),
    }
}

fn svg_rect(rect: Rect) -> SvgRect {
    let x2 = rect.x + rect.width;
    let y2 = rect.y + rect.height;
    SvgRect {
        x: f64::from(rect.x.min(x2)),
        y: f64::from(rect.y.min(y2)),
        width: f64::from(rect.width.abs()),
        height: f64::from(rect.height.abs()),
    }
}

fn object_rect(rect: objects::Rect) -> SvgRect {
    svg_rect(Rect {
        x: rect.x,
        y: rect.y,
        width: rect.width,
        height: rect.height,
    })
}

fn svg_transform(matrix: Matrix) -> SvgTransform {
    SvgTransform {
        a: f64::from(matrix.m11),
        b: f64::from(matrix.m12),
        c: f64::from(matrix.m21),
        d: f64::from(matrix.m22),
        e: f64::from(matrix.dx),
        f: f64::from(matrix.dy),
    }
}

fn object_transform(transform: objects::Transform) -> SvgTransform {
    SvgTransform {
        a: f64::from(transform.m11),
        b: f64::from(transform.m12),
        c: f64::from(transform.m21),
        d: f64::from(transform.m22),
        e: f64::from(transform.dx),
        f: f64::from(transform.dy),
    }
}

fn polygon_path(points: &[Point], close: bool) -> SvgPath {
    let mut path = SvgPath::new();
    if let Some((first, rest)) = points.split_first() {
        path.push(SvgPathCommand::MoveTo(svg_point(*first)));
        for point in rest {
            path.push(SvgPathCommand::LineTo(svg_point(*point)));
        }
        if close {
            path.push(SvgPathCommand::Close);
        }
    }
    path
}

fn bezier_path(points: &[Point]) -> SvgPath {
    let mut path = SvgPath::new();
    let Some(first) = points.first() else {
        return path;
    };
    path.push(SvgPathCommand::MoveTo(svg_point(*first)));
    for chunk in points[1..].as_chunks::<3>().0.iter() {
        path.push(SvgPathCommand::CubicTo {
            first: svg_point(chunk[0]),
            second: svg_point(chunk[1]),
            to: svg_point(chunk[2]),
        });
    }
    path
}

/// An arc of the ellipse inscribed in `rect`, optionally closed as a pie.
///
/// GDI+ measures `start` and `sweep` in degrees clockwise from the x axis as
/// angles of the points on the ellipse, not as ellipse parameters.
fn arc_path(rect: Rect, start: f32, sweep: f32, pie: bool) -> SvgPath {
    let center_x = f64::from(rect.x) + f64::from(rect.width) / 2.0;
    let center_y = f64::from(rect.y) + f64::from(rect.height) / 2.0;
    let radius_x = f64::from(rect.width.abs()) / 2.0;
    let radius_y = f64::from(rect.height.abs()) / 2.0;
    let sweep = f64::from(sweep).clamp(-360.0, 360.0);
    let start = f64::from(start).to_radians();
    let end = start + sweep.to_radians();
    // The ellipse parameter of the point at polar angle `angle`, kept within
    // one turn of the angle so that the sweep direction and size survive.
    let parameter = |angle: f64| {
        let value = (radius_x * angle.sin()).atan2(radius_y * angle.cos());
        value + ((angle - value) / std::f64::consts::TAU).round() * std::f64::consts::TAU
    };
    let point_at = |parameter: f64| SvgPoint {
        x: center_x + radius_x * parameter.cos(),
        y: center_y + radius_y * parameter.sin(),
    };
    let (first, last) = (parameter(start), parameter(end));
    let mut path = SvgPath::new();
    if pie {
        path.push(SvgPathCommand::MoveTo(SvgPoint {
            x: center_x,
            y: center_y,
        }));
        path.push(SvgPathCommand::LineTo(point_at(first)));
    } else {
        path.push(SvgPathCommand::MoveTo(point_at(first)));
    }
    // An SVG arc cannot end where it starts, so a full turn is two halves.
    let span = last - first;
    let segments: &[f64] = if span.abs() >= std::f64::consts::TAU - 1e-9 {
        &[0.5, 1.0]
    } else {
        &[1.0]
    };
    let mut from = first;
    for fraction in segments {
        let to = first + span * fraction;
        path.push(SvgPathCommand::ArcTo {
            rx: radius_x,
            ry: radius_y,
            rotation: 0.0,
            large_arc: (to - from).abs() > std::f64::consts::PI,
            sweep: span >= 0.0,
            to: point_at(to),
        });
        from = to;
    }
    if pie {
        path.push(SvgPathCommand::Close);
    }
    path
}

fn svg_object_path(path: &ObjectPath) -> SvgPath {
    let mut output = SvgPath::new();
    let segments = path.segments();
    let mut index = 0usize;
    while index < segments.len() {
        let segment = segments[index];
        match segment.kind {
            0 => output.push(SvgPathCommand::MoveTo(SvgPoint {
                x: f64::from(segment.point.x),
                y: f64::from(segment.point.y),
            })),
            1 => output.push(SvgPathCommand::LineTo(SvgPoint {
                x: f64::from(segment.point.x),
                y: f64::from(segment.point.y),
            })),
            3 if index + 2 < segments.len() => {
                output.push(SvgPathCommand::CubicTo {
                    first: SvgPoint {
                        x: f64::from(segment.point.x),
                        y: f64::from(segment.point.y),
                    },
                    second: SvgPoint {
                        x: f64::from(segments[index + 1].point.x),
                        y: f64::from(segments[index + 1].point.y),
                    },
                    to: SvgPoint {
                        x: f64::from(segments[index + 2].point.x),
                        y: f64::from(segments[index + 2].point.y),
                    },
                });
                index += 2;
            },
            _ => {},
        }
        if segment.flags & 0x80 != 0 {
            output.push(SvgPathCommand::Close);
        }
        index += 1;
    }
    output
}

fn bounds(points: &[Point]) -> Option<Rect> {
    let first = *points.first()?;
    let (mut min_x, mut min_y, mut max_x, mut max_y) = (first.x, first.y, first.x, first.y);
    for point in &points[1..] {
        min_x = min_x.min(point.x);
        min_y = min_y.min(point.y);
        max_x = max_x.max(point.x);
        max_y = max_y.max(point.y);
    }
    Some(Rect {
        x: min_x,
        y: min_y,
        width: max_x - min_x,
        height: max_y - min_y,
    })
}

fn image_mime(bytes: &[u8]) -> Option<SvgImageMime> {
    if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        Some(SvgImageMime::Png)
    } else if bytes.starts_with(&[0xff, 0xd8, 0xff]) {
        Some(SvgImageMime::Jpeg)
    } else if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        Some(SvgImageMime::Gif)
    } else if bytes.len() >= 12 && &bytes[..4] == b"RIFF" && &bytes[8..12] == b"WEBP" {
        Some(SvgImageMime::Webp)
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// `EmfPlusHeader` data of a metafile recorded for a 96 DPI display.
    fn video_header() -> Vec<u8> {
        [0xdbc0_1002_u32, 1, 96, 96]
            .iter()
            .flat_map(|value| value.to_le_bytes())
            .collect()
    }

    fn record(kind: RecordType, flags: u16, data: &[u8]) -> Vec<u8> {
        let size = 12_u32 + u32::try_from(data.len()).unwrap();
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&kind.raw().to_le_bytes());
        bytes.extend_from_slice(&flags.to_le_bytes());
        bytes.extend_from_slice(&size.to_le_bytes());
        bytes.extend_from_slice(&u32::try_from(data.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(data);
        bytes
    }

    fn comment(payload: &[u8]) -> Vec<u8> {
        let data_size = 4_u32 + u32::try_from(payload.len()).unwrap();
        let mut body = Vec::new();
        body.extend_from_slice(&data_size.to_le_bytes());
        body.extend_from_slice(&super::super::EMFPLUS_COMMENT_IDENTIFIER.to_le_bytes());
        body.extend_from_slice(payload);
        while (body.len() + 8) % 4 != 0 {
            body.push(0);
        }
        body
    }

    #[test]
    fn direct_comment_api_returns_svg_and_fragment() {
        let mut payload = record(RecordType::Header, 1, &video_header());
        payload.extend(record(
            RecordType::Clear,
            0,
            &0xff_11_22_33_u32.to_le_bytes(),
        ));
        payload.extend(record(RecordType::EndOfFile, 0, &[]));
        let comment = comment(&payload);
        let output =
            render_emfplus_comments_to_svg(40.0, 30.0, [&comment[..]], RendererLimits::default())
                .unwrap();

        assert_eq!(output.mux.kind, Some(MetafileKind::Dual));
        assert!(!output.mux.classic_emf_enabled);
        assert!(output.svg().contains("<rect"));
        assert!(output.fragment().body().contains("width=\"40\""));
    }

    #[test]
    fn get_dc_mux_interval_is_reported_for_emfplus_only() {
        let mut renderer = EmfPlusSvgRenderer::new(10.0, 10.0, RendererLimits::default()).unwrap();
        let mut first = record(RecordType::Header, 0, &video_header());
        first.extend(record(RecordType::GetDc, 0, &[]));
        let update = renderer.push_comment_body(&comment(&first)).unwrap();
        assert!(update.mux.classic_emf_enabled);
        assert_eq!(update.mux.get_dc_count, 1);

        let eof = record(RecordType::EndOfFile, 0, &[]);
        let update = renderer.push_comment_body(&comment(&eof)).unwrap();
        assert!(!update.mux.classic_emf_enabled);
        assert_eq!(
            renderer.finish().unwrap().mux.kind,
            Some(MetafileKind::EmfPlusOnly)
        );
    }

    fn words(values: &[u32]) -> Vec<u8> {
        values
            .iter()
            .flat_map(|value| value.to_le_bytes())
            .collect()
    }

    fn floats(values: &[f32]) -> Vec<u8> {
        values
            .iter()
            .flat_map(|value| value.to_le_bytes())
            .collect()
    }

    /// Render `records` between a display header and EndOfFile into a
    /// 100 x 100 picture.
    fn render_records(records: &[Vec<u8>]) -> RenderOutput {
        let mut payload = record(RecordType::Header, 1, &video_header());
        for value in records {
            payload.extend_from_slice(value);
        }
        payload.extend(record(RecordType::EndOfFile, 0, &[]));
        render_emfplus_comments_to_svg(
            100.0,
            100.0,
            [&comment(&payload)[..]],
            RendererLimits::default(),
        )
        .unwrap()
    }

    fn codes(output: &RenderOutput) -> Vec<&'static str> {
        output
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.code)
            .collect()
    }

    /// A black pen object in `slot` with `flags`, `width` and the optional
    /// fields that `flags` selects, in record order.
    fn pen_object(slot: u16, flags: u32, width: f32, optional: &[u8]) -> Vec<u8> {
        let mut data = words(&[0xdbc0_1002, 0, flags, 0]);
        data.extend(floats(&[width]));
        data.extend_from_slice(optional);
        data.extend(words(&[0xdbc0_1002, 0, 0xff00_0000]));
        record(RecordType::Object, 0x0200 | slot, &data)
    }

    /// A path object in `slot` with one closed triangle.
    fn path_object(slot: u16, flags: u32, points: [[f32; 2]; 3]) -> Vec<u8> {
        let mut data = words(&[0xdbc0_1002, 3, flags]);
        for point in points {
            data.extend(floats(&point));
        }
        data.extend_from_slice(&[0, 1, 0x81, 0]);
        record(RecordType::Object, 0x0300 | slot, &data)
    }

    fn fill_rect(color: u32, rect: [f32; 4]) -> Vec<u8> {
        let mut data = words(&[color, 1]);
        data.extend(floats(&rect));
        record(RecordType::FillRects, 0x8000, &data)
    }

    #[test]
    fn view_origin_moves_the_view_box_and_the_clear_rectangle() {
        let mut payload = record(RecordType::Header, 1, &video_header());
        payload.extend(record(RecordType::Clear, 0, &0xff11_2233_u32.to_le_bytes()));
        payload.extend(record(RecordType::EndOfFile, 0, &[]));
        let mut renderer = EmfPlusSvgRenderer::new(40.0, 30.0, RendererLimits::default())
            .unwrap()
            .with_view_origin(5.0, -7.0);
        renderer.push_comment_body(&comment(&payload)).unwrap();
        let svg = renderer.finish().unwrap().document.into_string();
        assert!(svg.contains("viewBox=\"5 -7 40 30\""), "{svg}");
        assert!(svg.contains("<rect x=\"5\" y=\"-7\" width=\"40\""), "{svg}");
    }

    #[test]
    fn page_transform_scales_world_coordinates_to_device_pixels() {
        // Millimeters at 96 DPI with a page scale of two.
        let output = render_records(&[
            record(RecordType::SetPageTransform, 6, &floats(&[2.0])),
            fill_rect(0xff00_0000, [1.0, 1.0, 2.0, 2.0]),
        ]);
        let scale = 2.0 * 96.0 / 25.4;
        let body = output.fragment().body();
        assert!(
            body.contains(&format!("transform=\"matrix({scale} 0 0 {scale} 0 0)\"")),
            "{body}"
        );
    }

    #[test]
    fn display_page_unit_of_a_printer_reference_is_a_hundredth_inch() {
        let mut payload = record(RecordType::Header, 1, &words(&[0xdbc0_1002, 0, 600, 600]));
        payload.extend(fill_rect(0xff00_0000, [0.0, 0.0, 1.0, 1.0]));
        payload.extend(record(RecordType::EndOfFile, 0, &[]));
        let output = render_emfplus_comments_to_svg(
            10.0,
            10.0,
            [&comment(&payload)[..]],
            RendererLimits::default(),
        )
        .unwrap();
        assert!(
            output
                .fragment()
                .body()
                .contains("transform=\"matrix(6 0 0 6 0 0)\""),
            "{}",
            output.fragment().body()
        );
    }

    #[test]
    fn pen_caps_joins_and_dashes_follow_the_gdiplus_enumerations() {
        // StartCap round, EndCap round, LineJoin round, DashedLineCap round,
        // and a custom 2-on 1-off dash pattern in pen widths.
        let mut optional = words(&[2, 2, 2, 2]);
        optional.extend(words(&[2]));
        optional.extend(floats(&[2.0, 1.0]));
        let pen = pen_object(0, 0x2 | 0x4 | 0x8 | 0x40 | 0x100, 3.0, &optional);
        let mut line = words(&[2]);
        line.extend(floats(&[0.0, 0.0, 50.0, 0.0]));
        let output = render_records(&[pen, record(RecordType::DrawLines, 0, &line)]);
        let body = output.fragment().body();
        assert!(
            body.contains("stroke-width=\"3\" stroke-linecap=\"round\" stroke-linejoin=\"round\""),
            "{body}"
        );
        assert!(body.contains("stroke-dasharray=\"6 3\""), "{body}");
        assert!(output.diagnostics.is_empty(), "{:?}", output.diagnostics);
    }

    #[test]
    fn default_pen_uses_the_gdiplus_miter_limit_and_square_caps_are_square() {
        let pen = pen_object(0, 0x2 | 0x4, 1.0, &words(&[1, 1]));
        let mut line = words(&[2]);
        line.extend(floats(&[0.0, 0.0, 50.0, 0.0]));
        let output = render_records(&[pen, record(RecordType::DrawLines, 0, &line)]);
        let body = output.fragment().body();
        assert!(body.contains("stroke-linecap=\"square\""), "{body}");
        assert!(body.contains("stroke-miterlimit=\"10\""), "{body}");
    }

    #[test]
    fn zero_width_pen_is_one_device_pixel_wide() {
        let pen = pen_object(0, 0, 0.0, &[]);
        let mut line = words(&[2]);
        line.extend(floats(&[0.0, 0.0, 10.0, 0.0]));
        let output = render_records(&[
            record(RecordType::ScaleWorldTransform, 0, &floats(&[4.0, 4.0])),
            pen,
            record(RecordType::DrawLines, 0, &line),
        ]);
        let body = output.fragment().body();
        assert!(body.contains("stroke-width=\"0.25\""), "{body}");
    }

    #[test]
    fn unrepresentable_pen_features_are_reported() {
        // StartCap flat, EndCap triangle, PenAlignmentInset.
        let pen = pen_object(0, 0x2 | 0x4 | 0x200, 1.0, &words(&[0, 3, 1]));
        let mut line = words(&[2]);
        line.extend(floats(&[0.0, 0.0, 50.0, 0.0]));
        let output = render_records(&[pen, record(RecordType::DrawLines, 0, &line)]);
        let codes = codes(&output);
        assert!(codes.contains(&"pen_line_caps"), "{codes:?}");
        assert!(codes.contains(&"pen_alignment"), "{codes:?}");
    }

    #[test]
    fn fill_path_uses_the_fill_rule_recorded_on_the_path() {
        let triangle = [[0.0, 0.0], [10.0, 0.0], [0.0, 10.0]];
        let alternate = render_records(&[
            path_object(0, 0, triangle),
            record(RecordType::FillPath, 0x8000, &0xff00_0000_u32.to_le_bytes()),
        ]);
        assert!(
            alternate
                .fragment()
                .body()
                .contains("fill-rule=\"evenodd\"")
        );
        let winding = render_records(&[
            path_object(0, PATH_WINDING_FILL, triangle),
            record(RecordType::FillPath, 0x8000, &0xff00_0000_u32.to_le_bytes()),
        ]);
        assert!(!winding.fragment().body().contains("fill-rule"));
    }

    #[test]
    fn clip_keeps_the_transform_and_path_in_effect_when_it_was_set() {
        let output = render_records(&[
            path_object(0, 0, [[1.0, 1.0], [9.0, 1.0], [1.0, 9.0]]),
            record(RecordType::SetClipPath, 0, &[]),
            // The slot is reused and the transform changes before drawing.
            path_object(0, 0, [[50.0, 50.0], [60.0, 50.0], [50.0, 60.0]]),
            record(RecordType::ScaleWorldTransform, 0, &floats(&[2.0, 2.0])),
            fill_rect(0xff00_0000, [0.0, 0.0, 5.0, 5.0]),
        ]);
        let defs = output.fragment().defs();
        let body = output.fragment().body();
        assert!(defs.contains("d=\"M1 1L9 1L1 9Z\""), "{defs}");
        assert!(defs.contains("transform=\"matrix(1 0 0 1 0 0)\""), "{defs}");
        // The clip applies in device space, outside the element transform.
        assert!(
            body.starts_with("<g clip-path=\"url(#") && body.contains("matrix(2 0 0 2 0 0)"),
            "{body}"
        );
    }

    #[test]
    fn intersected_clips_nest_their_definitions() {
        let mut first = floats(&[0.0, 0.0, 50.0, 50.0]);
        let second = floats(&[10.0, 10.0, 50.0, 50.0]);
        let output = render_records(&[
            record(RecordType::SetClipRect, 0, &std::mem::take(&mut first)),
            // Rectangles set under one transform would merge instead.
            record(RecordType::ScaleWorldTransform, 0, &floats(&[2.0, 2.0])),
            // CombineModeIntersect in the second flag byte.
            record(RecordType::SetClipRect, 0x0100, &second),
            fill_rect(0xff00_0000, [0.0, 0.0, 100.0, 100.0]),
        ]);
        let defs = output.fragment().defs();
        assert!(
            defs.contains("<clipPath id=\"emfplus-clip-2\" clip-path=\"url(#emfplus-clip-1)\">"),
            "{defs}"
        );
    }

    #[test]
    fn intersected_rectangles_in_one_space_merge_and_other_clips_stop_at_a_depth() {
        let rect = floats(&[0.0, 0.0, 50.0, 50.0]);
        let mut records = vec![record(RecordType::SetClipRect, 0, &rect)];
        records.extend(std::iter::repeat_n(
            record(
                RecordType::SetClipRect,
                0x0100,
                &floats(&[10.0, 0.0, 50.0, 50.0]),
            ),
            100_000,
        ));
        records.push(fill_rect(0xff00_0000, [0.0, 0.0, 100.0, 100.0]));
        let output = render_records(&records);
        let defs = output.fragment().defs();
        assert_eq!(defs.matches("<clipPath").count(), 1, "{defs}");
        assert!(
            defs.contains("x=\"10\" y=\"0\" width=\"40\" height=\"50\""),
            "{defs}"
        );
        assert!(codes(&output).is_empty());

        // Rectangles set under other transforms do not merge.
        let mut records = vec![record(RecordType::SetClipRect, 0, &rect)];
        for _ in 0..100_000 {
            records.push(record(
                RecordType::TranslateWorldTransform,
                0,
                &floats(&[1.0, 0.0]),
            ));
            records.push(record(RecordType::SetClipRect, 0x0100, &rect));
            records.push(record(RecordType::OffsetClip, 0, &floats(&[0.0, 0.0])));
        }
        records.push(fill_rect(0xff00_0000, [0.0, 0.0, 100.0, 100.0]));
        let output = render_records(&records);
        let defs = output.fragment().defs();
        assert_eq!(
            defs.matches("<clipPath").count(),
            crate::emfplus::playback::MAX_CLIP_INTERSECTIONS
        );
        assert!(codes(&output).contains(&"clip_depth_limit"));
    }

    #[test]
    fn clip_from_a_missing_object_is_ignored_as_noncanonical() {
        let output = render_records(&[
            record(RecordType::SetClipPath, 5, &[]),
            fill_rect(0xff00_0000, [0.0, 0.0, 5.0, 5.0]),
        ]);
        assert!(!output.fragment().body().contains("clip-path"));
        assert_eq!(
            codes(&output),
            vec!["noncanonical-emfplus-missing-clip-object"]
        );
    }

    #[test]
    fn linear_gradient_runs_across_its_rectangle_and_tiles() {
        // Brush version, LinearGradient type, BrushDataFlags, WrapModeTile,
        // RectF, start and end colors, reserved.
        let mut brush = words(&[0xdbc0_1002, 4, 0, 0]);
        brush.extend(floats(&[10.0, 20.0, 30.0, 40.0]));
        brush.extend(words(&[0xffff_0000, 0xff00_00ff, 0, 0]));
        let mut fill = words(&[0, 1]);
        fill.extend(floats(&[0.0, 0.0, 100.0, 100.0]));
        let output = render_records(&[
            record(RecordType::Object, 0x0100, &brush),
            record(RecordType::FillRects, 0, &fill),
        ]);
        let defs = output.fragment().defs();
        assert!(
            defs.contains(
                "x1=\"10\" y1=\"20\" x2=\"40\" y2=\"20\" gradientUnits=\"userSpaceOnUse\" spreadMethod=\"repeat\""
            ),
            "{defs}"
        );
    }

    #[test]
    fn arc_angles_are_angles_of_points_on_the_ellipse() {
        let pen = pen_object(0, 0, 1.0, &[]);
        // Start 0 and sweep 45 degrees on a 200 x 100 ellipse; the end point is
        // where the 45 degree ray meets the ellipse.
        let mut arc = floats(&[0.0, 45.0]);
        arc.extend(floats(&[0.0, 0.0, 200.0, 100.0]));
        let output = render_records(&[pen, record(RecordType::DrawArc, 0, &arc)]);
        let body = output.fragment().body();
        let end = 100.0 * 50.0 / 100.0_f64.hypot(50.0);
        let expected = format!("{} {}", 100.0 + end, 50.0 + end);
        let path = body.split("d=\"").nth(1).unwrap();
        let tail: Vec<f64> = path
            .split('"')
            .next()
            .unwrap()
            .split(|c: char| c == ' ' || c.is_ascii_alphabetic())
            .filter(|value| !value.is_empty())
            .map(|value| value.parse().unwrap())
            .collect();
        let (x, y) = (tail[tail.len() - 2], tail[tail.len() - 1]);
        assert!(
            (x - (100.0 + end)).abs() < 1e-3 && (y - (50.0 + end)).abs() < 1e-3,
            "{expected} vs {body}"
        );
    }

    #[test]
    fn full_circle_arc_is_drawn_as_two_halves() {
        let path = arc_path(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 10.0,
            },
            0.0,
            360.0,
            false,
        );
        let arcs = path
            .commands()
            .iter()
            .filter(|command| matches!(command, SvgPathCommand::ArcTo { .. }))
            .count();
        assert_eq!(arcs, 2);
    }

    #[test]
    fn draw_string_layout_is_reported() {
        let mut text = words(&[0xff00_0000, 0xffff_ffff, 2]);
        text.extend(floats(&[0.0, 0.0, 50.0, 20.0]));
        for unit in "Hi".encode_utf16() {
            text.extend_from_slice(&unit.to_le_bytes());
        }
        let output =
            render_records(&[font_object(), record(RecordType::DrawString, 0x8000, &text)]);
        assert!(codes(&output).contains(&"draw_string_layout"));
        assert!(output.fragment().body().contains(">Hi</text>"));
    }

    /// An uncompressed bitmap image object in slot 1.
    fn raw_bitmap_object(width: u32, pixel_format: u32, pixels: &[u8]) -> Vec<u8> {
        let stride = (pixels.len() as u32).max(4);
        let mut data = words(&[0xdbc0_1002, 1, width, 1, stride, pixel_format, 0]);
        data.extend_from_slice(pixels);
        while data.len() % 4 != 0 {
            data.push(0);
        }
        record(RecordType::Object, 0x0500 | 1, &data)
    }

    fn draw_bitmap(width: f32) -> Vec<u8> {
        let mut draw = words(&[u32::MAX, 2]);
        draw.extend(floats(&[0.0, 0.0, width, 1.0, 0.0, 0.0, 10.0, 10.0]));
        record(RecordType::DrawImage, 1, &draw)
    }

    fn embedded_png(output: &RenderOutput) -> image::RgbaImage {
        let body = output.fragment().body();
        let data = body
            .split("data:image/png;base64,")
            .nth(1)
            .and_then(|rest| rest.split('"').next())
            .unwrap_or_else(|| panic!("{body}"));
        let bytes =
            base64::Engine::decode(&base64::engine::general_purpose::STANDARD, data).unwrap();
        image::load_from_memory(&bytes).unwrap().to_rgba8()
    }

    #[test]
    fn raw_rgb_bitmap_is_embedded_as_png() {
        // Two 24-bit pixels in blue, green, red order: red, then blue.
        let output = render_records(&[
            raw_bitmap_object(2, PIXEL_FORMAT_24BPP_RGB, &[0, 0, 255, 255, 0, 0, 0, 0]),
            draw_bitmap(2.0),
        ]);
        let image = embedded_png(&output);
        assert_eq!(image.get_pixel(0, 0).0, [255, 0, 0, 255]);
        assert_eq!(image.get_pixel(1, 0).0, [0, 0, 255, 255]);
        assert!(output.diagnostics.is_empty(), "{:?}", output.diagnostics);
    }

    #[test]
    fn raw_indexed_bitmap_uses_its_palette() {
        // PaletteStyleFlags, PaletteCount, then black and white, then one row
        // of 1-bit pixels: white, black.
        let mut pixels = words(&[0, 2, 0xff00_0000, 0xffff_ffff]);
        pixels.extend_from_slice(&[0b1000_0000, 0, 0, 0]);
        let output = render_records(&[
            raw_bitmap_object(2, PIXEL_FORMAT_1BPP_INDEXED, &pixels),
            draw_bitmap(2.0),
        ]);
        let image = embedded_png(&output);
        assert_eq!(image.get_pixel(0, 0).0, [255, 255, 255, 255]);
        assert_eq!(image.get_pixel(1, 0).0, [0, 0, 0, 255]);
    }

    /// A font object in slot 0 with a four-unit family name.
    fn font_object() -> Vec<u8> {
        let mut data = Vec::new();
        for value in [0xdbc0_1002_u32, 16.0_f32.to_bits(), 0, 0, 0, 4] {
            data.extend_from_slice(&value.to_le_bytes());
        }
        for unit in "Aria".encode_utf16() {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        record(
            RecordType::Object,
            0x0600, // ObjectID 0, ObjectType Font
            &data,
        )
    }

    /// Render one red `AV` driver string at (10, 20) and (40, 20).
    fn render_driver_string(options: u32, matrix: Option<[f32; 6]>) -> RenderOutput {
        let mut payload = record(RecordType::Header, 1, &video_header());
        payload.extend(font_object());
        // BrushId (solid red), Options, MatrixPresent, GlyphCount.
        let mut draw = Vec::new();
        draw.extend_from_slice(&0xffff_0000_u32.to_le_bytes());
        draw.extend_from_slice(&options.to_le_bytes());
        draw.extend_from_slice(&u32::from(matrix.is_some()).to_le_bytes());
        draw.extend_from_slice(&2_u32.to_le_bytes());
        for glyph in [b'A' as u16, b'V' as u16] {
            draw.extend_from_slice(&glyph.to_le_bytes());
        }
        for (x, y) in [(10.0_f32, 20.0_f32), (40.0, 20.0)] {
            draw.extend_from_slice(&x.to_le_bytes());
            draw.extend_from_slice(&y.to_le_bytes());
        }
        for value in matrix.into_iter().flatten() {
            draw.extend_from_slice(&value.to_le_bytes());
        }
        payload.extend(record(RecordType::DrawDriverString, 0x8000, &draw));
        payload.extend(record(RecordType::EndOfFile, 0, &[]));
        render_emfplus_comments_to_svg(
            100.0,
            50.0,
            [&comment(&payload)[..]],
            RendererLimits::default(),
        )
        .unwrap()
    }

    #[test]
    fn driver_string_glyphs_keep_their_own_positions() {
        let output = render_driver_string(DRIVER_STRING_CMAP_LOOKUP, None);
        let body = output.fragment().body();
        assert!(
            body.contains("x=\"10 40\" y=\"20 20\" xml:space=\"preserve\" font-family=\"Aria\"")
                && body.contains(">AV</text>"),
            "{body}"
        );
        assert!(output.diagnostics.is_empty(), "{:?}", output.diagnostics);
    }

    #[test]
    fn driver_string_realized_advance_uses_the_first_position() {
        let output = render_driver_string(
            DRIVER_STRING_CMAP_LOOKUP | DRIVER_STRING_REALIZED_ADVANCE,
            None,
        );
        let body = output.fragment().body();
        assert!(body.contains("<text x=\"10\" y=\"20\""), "{body}");
    }

    #[test]
    fn driver_string_matrix_applies_before_the_world_transform() {
        let output = render_driver_string(
            DRIVER_STRING_CMAP_LOOKUP,
            Some([2.0, 0.0, 0.0, 2.0, 5.0, 0.0]),
        );
        let body = output.fragment().body();
        assert!(body.contains("transform=\"matrix(2 0 0 2 5 0)\""), "{body}");
    }

    #[test]
    fn driver_string_glyph_indices_are_reported_and_not_drawn() {
        let output = render_driver_string(0, None);
        assert!(!output.fragment().body().contains("<text"));
        assert!(
            output
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.code == "driver_string_glyph_indices")
        );
    }

    /// A PNG header with only the signature and `IHDR` chunk, which is all the
    /// renderer reads to place a compressed image.
    fn png_header(width: u32, height: u32) -> Vec<u8> {
        let mut bytes = b"\x89PNG\r\n\x1a\n".to_vec();
        bytes.extend_from_slice(&13_u32.to_be_bytes());
        bytes.extend_from_slice(b"IHDR");
        bytes.extend_from_slice(&width.to_be_bytes());
        bytes.extend_from_slice(&height.to_be_bytes());
        bytes.extend_from_slice(&[8, 6, 0, 0, 0]);
        bytes.extend_from_slice(&[0; 4]);
        bytes
    }

    /// A compressed image object in slot 1.
    fn image_object(png: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        for value in [0xdbc0_1002_u32, 1, 2, 2, 8, 0, 1] {
            data.extend_from_slice(&value.to_le_bytes());
        }
        data.extend_from_slice(png);
        while data.len() % 4 != 0 {
            data.push(0);
        }
        record(
            RecordType::Object,
            0x0500 | 1, // ObjectID 1, ObjectType Image
            &data,
        )
    }

    fn image_points_record(src: [f32; 4], points: &[[f32; 2]]) -> Vec<u8> {
        let mut draw = Vec::new();
        draw.extend_from_slice(&u32::MAX.to_le_bytes());
        draw.extend_from_slice(&2_u32.to_le_bytes());
        for value in src {
            draw.extend_from_slice(&value.to_le_bytes());
        }
        draw.extend_from_slice(&u32::try_from(points.len()).unwrap().to_le_bytes());
        for point in points {
            for value in point {
                draw.extend_from_slice(&value.to_le_bytes());
            }
        }
        record(RecordType::DrawImagePoints, 1, &draw)
    }

    fn render_image_points(src: [f32; 4], points: &[[f32; 2]]) -> RenderOutput {
        let mut payload = record(RecordType::Header, 1, &video_header());
        payload.extend(image_object(&png_header(2, 2)));
        payload.extend(image_points_record(src, points));
        payload.extend(record(RecordType::EndOfFile, 0, &[]));
        render_emfplus_comments_to_svg(
            100.0,
            100.0,
            [&comment(&payload)[..]],
            RendererLimits::default(),
        )
        .unwrap()
    }

    #[test]
    fn image_points_draw_their_parallelogram() {
        let output = render_image_points(
            [0.0, 0.0, 2.0, 2.0],
            &[[10.0, 20.0], [30.0, 20.0], [10.0, 50.0]],
        );
        let body = output.fragment().body();
        assert!(body.contains("preserveAspectRatio=\"none\""), "{body}");
        assert!(
            body.contains("transform=\"matrix(20 0 0 30 10 20)\""),
            "{body}"
        );
        assert!(
            !output
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.code == "image_parallelogram")
        );
    }

    #[test]
    fn cropped_image_source_is_scaled_and_clipped() {
        let output = render_image_points(
            [0.0, 0.0, 1.0, 1.0],
            &[[10.0, 20.0], [30.0, 20.0], [10.0, 50.0]],
        );
        let body = output.fragment().body();
        assert!(
            body.contains("transform=\"matrix(40 0 0 60 10 20)\""),
            "{body}"
        );
        assert!(body.contains("clip-path=\"url(#"), "{body}");
        assert!(output.fragment().defs().contains("<clipPath"), "{body}");
    }

    #[test]
    fn short_ordinary_comment_is_ignored() {
        let mut renderer = EmfPlusSvgRenderer::new(10.0, 10.0, RendererLimits::default()).unwrap();
        assert!(
            !renderer
                .push_comment_body(&0_u32.to_le_bytes())
                .unwrap()
                .was_emfplus
        );
    }
}
