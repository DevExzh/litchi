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
    Brush, Clip, CombineMode, CompositingMode, DrawCommand, DrawCommandKind, GraphicsState, Matrix,
    PlaybackEngine, PlaybackLimits, Point, Rect,
};
use super::svg::{
    SvgBuilder, SvgColor, SvgCompositingMode, SvgDocument, SvgGradientStop, SvgId, SvgImage,
    SvgImageMime, SvgImageSource, SvgLineCap, SvgLineJoin, SvgLinearGradient, SvgPaint, SvgPath,
    SvgPathCommand, SvgPoint, SvgRect, SvgStroke, SvgStyle, SvgText, SvgTransform,
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
        })
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
        } else if record.header.record_type == RecordType::GetDc {
            self.mux.get_dc_count = self.mux.get_dc_count.saturating_add(1);
            self.mux.classic_emf_enabled = self.mux.kind == Some(MetafileKind::EmfPlusOnly);
        }
    }

    fn render_command(&mut self, command: &DrawCommand) -> Result<()> {
        let transform = Some(svg_transform(command.state.world_transform));
        let clip = self.clip_id(&command.state)?;
        match &command.kind {
            DrawCommandKind::Clear { color } => {
                let style =
                    self.fill_style(Brush::Solid(*color), &command.state, command.offset)?;
                self.svg.rect(
                    SvgRect {
                        x: 0.0,
                        y: 0.0,
                        width: self.width,
                        height: self.height,
                    },
                    &style,
                    None,
                    None,
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
                if *fill_mode == 0 {
                    self.diag(
                        "alternate_fill_rule",
                        "SVG emitter uses its safe default fill rule",
                        command.offset,
                    );
                }
                let path = polygon_path(points, true);
                let style = self.fill_style(*brush, &command.state, command.offset)?;
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
                let path = self.object_path(*path)?;
                let style = self.fill_style(*brush, &command.state, command.offset)?;
                self.svg.path(&path, &style, transform, clip.as_ref())?;
            },
            DrawCommandKind::DrawPath { pen, path } => {
                let path = self.object_path(*path)?;
                let style = self.stroke_style(*pen, &command.state, command.offset)?;
                self.svg.path(&path, &style, transform, clip.as_ref())?;
            },
            DrawCommandKind::FillClosedCurve { brush, points, .. } => {
                self.diag(
                    "curve_approximation",
                    "cardinal curve rendered as a closed polygon",
                    command.offset,
                );
                let style = self.fill_style(*brush, &command.state, command.offset)?;
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
            DrawCommandKind::DrawImage { image, dest, .. } => {
                self.render_image(*image, *dest, transform, clip.as_ref(), command.offset)?;
            },
            DrawCommandKind::DrawImagePoints { image, points, .. } => {
                let dest = bounds(points).ok_or_else(|| {
                    Error::ParseError("DrawImagePoints has no destination".into())
                })?;
                self.diag(
                    "image_parallelogram",
                    "image parallelogram/cropping represented by its bounds",
                    command.offset,
                );
                self.render_image(*image, dest, transform, clip.as_ref(), command.offset)?;
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
                    Point {
                        x: layout.x,
                        y: layout.y + layout.height,
                    },
                    &command.state,
                    clip.as_ref(),
                    command.offset,
                )?;
            },
            DrawCommandKind::DrawDriverString {
                brush,
                font,
                glyphs,
                positions,
                ..
            } => {
                let text = String::from_utf16_lossy(glyphs);
                let origin = positions.first().copied().unwrap_or_default();
                if positions.len() > 1 {
                    self.diag(
                        "driver_string_positions",
                        "per-glyph positions represented by one text origin",
                        command.offset,
                    );
                }
                self.render_text(
                    *brush,
                    *font,
                    text,
                    origin,
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
        let line_cap = match pen.start_cap.unwrap_or(0) {
            2 => SvgLineCap::Square,
            1 | 3 => SvgLineCap::Round,
            _ => SvgLineCap::Butt,
        };
        let line_join = match pen.join.unwrap_or(0) {
            2 => SvgLineJoin::Bevel,
            1 => SvgLineJoin::Round,
            _ => SvgLineJoin::Miter,
        };
        Ok(SvgStyle {
            fill: SvgPaint::None,
            stroke: Some(SvgStroke {
                paint,
                width: f64::from(pen.width.abs()),
                line_cap,
                line_join,
                dashes: pen
                    .dashes
                    .iter()
                    .map(|value| f64::from(value.abs()))
                    .collect(),
                dash_offset: f64::from(pen.dash_offset.unwrap_or(0.0)),
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
                let start = SvgPoint {
                    x: f64::from(gradient.rect.x),
                    y: f64::from(gradient.rect.y),
                };
                let end = SvgPoint {
                    x: f64::from(gradient.rect.x + gradient.rect.width),
                    y: f64::from(gradient.rect.y + gradient.rect.height),
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

    fn object_path(&self, id: ObjectId) -> Result<SvgPath> {
        let GraphicsObject::Path(path) = self.decoded(id)? else {
            return Err(Error::ParseError(
                "EMF+ path reference has the wrong object type".into(),
            ));
        };
        Ok(svg_object_path(&path))
    }

    fn clip_id(&mut self, state: &GraphicsState) -> Result<Option<SvgId>> {
        match &state.clip {
            Clip::Infinite => Ok(None),
            Clip::Rect { mode, rect } => {
                self.combine_diagnostic(*mode);
                Ok(Some(self.svg.define_clip_rect(
                    svg_rect(*rect),
                    Some(svg_transform(state.world_transform)),
                )?))
            },
            Clip::Path { mode, path } => {
                self.combine_diagnostic(*mode);
                let path = self.object_path(*path)?;
                Ok(Some(self.svg.define_clip_path(
                    &path,
                    Some(svg_transform(state.world_transform)),
                )?))
            },
            Clip::Region { mode, region } => {
                self.combine_diagnostic(*mode);
                let GraphicsObject::Region(region) = self.decoded(*region)? else {
                    return Err(Error::ParseError(
                        "EMF+ clip region has the wrong object type".into(),
                    ));
                };
                match region.root {
                    RegionNode::Rect(rect) => Ok(Some(self.svg.define_clip_rect(
                        object_rect(rect),
                        Some(svg_transform(state.world_transform)),
                    )?)),
                    RegionNode::Path(path) => Ok(Some(self.svg.define_clip_path(
                        &svg_object_path(&path),
                        Some(svg_transform(state.world_transform)),
                    )?)),
                    RegionNode::Infinite => Ok(None),
                    RegionNode::And(_, _)
                    | RegionNode::Or(_, _)
                    | RegionNode::Xor(_, _)
                    | RegionNode::Exclude(_, _)
                    | RegionNode::Complement(_, _)
                    | RegionNode::Empty => {
                        self.diag(
                            "complex_region_clip",
                            "complex Boolean region clip was not applied",
                            0,
                        );
                        Ok(None)
                    },
                }
            },
            Clip::Offset { clip, .. } => {
                self.diag(
                    "offset_clip",
                    "offset clip represented without its offset",
                    0,
                );
                let nested = GraphicsState {
                    clip: (**clip).clone(),
                    ..state.clone()
                };
                self.clip_id(&nested)
            },
        }
    }

    fn combine_diagnostic(&mut self, mode: CombineMode) {
        if !matches!(mode, CombineMode::Replace | CombineMode::Intersect) {
            self.diag(
                "clip_combine_approximation",
                "SVG clip cannot represent this Boolean combine mode",
                0,
            );
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
                self.svg
                    .path(&svg_object_path(&path), style, transform, clip)
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
        dest: Rect,
        transform: Option<SvgTransform>,
        clip: Option<&SvgId>,
        offset: usize,
    ) -> Result<()> {
        let GraphicsObject::Image(image) = self.decoded(id)? else {
            return Err(Error::ParseError(
                "EMF+ image reference has the wrong object type".into(),
            ));
        };
        let ImageKind::Compressed(bytes) = image.kind else {
            self.diag(
                "unsupported_image",
                "raw bitmap or nested metafile image requires a raster decoder",
                offset,
            );
            return Ok(());
        };
        let Some(mime) = image_mime(&bytes) else {
            self.diag(
                "unknown_image_encoding",
                "compressed image signature is unsupported",
                offset,
            );
            return Ok(());
        };
        self.svg.image(
            &SvgImage {
                rect: svg_rect(dest),
                source: SvgImageSource::Embedded { mime, bytes },
                transform,
                opacity: 1.0,
            },
            clip,
        )
    }

    fn render_text(
        &mut self,
        brush: Brush,
        font: ObjectId,
        value: String,
        origin: Point,
        state: &GraphicsState,
        clip: Option<&SvgId>,
        offset: usize,
    ) -> Result<()> {
        let GraphicsObject::Font(font) = self.decoded(font)? else {
            return Err(Error::ParseError(
                "EMF+ font reference has the wrong object type".into(),
            ));
        };
        let style = self.fill_style(brush, state, offset)?;
        self.svg.text(
            &SvgText {
                value,
                origin: svg_point(origin),
                font_family: font.family,
                font_size: f64::from(font.em_size.abs()),
                style,
                transform: Some(svg_transform(state.world_transform)),
            },
            clip,
        )
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
    for chunk in points[1..].chunks_exact(3) {
        path.push(SvgPathCommand::CubicTo {
            first: svg_point(chunk[0]),
            second: svg_point(chunk[1]),
            to: svg_point(chunk[2]),
        });
    }
    path
}

fn arc_path(rect: Rect, start: f32, sweep: f32, pie: bool) -> SvgPath {
    let center = Point {
        x: rect.x + rect.width / 2.0,
        y: rect.y + rect.height / 2.0,
    };
    let radius_x = rect.width.abs() / 2.0;
    let radius_y = rect.height.abs() / 2.0;
    let point_at = |degrees: f32| {
        let radians = degrees.to_radians();
        SvgPoint {
            x: f64::from(center.x + radius_x * radians.cos()),
            y: f64::from(center.y + radius_y * radians.sin()),
        }
    };
    let mut path = SvgPath::new();
    if pie {
        path.push(SvgPathCommand::MoveTo(svg_point(center)));
        path.push(SvgPathCommand::LineTo(point_at(start)));
    } else {
        path.push(SvgPathCommand::MoveTo(point_at(start)));
    }
    path.push(SvgPathCommand::ArcTo {
        rx: f64::from(radius_x),
        ry: f64::from(radius_y),
        rotation: 0.0,
        large_arc: sweep.abs() > 180.0,
        sweep: sweep >= 0.0,
        to: point_at(start + sweep),
    });
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
        let mut payload = record(RecordType::Header, 1, &[0; 16]);
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
        let mut first = record(RecordType::Header, 0, &[0; 16]);
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
