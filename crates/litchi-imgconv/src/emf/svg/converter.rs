//! Strict classic EMF playback into SVG.

use super::{
    buffer::ElementBuffer,
    path::PathBuilder,
    state::{Brush, DeviceContext, Font, GdiObject, Pen, RenderState},
};
use crate::emf::{
    parser::{EmfParser, EmfRecord},
    record_parser::{parse_pointl_array, parse_points_array, parse_poly_counts},
    records::*,
};
use crate::svg_utils::{write_num, write_xml_escaped};
use crate::{
    dib::DibLimits,
    metafile_bitmap::{BitmapOp, StretchPolicy},
};
use litchi_core::error::{Error, Result};
use std::{f64::consts::PI, fmt::Write, mem::size_of};
use zerocopy::FromBytes;

const MAX_RECORD_ITEMS: usize = 1_000_000;

pub struct EmfSvgConverter<'a> {
    parser: &'a EmfParser,
    limits: crate::Limits,
}

impl<'a> EmfSvgConverter<'a> {
    pub fn new(parser: &'a EmfParser) -> Self {
        Self {
            parser,
            limits: crate::Limits::default(),
        }
    }

    pub(crate) fn with_limits(parser: &'a EmfParser, limits: crate::Limits) -> Self {
        Self { parser, limits }
    }

    pub fn convert(&self) -> Result<String> {
        let (svg, diagnostics) = self.convert_with_diagnostics()?;
        if let Some(diagnostic) = diagnostics.first() {
            return Err(Error::Unsupported(format!(
                "EMF+ SVG playback is not exact ({}): {}",
                diagnostic.code, diagnostic.message
            )));
        }
        Ok(svg)
    }

    /// Converts while returning every non-fatal EMF+ approximation to the caller.
    pub fn convert_with_diagnostics(
        &self,
    ) -> Result<(String, Vec<crate::emfplus::renderer::RendererDiagnostic>)> {
        if let Some(output) = self.convert_emfplus_track()? {
            return Ok((output.document.into_string(), output.diagnostics));
        }
        let header = &self.parser.header;
        let dpi_x = dpi(header.device_width, header.device_width_mm);
        let dpi_y = dpi(header.device_height, header.device_height_mm);
        let mut state = RenderState::with_device_metrics(dpi_x, dpi_y);
        let mut buffer = ElementBuffer::new();
        let mut generated_bytes = 0usize;
        let mut generated_elements = 0usize;
        let mut counted_definitions = 0usize;
        for record in &self.parser.records {
            let record_elements = self.process_record(record, &mut state)?;
            if state
                .path_builder
                .as_ref()
                .is_some_and(|path| path.len() > self.limits.max_path_points)
            {
                return Err(Error::ParseError(format!(
                    "EMF path command count exceeds limit {}",
                    self.limits.max_path_points
                )));
            }
            if state.objects.len() > self.limits.max_objects {
                return Err(Error::ParseError(format!(
                    "EMF retained object count exceeds limit {}",
                    self.limits.max_objects
                )));
            }
            if state.defs.len() > self.limits.max_objects {
                return Err(Error::ParseError(format!(
                    "EMF SVG definition count exceeds limit {}",
                    self.limits.max_objects
                )));
            }
            for definition in &state.defs[counted_definitions..] {
                generated_elements = generated_elements
                    .checked_add(svg_element_count(definition))
                    .ok_or_else(|| Error::ParseError("EMF SVG element count overflow".into()))?;
                generated_bytes = generated_bytes
                    .checked_add(definition.len())
                    .ok_or_else(|| Error::ParseError("EMF SVG size overflow".into()))?;
            }
            counted_definitions = state.defs.len();
            if generated_bytes > self.limits.max_output_bytes {
                return Err(Error::ParseError(format!(
                    "EMF SVG output exceeds limit {} bytes",
                    self.limits.max_output_bytes
                )));
            }
            if generated_elements > self.limits.max_svg_elements {
                return Err(Error::ParseError(format!(
                    "EMF SVG element count exceeds limit {}",
                    self.limits.max_svg_elements
                )));
            }
            for element in record_elements {
                generated_elements = generated_elements
                    .checked_add(svg_element_count(&element))
                    .ok_or_else(|| Error::ParseError("EMF SVG element count overflow".into()))?;
                generated_bytes = generated_bytes
                    .checked_add(element.len())
                    .ok_or_else(|| Error::ParseError("EMF SVG size overflow".into()))?;
                if generated_bytes > self.limits.max_output_bytes {
                    return Err(Error::ParseError(format!(
                        "EMF SVG output exceeds limit {} bytes",
                        self.limits.max_output_bytes
                    )));
                }
                buffer.add_element(element, &state.dc);
                if generated_elements > self.limits.max_svg_elements {
                    return Err(Error::ParseError(format!(
                        "EMF SVG element count exceeds limit {}",
                        self.limits.max_svg_elements
                    )));
                }
            }
        }
        buffer.flush();
        if generated_elements > self.limits.max_svg_elements {
            return Err(Error::ParseError(format!(
                "EMF SVG element count exceeds limit {}",
                self.limits.max_svg_elements
            )));
        }
        Ok((self.build_svg(&buffer.elements, &state)?, Vec::new()))
    }

    /// Select one complete EMF+ rendering track before classic playback.
    ///
    /// Dual metafiles deliberately use the EMF+ track so the equivalent
    /// classic fallback is not painted a second time. EMF+-only streams that
    /// use GetDC require true record-level interleaving; reject those until the
    /// two renderers share a common ordered sink.
    fn convert_emfplus_track(&self) -> Result<Option<crate::emfplus::renderer::RenderOutput>> {
        let mut renderer = None;
        for record in &self.parser.records {
            if record.record_type != crate::emfplus::EMR_COMMENT
                || record.data.get(4..8) != Some(b"EMF+".as_slice())
            {
                continue;
            }
            let emfplus = if let Some(value) = renderer.as_mut() {
                value
            } else {
                let header = &self.parser.header;
                renderer.insert(crate::emfplus::renderer::EmfPlusSvgRenderer::new(
                    f64::from(header.width().unsigned_abs().max(1)),
                    f64::from(header.height().unsigned_abs().max(1)),
                    emfplus_limits(self.limits),
                )?)
            };
            emfplus.push_comment_body(&record.data)?;
        }

        let Some(renderer) = renderer else {
            return Ok(None);
        };
        let output = renderer.finish()?;
        if output.mux.get_dc_count != 0 {
            return Err(Error::Unsupported(
                "EMF+ GetDC requires ordered classic/EMF+ SVG interleaving".into(),
            ));
        }
        Ok(Some(output))
    }

    fn process_record(&self, record: &EmfRecord, state: &mut RenderState) -> Result<Vec<String>> {
        let Some(kind) = EmrType::from_u32(record.record_type) else {
            return Err(Error::Unsupported(format!(
                "unknown EMF record type {}",
                record.record_type
            )));
        };
        let data: &[u8] = &record.data;
        match kind {
            EmrType::Header => Ok(Vec::new()),
            EmrType::Eof => {
                require_len(data, 12, kind)?;
                Ok(Vec::new())
            },

            // Mapping and graphics state.
            EmrType::SetWindowExtEx | EmrType::SetViewportExtEx => {
                let extent = read_record::<SizeL>(data, kind)?;
                if extent.cx == 0 || extent.cy == 0 {
                    return malformed(kind, "zero mapping extent");
                }
                if kind == EmrType::SetWindowExtEx {
                    state.dc.window_ext = (extent.cx, extent.cy);
                } else {
                    state.dc.viewport_ext = (extent.cx, extent.cy);
                }
                Ok(Vec::new())
            },
            EmrType::SetWindowOrgEx | EmrType::SetViewportOrgEx | EmrType::SetBrushOrgEx => {
                let point = read_record::<PointL>(data, kind)?;
                match kind {
                    EmrType::SetWindowOrgEx => state.dc.window_org = (point.x, point.y),
                    EmrType::SetViewportOrgEx => state.dc.viewport_org = (point.x, point.y),
                    _ => state.dc.brush_org = (point.x, point.y),
                }
                Ok(Vec::new())
            },
            EmrType::ScaleViewportExtEx | EmrType::ScaleWindowExtEx => {
                let scale = read_record::<EmrScaleExtEx>(data, kind)?;
                if scale.x_denom == 0 || scale.y_denom == 0 {
                    return malformed(kind, "zero scale denominator");
                }
                let target = if kind == EmrType::ScaleViewportExtEx {
                    &mut state.dc.viewport_ext
                } else {
                    &mut state.dc.window_ext
                };
                target.0 = scale_i32(target.0, scale.x_num, scale.x_denom, kind)?;
                target.1 = scale_i32(target.1, scale.y_num, scale.y_denom, kind)?;
                Ok(Vec::new())
            },
            EmrType::SetMapMode => {
                let mode = read_u32(data, 0, kind)?;
                if !(1..=8).contains(&mode) {
                    return malformed(kind, "invalid mapping mode");
                }
                state.dc.map_mode = mode;
                Ok(Vec::new())
            },
            EmrType::SetWorldTransform => {
                let xform = read_record::<XForm>(data, kind)?;
                validate_xform(&xform, kind)?;
                state.dc.world_transform = xform;
                Ok(Vec::new())
            },
            EmrType::ModifyWorldTransform => {
                let modification = read_record::<EmrModifyWorldTransform>(data, kind)?;
                validate_xform(&modification.xform, kind)?;
                state.dc.world_transform = match modification.mode {
                    1 => XForm::default(),
                    // With our column-vector representation, left multiplication
                    // is XForm * current and right multiplication the reverse.
                    2 => modification.xform.multiply(&state.dc.world_transform),
                    3 => state.dc.world_transform.multiply(&modification.xform),
                    _ => return malformed(kind, "invalid ModifyWorldTransform mode"),
                };
                Ok(Vec::new())
            },
            EmrType::SaveDc => {
                require_len_exact(data, 0, kind)?;
                if state.dc_stack.len() >= self.limits.max_state_depth {
                    return malformed(kind, "saved DC depth exceeds configured limit");
                }
                state.push_dc();
                Ok(Vec::new())
            },
            EmrType::RestoreDc => {
                let index = read_i32(data, 0, kind)?;
                if !state.pop_dc(index) {
                    return malformed(kind, "RestoreDC references a missing saved state");
                }
                Ok(Vec::new())
            },
            EmrType::SetTextColor | EmrType::SetBkColor => {
                let color = read_record::<ColorRef>(data, kind)?;
                if kind == EmrType::SetTextColor {
                    state.dc.text_color = color;
                } else {
                    state.dc.bg_color = color;
                    // A selected hatch definition captures its old background.
                    state.dc.brush.pattern_id = None;
                }
                Ok(Vec::new())
            },
            EmrType::SetBkMode => {
                let mode = read_u32(data, 0, kind)?;
                if mode != 1 && mode != 2 {
                    return malformed(kind, "invalid background mode");
                }
                state.dc.bg_mode = mode;
                state.dc.brush.pattern_id = None;
                Ok(Vec::new())
            },
            EmrType::SetPolyFillMode => {
                let mode = read_u32(data, 0, kind)?;
                if mode != 1 && mode != 2 {
                    return malformed(kind, "invalid polygon fill mode");
                }
                state.dc.poly_fill_mode = mode;
                Ok(Vec::new())
            },
            EmrType::SetTextAlign => {
                state.dc.text_align = read_u32(data, 0, kind)?;
                Ok(Vec::new())
            },
            EmrType::SetRop2 => {
                let mode = read_u32(data, 0, kind)?;
                if !matches!(mode, 11 | 13) {
                    return unsupported(
                        kind,
                        "binary ROP2 composition has no faithful SVG equivalent",
                    );
                }
                state.dc.rop2 = mode;
                Ok(Vec::new())
            },
            EmrType::SetStretchBltMode => {
                let mode = read_u32(data, 0, kind)?;
                if !(1..=4).contains(&mode) {
                    return malformed(kind, "invalid stretch mode");
                }
                state.dc.stretch_mode = mode;
                Ok(Vec::new())
            },
            EmrType::SetArcDirection => {
                state.dc.arc_direction = match read_u32(data, 0, kind)? {
                    1 => false,
                    2 => true,
                    _ => return malformed(kind, "invalid arc direction"),
                };
                Ok(Vec::new())
            },
            EmrType::SetMiterLimit => {
                let value = read_u32(data, 0, kind)?;
                if value == 0 {
                    return malformed(kind, "zero miter limit");
                }
                state.dc.miter_limit = f64::from(value);
                Ok(Vec::new())
            },
            EmrType::SetLayout => {
                state.dc.layout = read_u32(data, 0, kind)?;
                Ok(Vec::new())
            },

            // Object table operations. Creation never selects the object.
            EmrType::CreatePen => {
                let value = read_record::<EmrCreatePen>(data, kind)?;
                insert_object(
                    state,
                    value.object_index,
                    GdiObject::Pen(Pen::from_create_pen(
                        value.pen_style,
                        value.width,
                        value.color,
                    )),
                    kind,
                )?;
                Ok(Vec::new())
            },
            EmrType::CreateBrushIndirect => {
                let value = read_record::<EmrCreateBrushIndirect>(data, kind)?;
                if value.brush_style > brush_style::HATCHED {
                    return unsupported(kind, "bitmap/pattern brush requires DIB integration");
                }
                if value.brush_style == brush_style::HATCHED && value.brush_hatch > 5 {
                    return malformed(kind, "invalid hatch style");
                }
                insert_object(
                    state,
                    value.object_index,
                    GdiObject::Brush(Brush::from_create_brush(
                        value.brush_style,
                        value.color,
                        value.brush_hatch,
                    )),
                    kind,
                )?;
                Ok(Vec::new())
            },
            EmrType::ExtCreatePen => {
                let value = read_record::<EmrExtCreatePenHeader>(data, kind)?;
                if value.brush_style > brush_style::HATCHED
                    || value.cb_bmi != 0
                    || value.cb_bits != 0
                {
                    return unsupported(
                        kind,
                        "bitmap-backed extended pen requires DIB integration",
                    );
                }
                let mut pen = Pen::from_create_pen(value.pen_style, value.width, value.color);
                if value.num_style_entries > 0 {
                    let count = bounded_count(value.num_style_entries, kind)?;
                    if count > self.limits.max_path_points
                        || count.saturating_mul(12) > self.limits.max_output_bytes
                    {
                        return malformed(
                            kind,
                            "extended pen style array exceeds configured limits",
                        );
                    }
                    let offset = size_of::<EmrExtCreatePenHeader>();
                    let entries = read_u32_array(data, offset, count, kind)?;
                    if value.pen_style & 0xff == pen_style::USERSTYLE {
                        pen.dash_pattern = Some(
                            entries
                                .iter()
                                .map(u32::to_string)
                                .collect::<Vec<_>>()
                                .join(" ")
                                .into(),
                        );
                    }
                }
                insert_object(state, value.object_index, GdiObject::Pen(pen), kind)?;
                Ok(Vec::new())
            },
            EmrType::ExtCreateFontIndirectW => {
                let (handle, font) = parse_font(data, kind)?;
                insert_object(state, handle, GdiObject::Font(font), kind)?;
                Ok(Vec::new())
            },
            EmrType::CreatePalette => {
                let header = read_record::<EmrCreatePaletteHeader>(data, kind)?;
                let count = usize::from(header.num_entries);
                checked_range(
                    size_of::<EmrCreatePaletteHeader>(),
                    count,
                    size_of::<PaletteEntry>(),
                    data.len(),
                    kind,
                )?;
                insert_object(state, header.object_index, GdiObject::Palette, kind)?;
                Ok(Vec::new())
            },
            EmrType::SelectObject => {
                let handle = read_u32(data, 0, kind)?;
                if !state.select_object(handle) {
                    return malformed(kind, "unknown object handle");
                }
                state.prepare_brush_pattern();
                Ok(Vec::new())
            },
            EmrType::DeleteObject => {
                let handle = read_u32(data, 0, kind)?;
                // DeleteObject returns failure for an unknown, stock, or
                // currently selected object. That failure does not alter the
                // DC or subsequent drawing and is therefore a safe playback
                // no-op rather than a malformed metafile.
                state.delete_object(handle);
                Ok(Vec::new())
            },

            // Current point and path grammar.
            EmrType::MoveToEx => {
                let point = read_record::<PointL>(data, kind)?;
                state.dc.current_pos = (f64::from(point.x), f64::from(point.y));
                if state.in_path {
                    let transformed = state
                        .dc
                        .transform_point(state.dc.current_pos.0, state.dc.current_pos.1);
                    state
                        .path_builder
                        .as_mut()
                        .expect("path state")
                        .move_to(transformed.0, transformed.1);
                }
                Ok(Vec::new())
            },
            EmrType::LineTo => {
                let point = read_record::<PointL>(data, kind)?;
                let start = state
                    .dc
                    .transform_point(state.dc.current_pos.0, state.dc.current_pos.1);
                let logical_end = (f64::from(point.x), f64::from(point.y));
                let end = state.dc.transform_point(logical_end.0, logical_end.1);
                state.dc.current_pos = logical_end;
                if state.in_path {
                    let builder = state.path_builder.as_mut().expect("path state");
                    if builder.is_empty() {
                        builder.move_to(start.0, start.1);
                    }
                    builder.line_to(end.0, end.1);
                    Ok(Vec::new())
                } else if state.dc.rop2 == 11 {
                    Ok(Vec::new())
                } else {
                    Ok(vec![render_line(start, end, &state.dc)])
                }
            },
            EmrType::BeginPath => {
                require_len_exact(data, 0, kind)?;
                state.begin_path();
                Ok(Vec::new())
            },
            EmrType::EndPath => {
                require_len_exact(data, 0, kind)?;
                if !state.in_path {
                    return malformed(kind, "EndPath without BeginPath");
                }
                state.end_path();
                Ok(Vec::new())
            },
            EmrType::CloseFigure => {
                require_len_exact(data, 0, kind)?;
                if !state.in_path {
                    return malformed(kind, "CloseFigure outside path bracket");
                }
                state.path_builder.as_mut().expect("path state").close();
                Ok(Vec::new())
            },
            EmrType::AbortPath => {
                require_len_exact(data, 0, kind)?;
                state.abort_path();
                Ok(Vec::new())
            },
            EmrType::FillPath | EmrType::StrokePath | EmrType::StrokeAndFillPath => {
                require_len(data, 16, kind)?; // bounds
                let builder = state.take_path().ok_or_else(|| {
                    Error::ParseError(format!("{} without a completed path", kind.name()))
                })?;
                let fill = matches!(kind, EmrType::FillPath | EmrType::StrokeAndFillPath);
                let stroke = matches!(kind, EmrType::StrokePath | EmrType::StrokeAndFillPath);
                Ok(path_element(builder, &state.dc, fill, stroke)
                    .into_iter()
                    .collect())
            },
            EmrType::FlattenPath => {
                require_len_exact(data, 0, kind)?;
                // SVG renders the same geometry whether curves remain cubic.
                Ok(Vec::new())
            },
            EmrType::WidenPath => unsupported(kind, "widening a GDI path is output-affecting"),
            EmrType::SelectClipPath => {
                let mode = read_u32(data, 0, kind)?;
                let builder = state.take_path().ok_or_else(|| {
                    Error::ParseError("SelectClipPath without a completed path".into())
                })?;
                if !state.install_clip(&builder.build(), mode) {
                    return unsupported(
                        kind,
                        "requested region combine mode is not faithfully representable",
                    );
                }
                Ok(Vec::new())
            },

            // Poly records.
            EmrType::Polygon
            | EmrType::Polyline
            | EmrType::PolyBezier
            | EmrType::PolyBezierTo
            | EmrType::PolyLineTo => self.poly_record(data, kind, false, state),
            EmrType::Polygon16
            | EmrType::Polyline16
            | EmrType::PolyBezier16
            | EmrType::PolyBezierTo16
            | EmrType::PolyLineTo16 => self.poly_record(data, kind, true, state),
            EmrType::PolyPolyline | EmrType::PolyPolygon => {
                self.poly_poly_record(data, kind, false, state)
            },
            EmrType::PolyPolyline16 | EmrType::PolyPolygon16 => {
                self.poly_poly_record(data, kind, true, state)
            },
            EmrType::PolyDraw => self.poly_draw(data, kind, false, state),
            EmrType::PolyDraw16 => self.poly_draw(data, kind, true, state),

            // Primitive shapes.
            EmrType::Rectangle | EmrType::Ellipse => {
                let rect = read_record::<RectL>(data, kind)?;
                let builder = if kind == EmrType::Rectangle {
                    rectangle_path(rect, &state.dc)
                } else {
                    ellipse_path(rect, &state.dc)
                }?;
                self.record_or_render(builder, state, true, true)
            },
            EmrType::RoundRect => {
                let shape = read_record::<EmrRoundRect>(data, kind)?;
                let builder = round_rect_path(shape.rect, shape.corner, &state.dc)?;
                self.record_or_render(builder, state, true, true)
            },
            EmrType::Arc | EmrType::ArcTo | EmrType::Chord | EmrType::Pie => {
                let arc = read_record::<EmrArc>(data, kind)?;
                let (builder, end) = arc_path(arc, kind, &state.dc)?;
                if matches!(kind, EmrType::ArcTo) {
                    state.dc.current_pos = end;
                }
                let fill = matches!(kind, EmrType::Chord | EmrType::Pie);
                self.record_or_render(builder, state, fill, true)
            },
            EmrType::AngleArc => {
                let arc = read_record::<EmrAngleArc>(data, kind)?;
                let (builder, end) = angle_arc_path(arc, &state.dc)?;
                state.dc.current_pos = end;
                self.record_or_render(builder, state, false, true)
            },
            EmrType::SetPixelV => {
                let pixel = read_record::<EmrSetPixelV>(data, kind)?;
                let point = state
                    .dc
                    .transform_point(f64::from(pixel.point.x), f64::from(pixel.point.y));
                Ok(vec![format!(
                    "<rect x=\"{}\" y=\"{}\" width=\"1\" height=\"1\" fill=\"{}\" {}/>",
                    fmt(point.0),
                    fmt(point.1),
                    pixel.color.to_svg_color(),
                    state.dc.clip_attr()
                )])
            },
            EmrType::GradientFill => self.gradient_fill(data, kind, state),
            EmrType::ExtFloodFill => {
                unsupported(kind, "flood filling cannot be reconstructed in SVG")
            },

            // Clipping and regions.
            EmrType::IntersectClipRect | EmrType::ExcludeClipRect => {
                let rect = read_record::<RectL>(data, kind)?;
                let path = if kind == EmrType::IntersectClipRect {
                    rectangle_path(rect, &state.dc)?.build()
                } else {
                    exclusion_path(rect, &state.dc)?
                };
                if !state.install_clip(&path, 1) {
                    return unsupported(kind, "clip combine mode");
                }
                Ok(Vec::new())
            },
            EmrType::OffsetClipRgn => {
                let offset = read_record::<PointL>(data, kind)?;
                let transformed = state
                    .dc
                    .transform_vector(f64::from(offset.x), f64::from(offset.y));
                state.offset_clip(transformed.0, transformed.1);
                Ok(Vec::new())
            },
            EmrType::SetMetaRgn => {
                require_len_exact(data, 0, kind)?;
                // The effective clip is unchanged; only its internal ownership
                // moves from application clip to meta clip.
                Ok(Vec::new())
            },
            EmrType::ExtSelectClipRgn => {
                let size = usize::try_from(read_u32(data, 0, kind)?)
                    .map_err(|_| Error::ParseError("region size overflow".into()))?;
                let mode = read_u32(data, 4, kind)?;
                if size == 0 {
                    if mode != 5 {
                        return malformed(kind, "empty region requires RGN_COPY");
                    }
                    state.dc.clip_id = None;
                    return Ok(Vec::new());
                }
                let end = 8usize
                    .checked_add(size)
                    .ok_or_else(|| Error::ParseError("region range overflow".into()))?;
                if end > data.len() {
                    return malformed(kind, "region data exceeds record payload");
                }
                let path = region_path(
                    &data[8..end],
                    &state.dc,
                    kind,
                    self.limits.max_path_points,
                    self.limits.max_output_bytes,
                )?;
                if !state.install_clip(&path, mode) {
                    return unsupported(
                        kind,
                        "requested region combine mode is not faithfully representable",
                    );
                }
                Ok(Vec::new())
            },
            EmrType::FillRgn | EmrType::FrameRgn | EmrType::InvertRgn | EmrType::PaintRgn => {
                self.render_region(data, kind, state)
            },

            // Text.
            EmrType::ExtTextOutA | EmrType::ExtTextOutW => self.ext_text_out(data, kind, state),
            EmrType::PolyTextOutA | EmrType::PolyTextOutW => self.poly_text_out(data, kind, state),
            EmrType::SmallTextOut => self.small_text_out(data, kind, state),
            EmrType::SetTextJustification => {
                // SVG has no direct equivalent. Zero values have no visual effect.
                let value = read_record::<EmrSetTextJustification>(data, kind)?;
                if value.num_break_extra != 0 || value.num_break_count != 0 {
                    return unsupported(kind, "text justification spacing is output-affecting");
                }
                Ok(Vec::new())
            },

            EmrType::BitBlt
            | EmrType::StretchBlt
            | EmrType::MaskBlt
            | EmrType::PlgBlt
            | EmrType::SetDIBitsToDevice
            | EmrType::StretchDIBits
            | EmrType::AlphaBlend
            | EmrType::TransparentBlt => self.bitmap_record(data, kind, state),
            EmrType::CreateMonoBrush | EmrType::CreateDIBPatternBrushPt => unsupported(
                kind,
                "bitmap brush playback is not supported by the shared DIB renderer",
            ),

            // Validated non-visual records.
            EmrType::SetMapperFlags => {
                require_len(data, 4, kind)?;
                Ok(Vec::new())
            },
            EmrType::SetColorAdjustment => {
                require_len(data, size_of::<ColorAdjustment>(), kind)?;
                Ok(Vec::new())
            },
            EmrType::SelectPalette => {
                require_len(data, 4, kind)?;
                Ok(Vec::new())
            },
            EmrType::SetPaletteEntries => {
                require_len(data, 12, kind)?;
                let count = bounded_count(read_u32(data, 8, kind)?, kind)?;
                checked_range(12, count, 4, data.len(), kind)?;
                Ok(Vec::new())
            },
            EmrType::ResizePalette => {
                require_len(data, 8, kind)?;
                Ok(Vec::new())
            },
            EmrType::RealizePalette => {
                require_len_exact(data, 0, kind)?;
                Ok(Vec::new())
            },
            EmrType::SetIcmMode
            | EmrType::SetColorSpace
            | EmrType::DeleteColorSpace
            | EmrType::ForceUfiMapping
            | EmrType::SetLinkedUfis => {
                require_len(data, 4, kind)?;
                Ok(Vec::new())
            },
            EmrType::CreateColorSpace | EmrType::CreateColorSpaceW => {
                require_len(data, 4, kind)?;
                Ok(Vec::new())
            },
            EmrType::ColorCorrectPalette
            | EmrType::SetIcmProfileA
            | EmrType::SetIcmProfileW
            | EmrType::ColorMatchToTargetW
            | EmrType::StartDoc
            | EmrType::GlsRecord
            | EmrType::GlsBoundedRecord
            | EmrType::PixelFormat => {
                // Palette/ICM/print/query/OpenGL metadata does not have SVG
                // playback semantics, but the record must at least be DWORD aligned.
                if data.len() % 4 != 0 {
                    return malformed(kind, "payload is not DWORD aligned");
                }
                Ok(Vec::new())
            },
            EmrType::Comment => {
                require_len(data, 4, kind)?;
                let size = usize::try_from(read_u32(data, 0, kind)?)
                    .map_err(|_| Error::ParseError("comment size overflow".into()))?;
                if size > data.len().saturating_sub(4) {
                    return malformed(kind, "comment data exceeds payload");
                }
                if data.get(4..8) == Some(&[0x45, 0x4d, 0x46, 0x2b]) {
                    return unsupported(kind, "EMF+ comment requires EMF+ playback");
                }
                Ok(Vec::new())
            },

            // Escape programs are data/code boundaries and are never executed.
            EmrType::DrawEscape | EmrType::ExtEscape | EmrType::NamedEscape => {
                unsupported(kind, "escape/PostScript programs are never executed")
            },
        }
    }

    fn record_or_render(
        &self,
        builder: PathBuilder,
        state: &mut RenderState,
        fill: bool,
        stroke: bool,
    ) -> Result<Vec<String>> {
        if state.in_path {
            append_path(state.path_builder.as_mut().expect("path state"), builder);
            Ok(Vec::new())
        } else if state.dc.rop2 == 11 {
            Ok(Vec::new())
        } else {
            state.prepare_brush_pattern();
            Ok(path_element(builder, &state.dc, fill, stroke)
                .into_iter()
                .collect())
        }
    }

    fn bitmap_record(
        &self,
        data: &[u8],
        kind: EmrType,
        state: &RenderState,
    ) -> Result<Vec<String>> {
        let stretch = match state.dc.stretch_mode {
            1 => StretchPolicy::BlackOnWhite,
            2 => StretchPolicy::WhiteOnBlack,
            3 => StretchPolicy::ColorOnColor,
            4 => StretchPolicy::Halftone,
            _ => return malformed(kind, "invalid active stretch mode"),
        };
        let operation = BitmapOp::parse_emf(kind as u32, data, stretch, dib_limits(self.limits))?;
        let image = operation.to_svg_image()?;

        // The shared helper emits logical destination coordinates. Applying
        // the complete DC affine matrix as a wrapper preserves rectangular
        // mirroring and the three-point PlgBlt parallelogram under the active
        // world, page, and device transforms.
        let matrix = logical_to_device_matrix(&state.dc);
        let interpolation = match stretch {
            StretchPolicy::Halftone => "optimizeQuality",
            StretchPolicy::BlackOnWhite
            | StretchPolicy::WhiteOnBlack
            | StretchPolicy::ColorOnColor => "optimizeSpeed",
        };
        Ok(vec![format!(
            "<g transform=\"matrix({} {} {} {} {} {})\" image-rendering=\"{}\" {}>{}</g>",
            fmt(matrix[0]),
            fmt(matrix[1]),
            fmt(matrix[2]),
            fmt(matrix[3]),
            fmt(matrix[4]),
            fmt(matrix[5]),
            interpolation,
            state.dc.clip_attr(),
            image.element
        )])
    }

    fn poly_record(
        &self,
        data: &[u8],
        kind: EmrType,
        short: bool,
        state: &mut RenderState,
    ) -> Result<Vec<String>> {
        require_len(data, 20, kind)?;
        let count = bounded_count(read_u32(data, 16, kind)?, kind)?;
        if count > self.limits.max_path_points {
            return malformed(kind, "point count exceeds configured path limit");
        }
        let points = if short {
            parse_points_array(&data[20..], count)?
                .into_iter()
                .map(|p| (i32::from(p.x), i32::from(p.y)))
                .collect::<Vec<_>>()
        } else {
            parse_pointl_array(&data[20..], count)?
                .into_iter()
                .map(|p| (p.x, p.y))
                .collect::<Vec<_>>()
        };
        let is_to = matches!(
            kind,
            EmrType::PolyBezierTo
                | EmrType::PolyLineTo
                | EmrType::PolyBezierTo16
                | EmrType::PolyLineTo16
        );
        let bezier = matches!(
            kind,
            EmrType::PolyBezier
                | EmrType::PolyBezierTo
                | EmrType::PolyBezier16
                | EmrType::PolyBezierTo16
        );
        let mut builder = PathBuilder::new();
        let mut index = 0;
        if is_to {
            let start = state
                .dc
                .transform_point(state.dc.current_pos.0, state.dc.current_pos.1);
            builder.move_to(start.0, start.1);
        } else {
            let first = points
                .first()
                .ok_or_else(|| Error::ParseError(format!("{} has no points", kind.name())))?;
            let first = state
                .dc
                .transform_point(f64::from(first.0), f64::from(first.1));
            builder.move_to(first.0, first.1);
            index = 1;
        }
        if bezier {
            if (points.len() - index) % 3 != 0 {
                return malformed(kind, "Bezier point count is not a multiple of three");
            }
            for curve in points[index..].chunks_exact(3) {
                let a = transform(curve[0], &state.dc);
                let b = transform(curve[1], &state.dc);
                let c = transform(curve[2], &state.dc);
                builder.cubic_to(a.0, a.1, b.0, b.1, c.0, c.1);
            }
        } else {
            for point in &points[index..] {
                let point = transform(*point, &state.dc);
                builder.line_to(point.0, point.1);
            }
        }
        let polygon = matches!(kind, EmrType::Polygon | EmrType::Polygon16);
        if polygon {
            builder.close();
        }
        if is_to {
            if let Some(last) = points.last() {
                state.dc.current_pos = (f64::from(last.0), f64::from(last.1));
            }
        }
        self.record_or_render(builder, state, polygon, true)
    }

    fn poly_poly_record(
        &self,
        data: &[u8],
        kind: EmrType,
        short: bool,
        state: &mut RenderState,
    ) -> Result<Vec<String>> {
        require_len(data, 24, kind)?;
        let polygons = bounded_count(read_u32(data, 16, kind)?, kind)?;
        let total = bounded_count(read_u32(data, 20, kind)?, kind)?;
        if polygons > self.limits.max_path_points || total > self.limits.max_path_points {
            return malformed(kind, "poly-polygon count exceeds configured path limit");
        }
        let counts = parse_poly_counts(&data[24..], polygons)?;
        let sum = counts.iter().try_fold(0usize, |sum, &count| {
            sum.checked_add(usize::try_from(count).ok()?)
        });
        if sum != Some(total) {
            return malformed(kind, "polygon point counts do not equal total count");
        }
        let point_offset = 24usize
            .checked_add(
                polygons
                    .checked_mul(4)
                    .ok_or_else(|| Error::ParseError("count range overflow".into()))?,
            )
            .ok_or_else(|| Error::ParseError("point offset overflow".into()))?;
        let points = if short {
            parse_points_array(
                data.get(point_offset..)
                    .ok_or_else(|| Error::ParseError("point offset outside payload".into()))?,
                total,
            )?
            .into_iter()
            .map(|p| (i32::from(p.x), i32::from(p.y)))
            .collect::<Vec<_>>()
        } else {
            parse_pointl_array(
                data.get(point_offset..)
                    .ok_or_else(|| Error::ParseError("point offset outside payload".into()))?,
                total,
            )?
            .into_iter()
            .map(|p| (p.x, p.y))
            .collect::<Vec<_>>()
        };
        let mut builder = PathBuilder::new();
        let mut offset = 0usize;
        let filled = matches!(kind, EmrType::PolyPolygon | EmrType::PolyPolygon16);
        for count in counts {
            let count = usize::try_from(count)
                .map_err(|_| Error::ParseError("polygon count overflow".into()))?;
            let polygon = &points[offset..offset + count];
            if let Some(first) = polygon.first() {
                let first = transform(*first, &state.dc);
                builder.move_to(first.0, first.1);
                for point in &polygon[1..] {
                    let point = transform(*point, &state.dc);
                    builder.line_to(point.0, point.1);
                }
                if filled {
                    builder.close();
                }
            }
            offset += count;
        }
        self.record_or_render(builder, state, filled, true)
    }

    fn poly_draw(
        &self,
        data: &[u8],
        kind: EmrType,
        short: bool,
        state: &mut RenderState,
    ) -> Result<Vec<String>> {
        require_len(data, 20, kind)?;
        let count = bounded_count(read_u32(data, 16, kind)?, kind)?;
        if count > self.limits.max_path_points {
            return malformed(kind, "PolyDraw count exceeds configured path limit");
        }
        let point_bytes = count
            .checked_mul(if short { 4 } else { 8 })
            .ok_or_else(|| Error::ParseError("PolyDraw point range overflow".into()))?;
        let types_offset = 20usize
            .checked_add(point_bytes)
            .ok_or_else(|| Error::ParseError("PolyDraw type offset overflow".into()))?;
        let types = data
            .get(
                types_offset
                    ..types_offset
                        .checked_add(count)
                        .ok_or_else(|| Error::ParseError("PolyDraw types overflow".into()))?,
            )
            .ok_or_else(|| Error::ParseError("PolyDraw types exceed payload".into()))?;
        let points = if short {
            parse_points_array(&data[20..], count)?
                .into_iter()
                .map(|p| (i32::from(p.x), i32::from(p.y)))
                .collect::<Vec<_>>()
        } else {
            parse_pointl_array(&data[20..], count)?
                .into_iter()
                .map(|p| (p.x, p.y))
                .collect::<Vec<_>>()
        };
        let mut builder = PathBuilder::new();
        let start = state
            .dc
            .transform_point(state.dc.current_pos.0, state.dc.current_pos.1);
        builder.move_to(start.0, start.1);
        let mut i = 0usize;
        while i < count {
            let base = types[i] & !point_type::CLOSEFIGURE;
            match base {
                point_type::MOVETO => {
                    let p = transform(points[i], &state.dc);
                    builder.move_to(p.0, p.1);
                    i += 1;
                },
                point_type::LINETO => {
                    let p = transform(points[i], &state.dc);
                    builder.line_to(p.0, p.1);
                    i += 1;
                },
                point_type::BEZIERTO => {
                    if i + 3 > count
                        || types[i..i + 3]
                            .iter()
                            .any(|value| value & !point_type::CLOSEFIGURE != point_type::BEZIERTO)
                    {
                        return malformed(kind, "BezierTo must occur in groups of three");
                    }
                    let a = transform(points[i], &state.dc);
                    let b = transform(points[i + 1], &state.dc);
                    let c = transform(points[i + 2], &state.dc);
                    builder.cubic_to(a.0, a.1, b.0, b.1, c.0, c.1);
                    i += 3;
                },
                _ => return malformed(kind, "invalid PolyDraw point type"),
            }
            if types[i - 1] & point_type::CLOSEFIGURE != 0 {
                builder.close();
            }
        }
        if let Some(last) = points.last() {
            state.dc.current_pos = (f64::from(last.0), f64::from(last.1));
        }
        self.record_or_render(builder, state, false, true)
    }

    fn render_region(
        &self,
        data: &[u8],
        kind: EmrType,
        state: &mut RenderState,
    ) -> Result<Vec<String>> {
        require_len(data, 24, kind)?;
        if kind == EmrType::InvertRgn {
            return unsupported(
                kind,
                "region inversion requires destination-dependent raster composition",
            );
        }
        let region_size = usize::try_from(read_u32(data, 16, kind)?)
            .map_err(|_| Error::ParseError("region size overflow".into()))?;
        let region_offset = match kind {
            EmrType::FrameRgn => 32,
            EmrType::FillRgn => 24,
            EmrType::InvertRgn | EmrType::PaintRgn => 20,
            _ => unreachable!(),
        };
        if region_offset + region_size > data.len() {
            return malformed(kind, "region exceeds payload");
        }
        let path = region_path(
            &data[region_offset..region_offset + region_size],
            &state.dc,
            kind,
            self.limits.max_path_points,
            self.limits.max_output_bytes,
        )?;
        let saved_brush = state.dc.brush.clone();
        let saved_brush_handle = state.dc.brush_handle;
        if matches!(kind, EmrType::FillRgn | EmrType::FrameRgn) {
            let brush_handle = read_u32(data, 20, kind)?;
            if !state.select_brush(brush_handle) {
                return malformed(kind, "unknown region brush handle");
            }
            state.prepare_brush_pattern();
        }
        let attrs = match kind {
            EmrType::FrameRgn => {
                let width = read_i32(data, 24, kind)?.unsigned_abs().max(1);
                format!(
                    "fill=\"none\" stroke=\"{}\" stroke-width=\"{}\" {}",
                    state.dc.brush.color.to_svg_color(),
                    width,
                    state.dc.clip_attr()
                )
            },
            EmrType::InvertRgn => unreachable!(),
            _ => format!("{} {}", state.dc.get_fill_attr(), state.dc.clip_attr()),
        };
        state.dc.brush = saved_brush;
        state.dc.brush_handle = saved_brush_handle;
        Ok(vec![format!(
            "<path d=\"{}\" {} fill-rule=\"evenodd\"/>",
            path, attrs
        )])
    }

    fn gradient_fill(
        &self,
        data: &[u8],
        kind: EmrType,
        state: &mut RenderState,
    ) -> Result<Vec<String>> {
        let header = read_record::<EmrGradientFillHeader>(data, kind)?;
        let vertices_count = bounded_count(header.num_vertices, kind)?;
        let primitives_count = bounded_count(header.num_triangles, kind)?;
        if vertices_count > self.limits.max_path_points {
            return malformed(kind, "gradient vertex count exceeds configured path limit");
        }
        if primitives_count > self.limits.max_svg_elements
            || primitives_count > self.limits.max_objects
        {
            return malformed(
                kind,
                "gradient primitive count exceeds configured SVG/object limit",
            );
        }
        let estimated = primitives_count
            .checked_mul(512)
            .ok_or_else(|| Error::ParseError("gradient output estimate overflow".into()))?;
        if estimated > self.limits.max_output_bytes {
            return malformed(kind, "gradient output exceeds configured byte limit");
        }
        let vertices_offset = size_of::<EmrGradientFillHeader>();
        let vertices_end = checked_range(
            vertices_offset,
            vertices_count,
            size_of::<TriVertex>(),
            data.len(),
            kind,
        )?;
        let vertices = (0..vertices_count)
            .map(|index| {
                read_record::<TriVertex>(
                    &data[vertices_offset + index * size_of::<TriVertex>()..],
                    kind,
                )
            })
            .collect::<Result<Vec<_>>>()?;
        let mut output = Vec::new();
        match header.mode {
            0 | 1 => {
                checked_range(
                    vertices_end,
                    primitives_count,
                    size_of::<GradientRect>(),
                    data.len(),
                    kind,
                )?;
                for index in 0..primitives_count {
                    let primitive = read_record::<GradientRect>(
                        &data[vertices_end + index * size_of::<GradientRect>()..],
                        kind,
                    )?;
                    let a = vertex(&vertices, primitive.upper_left, kind)?;
                    let b = vertex(&vertices, primitive.lower_right, kind)?;
                    let id = state.fresh_id("gradient");
                    let (x1, y1, x2, y2) = if header.mode == 0 {
                        ("0%", "0%", "100%", "0%")
                    } else {
                        ("0%", "0%", "0%", "100%")
                    };
                    state.add_definition(format!(
                        "<linearGradient id=\"{}\" x1=\"{}\" y1=\"{}\" x2=\"{}\" y2=\"{}\"><stop offset=\"0\" stop-color=\"{}\" stop-opacity=\"{}\"/><stop offset=\"1\" stop-color=\"{}\" stop-opacity=\"{}\"/></linearGradient>",
                        id, x1, y1, x2, y2, vertex_color(a), vertex_alpha(a), vertex_color(b), vertex_alpha(b)
                    ));
                    let p1 = state.dc.transform_point(f64::from(a.x), f64::from(a.y));
                    let p2 = state.dc.transform_point(f64::from(b.x), f64::from(b.y));
                    output.push(format!(
                        "<rect x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\" fill=\"url(#{})\" {}/>",
                        fmt(p1.0.min(p2.0)),
                        fmt(p1.1.min(p2.1)),
                        fmt((p2.0 - p1.0).abs()),
                        fmt((p2.1 - p1.1).abs()),
                        id,
                        state.dc.clip_attr()
                    ));
                }
            },
            2 => return unsupported(kind, "SVG 1.1 has no interoperable Gouraud triangle mesh"),
            _ => return malformed(kind, "invalid gradient fill mode"),
        }
        Ok(output)
    }

    fn ext_text_out(
        &self,
        data: &[u8],
        kind: EmrType,
        state: &mut RenderState,
    ) -> Result<Vec<String>> {
        let header = read_record::<EmrExtTextOutHeader>(data, kind)?;
        self.render_emr_text(
            data,
            &header.text,
            kind == EmrType::ExtTextOutW,
            kind,
            state,
        )
    }

    fn poly_text_out(
        &self,
        data: &[u8],
        kind: EmrType,
        state: &mut RenderState,
    ) -> Result<Vec<String>> {
        let header = read_record::<EmrPolyTextOutHeader>(data, kind)?;
        let count = bounded_count(header.num_strings, kind)?;
        if count > self.limits.max_svg_elements {
            return malformed(kind, "PolyTextOut string count exceeds SVG element limit");
        }
        let base = size_of::<EmrPolyTextOutHeader>();
        checked_range(base, count, size_of::<EmrTextInfo>(), data.len(), kind)?;

        // Preflight the whole record before decoding any repeated string or
        // spacing range. A compact PolyTextOut record may legally point many
        // entries at the same bytes, so per-entry limits alone do not bound
        // the aggregate temporary allocations or generated SVG.
        let unicode = kind == EmrType::PolyTextOutW;
        let mut aggregate_chars = 0usize;
        let mut aggregate_svg_bytes = 0usize;
        for index in 0..count {
            let info =
                read_record::<EmrTextInfo>(&data[base + index * size_of::<EmrTextInfo>()..], kind)?;
            let chars = bounded_count(info.num_chars, kind)?;
            if chars > self.limits.max_path_points {
                return malformed(kind, "text character count exceeds configured limit");
            }
            aggregate_chars = aggregate_chars
                .checked_add(chars)
                .ok_or_else(|| Error::ParseError("PolyTextOut character count overflow".into()))?;
            if aggregate_chars > self.limits.max_path_points {
                return malformed(
                    kind,
                    "PolyTextOut aggregate character count exceeds configured path limit",
                );
            }

            let string_offset = payload_offset(info.off_string, kind)?;
            checked_range(
                string_offset,
                chars,
                if unicode { 2 } else { 1 },
                data.len(),
                kind,
            )?;
            if info.off_dx != 0 {
                let spacing_count = chars
                    .checked_mul(if info.options & text_options::PDY != 0 {
                        2
                    } else {
                        1
                    })
                    .ok_or_else(|| Error::ParseError("text spacing count overflow".into()))?;
                checked_range(
                    payload_offset(info.off_dx, kind)?,
                    spacing_count,
                    size_of::<i32>(),
                    data.len(),
                    kind,
                )?;
            }

            // Numeric position lists, escaped UTF-8, attributes, optional
            // background and clip markup all fit within this conservative
            // per-entry bound. This is deliberately checked before rendering.
            let entry_bound = chars
                .checked_mul(128)
                .and_then(|bytes| bytes.checked_add(1024))
                .ok_or_else(|| Error::ParseError("PolyTextOut SVG size overflow".into()))?;
            aggregate_svg_bytes = aggregate_svg_bytes
                .checked_add(entry_bound)
                .ok_or_else(|| Error::ParseError("PolyTextOut SVG size overflow".into()))?;
            if aggregate_svg_bytes > self.limits.max_output_bytes {
                return malformed(
                    kind,
                    "PolyTextOut expansion exceeds configured output limit",
                );
            }
        }

        let mut output = Vec::new();
        for index in 0..count {
            let info =
                read_record::<EmrTextInfo>(&data[base + index * size_of::<EmrTextInfo>()..], kind)?;
            output.extend(self.render_emr_text(data, &info, unicode, kind, state)?);
        }
        Ok(output)
    }

    fn small_text_out(
        &self,
        data: &[u8],
        kind: EmrType,
        state: &mut RenderState,
    ) -> Result<Vec<String>> {
        let header = read_record::<EmrSmallTextOutHeader>(data, kind)?;
        let count = bounded_count(header.num_chars, kind)?;
        if count > self.limits.max_path_points || count > self.limits.max_output_bytes {
            return malformed(
                kind,
                "SmallTextOut character count exceeds configured limit",
            );
        }
        let unicode = header.fu_options & 0x0200 == 0;
        let offset = size_of::<EmrSmallTextOutHeader>();
        let text = decode_text(data, offset, count, unicode, state.dc.font.charset, kind)?;
        let reference = PointL {
            x: header.x,
            y: header.y,
        };
        render_text_element(&text, reference, header.fu_options, None, None, state, kind)
    }

    fn render_emr_text(
        &self,
        data: &[u8],
        info: &EmrTextInfo,
        unicode: bool,
        kind: EmrType,
        state: &mut RenderState,
    ) -> Result<Vec<String>> {
        if info.options & text_options::GLYPH_INDEX != 0 {
            return unsupported(
                kind,
                "glyph-index text requires the original font glyph map",
            );
        }
        let count = bounded_count(info.num_chars, kind)?;
        if count > self.limits.max_path_points || count > self.limits.max_output_bytes {
            return malformed(kind, "text character count exceeds configured limit");
        }
        let string_offset = payload_offset(info.off_string, kind)?;
        let text = decode_text(
            data,
            string_offset,
            count,
            unicode,
            state.dc.font.charset,
            kind,
        )?;
        let dx = if info.off_dx == 0 {
            None
        } else {
            let offset = payload_offset(info.off_dx, kind)?;
            let values = count
                .checked_mul(if info.options & text_options::PDY != 0 {
                    2
                } else {
                    1
                })
                .ok_or_else(|| Error::ParseError("text spacing count overflow".into()))?;
            Some(read_i32_array(data, offset, values, kind)?)
        };
        render_text_element(
            &text,
            info.reference,
            info.options,
            Some(info.rectangle),
            dx.as_deref(),
            state,
            kind,
        )
    }

    fn build_svg(&self, elements: &[String], state: &RenderState) -> Result<String> {
        let header = &self.parser.header;
        let width = header.width().unsigned_abs().max(1);
        let height = header.height().unsigned_abs().max(1);
        let estimated = state
            .defs
            .iter()
            .chain(elements)
            .try_fold(256usize, |total, value| total.checked_add(value.len()))
            .ok_or_else(|| Error::ParseError("EMF SVG size overflow".into()))?;
        if estimated > self.limits.max_output_bytes {
            return Err(Error::ParseError(format!(
                "EMF SVG output exceeds limit {} bytes",
                self.limits.max_output_bytes
            )));
        }
        let mut svg = String::with_capacity(estimated);
        write!(
            &mut svg,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{}\" height=\"{}\" viewBox=\"{} {} {} {}\">",
            width, height, header.bounds.0, header.bounds.1, width, height
        )
        .map_err(|_| Error::ParseError("failed to write EMF SVG header".into()))?;
        if !state.defs.is_empty() {
            svg.push_str("<defs>");
            for definition in &state.defs {
                svg.push_str(definition);
            }
            svg.push_str("</defs>");
        }
        for element in elements {
            svg.push_str(element);
        }
        svg.push_str("</svg>");
        if svg.len() > self.limits.max_output_bytes {
            return Err(Error::ParseError(format!(
                "EMF SVG output exceeds limit {} bytes",
                self.limits.max_output_bytes
            )));
        }
        Ok(svg)
    }
}

fn dib_limits(limits: crate::Limits) -> DibLimits {
    DibLimits {
        max_input_bytes: limits.max_uncompressed_bytes,
        max_width: limits.max_width,
        max_height: limits.max_height,
        max_pixels: limits.max_pixels,
        max_palette_entries: 4096,
        max_decoded_bytes: usize::try_from(limits.max_pixels.saturating_mul(4))
            .unwrap_or(usize::MAX)
            .min(limits.max_uncompressed_bytes),
        max_output_bytes: limits.max_output_bytes,
    }
}

fn emfplus_limits(limits: crate::Limits) -> crate::emfplus::renderer::RendererLimits {
    let object_slots = limits
        .max_objects
        .min(crate::emfplus::MAX_EMFPLUS_OBJECT_SLOTS);
    crate::emfplus::renderer::RendererLimits {
        parser: crate::emfplus::ParserLimits {
            max_bytes: limits.max_uncompressed_bytes,
            max_records: limits.max_records,
            max_object_slots: object_slots,
        },
        playback: crate::emfplus::playback::PlaybackLimits {
            max_points: limits.max_path_points,
            max_output: limits.max_svg_elements,
            max_depth: limits.max_state_depth,
            max_objects: object_slots,
            max_bytes: limits.max_uncompressed_bytes,
        },
        objects: crate::emfplus::objects::DecodeLimits {
            max_bytes: limits.max_uncompressed_bytes,
            max_points: limits.max_path_points,
            max_recursion: limits.max_state_depth,
        },
        svg: crate::emfplus::svg::SvgLimits {
            max_output_bytes: limits.max_output_bytes,
            max_elements: limits.max_svg_elements,
            max_definitions: limits.max_objects,
            max_path_commands: limits.max_path_points,
            max_image_bytes: limits.max_uncompressed_bytes,
            max_diagnostics: limits.max_records.min(1024),
        },
    }
}

fn svg_element_count(fragment: &str) -> usize {
    fragment
        .as_bytes()
        .windows(2)
        .filter(|pair| pair[0] == b'<' && !matches!(pair[1], b'/' | b'!' | b'?'))
        .count()
}

fn render_line(start: (f64, f64), end: (f64, f64), dc: &DeviceContext) -> String {
    format!(
        "<line x1=\"{}\" y1=\"{}\" x2=\"{}\" y2=\"{}\" {} {}/>",
        fmt(start.0),
        fmt(start.1),
        fmt(end.0),
        fmt(end.1),
        dc.get_stroke_attrs(),
        dc.clip_attr()
    )
}

fn path_element(
    mut builder: PathBuilder,
    dc: &DeviceContext,
    fill: bool,
    stroke: bool,
) -> Option<String> {
    builder.optimize();
    if builder.is_empty() {
        return None;
    }
    let fill_attrs = if fill {
        dc.get_fill_attr()
    } else {
        "fill=\"none\"".to_string()
    };
    let stroke_attrs = if stroke {
        dc.get_stroke_attrs()
    } else {
        "stroke=\"none\"".to_string()
    };
    let rule = if fill {
        dc.get_fill_rule().unwrap_or_default()
    } else {
        String::new()
    };
    Some(format!(
        "<path d=\"{}\" {} {} {} {}/>",
        builder.build(),
        fill_attrs,
        stroke_attrs,
        rule,
        dc.clip_attr()
    ))
}

fn append_path(target: &mut PathBuilder, source: PathBuilder) {
    // Replaying the built SVG grammar would be fragile. PathBuilder exposes
    // typed append specifically for this operation.
    target.append(source);
}

fn rectangle_path(rect: RectL, dc: &DeviceContext) -> Result<PathBuilder> {
    if rect.left == rect.right || rect.top == rect.bottom {
        return Err(Error::ParseError("empty EMF rectangle".into()));
    }
    let points = [
        (rect.left, rect.top),
        (rect.right, rect.top),
        (rect.right, rect.bottom),
        (rect.left, rect.bottom),
    ];
    let mut builder = PathBuilder::new();
    for (index, point) in points.into_iter().enumerate() {
        let point = transform(point, dc);
        if index == 0 {
            builder.move_to(point.0, point.1);
        } else {
            builder.line_to(point.0, point.1);
        }
    }
    builder.close();
    Ok(builder)
}

fn ellipse_path(rect: RectL, dc: &DeviceContext) -> Result<PathBuilder> {
    let cx = (f64::from(rect.left) + f64::from(rect.right)) / 2.0;
    let cy = (f64::from(rect.top) + f64::from(rect.bottom)) / 2.0;
    let rx = (f64::from(rect.right) - f64::from(rect.left)).abs() / 2.0;
    let ry = (f64::from(rect.bottom) - f64::from(rect.top)).abs() / 2.0;
    if rx == 0.0 || ry == 0.0 {
        return Err(Error::ParseError("empty EMF ellipse".into()));
    }
    let k = 0.552_284_749_830_793_6;
    let logical = [
        (
            (cx + rx, cy),
            (cx + rx, cy + k * ry),
            (cx + k * rx, cy + ry),
            (cx, cy + ry),
        ),
        (
            (cx, cy + ry),
            (cx - k * rx, cy + ry),
            (cx - rx, cy + k * ry),
            (cx - rx, cy),
        ),
        (
            (cx - rx, cy),
            (cx - rx, cy - k * ry),
            (cx - k * rx, cy - ry),
            (cx, cy - ry),
        ),
        (
            (cx, cy - ry),
            (cx + k * rx, cy - ry),
            (cx + rx, cy - k * ry),
            (cx + rx, cy),
        ),
    ];
    let mut builder = PathBuilder::new();
    let start = dc.transform_point(logical[0].0.0, logical[0].0.1);
    builder.move_to(start.0, start.1);
    for (_, a, b, c) in logical {
        let a = dc.transform_point(a.0, a.1);
        let b = dc.transform_point(b.0, b.1);
        let c = dc.transform_point(c.0, c.1);
        builder.cubic_to(a.0, a.1, b.0, b.1, c.0, c.1);
    }
    builder.close();
    Ok(builder)
}

fn round_rect_path(rect: RectL, corner: SizeL, dc: &DeviceContext) -> Result<PathBuilder> {
    let left = f64::from(rect.left.min(rect.right));
    let right = f64::from(rect.left.max(rect.right));
    let top = f64::from(rect.top.min(rect.bottom));
    let bottom = f64::from(rect.top.max(rect.bottom));
    let rx = (f64::from(corner.cx).abs() / 2.0).min((right - left) / 2.0);
    let ry = (f64::from(corner.cy).abs() / 2.0).min((bottom - top) / 2.0);
    if rx == 0.0 || ry == 0.0 {
        return rectangle_path(rect, dc);
    }
    let mut b = PathBuilder::new();
    let k = 0.552_284_749_830_793_6;
    let p = |x, y| dc.transform_point(x, y);
    let s = p(left + rx, top);
    b.move_to(s.0, s.1);
    let q = p(right - rx, top);
    b.line_to(q.0, q.1);
    let a = p(right - rx + k * rx, top);
    let c = p(right, top + ry - k * ry);
    let e = p(right, top + ry);
    b.cubic_to(a.0, a.1, c.0, c.1, e.0, e.1);
    let q = p(right, bottom - ry);
    b.line_to(q.0, q.1);
    let a = p(right, bottom - ry + k * ry);
    let c = p(right - rx + k * rx, bottom);
    let e = p(right - rx, bottom);
    b.cubic_to(a.0, a.1, c.0, c.1, e.0, e.1);
    let q = p(left + rx, bottom);
    b.line_to(q.0, q.1);
    let a = p(left + rx - k * rx, bottom);
    let c = p(left, bottom - ry + k * ry);
    let e = p(left, bottom - ry);
    b.cubic_to(a.0, a.1, c.0, c.1, e.0, e.1);
    let q = p(left, top + ry);
    b.line_to(q.0, q.1);
    let a = p(left, top + ry - k * ry);
    let c = p(left + rx - k * rx, top);
    let e = p(left + rx, top);
    b.cubic_to(a.0, a.1, c.0, c.1, e.0, e.1);
    b.close();
    Ok(b)
}

fn arc_path(arc: EmrArc, kind: EmrType, dc: &DeviceContext) -> Result<(PathBuilder, (f64, f64))> {
    let cx = (f64::from(arc.rect.left) + f64::from(arc.rect.right)) / 2.0;
    let cy = (f64::from(arc.rect.top) + f64::from(arc.rect.bottom)) / 2.0;
    let rx = (f64::from(arc.rect.right) - f64::from(arc.rect.left)).abs() / 2.0;
    let ry = (f64::from(arc.rect.bottom) - f64::from(arc.rect.top)).abs() / 2.0;
    if rx == 0.0 || ry == 0.0 {
        return Err(Error::ParseError(format!(
            "{} has an empty bounds rectangle",
            kind.name()
        )));
    }
    let start_angle =
        ((f64::from(arc.start.y) - cy) / ry).atan2((f64::from(arc.start.x) - cx) / rx);
    let end_angle = ((f64::from(arc.end.y) - cy) / ry).atan2((f64::from(arc.end.x) - cx) / rx);
    let mut sweep = end_angle - start_angle;
    if dc.arc_direction {
        while sweep <= 0.0 {
            sweep += 2.0 * PI;
        }
    } else {
        while sweep >= 0.0 {
            sweep -= 2.0 * PI;
        }
    }
    let steps = ((sweep.abs() / (PI / 16.0)).ceil() as usize).max(1);
    let logical_start = (cx + rx * start_angle.cos(), cy + ry * start_angle.sin());
    let logical_end = (cx + rx * end_angle.cos(), cy + ry * end_angle.sin());
    let mut b = PathBuilder::new();
    if kind == EmrType::Pie {
        let p = dc.transform_point(cx, cy);
        b.move_to(p.0, p.1);
        let p = dc.transform_point(logical_start.0, logical_start.1);
        b.line_to(p.0, p.1);
    } else if kind == EmrType::ArcTo {
        let p = dc.transform_point(dc.current_pos.0, dc.current_pos.1);
        b.move_to(p.0, p.1);
        let p = dc.transform_point(logical_start.0, logical_start.1);
        b.line_to(p.0, p.1);
    } else {
        let p = dc.transform_point(logical_start.0, logical_start.1);
        b.move_to(p.0, p.1);
    }
    for i in 1..=steps {
        let a = start_angle + sweep * (i as f64 / steps as f64);
        let p = dc.transform_point(cx + rx * a.cos(), cy + ry * a.sin());
        b.line_to(p.0, p.1);
    }
    if kind == EmrType::Pie {
        let p = dc.transform_point(cx, cy);
        b.line_to(p.0, p.1);
        b.close();
    } else if kind == EmrType::Chord {
        b.close();
    }
    Ok((b, logical_end))
}

fn angle_arc_path(arc: EmrAngleArc, dc: &DeviceContext) -> Result<(PathBuilder, (f64, f64))> {
    if arc.radius == 0 || !arc.start_angle.is_finite() || !arc.sweep_angle.is_finite() {
        return Err(Error::ParseError("invalid AngleArc".into()));
    }
    let cx = f64::from(arc.center.x);
    let cy = f64::from(arc.center.y);
    let radius = f64::from(arc.radius);
    let start = f64::from(arc.start_angle).to_radians();
    let sweep = f64::from(arc.sweep_angle).to_radians();
    let steps = ((sweep.abs() / (PI / 16.0)).ceil() as usize).max(1);
    let first = (cx + radius * start.cos(), cy - radius * start.sin());
    let end_angle = start + sweep;
    let end = (cx + radius * end_angle.cos(), cy - radius * end_angle.sin());
    let mut b = PathBuilder::new();
    let current = dc.transform_point(dc.current_pos.0, dc.current_pos.1);
    b.move_to(current.0, current.1);
    let p = dc.transform_point(first.0, first.1);
    b.line_to(p.0, p.1);
    for i in 1..=steps {
        let a = start + sweep * (i as f64 / steps as f64);
        let p = dc.transform_point(cx + radius * a.cos(), cy - radius * a.sin());
        b.line_to(p.0, p.1);
    }
    Ok((b, end))
}

fn exclusion_path(rect: RectL, dc: &DeviceContext) -> Result<String> {
    let mut outer = PathBuilder::new();
    outer.move_to(-1e9, -1e9);
    outer.line_to(1e9, -1e9);
    outer.line_to(1e9, 1e9);
    outer.line_to(-1e9, 1e9);
    outer.close();
    let inner = rectangle_path(
        RectL {
            left: rect.left,
            top: rect.bottom,
            right: rect.right,
            bottom: rect.top,
        },
        dc,
    )?;
    outer.append(inner);
    Ok(outer.build())
}

fn region_path(
    data: &[u8],
    dc: &DeviceContext,
    kind: EmrType,
    max_path_points: usize,
    max_output_bytes: usize,
) -> Result<String> {
    require_len(data, 32, kind)?;
    let header_size = usize::try_from(read_u32(data, 0, kind)?)
        .map_err(|_| Error::ParseError("region header size overflow".into()))?;
    if header_size < 32 {
        return malformed(kind, "short region header");
    }
    let count = bounded_count(read_u32(data, 8, kind)?, kind)?;
    let commands = count
        .checked_mul(5)
        .ok_or_else(|| Error::ParseError("region path command count overflow".into()))?;
    if commands > max_path_points {
        return malformed(kind, "region path exceeds configured point limit");
    }
    let estimated = count
        .checked_mul(128)
        .ok_or_else(|| Error::ParseError("region SVG size estimate overflow".into()))?;
    if estimated > max_output_bytes {
        return malformed(kind, "region path exceeds configured output limit");
    }
    let declared = usize::try_from(read_u32(data, 12, kind)?)
        .map_err(|_| Error::ParseError("region data size overflow".into()))?;
    let bytes = count
        .checked_mul(16)
        .ok_or_else(|| Error::ParseError("region rectangles overflow".into()))?;
    if declared < bytes {
        return malformed(kind, "region data size is smaller than its rectangles");
    }
    checked_range(header_size, count, 16, data.len(), kind)?;
    let mut path = PathBuilder::new();
    for i in 0..count {
        let rect = read_record::<RectL>(&data[header_size + i * 16..], kind)?;
        path.append(rectangle_path(rect, dc)?);
    }
    Ok(path.build())
}

fn render_text_element(
    text: &str,
    reference: PointL,
    options: u32,
    rectangle: Option<RectL>,
    dx: Option<&[i32]>,
    state: &mut RenderState,
    kind: EmrType,
) -> Result<Vec<String>> {
    let update = state.dc.text_align & text_align::UPDATECP != 0;
    let logical = if update {
        state.dc.current_pos
    } else {
        (f64::from(reference.x), f64::from(reference.y))
    };
    let origin = state.dc.transform_point(logical.0, logical.1);
    let mut output = Vec::new();
    if options & text_options::OPAQUE != 0 || (state.dc.bg_mode == 2 && rectangle.is_some()) {
        if let Some(rect) = rectangle {
            let p = rectangle_path(rect, &state.dc)?;
            output.push(format!(
                "<path d=\"{}\" fill=\"{}\" stroke=\"none\" {}/>",
                p.build(),
                state.dc.bg_color.to_svg_color(),
                state.dc.clip_attr()
            ));
        }
    }
    let anchor = match state.dc.text_align & 6 {
        6 => "middle",
        2 => "end",
        _ => "start",
    };
    let baseline = if state.dc.text_align & 0x18 == 0x18 {
        "alphabetic"
    } else if state.dc.text_align & 8 != 0 {
        "text-after-edge"
    } else {
        "text-before-edge"
    };
    let direction = if state.dc.text_align & text_align::RTLREADING != 0
        || options & text_options::RTLREADING != 0
        || state.dc.layout & 1 != 0
    {
        "rtl"
    } else {
        "ltr"
    };
    let mut attrs = String::new();
    if options & text_options::CLIPPED != 0 {
        let rect = rectangle.ok_or_else(|| {
            Error::ParseError(format!("{} ETO_CLIPPED without rectangle", kind.name()))
        })?;
        let id = state.fresh_id("text-clip");
        state.add_definition(format!(
            "<clipPath id=\"{}\"><path d=\"{}\"/></clipPath>",
            id,
            rectangle_path(rect, &state.dc)?.build()
        ));
        write!(&mut attrs, " clip-path=\"url(#{})\"", id).ok();
    } else {
        write!(&mut attrs, " {}", state.dc.clip_attr()).ok();
    }
    if state.dc.font.escapement != 0.0 {
        write!(
            &mut attrs,
            " transform=\"rotate({} {} {})\"",
            -state.dc.font.escapement / 10.0,
            fmt(origin.0),
            fmt(origin.1)
        )
        .ok();
    }
    let mut x_values = vec![origin.0];
    let mut y_values = vec![origin.1];
    let mut advance = (0.0, 0.0);
    if let Some(values) = dx {
        let pdy = options & text_options::PDY != 0;
        let stride = if pdy { 2 } else { 1 };
        for chunk in values
            .chunks_exact(stride)
            .take(text.chars().count().saturating_sub(1))
        {
            advance.0 += f64::from(chunk[0]);
            if pdy {
                advance.1 += f64::from(chunk[1]);
            }
            let p = state
                .dc
                .transform_point(logical.0 + advance.0, logical.1 + advance.1);
            x_values.push(p.0);
            y_values.push(p.1);
        }
    }
    let x = x_values
        .iter()
        .map(|v| fmt(*v))
        .collect::<Vec<_>>()
        .join(" ");
    let y = y_values
        .iter()
        .map(|v| fmt(*v))
        .collect::<Vec<_>>()
        .join(" ");
    let mut escaped_text = String::with_capacity(text.len());
    write_xml_escaped(&mut escaped_text, text);
    output.push(format!("<text x=\"{}\" y=\"{}\" fill=\"{}\" {} text-anchor=\"{}\" dominant-baseline=\"{}\" direction=\"{}\"{}>{}</text>",x,y,state.dc.text_color.to_svg_color(),state.dc.font.to_svg_attrs(),anchor,baseline,direction,attrs,escaped_text));
    if update {
        let estimated = if let Some(values) = dx {
            let stride = if options & text_options::PDY != 0 {
                2
            } else {
                1
            };
            values.iter().step_by(stride).map(|v| f64::from(*v)).sum()
        } else {
            state.dc.font.height.abs() * 0.6 * text.chars().count() as f64
        };
        state.dc.current_pos = (logical.0 + estimated, logical.1);
    }
    Ok(output)
}

fn decode_text(
    data: &[u8],
    offset: usize,
    count: usize,
    unicode: bool,
    charset: u8,
    kind: EmrType,
) -> Result<String> {
    let bytes = count
        .checked_mul(if unicode { 2 } else { 1 })
        .ok_or_else(|| Error::ParseError("text byte length overflow".into()))?;
    let end = offset
        .checked_add(bytes)
        .ok_or_else(|| Error::ParseError("text range overflow".into()))?;
    let raw = data
        .get(offset..end)
        .ok_or_else(|| Error::ParseError(format!("{} text exceeds payload", kind.name())))?;
    if unicode {
        let units = raw
            .chunks_exact(2)
            .map(|v| u16::from_le_bytes([v[0], v[1]]))
            .collect::<Vec<_>>();
        Ok(String::from_utf16_lossy(&units)
            .trim_end_matches('\0')
            .to_string())
    } else {
        decode_ansi(raw, charset, kind)
    }
}

fn decode_ansi(raw: &[u8], charset: u8, kind: EmrType) -> Result<String> {
    match charset {
        0 | 1 | 255 => Ok(raw
            .iter()
            .filter(|&&b| b != 0)
            .map(|&b| cp1252(b))
            .collect()),
        2 => Ok(raw
            .iter()
            .filter(|&&b| b != 0)
            .map(|&b| char::from_u32(0xf000 + u32::from(b)).unwrap_or('\u{fffd}'))
            .collect()),
        128 | 129 | 134 | 136 => unsupported(
            kind,
            "multibyte ANSI charset needs litchi-codepage integration",
        ),
        _ => unsupported(kind, "unsupported LOGFONT character set"),
    }
}
fn cp1252(byte: u8) -> char {
    const SPECIAL: [char; 32] = [
        '€', '\u{0081}', '‚', 'ƒ', '„', '…', '†', '‡', 'ˆ', '‰', 'Š', '‹', 'Œ', '\u{008d}', 'Ž',
        '\u{008f}', '\u{0090}', '‘', '’', '“', '”', '•', '–', '—', '˜', '™', 'š', '›', 'œ',
        '\u{009d}', 'ž', 'Ÿ',
    ];
    if (0x80..=0x9f).contains(&byte) {
        SPECIAL[usize::from(byte - 0x80)]
    } else {
        char::from(byte)
    }
}

fn parse_font(data: &[u8], kind: EmrType) -> Result<(u32, Font)> {
    require_len(data, 4 + size_of::<LogFontW>() + 64, kind)?;
    let handle = read_u32(data, 0, kind)?;
    let log = read_record::<LogFontW>(&data[4..], kind)?;
    let face_offset = 4 + size_of::<LogFontW>();
    let units = data[face_offset..face_offset + 64]
        .chunks_exact(2)
        .map(|v| u16::from_le_bytes([v[0], v[1]]))
        .take_while(|&v| v != 0)
        .collect::<Vec<_>>();
    let face_name = String::from_utf16_lossy(&units);
    Ok((
        handle,
        Font {
            height: f64::from(log.height),
            width: f64::from(log.width),
            escapement: f64::from(log.escapement),
            orientation: f64::from(log.orientation),
            weight: log.weight,
            italic: log.italic != 0,
            underline: log.underline != 0,
            strike_out: log.strike_out != 0,
            charset: log.char_set,
            face_name,
        },
    ))
}
fn insert_object(
    state: &mut RenderState,
    handle: u32,
    object: GdiObject,
    kind: EmrType,
) -> Result<()> {
    if stock_objects::is_stock_object(handle) || !state.insert_object(handle, object) {
        return malformed(kind, "duplicate or stock object handle");
    }
    Ok(())
}
fn vertex(vertices: &[TriVertex], index: u32, kind: EmrType) -> Result<&TriVertex> {
    vertices
        .get(
            usize::try_from(index)
                .map_err(|_| Error::ParseError("gradient vertex index overflow".into()))?,
        )
        .ok_or_else(|| {
            Error::ParseError(format!(
                "{} gradient vertex index out of range",
                kind.name()
            ))
        })
}
fn vertex_color(v: &TriVertex) -> String {
    format!("#{:02x}{:02x}{:02x}", v.red >> 8, v.green >> 8, v.blue >> 8)
}
fn vertex_alpha(v: &TriVertex) -> String {
    format!("{:.3}", f64::from(v.alpha) / 65535.0)
}
fn transform(point: (i32, i32), dc: &DeviceContext) -> (f64, f64) {
    dc.transform_point(f64::from(point.0), f64::from(point.1))
}
fn logical_to_device_matrix(dc: &DeviceContext) -> [f64; 6] {
    let origin = dc.transform_point(0.0, 0.0);
    let x_axis = dc.transform_point(1.0, 0.0);
    let y_axis = dc.transform_point(0.0, 1.0);
    [
        x_axis.0 - origin.0,
        x_axis.1 - origin.1,
        y_axis.0 - origin.0,
        y_axis.1 - origin.1,
        origin.0,
        origin.1,
    ]
}
fn dpi(pixels: i32, mm: i32) -> f64 {
    if pixels > 0 && mm > 0 {
        f64::from(pixels) * 25.4 / f64::from(mm)
    } else {
        96.0
    }
}
fn fmt(value: f64) -> String {
    let mut s = String::new();
    write_num(&mut s, value);
    s
}
fn payload_offset(offset: u32, kind: EmrType) -> Result<usize> {
    if offset < 8 {
        return malformed(kind, "record-relative offset points into EMR header");
    }
    usize::try_from(offset - 8).map_err(|_| Error::ParseError("record offset overflow".into()))
}
fn validate_xform(x: &XForm, kind: EmrType) -> Result<()> {
    if [x.m11, x.m12, x.m21, x.m22, x.dx, x.dy]
        .iter()
        .all(|v| v.is_finite())
    {
        Ok(())
    } else {
        malformed(kind, "non-finite world transform")
    }
}
fn scale_i32(value: i32, num: i32, den: i32, kind: EmrType) -> Result<i32> {
    let scaled = i64::from(value)
        .checked_mul(i64::from(num))
        .ok_or_else(|| Error::ParseError(format!("{} extent overflow", kind.name())))?
        / i64::from(den);
    i32::try_from(scaled)
        .map_err(|_| Error::ParseError(format!("{} extent outside i32", kind.name())))
}
fn bounded_count(value: u32, kind: EmrType) -> Result<usize> {
    let count =
        usize::try_from(value).map_err(|_| Error::ParseError("item count overflow".into()))?;
    if count > MAX_RECORD_ITEMS {
        return malformed(kind, "record item count exceeds safety limit");
    }
    Ok(count)
}
fn checked_range(
    offset: usize,
    count: usize,
    stride: usize,
    len: usize,
    kind: EmrType,
) -> Result<usize> {
    let bytes = count
        .checked_mul(stride)
        .ok_or_else(|| Error::ParseError("record array length overflow".into()))?;
    let end = offset
        .checked_add(bytes)
        .ok_or_else(|| Error::ParseError("record array range overflow".into()))?;
    if end > len {
        return malformed(kind, "record array exceeds payload");
    }
    Ok(end)
}
fn read_u32(data: &[u8], offset: usize, kind: EmrType) -> Result<u32> {
    let bytes = data
        .get(
            offset
                ..offset
                    .checked_add(4)
                    .ok_or_else(|| Error::ParseError("field offset overflow".into()))?,
        )
        .ok_or_else(|| Error::ParseError(format!("{} has a truncated u32 field", kind.name())))?;
    Ok(u32::from_le_bytes(bytes.try_into().expect("four bytes")))
}
fn read_i32(data: &[u8], offset: usize, kind: EmrType) -> Result<i32> {
    Ok(read_u32(data, offset, kind)? as i32)
}
fn read_u32_array(data: &[u8], offset: usize, count: usize, kind: EmrType) -> Result<Vec<u32>> {
    checked_range(offset, count, 4, data.len(), kind)?;
    (0..count)
        .map(|i| read_u32(data, offset + i * 4, kind))
        .collect()
}
fn read_i32_array(data: &[u8], offset: usize, count: usize, kind: EmrType) -> Result<Vec<i32>> {
    read_u32_array(data, offset, count, kind).map(|v| v.into_iter().map(|x| x as i32).collect())
}
fn read_record<T: FromBytes>(data: &[u8], kind: EmrType) -> Result<T> {
    T::read_from_prefix(data)
        .map(|(v, _)| v)
        .map_err(|_| Error::ParseError(format!("{} payload is truncated", kind.name())))
}
fn require_len(data: &[u8], needed: usize, kind: EmrType) -> Result<()> {
    if data.len() < needed {
        malformed(kind, "payload is truncated")
    } else {
        Ok(())
    }
}
fn require_len_exact(data: &[u8], needed: usize, kind: EmrType) -> Result<()> {
    if data.len() != needed {
        malformed(kind, "unexpected payload length")
    } else {
        Ok(())
    }
}
fn malformed<T>(kind: EmrType, message: &str) -> Result<T> {
    Err(Error::ParseError(format!("{}: {}", kind.name(), message)))
}
fn unsupported<T>(kind: EmrType, message: &str) -> Result<T> {
    Err(Error::Unsupported(format!("{}: {}", kind.name(), message)))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn dword(out: &mut Vec<u8>, value: u32) {
        out.extend_from_slice(&value.to_le_bytes());
    }

    fn sdword(out: &mut Vec<u8>, value: i32) {
        out.extend_from_slice(&value.to_le_bytes());
    }

    fn record(kind: EmrType, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        dword(&mut out, kind as u32);
        dword(&mut out, u32::try_from(8 + payload.len()).unwrap());
        out.extend_from_slice(payload);
        while out.len() % 4 != 0 {
            out.push(0);
        }
        let size = u32::try_from(out.len()).unwrap();
        out[4..8].copy_from_slice(&size.to_le_bytes());
        out
    }

    fn classic_emf(records: Vec<Vec<u8>>) -> Vec<u8> {
        let eof = record(EmrType::Eof, &[0, 0, 0, 0, 0, 0, 0, 0, 20, 0, 0, 0]);
        let total = 88 + records.iter().map(Vec::len).sum::<usize>() + eof.len();
        let mut emf = Vec::with_capacity(total);
        dword(&mut emf, EmrType::Header as u32);
        dword(&mut emf, 88);
        for value in [0, 0, 100, 100, 0, 0, 2646, 2646] {
            sdword(&mut emf, value);
        }
        dword(&mut emf, 0x464d_4520);
        dword(&mut emf, 0x0001_0000);
        dword(&mut emf, u32::try_from(total).unwrap());
        dword(&mut emf, u32::try_from(records.len() + 2).unwrap());
        emf.extend_from_slice(&2u16.to_le_bytes());
        emf.extend_from_slice(&0u16.to_le_bytes());
        dword(&mut emf, 0);
        dword(&mut emf, 0);
        dword(&mut emf, 0);
        sdword(&mut emf, 100);
        sdword(&mut emf, 100);
        sdword(&mut emf, 26);
        sdword(&mut emf, 26);
        assert_eq!(emf.len(), 88);
        for value in records {
            emf.extend_from_slice(&value);
        }
        emf.extend_from_slice(&eof);
        emf
    }

    fn put_u32(data: &mut [u8], offset: usize, value: u32) {
        data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    }

    fn put_i32(data: &mut [u8], offset: usize, value: i32) {
        data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    }

    fn put_f32(data: &mut [u8], offset: usize, value: f32) {
        data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    }

    fn one_pixel_dib_parts() -> ([u8; 40], [u8; 4]) {
        let mut info = [0u8; 40];
        put_u32(&mut info, 0, 40);
        put_i32(&mut info, 4, 1);
        put_i32(&mut info, 8, 1);
        info[12..14].copy_from_slice(&1u16.to_le_bytes());
        info[14..16].copy_from_slice(&24u16.to_le_bytes());
        put_u32(&mut info, 20, 4);
        (info, [0, 0, 255, 0])
    }

    fn stretch_dib_payload() -> Vec<u8> {
        const FIXED: usize = 72;
        let (info, bits) = one_pixel_dib_parts();
        let mut payload = vec![0u8; FIXED + info.len() + bits.len()];
        put_i32(&mut payload, 16, 10);
        put_i32(&mut payload, 20, 20);
        put_i32(&mut payload, 24, 0);
        put_i32(&mut payload, 28, 0);
        put_i32(&mut payload, 32, 1);
        put_i32(&mut payload, 36, 1);
        put_u32(&mut payload, 40, (8 + FIXED) as u32);
        put_u32(&mut payload, 44, info.len() as u32);
        put_u32(&mut payload, 48, (8 + FIXED + info.len()) as u32);
        put_u32(&mut payload, 52, bits.len() as u32);
        put_u32(&mut payload, 56, 0);
        put_u32(&mut payload, 60, 0x00cc_0020);
        put_i32(&mut payload, 64, 20);
        put_i32(&mut payload, 68, 30);
        payload[FIXED..FIXED + info.len()].copy_from_slice(&info);
        payload[FIXED + info.len()..].copy_from_slice(&bits);
        payload
    }

    fn plg_blt_payload() -> Vec<u8> {
        const FIXED: usize = 132;
        let (info, bits) = one_pixel_dib_parts();
        let mut payload = vec![0u8; FIXED + info.len() + bits.len()];
        for (offset, point) in [(16, (10, 20)), (24, (30, 20)), (32, (10, 50))] {
            put_i32(&mut payload, offset, point.0);
            put_i32(&mut payload, offset + 4, point.1);
        }
        put_i32(&mut payload, 40, 0);
        put_i32(&mut payload, 44, 0);
        put_i32(&mut payload, 48, 1);
        put_i32(&mut payload, 52, 1);
        put_f32(&mut payload, 56, 1.0);
        put_f32(&mut payload, 68, 1.0);
        put_u32(&mut payload, 84, 0);
        put_u32(&mut payload, 88, (8 + FIXED) as u32);
        put_u32(&mut payload, 92, info.len() as u32);
        put_u32(&mut payload, 96, (8 + FIXED + info.len()) as u32);
        put_u32(&mut payload, 100, bits.len() as u32);
        payload[FIXED..FIXED + info.len()].copy_from_slice(&info);
        payload[FIXED + info.len()..].copy_from_slice(&bits);
        payload
    }

    #[test]
    fn bundled_classic_emf_plays_object_selection_and_line() {
        let mut pen = Vec::new();
        dword(&mut pen, 1); // explicit handle
        dword(&mut pen, pen_style::SOLID);
        dword(&mut pen, 2);
        dword(&mut pen, 0);
        dword(&mut pen, 0x0000_00ff); // red COLORREF
        let mut select = Vec::new();
        dword(&mut select, 1);
        let mut move_to = Vec::new();
        sdword(&mut move_to, 10);
        sdword(&mut move_to, 20);
        let mut line_to = Vec::new();
        sdword(&mut line_to, 80);
        sdword(&mut line_to, 90);
        let bytes = classic_emf(vec![
            record(EmrType::CreatePen, &pen),
            record(EmrType::SelectObject, &select),
            record(EmrType::MoveToEx, &move_to),
            record(EmrType::LineTo, &line_to),
        ]);
        let parser = EmfParser::new(&bytes).unwrap();
        let svg = EmfSvgConverter::new(&parser).convert().unwrap();
        assert!(svg.starts_with("<svg "));
        assert!(svg.contains("stroke=\"red\""));
        assert!(svg.contains("stroke-width=\"2\""));
        assert!(svg.contains("x1=\"10\" y1=\"20\""));
    }

    #[test]
    fn stretch_dib_embeds_png_under_current_world_transform() {
        let mut transform = vec![0u8; 24];
        put_f32(&mut transform, 0, 0.0);
        put_f32(&mut transform, 4, 1.0);
        put_f32(&mut transform, 8, -1.0);
        put_f32(&mut transform, 12, 0.0);
        put_f32(&mut transform, 16, 7.0);
        put_f32(&mut transform, 20, 9.0);
        let bytes = classic_emf(vec![
            record(EmrType::SetWorldTransform, &transform),
            record(EmrType::StretchDIBits, &stretch_dib_payload()),
        ]);
        let parser = EmfParser::new(&bytes).unwrap();
        let svg = EmfSvgConverter::new(&parser).convert().unwrap();
        assert!(svg.contains("data:image/png;base64,"));
        assert!(svg.contains("transform=\"matrix(0 1 -1 0 7 9)\""));
        assert!(svg.contains("x=\"10\" y=\"20\" width=\"20\" height=\"30\""));
    }

    #[test]
    fn plg_blt_keeps_affine_destination_parallelogram() {
        let bytes = classic_emf(vec![record(EmrType::PlgBlt, &plg_blt_payload())]);
        let parser = EmfParser::new(&bytes).unwrap();
        let svg = EmfSvgConverter::new(&parser).convert().unwrap();
        assert!(svg.contains("data:image/png;base64,"));
        assert!(svg.contains("transform=\"matrix(20 0 0 30 10 20)\""));
        assert!(svg.contains("transform=\"matrix(1 0 0 1 0 0)\""));
    }

    #[test]
    fn bundled_wrench_tolerates_failed_delete_object() {
        let bytes = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/images/emf/wrench.emf"
        ));
        let parser = EmfParser::new(bytes).unwrap();
        let svg = EmfSvgConverter::new(&parser).convert().unwrap();
        assert!(svg.starts_with("<svg "));
        assert!(svg.len() > 100);
    }
    #[test]
    fn payload_relative_text_offset() {
        assert_eq!(payload_offset(8, EmrType::ExtTextOutW).unwrap(), 0);
        assert!(payload_offset(4, EmrType::ExtTextOutW).is_err());
    }
    #[test]
    fn ansi_is_not_decoded_as_utf8() {
        assert_eq!(
            decode_ansi(&[0x80, 0x93, 0x94], 0, EmrType::ExtTextOutA).unwrap(),
            "€“”"
        );
    }

    #[test]
    fn gradient_and_region_expansion_honor_caller_limits_before_emission() {
        let mut gradient = vec![0_u8; size_of::<EmrGradientFillHeader>()];
        gradient[16..20].copy_from_slice(&2_u32.to_le_bytes());
        gradient[20..24].copy_from_slice(&2_u32.to_le_bytes());
        gradient.extend_from_slice(&[0_u8; 2 * size_of::<TriVertex>()]);
        for _ in 0..2 {
            gradient.extend_from_slice(&0_u32.to_le_bytes());
            gradient.extend_from_slice(&1_u32.to_le_bytes());
        }
        let bytes = classic_emf(vec![record(EmrType::GradientFill, &gradient)]);
        let parser = EmfParser::new(&bytes).unwrap();
        let limits = crate::Limits {
            max_svg_elements: 1,
            ..crate::Limits::default()
        };
        assert!(
            EmfSvgConverter::with_limits(&parser, limits)
                .convert()
                .is_err()
        );

        let mut region = vec![0_u8; 32 + 2 * size_of::<RectL>()];
        region[0..4].copy_from_slice(&32_u32.to_le_bytes());
        region[8..12].copy_from_slice(&2_u32.to_le_bytes());
        region[12..16].copy_from_slice(&32_u32.to_le_bytes());
        assert!(
            region_path(
                &region,
                &DeviceContext::default(),
                EmrType::FillRgn,
                1,
                1024
            )
            .is_err()
        );
    }

    #[test]
    fn poly_text_shared_offsets_are_bounded_before_rendering() {
        const STRINGS: usize = 4;
        let header_size = size_of::<EmrPolyTextOutHeader>();
        let info_size = size_of::<EmrTextInfo>();
        let string_offset = header_size + STRINGS * info_size;
        let mut payload = vec![0_u8; string_offset + 2];
        put_u32(&mut payload, 28, STRINGS as u32);
        for index in 0..STRINGS {
            let info = header_size + index * info_size;
            put_u32(&mut payload, info + 8, 1);
            put_u32(&mut payload, info + 12, (8 + string_offset) as u32);
        }
        payload[string_offset..].copy_from_slice(&[b'A', 0]);

        let bytes = classic_emf(vec![record(EmrType::PolyTextOutW, &payload)]);
        let parser = EmfParser::new(&bytes).unwrap();
        let limits = crate::Limits {
            max_output_bytes: 1500,
            ..crate::Limits::default()
        };
        let error = EmfSvgConverter::with_limits(&parser, limits)
            .convert()
            .unwrap_err();
        assert!(
            matches!(error, Error::ParseError(message) if message.contains("PolyTextOut expansion"))
        );
    }

    #[test]
    fn transformed_rectangle_keeps_rotation() {
        let mut dc = DeviceContext::default();
        dc.world_transform = XForm {
            m11: 0.0,
            m12: 1.0,
            m21: -1.0,
            m22: 0.0,
            dx: 0.0,
            dy: 0.0,
        };
        let path = rectangle_path(
            RectL {
                left: 0,
                top: 0,
                right: 10,
                bottom: 20,
            },
            &dc,
        )
        .unwrap()
        .build();
        assert!(path.contains("L0 10"));
        // Consecutive LineTo commands share one `L` command in compact SVG
        // grammar: M0 0L0 10 -20 10 -20 0z.
        assert!(path.contains("-20 10"), "{path}");
    }
}
