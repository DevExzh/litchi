//! Strict classic EMF playback into SVG.

use super::{
    buffer::ElementBuffer,
    path::PathBuilder,
    state::{
        Brush, Canvas, ClipRule, DashPattern, DeviceContext, Font, GdiObject, Pen, RenderState,
    },
};
use crate::emf::{
    parser::{EmfParser, EmfRecord},
    record_parser::{parse_pointl_array, parse_points_array, parse_poly_counts},
    records::*,
};
use crate::svg_utils::{ClipGrouper, write_num, write_xml_escaped};
use crate::{
    dib::{ColorUsage, Dib, DibLimits},
    metafile_bitmap::{
        BitmapOp, BitmapPaint, BitmapSource, BlitGrid, BrushPattern, MAX_BRACKET_OUTLINE_VERTICES,
        MAX_BRACKET_SHAPES, RasterComposite, Rop2, Rop2Paint, Rop3, StretchPolicy,
        convex_outlines_overlap, is_convex, parse_emf_object_dib, xor_bracket_paint,
    },
    text_charset::DecodedText,
};
use bytes::Bytes;
use litchi_core::error::{Error, Result};
use std::{f64::consts::PI, fmt::Write, mem::size_of};
use zerocopy::FromBytes;

const MAX_RECORD_ITEMS: usize = 1_000_000;

pub struct EmfSvgConverter<'a> {
    parser: &'a EmfParser,
    limits: crate::Limits,
}

/// One completed classic EMF playback.
struct ClassicTrack {
    svg: String,
    diagnostics: Vec<crate::emfplus::renderer::RendererDiagnostic>,
    /// Body elements emitted; zero means the track draws an empty picture.
    elements: usize,
}

impl ClassicTrack {
    fn into_parts(self) -> (String, Vec<crate::emfplus::renderer::RendererDiagnostic>) {
        (self.svg, self.diagnostics)
    }
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
        if let Some(diagnostic) = diagnostics
            .iter()
            .find(|diagnostic| crate::codec::is_lossy_diagnostic(diagnostic.code))
        {
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
        let facts = crate::emfplus::track::scan_stream(
            self.parser
                .records
                .iter()
                .map(|record| (record.record_type, &record.data[..])),
            emfplus_limits(self.limits).parser,
        );
        let choice = crate::emfplus::track::choose_track(facts)
            .map_err(|message| Error::Unsupported(message.into()))?;
        match choice {
            crate::emfplus::track::TrackChoice::Classic { note } => Ok(self
                .convert_classic_track(note.map(|(code, message)| diagnostic(code, message)))?
                .into_parts()),
            crate::emfplus::track::TrackChoice::EmfPlus { classic_fallback } => {
                self.convert_emfplus_track(classic_fallback)
            },
        }
    }

    /// Play every classic EMF record, with any selection facts reported first.
    fn convert_classic_track(
        &self,
        selection: Option<crate::emfplus::renderer::RendererDiagnostic>,
    ) -> Result<ClassicTrack> {
        let header = &self.parser.header;
        let dpi_x = dpi(header.device_width, header.device_width_mm);
        let dpi_y = dpi(header.device_height, header.device_height_mm);
        let mut state = RenderState::with_device_metrics(dpi_x, dpi_y);
        if let Some(selection) = selection {
            state
                .diagnostics
                .insert(0, (selection.code, selection.message));
        }
        let mut buffer = ElementBuffer::new();
        let mut generated_bytes = 0usize;
        let mut generated_elements = 0usize;
        let mut counted_definitions = 0usize;
        let mut records = self.parser.records.iter();
        loop {
            let record_elements = match records.next() {
                Some(record) => self.play_record(record, &mut state)?,
                None if state.xor_bracket.is_some() => return Err(xor_bracket_failure()),
                // Deferred bitmap records paint at the end of playback.
                None if state.deferred_blits.is_some() => self.flush_blits(&mut state)?,
                None => break,
            };
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
        let svg = self.build_svg(&buffer.elements, &state, generated_elements)?;
        let diagnostics = std::mem::take(&mut state.diagnostics)
            .into_iter()
            .map(
                |(code, message)| crate::emfplus::renderer::RendererDiagnostic {
                    code,
                    message,
                    record_offset: None,
                },
            )
            .collect();
        Ok(ClassicTrack {
            svg,
            diagnostics,
            elements: buffer.elements.len(),
        })
    }

    /// Play the EMF+ records, falling back to the complete classic track of a
    /// dual metafile when EMF+ playback fails or would need an approximation.
    ///
    /// [MS-EMFPLUS] 1.3.1: a dual metafile's classic records alone define the
    /// whole picture, and EMF+ playback of a dual metafile is the same as that
    /// of an EMF+-only metafile. EMF+-only metafiles have no complete classic
    /// rendering, so their EMF+ failures stay failures.
    fn convert_emfplus_track(
        &self,
        classic_fallback: bool,
    ) -> Result<(String, Vec<crate::emfplus::renderer::RendererDiagnostic>)> {
        match self.render_emfplus_track() {
            Ok(output) => {
                let approximated = output
                    .diagnostics
                    .iter()
                    .find(|diagnostic| crate::codec::is_lossy_diagnostic(diagnostic.code))
                    .map(|diagnostic| diagnostic.code);
                let Some(approximated) = approximated.filter(|_| classic_fallback) else {
                    return Ok((output.document.into_string(), output.diagnostics));
                };
                let mut diagnostics = output.diagnostics;
                let message = format!(
                    "EMF+ playback needs the reported approximation {approximated}; the complete classic EMF track of the dual metafile was rendered instead"
                );
                let selection = diagnostic("noncanonical-emfplus-approximated", &message);
                let classic = match self.convert_classic_track(Some(selection)) {
                    Ok(classic) if classic.elements > 0 => classic,
                    Ok(_) => {
                        diagnostics.push(diagnostic(
                            "noncanonical-classic-track-empty",
                            "the classic EMF track of the dual metafile draws nothing, so the EMF+ picture is kept",
                        ));
                        return Ok((output.document.into_string(), diagnostics));
                    },
                    Err(error) => {
                        // The EMF+ picture is still the better answer than no
                        // picture, but the failed fallback must be reported.
                        diagnostics.push(diagnostic(
                            "noncanonical-classic-track-failed",
                            &error.to_string(),
                        ));
                        return Ok((output.document.into_string(), diagnostics));
                    },
                };
                Ok(classic.into_parts())
            },
            Err(error) if classic_fallback => {
                let message = format!(
                    "EMF+ playback failed ({error}); the complete classic EMF track of the dual metafile was rendered instead"
                );
                match self.convert_classic_track(Some(diagnostic(
                    "noncanonical-emfplus-unreadable",
                    &message,
                ))) {
                    Ok(classic) if classic.elements > 0 => Ok(classic.into_parts()),
                    Ok(_) => Err(error),
                    Err(classic_error) => Err(classic_error),
                }
            },
            Err(error) => Err(error),
        }
    }

    fn render_emfplus_track(&self) -> Result<crate::emfplus::renderer::RenderOutput> {
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
                // EMF+ device space is the EMF device space, so the EMF+
                // picture shares the bounds of the classic picture.
                renderer.insert(
                    crate::emfplus::renderer::EmfPlusSvgRenderer::new(
                        f64::from(header.width().unsigned_abs().max(1)),
                        f64::from(header.height().unsigned_abs().max(1)),
                        emfplus_limits(self.limits),
                    )?
                    .with_view_origin(f64::from(header.bounds.0), f64::from(header.bounds.1)),
                )
            };
            emfplus.push_comment_body(&record.data)?;
        }

        let Some(renderer) = renderer else {
            return Err(Error::ParseError(
                "EMF+ track requires at least one EMF+ comment".into(),
            ));
        };
        renderer.finish()
    }

    /// Device-space rectangle covered by the SVG canvas the converter emits.
    ///
    /// Clips are bounded by it instead of by an unbounded sentinel box.
    fn canvas(&self) -> Canvas {
        let bounds = self.parser.header.canvas();
        Canvas {
            left: f64::from(bounds.0),
            top: f64::from(bounds.1),
            right: f64::from(bounds.0)
                + f64::from(self.parser.header.width().unsigned_abs().max(1)),
            bottom: f64::from(bounds.1)
                + f64::from(self.parser.header.height().unsigned_abs().max(1)),
        }
    }

    /// Plays one record, painting deferred bitmap records before anything
    /// another record draws.
    fn play_record(&self, record: &EmfRecord, state: &mut RenderState) -> Result<Vec<String>> {
        state.bracket_shape = None;
        let mut elements = self.process_record(record, state)?;
        if let Some((shape, shape_elements)) = state.bracket_shape.take() {
            let outline = logical_outline(record);
            elements = self.capture_bracket_shape(state, shape, shape_elements, outline)?;
        } else if state.xor_bracket.is_some() && !elements.is_empty() {
            return Err(xor_bracket_failure());
        }
        let bitmap = matches!(
            EmrType::from_u32(record.record_type),
            Some(
                EmrType::BitBlt
                    | EmrType::StretchBlt
                    | EmrType::MaskBlt
                    | EmrType::PlgBlt
                    | EmrType::SetDIBitsToDevice
                    | EmrType::StretchDIBits
                    | EmrType::AlphaBlend
                    | EmrType::TransparentBlt
            )
        );
        if !bitmap && !elements.is_empty() && state.deferred_blits.is_some() {
            let mut output = self.flush_blits(state)?;
            output.append(&mut elements);
            elements = output;
        }
        Ok(elements)
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
            // The parser has already validated the terminator; playback ends at
            // its record type.
            EmrType::Eof => Ok(Vec::new()),

            // Mapping and graphics state.
            EmrType::SetWindowExtEx | EmrType::SetViewportExtEx => {
                let extent = read_record::<SizeL>(data, kind)?;
                if extent.cx == 0 || extent.cy == 0 {
                    // GDI refuses a zero extent and keeps the previous one.
                    state.note(
                        "noncanonical-emf-zero-mapping-extent",
                        format!("{} with a zero extent is ignored", kind.name()),
                    );
                    return Ok(Vec::new());
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
                // MWT_IDENTITY ignores the transform, so only the other modes
                // need a valid one.
                if modification.mode != 1 {
                    validate_xform(&modification.xform, kind)?;
                }
                state.dc.world_transform = match modification.mode {
                    1 => XForm::default(),
                    // With our column-vector representation, left multiplication
                    // is XForm * current and right multiplication the reverse.
                    2 => modification.xform.multiply(&state.dc.world_transform),
                    3 => state.dc.world_transform.multiply(&modification.xform),
                    // MWT_SET replaces the transform, as SetWorldTransform does.
                    4 => modification.xform,
                    mode => {
                        // GDI fails an unknown mode and keeps the transform.
                        state.note(
                            "noncanonical-emf-modify-world-transform-mode",
                            format!("{} with unknown mode {mode} is ignored", kind.name()),
                        );
                        return Ok(Vec::new());
                    },
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
                    // GDI fails a RestoreDC that names no saved state and
                    // leaves the device context unchanged.
                    state.note(
                        "noncanonical-emf-restore-dc-level",
                        format!(
                            "{} level {index} names no saved state and is ignored",
                            kind.name()
                        ),
                    );
                }
                Ok(Vec::new())
            },
            EmrType::SetTextColor | EmrType::SetBkColor => {
                let color = read_record::<ColorRef>(data, kind)?;
                if kind == EmrType::SetTextColor {
                    state.dc.text_color = color;
                    // A monochrome pattern definition captures its old text
                    // colour.
                    if state
                        .dc
                        .brush
                        .pattern
                        .as_ref()
                        .is_some_and(|pattern| pattern.as_ref().is_ok_and(|p| p.is_monochrome()))
                    {
                        state.dc.brush.pattern_id = None;
                    }
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
                    // GDI fails a mode that is neither TRANSPARENT nor OPAQUE
                    // and keeps the background mode.
                    state.note(
                        "noncanonical-emf-background-mode",
                        format!("{} with unknown mode {mode} is ignored", kind.name()),
                    );
                    return Ok(Vec::new());
                }
                state.dc.bg_mode = mode;
                state.dc.brush.pattern_id = None;
                Ok(Vec::new())
            },
            EmrType::SetPolyFillMode => {
                let mode = read_u32(data, 0, kind)?;
                if mode != 1 && mode != 2 {
                    // GDI fails a mode that is neither ALTERNATE nor WINDING
                    // and keeps the fill mode.
                    state.note(
                        "noncanonical-emf-polygon-fill-mode",
                        format!("{} with unknown mode {mode} is ignored", kind.name()),
                    );
                    return Ok(Vec::new());
                }
                state.dc.poly_fill_mode = mode;
                Ok(Vec::new())
            },
            EmrType::SetTextAlign => {
                state.dc.text_align = read_u32(data, 0, kind)?;
                Ok(Vec::new())
            },
            EmrType::SetRop2 => {
                // The mix mode only matters to records that paint with it,
                // so it is resolved per drawing.
                let mode = read_u32(data, 0, kind)?;
                if Rop2::from_mode(mode).is_some() {
                    state.dc.rop2 = mode;
                } else {
                    // SetROP2 fails for an undefined mode and keeps the
                    // current one.
                    state.note(
                        "noncanonical-emf-invalid-rop2",
                        format!("EMR_SETROP2 ignores undefined mix mode {mode}"),
                    );
                }
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
                        logpen_width(value.width),
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
                let mut pen = Pen::from_ext_create_pen(value.pen_style, value.width, value.color);
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
                        pen.dash_pattern = Some(DashPattern::Units(
                            entries
                                .iter()
                                .map(|value| f64::from(*value))
                                .collect::<Vec<_>>()
                                .into(),
                        ));
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
                    // SelectObject fails for a handle without a selectable
                    // object and leaves the selection unchanged.
                    state.note(
                        "noncanonical-emf-select-unknown-object",
                        format!("{} of handle {handle:#x} selects nothing", kind.name()),
                    );
                    return Ok(Vec::new());
                }
                Ok(Vec::new())
            },
            EmrType::DeleteObject => {
                let handle = read_u32(data, 0, kind)?;
                // DeleteObject of an unknown or stock handle fails without
                // altering the DC or subsequent drawing.
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
                    active_path(state)?.move_to(transformed.0, transformed.1);
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
                    let builder = active_path(state)?;
                    if builder.is_empty() {
                        builder.move_to(start.0, start.1);
                    }
                    builder.line_to(end.0, end.1);
                    Ok(Vec::new())
                } else {
                    self.paint_shape(kind, state, false, true, |state| {
                        Ok(vec![render_line(start, end, &state.dc)])
                    })
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
                    // GDI fails CloseFigure outside a path bracket.
                    state.note(
                        "noncanonical-emf-close-figure",
                        "CloseFigure outside a path bracket is ignored".to_owned(),
                    );
                    return Ok(Vec::new());
                }
                active_path(state)?.close();
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
                self.paint_shape(kind, state, fill, stroke, |state| {
                    Ok(path_element(builder, &state.dc, fill, stroke)
                        .into_iter()
                        .collect())
                })
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
                if !state.install_clip(&builder.build(), ClipRule::NonZero, mode) {
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
                if rect.left == rect.right || rect.top == rect.bottom {
                    // A shape with no area draws nothing.
                    state.note(
                        "noncanonical-emf-empty-shape",
                        format!(
                            "{} with an empty bounding rectangle draws nothing",
                            kind.name()
                        ),
                    );
                    return Ok(Vec::new());
                }
                let builder = if kind == EmrType::Rectangle {
                    rectangle_path(rect, &state.dc)
                } else {
                    ellipse_path(rect, &state.dc)
                }?;
                self.record_or_render(builder, kind, state, true, true)
            },
            EmrType::RoundRect => {
                let shape = read_record::<EmrRoundRect>(data, kind)?;
                let builder = round_rect_path(shape.rect, shape.corner, &state.dc)?;
                self.record_or_render(builder, kind, state, true, true)
            },
            EmrType::Arc | EmrType::ArcTo | EmrType::Chord | EmrType::Pie => {
                let arc = read_record::<EmrArc>(data, kind)?;
                let (builder, end) = arc_path(arc, kind, &state.dc)?;
                if matches!(kind, EmrType::ArcTo) {
                    state.dc.current_pos = end;
                }
                let fill = matches!(kind, EmrType::Chord | EmrType::Pie);
                self.record_or_render(builder, kind, state, fill, true)
            },
            EmrType::AngleArc => {
                let arc = read_record::<EmrAngleArc>(data, kind)?;
                let (builder, end) = angle_arc_path(arc, &state.dc)?;
                state.dc.current_pos = end;
                self.record_or_render(builder, kind, state, false, true)
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
                let (path, rule) = if kind == EmrType::IntersectClipRect {
                    // An empty rectangle leaves an empty clip region, which
                    // hides everything drawn through it.
                    (
                        rectangle_outline(rect, &state.dc).build(),
                        ClipRule::NonZero,
                    )
                } else {
                    (
                        exclusion_path(rect, &state.dc, self.canvas())?,
                        ClipRule::EvenOdd,
                    )
                };
                if !state.install_clip(&path, rule, 1) {
                    return unsupported(kind, "clip combine mode");
                }
                Ok(Vec::new())
            },
            EmrType::OffsetClipRgn => {
                let offset = read_record::<PointL>(data, kind)?;
                let transformed = state
                    .dc
                    .transform_vector(f64::from(offset.x), f64::from(offset.y));
                state.offset_clip(transformed.0, transformed.1, self.canvas());
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
                if !state.install_clip(&path, ClipRule::NonZero, mode) {
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
                // The extra space only matters for text drawn without explicit
                // character displacements; see `render_text_element`.
                let value = read_record::<EmrSetTextJustification>(data, kind)?;
                state.dc.text_break_extra = if value.num_break_count != 0 {
                    value.num_break_extra
                } else {
                    0
                };
                Ok(Vec::new())
            },

            EmrType::BitBlt
            | EmrType::StretchBlt
            | EmrType::MaskBlt
            | EmrType::PlgBlt
            | EmrType::SetDIBitsToDevice
            | EmrType::StretchDIBits
            | EmrType::AlphaBlend
            | EmrType::TransparentBlt => self.bitmap_record(&record.data, kind, state),
            EmrType::CreateMonoBrush | EmrType::CreateDIBPatternBrushPt => {
                // Both records share one header: object index, colour usage,
                // then the bitmap-info and bitmap-bits ranges.
                let value = read_record::<EmrCreateDIBPatternBrushPtHeader>(data, kind)?;
                let dib = ColorUsage::from_raw(value.usage).and_then(|usage| {
                    parse_emf_object_dib(
                        data,
                        [value.off_bmi, value.cb_bmi, value.off_bits, value.cb_bits],
                        usage,
                        dib_limits(self.limits),
                    )
                });
                // A monochrome brush takes its colours from the device
                // context; a DIB pattern brush from its own color table.
                let pattern = dib.and_then(|dib| {
                    note_dib_tolerances(state, &dib);
                    if kind == EmrType::CreateMonoBrush {
                        BrushPattern::from_monochrome_dib(&dib)
                    } else {
                        BrushPattern::from_dib(&dib)
                    }
                });
                // The object keeps its table slot; painting with it fails.
                let pattern = pattern
                    .map(std::sync::Arc::new)
                    .map_err(|error| std::sync::Arc::from(error.to_string().as_str()));
                insert_object(
                    state,
                    value.object_index,
                    GdiObject::Brush(Brush::from_pattern(brush_style::PATTERN, pattern)),
                    kind,
                )?;
                Ok(Vec::new())
            },

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
                // An EMF player ignores the EMF+ records embedded in comments;
                // playing them is the job of the EMF+ track.
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
        kind: EmrType,
        state: &mut RenderState,
        fill: bool,
        stroke: bool,
    ) -> Result<Vec<String>> {
        if state.in_path {
            append_path(active_path(state)?, builder);
            return Ok(Vec::new());
        }
        self.paint_shape(kind, state, fill, stroke, |state| {
            Ok(path_element(builder, &state.dc, fill, stroke)
                .into_iter()
                .collect())
        })
    }

    /// Renders a shape with the pen and brush as the active mix mode paints
    /// them. Nothing is emitted when the mode leaves every pixel the shape
    /// would paint unchanged.
    fn paint_shape(
        &self,
        kind: EmrType,
        state: &mut RenderState,
        fill: bool,
        stroke: bool,
        render: impl FnOnce(&mut RenderState) -> Result<Vec<String>>,
    ) -> Result<Vec<String>> {
        let rop2 = Rop2::from_mode(state.dc.rop2).unwrap_or(Rop2::COPY_PEN);
        if let Some(shape) = bracket_paint(state, rop2, fill, stroke) {
            // The bracket decides the fill once its second pass is complete.
            let saved = (state.dc.pen.clone(), state.dc.brush.clone());
            state.dc.pen = Pen::from_create_pen(pen_style::NULL, 0, ColorRef::from_rgb(0, 0, 0));
            state.dc.brush = xor_fill_brush();
            let result = render(state);
            (state.dc.pen, state.dc.brush) = saved;
            state.bracket_shape = Some((shape, result?));
            return Ok(Vec::new());
        }
        if rop2 == Rop2::COPY_PEN {
            if fill {
                state.prepare_brush_pattern(self.limits.max_output_bytes)?;
            }
            return render(state);
        }
        let dependent = || {
            Error::Unsupported(format!(
                "{}: ROP2 mode {} depends on destination pixels and has no faithful SVG equivalent",
                kind.name(),
                rop2.mode()
            ))
        };
        let pen = if stroke {
            Some(mixed_pen(rop2, &state.dc.pen).ok_or_else(dependent)?)
        } else {
            None
        };
        let brush = if fill {
            Some(mixed_brush(rop2, &state.dc).ok_or_else(dependent)?)
        } else {
            None
        };
        let paints_stroke = pen.as_ref().is_some_and(|pen| pen.style != pen_style::NULL);
        let paints_fill = brush
            .as_ref()
            .is_some_and(|(brush, _)| brush.style != brush_style::NULL);
        if !paints_stroke && !paints_fill {
            return Ok(Vec::new());
        }
        let saved = (
            state.dc.pen.clone(),
            state.dc.brush.clone(),
            state.dc.bg_mode,
            state.dc.bg_color,
        );
        if let Some(pen) = pen {
            state.dc.pen = pen;
        }
        if let Some((brush, background)) = brush {
            state.dc.brush = brush;
            match background {
                Some(color) => state.dc.bg_color = color,
                None => state.dc.bg_mode = 1,
            }
        }
        let result = if fill {
            state.prepare_brush_pattern(self.limits.max_output_bytes)
        } else {
            Ok(())
        }
        .and_then(|()| render(state));
        (
            state.dc.pen,
            state.dc.brush,
            state.dc.bg_mode,
            state.dc.bg_color,
        ) = saved;
        result
    }

    fn bitmap_record(
        &self,
        data: &Bytes,
        kind: EmrType,
        state: &mut RenderState,
    ) -> Result<Vec<String>> {
        let stretch = match state.dc.stretch_mode {
            1 => StretchPolicy::BlackOnWhite,
            2 => StretchPolicy::WhiteOnBlack,
            3 => StretchPolicy::ColorOnColor,
            4 => StretchPolicy::Halftone,
            _ => return malformed(kind, "invalid active stretch mode"),
        };
        let operation = BitmapOp::parse_emf(kind as u32, data, stretch, dib_limits(self.limits))?;
        if let Some(BitmapSource::Dib(dib)) = operation.source {
            note_dib_tolerances(state, &dib);
        }
        let paint = operation.paint()?;
        let matrix = logical_to_device_matrix(&state.dc);
        if !operation.rop3.depends_on_source() && paint != BitmapPaint::Nothing {
            let destination = operation.destination_rect()?;
            if let Some(output) = self.bracket_blit(state, operation.rop3, &matrix, destination)? {
                return Ok(output);
            }
        }
        if state.xor_bracket.is_some() {
            return Err(xor_bracket_failure());
        }
        match paint {
            BitmapPaint::Nothing => {
                state.note(
                    "noncanonical-emf-empty-bitmap-record",
                    format!("{} paints no pixels", kind.name()),
                );
                Ok(Vec::new())
            },
            BitmapPaint::Solid(color) => {
                // The ROP3 ignores source and destination and paints the whole
                // destination rectangle with one colour.
                let element = format!(
                    "<path d=\"{}\" fill=\"{}\" {}/>",
                    device_rect_path(&matrix, operation.destination_rect()?),
                    ColorRef { value: color }.to_svg_color(),
                    state.dc.clip_attr()
                );
                self.paint_after_blits(state, element)
            },
            BitmapPaint::Brush => {
                if state.dc.brush.style == brush_style::NULL {
                    return Ok(Vec::new());
                }
                // GDI aligns brush patterns to device pixels, so the
                // rectangle is emitted in device coordinates.
                state.prepare_brush_pattern(self.limits.max_output_bytes)?;
                let element = format!(
                    "<path d=\"{}\" {} {}/>",
                    device_rect_path(&matrix, operation.destination_rect()?),
                    state.dc.get_fill_attr(),
                    state.dc.clip_attr()
                );
                self.paint_after_blits(state, element)
            },
            BitmapPaint::Image | BitmapPaint::Composite => {
                let transform = matrix_attr(&matrix);
                let interpolation = interpolation(stretch);
                if !operation.is_plain_blit() {
                    // Alpha, colour-key and parallelogram blits paint their
                    // source directly.
                    let Some(image) = operation.to_svg_image()? else {
                        return Ok(Vec::new());
                    };
                    let element = bitmap_group(
                        &state.dc.clip_attr(),
                        &transform,
                        interpolation,
                        &image.element,
                    );
                    return self.paint_after_blits(state, element);
                }
                let destination = operation.destination_rect()?;
                let destination = [
                    f64::from(destination.x),
                    f64::from(destination.y),
                    f64::from(destination.width),
                    f64::from(destination.height),
                ];
                let (painted, pixels) = if operation.rop3.depends_on_source() {
                    let Some(window) = operation.blit_window()? else {
                        state.note(
                            "noncanonical-emf-blit-outside-bitmap",
                            format!("{} reads no pixel inside its bitmap", kind.name()),
                        );
                        return Ok(Vec::new());
                    };
                    (window.painted(destination), Some(window.pixels()))
                } else {
                    (destination, None)
                };
                let blit = DeferredBlit {
                    kind,
                    data: data.clone(),
                    stretch,
                    rop3: operation.rop3,
                    destination,
                    brush: (state.dc.brush.style == brush_style::SOLID
                        && state.dc.brush.pattern.is_none())
                    .then_some(state.dc.brush.color.value),
                };
                let context = BlitContext {
                    clip: state.dc.clip_attr(),
                    transform,
                    interpolation,
                };
                self.defer_blit(state, context, blit, painted, pixels)
            },
        }
    }

    /// Adds a blit without a source to an XOR bracket: a PATINVERT with a
    /// solid brush is an XOR fill, and another brush raster operation right
    /// after a first XOR pass is the middle fill. Returns the bracket's
    /// output, or `None` when the blit is not part of a bracket.
    fn bracket_blit(
        &self,
        state: &mut RenderState,
        rop3: Rop3,
        matrix: &[f64; 6],
        destination: crate::metafile_bitmap::SignedRect,
    ) -> Result<Option<Vec<String>>> {
        let shape = if rop3.table() == 0x5a {
            let brush = &state.dc.brush;
            if brush.style != brush_style::SOLID || brush.pattern.is_some() {
                return Ok(None);
            }
            let color = brush.color.value & 0x00ff_ffff;
            if color == 0 {
                // XOR with black keeps every pixel.
                return Ok(Some(Vec::new()));
            }
            BracketShape::Xor(color)
        } else {
            let Some(bracket) = &state.xor_bracket else {
                return Ok(None);
            };
            if bracket.middle.is_some() || bracket.matched != 0 {
                return Ok(None);
            }
            let rop2 = rop3_as_rop2(rop3);
            if rop2.ignores_color() && rop2.paint(0) == Rop2Paint::Unchanged {
                return Ok(Some(Vec::new()));
            }
            match bracket_middle(state, rop2) {
                Some(middle) => BracketShape::Middle(middle),
                None => return Ok(None),
            }
        };
        let element = format!(
            "<path d=\"{}\" fill=\"url(#{XOR_FILL})\" {}/>",
            device_rect_path(matrix, destination),
            state.dc.clip_attr()
        );
        let (x, y) = (i64::from(destination.x), i64::from(destination.y));
        let (right, bottom) = (
            x + i64::from(destination.width),
            y + i64::from(destination.height),
        );
        let outline = vec![(x, y), (right, y), (right, bottom), (x, bottom)];
        self.capture_bracket_shape(state, shape, vec![element], Some(outline))
            .map(Some)
    }

    /// Adds a shape to the XOR bracket, returning the bracket's output once
    /// its second pass is complete.
    fn capture_bracket_shape(
        &self,
        state: &mut RenderState,
        shape: BracketShape,
        elements: Vec<String>,
        outline: Option<Vec<(i64, i64)>>,
    ) -> Result<Vec<String>> {
        let matrix = logical_to_device_matrix(&state.dc);
        let element = elements.concat();
        match shape {
            BracketShape::Xor(color) => {
                let shape = XorShape {
                    element,
                    color,
                    outline,
                };
                let Some(bracket) = &mut state.xor_bracket else {
                    if shape.element.len() > self.limits.max_output_bytes {
                        return Err(xor_bracket_failure());
                    }
                    state.xor_bracket = Some(XorBracket {
                        matrix,
                        bytes: shape.element.len(),
                        first: vec![shape],
                        middle: None,
                        matched: 0,
                    });
                    return Ok(Vec::new());
                };
                if bracket.matrix != matrix {
                    state.xor_bracket = None;
                    return Err(xor_bracket_failure());
                }
                if bracket.middle.is_none() {
                    // The fills wait outside every output limit until the
                    // bracket completes, so the bracket bounds them itself.
                    bracket.bytes = bracket.bytes.saturating_add(shape.element.len());
                    if bracket.first.len() >= MAX_BRACKET_SHAPES
                        || bracket.bytes > self.limits.max_output_bytes
                    {
                        state.xor_bracket = None;
                        return Err(xor_bracket_failure());
                    }
                    bracket.first.push(shape);
                    return Ok(Vec::new());
                }
                let expected = &bracket.first[bracket.matched];
                if expected.element != shape.element || expected.color != shape.color {
                    state.xor_bracket = None;
                    return Err(xor_bracket_failure());
                }
                bracket.matched += 1;
                if bracket.matched < bracket.first.len() {
                    return Ok(Vec::new());
                }
                let Some(bracket) = state.xor_bracket.take() else {
                    return Ok(Vec::new());
                };
                let output = self.resolve_xor_bracket(state, bracket)?;
                let mut elements = self.flush_blits(state)?;
                elements.extend(output);
                Ok(elements)
            },
            BracketShape::Middle(middle) => {
                if let Some(bracket) = &mut state.xor_bracket {
                    bracket.bytes = bracket.bytes.saturating_add(element.len());
                    if bracket.bytes > self.limits.max_output_bytes {
                        state.xor_bracket = None;
                        return Err(xor_bracket_failure());
                    }
                    bracket.middle = Some((element, middle, outline));
                }
                Ok(Vec::new())
            },
        }
    }

    /// Paints a complete XOR bracket: the middle fill where no XOR fill
    /// covers it, then every XOR fill clipped to the middle fill.
    fn resolve_xor_bracket(
        &self,
        state: &mut RenderState,
        bracket: XorBracket,
    ) -> Result<Vec<String>> {
        let Some((middle_element, middle, middle_outline)) = bracket.middle else {
            return Err(xor_bracket_failure());
        };
        // Each pixel takes the colour of the one XOR fill covering it only
        // when no two fills overlap.
        if bracket.first.len() > 1 {
            let outlines = bracket
                .first
                .iter()
                .map(|shape| shape.outline.as_deref())
                .collect::<Option<Vec<_>>>()
                .ok_or_else(xor_bracket_failure)?;
            if convex_outlines_overlap(&outlines) {
                return Err(xor_bracket_failure());
            }
        }
        let token = format!("url(#{XOR_FILL})");
        let mut output = Vec::new();
        // The middle fill outside every XOR fill. When the XOR fills tile a
        // rectangle that holds the middle fill nothing is left of it;
        // otherwise a mask keeps it from painting under the XOR fills, which
        // paint those pixels themselves.
        if !xor_fills_cover(&bracket.first, middle_outline.as_deref())
            && let Some(paint) = self.bracket_region_paint(state, &middle, 0)?
        {
            let id = state.fresh_id("xor-mask");
            let mut hidden = String::new();
            for shape in &bracket.first {
                hidden.push_str(&shape.element.replace(&token, "black"));
            }
            // Nothing outside the canvas is visible, so the canvas bounds the
            // mask exactly.
            let canvas = self.canvas();
            let (x, y) = (fmt(canvas.left), fmt(canvas.top));
            let (width, height) = (
                fmt(canvas.right - canvas.left),
                fmt(canvas.bottom - canvas.top),
            );
            state.add_definition(format!(
                "<mask id=\"{id}\" maskUnits=\"userSpaceOnUse\" x=\"{x}\" y=\"{y}\" width=\"{width}\" height=\"{height}\"><rect x=\"{x}\" y=\"{y}\" width=\"{width}\" height=\"{height}\" fill=\"white\"/>{hidden}</mask>"
            ));
            output.push(format!(
                "<g mask=\"url(#{id})\">{}</g>",
                middle_element.replace(&token, &paint)
            ));
        }
        let mut strips = String::new();
        for shape in &bracket.first {
            if let Some(paint) = self.bracket_region_paint(state, &middle, shape.color)? {
                strips.push_str(&shape.element.replace(&token, &paint));
            }
        }
        if !strips.is_empty() {
            let id = state.fresh_id("xor-clip");
            let geometry = middle_element
                .replace(&token, "black")
                .replace(" fill-rule=", " clip-rule=");
            state.add_definition(format!(
                "<clipPath id=\"{id}\" clipPathUnits=\"userSpaceOnUse\">{geometry}</clipPath>"
            ));
            // The XOR fills tile without overlap; edges without antialiasing
            // keep the seams between them closed, as GDI draws them.
            output.push(format!(
                "<g clip-path=\"url(#{id})\"><g shape-rendering=\"crispEdges\">{strips}</g></g>"
            ));
        }
        Ok(output)
    }

    /// The SVG paint of the middle region whose XOR fills total `xor`, or
    /// `None` when the region keeps the destination.
    fn bracket_region_paint(
        &self,
        state: &mut RenderState,
        middle: &XorMiddle,
        xor: u32,
    ) -> Result<Option<String>> {
        match middle {
            XorMiddle::Solid { rop2, color } => Ok(xor_bracket_paint(*rop2, *color, xor)
                .ok_or_else(xor_bracket_failure)?
                .map(|value| ColorRef { value }.to_svg_color())),
            XorMiddle::Pattern {
                rop2,
                pattern,
                text_color,
                background_color,
                origin,
            } => {
                let png = pattern
                    .mapped(*text_color, *background_color, |color| {
                        xor_bracket_paint(*rop2, color, xor)
                    })
                    .ok_or_else(xor_bracket_failure)?
                    .tile_png(0, 0, self.limits.max_output_bytes)?;
                let size = (f64::from(pattern.width()), f64::from(pattern.height()));
                let id = state.pattern_definition(png, size, *origin)?;
                Ok(Some(format!("url(#{id})")))
            },
        }
    }

    /// Paints the deferred bitmap records, then `element`.
    fn paint_after_blits(&self, state: &mut RenderState, element: String) -> Result<Vec<String>> {
        let mut output = self.flush_blits(state)?;
        output.push(element);
        Ok(output)
    }

    /// Defers a bitmap record that may combine with the records after it,
    /// returning the output of the deferred records it cannot join.
    fn defer_blit(
        &self,
        state: &mut RenderState,
        context: BlitContext,
        blit: DeferredBlit,
        painted: [f64; 4],
        pixels: Option<(u32, u32)>,
    ) -> Result<Vec<String>> {
        let composite =
            blit.rop3.depends_on_destination() || !matches!(blit.rop3.table(), 0xcc | 0x33);
        let max_cells = usize::try_from(self.limits.max_pixels).unwrap_or(usize::MAX);
        if let Some(group) = &mut state.deferred_blits
            && group.context == context
            && group.grid.extend(painted, pixels, max_cells)
        {
            group.composite |= composite;
            group.blits.push(blit);
            return Ok(Vec::new());
        }
        let flushed = self.flush_blits(state)?;
        let grid = match pixels {
            Some((columns, rows)) => BlitGrid::for_pixels(painted, columns, rows),
            None => BlitGrid::for_area(painted),
        };
        if let Some(grid) = grid {
            state.deferred_blits = Some(DeferredBlits {
                context,
                grid,
                blits: vec![blit],
                composite,
            });
        }
        Ok(flushed)
    }

    /// Paints the deferred bitmap records.
    fn flush_blits(&self, state: &mut RenderState) -> Result<Vec<String>> {
        let Some(group) = state.deferred_blits.take() else {
            return Ok(Vec::new());
        };
        let limits = dib_limits(self.limits);
        fn parse(blit: &DeferredBlit, limits: DibLimits) -> Result<BitmapOp<'_>> {
            BitmapOp::parse_emf(blit.kind as u32, &blit.data, blit.stretch, limits)
        }
        let context = &group.context;
        if !group.composite {
            // Plain source copies paint exactly as they would one by one.
            let mut output = Vec::new();
            for blit in &group.blits {
                if let Some(image) = parse(blit, limits)?.to_svg_image()? {
                    output.push(bitmap_group(
                        &context.clip,
                        &context.transform,
                        context.interpolation,
                        &image.element,
                    ));
                }
            }
            return Ok(output);
        }
        let mut composite = RasterComposite::new(group.grid)?;
        let mut brush_approximated = false;
        for blit in &group.blits {
            let pattern = blit.brush.unwrap_or_else(|| {
                brush_approximated |= blit.rop3.depends_on_pattern();
                0
            });
            if blit.rop3.depends_on_source() {
                let Some(pixels) = parse(blit, limits)?.blit_pixels()? else {
                    continue;
                };
                composite.apply(
                    blit.rop3,
                    pixels.window.painted(blit.destination),
                    Some(&pixels.image),
                    pattern,
                )?;
            } else {
                composite.apply(blit.rop3, blit.destination, None, pattern)?;
            }
        }
        let image = composite.to_svg_image(self.limits.max_output_bytes)?;
        if image.approximated || brush_approximated {
            state.note(
                "emf-raster-operation-approximated",
                "raster operations that combine with destination pixels were painted as on a white destination"
                    .to_owned(),
            );
        }
        Ok(vec![bitmap_group(
            &context.clip,
            &context.transform,
            context.interpolation,
            &image.image.element,
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
            parse_points_array(&data[20..], count, &self.limits)?
                .into_iter()
                .map(|p| (i32::from(p.x), i32::from(p.y)))
                .collect::<Vec<_>>()
        } else {
            parse_pointl_array(&data[20..], count, &self.limits)?
                .into_iter()
                .map(|p| (p.x, p.y))
                .collect::<Vec<_>>()
        };
        if points.is_empty() {
            // A poly record without points draws nothing and moves nothing.
            state.note(
                "noncanonical-emf-empty-poly-record",
                format!("{} without points is ignored", kind.name()),
            );
            return Ok(Vec::new());
        }
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
            for curve in points[index..].as_chunks::<3>().0.iter() {
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
        self.record_or_render(builder, kind, state, polygon, true)
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
        let counts = parse_poly_counts(&data[24..], polygons, &self.limits)?;
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
                &self.limits,
            )?
            .into_iter()
            .map(|p| (i32::from(p.x), i32::from(p.y)))
            .collect::<Vec<_>>()
        } else {
            parse_pointl_array(
                data.get(point_offset..)
                    .ok_or_else(|| Error::ParseError("point offset outside payload".into()))?,
                total,
                &self.limits,
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
        self.record_or_render(builder, kind, state, filled, true)
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
            parse_points_array(&data[20..], count, &self.limits)?
                .into_iter()
                .map(|p| (i32::from(p.x), i32::from(p.y)))
                .collect::<Vec<_>>()
        } else {
            parse_pointl_array(&data[20..], count, &self.limits)?
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
        self.record_or_render(builder, kind, state, false, true)
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
                // GDI fails the record for an invalid brush handle and keeps
                // painting the rest of the picture.
                state.note(
                    "noncanonical-emf-region-brush-missing",
                    format!("{} references missing brush {brush_handle}", kind.name()),
                );
                return Ok(Vec::new());
            }
        }
        let frame_width = if kind == EmrType::FrameRgn {
            Some(read_i32(data, 24, kind)?.unsigned_abs().max(1))
        } else {
            None
        };
        let result = self.paint_shape(kind, state, true, false, |state| {
            let attrs = match frame_width {
                // FrameRgn strokes the region outline with the brush.
                Some(width) => format!(
                    "fill=\"none\" stroke=\"{}\" stroke-width=\"{width}\"",
                    state.dc.get_fill_attr()["fill=\"".len()..].trim_end_matches('"')
                ),
                None => state.dc.get_fill_attr(),
            };
            Ok(vec![format!(
                "<path d=\"{path}\" {attrs} {} fill-rule=\"evenodd\"/>",
                state.dc.clip_attr()
            )])
        });
        state.dc.brush = saved_brush;
        state.dc.brush_handle = saved_brush_handle;
        result
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
        let decoded = decode_text(data, offset, count, unicode, &state.dc.font, kind)?;
        note_inexact_text(state, &decoded);
        let reference = PointL {
            x: header.x,
            y: header.y,
        };
        render_text_element(
            decoded,
            reference,
            header.fu_options,
            None,
            None,
            self.canvas(),
            state,
            kind,
        )
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
            // The string holds glyph numbers of the recording system's font,
            // which no character mapping can recover; the text is omitted.
            state.note(
                "emf-glyph-index-text-omitted",
                format!(
                    "{} text given as font glyph indices is not drawn",
                    kind.name()
                ),
            );
            return Ok(Vec::new());
        }
        let count = bounded_count(info.num_chars, kind)?;
        if count > self.limits.max_path_points || count > self.limits.max_output_bytes {
            return malformed(kind, "text character count exceeds configured limit");
        }
        let string_offset = payload_offset(info.off_string, kind)?;
        let decoded = decode_text(data, string_offset, count, unicode, &state.dc.font, kind)?;
        note_inexact_text(state, &decoded);
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
            decoded,
            info.reference,
            info.options,
            Some(info.rectangle),
            dx.as_deref(),
            self.canvas(),
            state,
            kind,
        )
    }

    fn build_svg(
        &self,
        elements: &[String],
        state: &RenderState,
        generated_elements: usize,
    ) -> Result<String> {
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
            width,
            height,
            header.canvas().0,
            header.canvas().1,
            width,
            height
        )
        .map_err(|_| Error::ParseError("failed to write EMF SVG header".into()))?;
        if !state.defs.is_empty() {
            svg.push_str("<defs>");
            for definition in &state.defs {
                svg.push_str(definition);
            }
            svg.push_str("</defs>");
        }
        let mut groups = ClipGrouper::new();
        for element in elements {
            groups.push(&mut svg, element);
        }
        groups.finish(&mut svg);
        if generated_elements.saturating_add(groups.groups()) > self.limits.max_svg_elements {
            return Err(Error::ParseError(format!(
                "EMF SVG element count exceeds limit {}",
                self.limits.max_svg_elements
            )));
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

/// A classic-playback diagnostic with no EMF+ record offset.
fn diagnostic(code: &'static str, message: &str) -> crate::emfplus::renderer::RendererDiagnostic {
    crate::emfplus::renderer::RendererDiagnostic {
        code,
        message: message.to_owned(),
        record_offset: None,
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

/// Stands in for the fill of a shape whose paint is known only once its XOR
/// bracket is complete.
const XOR_FILL: &str = "@xor-fill@";

/// A fill painted in an XOR mix mode, waiting for the pass that undoes it.
#[derive(Debug, Clone)]
pub(super) struct XorShape {
    /// The elements, with [`XOR_FILL`] as their fill.
    element: String,
    /// The colour the fill XORs into the destination.
    color: u32,
    /// Convex outline in logical coordinates, when the record has one.
    outline: Option<Vec<(i64, i64)>>,
}

/// How the fill between the two XOR passes paints.
#[derive(Debug, Clone)]
pub(super) enum XorMiddle {
    Solid {
        rop2: Rop2,
        color: u32,
    },
    Pattern {
        rop2: Rop2,
        pattern: std::sync::Arc<BrushPattern>,
        text_color: u32,
        background_color: u32,
        origin: (i32, i32),
    },
}

/// Fills XORed into the destination, the one fill drawn after them, and the
/// pass that XORs the same fills again.
///
/// The second pass cancels the first wherever the middle fill did not
/// paint. Inside the middle fill every pixel becomes `m(D ^ a) ^ a`, where
/// `a` is the colour of the XOR fill covering it and `m` the paint of the
/// middle fill. Producers use this to clip gradients and patterns to a
/// shape.
#[derive(Debug, Clone)]
pub(super) struct XorBracket {
    matrix: [f64; 6],
    /// Bytes of the waiting fill and middle elements.
    bytes: usize,
    first: Vec<XorShape>,
    /// The middle elements, with [`XOR_FILL`] as their fill, their paint and
    /// their logical outline, when the record has one.
    middle: Option<(String, XorMiddle, Option<Vec<(i64, i64)>>)>,
    /// Shapes of the second pass matched so far.
    matched: usize,
}

/// What the current record's shape contributes to an XOR bracket.
#[derive(Debug, Clone)]
pub(super) enum BracketShape {
    Xor(u32),
    Middle(XorMiddle),
}

/// Whether the XOR fills tile the rectangle that bounds them and that
/// rectangle holds the middle fill's outline.
fn xor_fills_cover(fills: &[XorShape], middle: Option<&[(i64, i64)]>) -> bool {
    let bounds = |outline: &[(i64, i64)]| {
        outline.iter().fold(
            [i64::MAX, i64::MAX, i64::MIN, i64::MIN],
            |[left, top, right, bottom], &(x, y)| {
                [left.min(x), top.min(y), right.max(x), bottom.max(y)]
            },
        )
    };
    // Exact for `i32` coordinates: each product stays below 2^62 and at
    // most [`MAX_BRACKET_OUTLINE_VERTICES`] terms are added.
    let doubled_area = |outline: &[(i64, i64)]| {
        let mut sum = 0i128;
        for index in 0..outline.len() {
            let (x0, y0) = outline[index];
            let (x1, y1) = outline[(index + 1) % outline.len()];
            sum += i128::from(x0) * i128::from(y1) - i128::from(x1) * i128::from(y0);
        }
        sum.abs()
    };
    let (Some(middle), Some(outlines)) = (
        middle,
        fills
            .iter()
            .map(|fill| fill.outline.as_deref())
            .collect::<Option<Vec<_>>>(),
    ) else {
        return false;
    };
    let Some(cover) = outlines
        .iter()
        .map(|outline| bounds(outline))
        .reduce(|a, b| {
            [
                a[0].min(b[0]),
                a[1].min(b[1]),
                a[2].max(b[2]),
                a[3].max(b[3]),
            ]
        })
    else {
        return false;
    };
    // Disjoint fills tile their bounds exactly when their areas add up to
    // its area; both sides are doubled to stay in integers.
    let area = outlines
        .iter()
        .map(|outline| doubled_area(outline))
        .sum::<i128>();
    let cover_area = 2 * i128::from(cover[2] - cover[0]) * i128::from(cover[3] - cover[1]);
    let inner = bounds(middle);
    area == cover_area
        && inner[0] >= cover[0]
        && inner[1] >= cover[1]
        && inner[2] <= cover[2]
        && inner[3] <= cover[3]
}

fn xor_bracket_failure() -> Error {
    Error::Unsupported(
        "ROP2 mode 7 depends on destination pixels and has no faithful SVG equivalent".into(),
    )
}

/// A solid brush whose fill attribute is the [`XOR_FILL`] placeholder.
fn xor_fill_brush() -> Brush {
    let mut brush = Brush::from_create_brush(brush_style::SOLID, ColorRef::from_rgb(0, 0, 0), 0);
    brush.pattern_id = Some(XOR_FILL.to_owned());
    brush
}

/// Identifies a fill that belongs to an XOR bracket: an XOR fill, or the
/// one fill after a first XOR pass.
fn bracket_paint(
    state: &RenderState,
    rop2: Rop2,
    fill: bool,
    stroke: bool,
) -> Option<BracketShape> {
    if !fill || state.in_path {
        return None;
    }
    if stroke && mixed_pen(rop2, &state.dc.pen).is_none_or(|pen| pen.style != pen_style::NULL) {
        return None;
    }
    if matches!(rop2.mode(), 7 | 10) {
        let brush = &state.dc.brush;
        if brush.style != brush_style::SOLID || brush.pattern.is_some() {
            return None;
        }
        // R2_NOTXORPEN XORs the inverse colour.
        let color = if rop2.mode() == 7 {
            brush.color.value
        } else {
            !brush.color.value
        } & 0x00ff_ffff;
        return (color != 0).then_some(BracketShape::Xor(color));
    }
    let bracket = state.xor_bracket.as_ref()?;
    if bracket.middle.is_some() || bracket.matched != 0 {
        return None;
    }
    if rop2.ignores_color() && rop2.paint(0) == Rop2Paint::Unchanged {
        return None;
    }
    bracket_middle(state, rop2).map(BracketShape::Middle)
}

/// The middle-fill paint of the selected brush under `rop2`.
fn bracket_middle(state: &RenderState, rop2: Rop2) -> Option<XorMiddle> {
    let brush = &state.dc.brush;
    match &brush.pattern {
        None if brush.style == brush_style::SOLID => Some(XorMiddle::Solid {
            rop2,
            color: brush.color.value & 0x00ff_ffff,
        }),
        Some(Ok(pattern)) => Some(XorMiddle::Pattern {
            rop2,
            pattern: std::sync::Arc::clone(pattern),
            text_color: state.dc.text_color.value & 0x00ff_ffff,
            background_color: state.dc.bg_color.value & 0x00ff_ffff,
            origin: state.dc.brush_org,
        }),
        _ => None,
    }
}

/// The mix mode that a ROP3 without a source applies to its brush.
fn rop3_as_rop2(rop3: Rop3) -> Rop2 {
    // The ROP2 truth table is indexed by (P, D); the ROP3 one by (P, S, D).
    let table = rop3.table();
    let mode = (table & 0b11) | ((table >> 2) & 0b1100);
    Rop2::from_mode(u32::from(mode) + 1).unwrap_or(Rop2::COPY_PEN)
}

/// The convex outline of a rectangle or polygon record in logical
/// coordinates, or `None` for other records, non-convex polygons and
/// polygons with more than [`MAX_BRACKET_OUTLINE_VERTICES`] vertices.
fn logical_outline(record: &EmfRecord) -> Option<Vec<(i64, i64)>> {
    let data: &[u8] = &record.data;
    let int = |offset: usize| -> Option<i64> {
        Some(i64::from(i32::from_le_bytes(
            data.get(offset..offset + 4)?.try_into().ok()?,
        )))
    };
    let short = |offset: usize| -> Option<i64> {
        Some(i64::from(i16::from_le_bytes(
            data.get(offset..offset + 2)?.try_into().ok()?,
        )))
    };
    let outline = match EmrType::from_u32(record.record_type)? {
        EmrType::Rectangle => {
            let (left, top, right, bottom) = (int(0)?, int(4)?, int(8)?, int(12)?);
            vec![(left, top), (right, top), (right, bottom), (left, bottom)]
        },
        EmrType::Polygon => {
            let count = usize::try_from(int(16)?).ok()?;
            if count > MAX_BRACKET_OUTLINE_VERTICES {
                return None;
            }
            (0..count)
                .map(|index| Some((int(20 + index * 8)?, int(24 + index * 8)?)))
                .collect::<Option<Vec<_>>>()?
        },
        EmrType::Polygon16 => {
            let count = usize::try_from(int(16)?).ok()?;
            if count > MAX_BRACKET_OUTLINE_VERTICES {
                return None;
            }
            (0..count)
                .map(|index| Some((short(20 + index * 4)?, short(22 + index * 4)?)))
                .collect::<Option<Vec<_>>>()?
        },
        _ => return None,
    };
    is_convex(&outline).then_some(outline)
}

/// A deferred bitmap record of classic EMF playback.
#[derive(Debug, Clone)]
pub(super) struct DeferredBlit {
    kind: EmrType,
    data: Bytes,
    stretch: StretchPolicy,
    rop3: Rop3,
    /// Normalized logical destination rectangle `[x, y, width, height]`.
    destination: [f64; 4],
    /// Solid brush `COLORREF`, or `None` for other brushes.
    brush: Option<u32>,
}

/// SVG context that every deferred record of one group shares.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct BlitContext {
    clip: String,
    transform: String,
    interpolation: &'static str,
}

/// Consecutive deferred bitmap records that share one context and pixel
/// grid, so that raster operations combining with each other's pixels are
/// composed into one image.
#[derive(Debug, Clone)]
pub(super) struct DeferredBlits {
    context: BlitContext,
    grid: BlitGrid,
    blits: Vec<DeferredBlit>,
    /// At least one record combines with destination pixels.
    composite: bool,
}

/// Groups a logical-coordinate bitmap element under the DC transform and the
/// device-space clip.
///
/// The clip sits on an outer group: a clip path on the transformed group
/// would be transformed with it.
fn bitmap_group(clip: &str, transform: &str, interpolation: &str, element: &str) -> String {
    let inner = format!("<g {transform} image-rendering=\"{interpolation}\">{element}</g>");
    if clip.is_empty() {
        inner
    } else {
        format!("<g {clip}>{inner}</g>")
    }
}

fn matrix_attr(matrix: &[f64; 6]) -> String {
    format!(
        "transform=\"matrix({} {} {} {} {} {})\"",
        fmt(matrix[0]),
        fmt(matrix[1]),
        fmt(matrix[2]),
        fmt(matrix[3]),
        fmt(matrix[4]),
        fmt(matrix[5])
    )
}

fn interpolation(stretch: StretchPolicy) -> &'static str {
    match stretch {
        StretchPolicy::Halftone => "optimizeQuality",
        StretchPolicy::BlackOnWhite | StretchPolicy::WhiteOnBlack | StretchPolicy::ColorOnColor => {
            "optimizeSpeed"
        },
    }
}

/// Path data of a logical rectangle mapped to device coordinates.
fn device_rect_path(matrix: &[f64; 6], rect: crate::metafile_bitmap::SignedRect) -> String {
    let x0 = f64::from(rect.x);
    let y0 = f64::from(rect.y);
    let x1 = x0 + f64::from(rect.width);
    let y1 = y0 + f64::from(rect.height);
    let map = |x: f64, y: f64| {
        (
            matrix[0] * x + matrix[2] * y + matrix[4],
            matrix[1] * x + matrix[3] * y + matrix[5],
        )
    };
    let mut path = String::with_capacity(96);
    for (index, (x, y)) in [map(x0, y0), map(x1, y0), map(x1, y1), map(x0, y1)]
        .into_iter()
        .enumerate()
    {
        path.push(if index == 0 { 'M' } else { 'L' });
        write_num(&mut path, x);
        path.push(' ');
        write_num(&mut path, y);
    }
    path.push('Z');
    path
}

/// Reports DIB header fields that GDI ignores.
fn note_dib_tolerances(state: &mut RenderState, dib: &Dib<'_>) {
    for (code, message) in dib.info().tolerances.diagnostics() {
        state.note(code, message.to_owned());
    }
}

/// The pen that paints like `pen` under `rop2`, or `None` when the result
/// depends on destination pixels.
fn mixed_pen(rop2: Rop2, pen: &Pen) -> Option<Pen> {
    if pen.style == pen_style::NULL {
        return Some(pen.clone());
    }
    match rop2.paint(pen.color.value) {
        Rop2Paint::Unchanged => Some(Pen {
            style: pen_style::NULL,
            ..pen.clone()
        }),
        Rop2Paint::Color(value) => Some(Pen {
            color: ColorRef { value },
            ..pen.clone()
        }),
        Rop2Paint::DestinationDependent => None,
    }
}

/// The brush and hatch background (`None` when transparent) that paint like
/// the selected brush under `rop2`, or `None` when the result depends on
/// destination pixels.
fn mixed_brush(rop2: Rop2, dc: &DeviceContext) -> Option<(Brush, Option<ColorRef>)> {
    let brush = &dc.brush;
    let background = (dc.bg_mode == 2).then_some(dc.bg_color);
    if brush.style == brush_style::NULL {
        return Some((brush.clone(), background));
    }
    let null = || Brush::from_create_brush(brush_style::NULL, ColorRef::from_rgb(0, 0, 0), 0);
    let solid = |value| Brush::from_create_brush(brush_style::SOLID, ColorRef { value }, 0);
    if let Some(Ok(pattern)) = &brush.pattern {
        // Every pixel of the pattern goes through the mix mode on its own.
        let pattern = pattern.under_rop2(
            rop2,
            dc.text_color.value & 0x00ff_ffff,
            dc.bg_color.value & 0x00ff_ffff,
        )?;
        return Some((
            Brush {
                pattern: Some(Ok(std::sync::Arc::new(pattern))),
                pattern_id: None,
                ..brush.clone()
            },
            None,
        ));
    }
    if brush.pattern.is_some() || !matches!(brush.style, brush_style::SOLID | brush_style::HATCHED)
    {
        // A brush without its pattern bitmap fails when it paints.
        return Some((brush.clone(), background));
    }
    let mix = |color: ColorRef| match rop2.paint(color.value) {
        Rop2Paint::Unchanged => Some(None),
        Rop2Paint::Color(value) => Some(Some(ColorRef { value })),
        Rop2Paint::DestinationDependent => None,
    };
    let foreground = mix(brush.color)?;
    if brush.style == brush_style::SOLID {
        return Some((
            foreground.map_or_else(null, |color| solid(color.value)),
            None,
        ));
    }
    let background = match background {
        Some(color) => mix(color)?,
        None => None,
    };
    match (foreground, background) {
        (Some(color), background) => Some((
            Brush {
                color,
                pattern_id: None,
                ..brush.clone()
            },
            background,
        )),
        (None, None) => Some((null(), None)),
        // A hatch whose lines keep the destination needs a pattern with
        // transparent lines, which the hatch definitions do not draw.
        (None, Some(_)) => None,
    }
}

fn append_path(target: &mut PathBuilder, source: PathBuilder) {
    // Replaying the built SVG grammar would be fragile. PathBuilder exposes
    // typed append specifically for this operation.
    target.append(source);
}

fn active_path(state: &mut RenderState) -> Result<&mut PathBuilder> {
    state.path_builder.as_mut().ok_or_else(|| {
        Error::ParseError("EMF path state is inconsistent with an open path bracket".into())
    })
}

fn rectangle_path(rect: RectL, dc: &DeviceContext) -> Result<PathBuilder> {
    if rect.left == rect.right || rect.top == rect.bottom {
        return Err(Error::ParseError("empty EMF rectangle".into()));
    }
    Ok(rectangle_outline(rect, dc))
}

/// The closed outline of a rectangle, which has no area when two opposite
/// sides coincide.
fn rectangle_outline(rect: RectL, dc: &DeviceContext) -> PathBuilder {
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
    builder
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

/// Width of an [MS-EMF] 2.2.19 LogPen, whose `Width` is the signed `x` of a
/// PointL; a negative width draws like a zero-width (cosmetic) pen.
fn logpen_width(raw: u32) -> u32 {
    u32::try_from(i32::from_le_bytes(raw.to_le_bytes())).unwrap_or(0)
}

fn angle_arc_path(arc: EmrAngleArc, dc: &DeviceContext) -> Result<(PathBuilder, (f64, f64))> {
    if arc.radius == 0 || !arc.start_angle.is_finite() || !arc.sweep_angle.is_finite() {
        return Err(Error::ParseError("invalid AngleArc".into()));
    }
    let cx = f64::from(arc.center.x);
    let cy = f64::from(arc.center.y);
    let radius = f64::from(arc.radius);
    let start = f64::from(arc.start_angle).to_radians();
    let full_sweep = f64::from(arc.sweep_angle).to_radians();
    // A sweep beyond one turn retraces the circle. One full turn plus the
    // remainder draws the same outline and ends at the same point.
    let sweep = if full_sweep.abs() > 2.0 * PI {
        full_sweep.signum() * (2.0 * PI + full_sweep.abs() % (2.0 * PI))
    } else {
        full_sweep
    };
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

/// Wraps an element that carries its own `transform` in a group that applies
/// the device-space `clip` attribute.
///
/// A `clip-path` on the transformed element itself would be read in that
/// element's transformed user space, not in device space.
fn device_clipped(clip: &str, element: &str) -> String {
    let mut output = String::with_capacity(clip.len() + element.len() + 8);
    output.push_str("<g ");
    output.push_str(clip);
    output.push('>');
    output.push_str(element);
    output.push_str("</g>");
    output
}

/// The canvas with `rect` cut out, to be filled with the even-odd rule so that
/// the result does not depend on the orientation the transforms give `rect`.
fn exclusion_path(rect: RectL, dc: &DeviceContext, canvas: Canvas) -> Result<String> {
    let mut outer = PathBuilder::new();
    // "Everything outside `rect`" only has to cover the canvas.
    outer.move_to(canvas.left, canvas.top);
    outer.line_to(canvas.right, canvas.top);
    outer.line_to(canvas.right, canvas.bottom);
    outer.line_to(canvas.left, canvas.bottom);
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
    mut decoded: DecodedText,
    reference: PointL,
    options: u32,
    rectangle: Option<RectL>,
    dx: Option<&[i32]>,
    canvas: Canvas,
    state: &mut RenderState,
    kind: EmrType,
) -> Result<Vec<String>> {
    if state.dc.text_break_extra != 0 && dx.is_none() && decoded.text.contains(' ') {
        state.note(
            "emf-text-justification-ignored",
            format!(
                "{} text is drawn without the {} units of justification added to its break characters",
                kind.name(),
                state.dc.text_break_extra
            ),
        );
    }
    let update = state.dc.text_align & text_align::UPDATECP != 0;
    let logical = if update {
        state.dc.current_pos
    } else {
        (f64::from(reference.x), f64::from(reference.y))
    };
    let origin = state.dc.transform_point(logical.0, logical.1);
    let mut output = Vec::new();
    if options & text_options::OPAQUE != 0 || (state.dc.bg_mode == 2 && rectangle.is_some()) {
        // A rectangle without area paints no background.
        if let Some(rect) =
            rectangle.filter(|rect| rect.left != rect.right && rect.top != rect.bottom)
        {
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
    let clip = if options & text_options::CLIPPED != 0 {
        let rect = rectangle.ok_or_else(|| {
            Error::ParseError(format!("{} ETO_CLIPPED without rectangle", kind.name()))
        })?;
        let id = state.fresh_id("text-clip");
        state.add_definition(format!(
            "<clipPath id=\"{}\"><path d=\"{}\"/></clipPath>",
            id,
            rectangle_path(rect, &state.dc)?.build()
        ));
        format!("clip-path=\"url(#{})\"", id)
    } else {
        state.dc.clip_attr()
    };
    let rotated = state.dc.font.escapement != 0.0;
    let mut attrs = String::new();
    if !rotated && !clip.is_empty() {
        write!(&mut attrs, " {}", clip).ok();
    }
    if rotated {
        write!(
            &mut attrs,
            " transform=\"rotate({} {} {})\"",
            -state.dc.font.escapement / 10.0,
            fmt(origin.0),
            fmt(origin.1)
        )
        .ok();
    }
    // A dx array holds the distance between the origins of adjacent
    // character cells ([MS-EMF] 2.2.5 EmrText), one value (or an x and a y
    // value with ETO_PDY) per code unit of the string. The characters are
    // placed at those absolute origins.
    let advances: Vec<(f64, f64)> = match dx {
        Some(values) if options & text_options::PDY != 0 => values
            .as_chunks::<2>()
            .0
            .iter()
            .map(|pair| (f64::from(pair[0]), f64::from(pair[1])))
            .collect(),
        Some(values) => values.iter().map(|&x| (f64::from(x), 0.0)).collect(),
        None => Vec::new(),
    };
    let mut origins = if advances.is_empty() {
        vec![origin]
    } else {
        decoded
            .origin_offsets(&advances)
            .into_iter()
            .map(|(x, y)| state.dc.transform_point(logical.0 + x, logical.1 + y))
            .collect()
    };
    let extent = advances
        .iter()
        .fold((0.0, 0.0), |sum, (x, y)| (sum.0 + x, sum.1 + y));
    // GDI aligns the whole string on the reference point, while SVG applies
    // text-anchor to every absolutely positioned character; with recorded
    // advances the string extent is known, so the origins move instead.
    let mut anchor = anchor;
    if origins.len() > 1 && anchor != "start" {
        let (width, height) = state.dc.transform_vector(extent.0, extent.1);
        let shift = if anchor == "middle" { 0.5 } else { 1.0 };
        for (x, y) in &mut origins {
            *x -= width * shift;
            *y -= height * shift;
        }
        anchor = "start";
    }
    let font_size = state.dc.font_device_size();
    let origins = decoded.retain_visible(&origins, |position| {
        crate::text_charset::origin_reaches_canvas(
            position,
            origin,
            (canvas.left, canvas.top, canvas.right, canvas.bottom),
            4.0 * font_size,
        )
    });
    let text = decoded.text.as_str();
    let (x_values, y_values): (Vec<f64>, Vec<f64>) = origins.into_iter().unzip();
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
    // GDI draws and advances over every space; XML would collapse runs of them
    // and shift the characters after them onto earlier positions.
    let element = format!(
        "<text x=\"{}\" y=\"{}\" fill=\"{}\" {} text-anchor=\"{}\" dominant-baseline=\"{}\" direction=\"{}\" xml:space=\"preserve\"{}>{}</text>",
        x,
        y,
        state.dc.text_color.to_svg_color(),
        state.dc.font.svg_attrs(font_size, !decoded.unicode_face),
        anchor,
        baseline,
        direction,
        attrs,
        escaped_text
    );
    if !text.is_empty() {
        output.push(if rotated && !clip.is_empty() {
            device_clipped(&clip, &element)
        } else {
            element
        });
    }
    if update {
        let advanced = if advances.is_empty() {
            (
                state.dc.font.height.abs() * 0.6 * text.chars().count() as f64,
                0.0,
            )
        } else {
            extent
        };
        state.dc.current_pos = (logical.0 + advanced.0, logical.1 + advanced.1);
    }
    Ok(output)
}

fn decode_text(
    data: &[u8],
    offset: usize,
    count: usize,
    unicode: bool,
    font: &Font,
    kind: EmrType,
) -> Result<DecodedText> {
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
            .as_chunks::<2>()
            .0
            .iter()
            .map(|v| u16::from_le_bytes([v[0], v[1]]))
            .collect::<Vec<_>>();
        Ok(crate::text_charset::decode_utf16_logfont_text(
            &units,
            &font.face_name,
        ))
    } else {
        Ok(crate::text_charset::decode_logfont_text(
            raw,
            font.charset,
            &font.face_name,
        ))
    }
}

/// Reports text whose character set could not be resolved exactly.
fn note_inexact_text(state: &mut RenderState, decoded: &DecodedText) {
    if !decoded.exact {
        let charset = state.dc.font.charset;
        state.note(
            "emf-text-charset-inexact",
            format!(
                "text in font charset {charset} could not be decoded exactly; Windows-1252 or replacement characters were used"
            ),
        );
    }
}

fn parse_font(data: &[u8], kind: EmrType) -> Result<(u32, Font)> {
    require_len(data, 4 + size_of::<LogFontW>() + 64, kind)?;
    let handle = read_u32(data, 0, kind)?;
    let log = read_record::<LogFontW>(&data[4..], kind)?;
    let face_offset = 4 + size_of::<LogFontW>();
    let units = data[face_offset..face_offset + 64]
        .as_chunks::<2>()
        .0
        .iter()
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
            pitch_and_family: log.pitch_and_family,
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
    if stock_objects::is_stock_object(handle) {
        // A stock object cannot be recreated; GDI refuses the record.
        state.note(
            "noncanonical-emf-stock-object-handle",
            format!(
                "{} names stock object handle {handle:#x}; the record is ignored",
                kind.name()
            ),
        );
        return Ok(());
    }
    if state.insert_object(handle, object) {
        // GDI overwrites the handle-table slot.
        state.note(
            "noncanonical-emf-object-handle-reused",
            format!("{} reuses live object handle {handle}", kind.name()),
        );
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
    let bytes: [u8; 4] = bytes
        .try_into()
        .map_err(|_| Error::ParseError(format!("{} has an invalid u32 field", kind.name())))?;
    Ok(u32::from_le_bytes(bytes))
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

    fn rect_payload(left: i32, top: i32, right: i32, bottom: i32) -> Vec<u8> {
        let mut out = Vec::new();
        for value in [left, top, right, bottom] {
            sdword(&mut out, value);
        }
        out
    }

    fn point(x: i32, y: i32) -> Vec<u8> {
        let mut out = Vec::new();
        sdword(&mut out, x);
        sdword(&mut out, y);
        out
    }

    fn clip_rect(left: i32, top: i32, right: i32, bottom: i32) -> Vec<u8> {
        let mut out = Vec::new();
        for value in [left, top, right, bottom] {
            sdword(&mut out, value);
        }
        record(EmrType::IntersectClipRect, &out)
    }

    fn clear_clip() -> Vec<u8> {
        let mut out = Vec::new();
        dword(&mut out, 0); // no region data
        dword(&mut out, 5); // RGN_COPY
        record(EmrType::ExtSelectClipRgn, &out)
    }

    fn line_to(x: i32, y: i32) -> Vec<u8> {
        record(EmrType::LineTo, &point(x, y))
    }

    fn offset_clip(x: i32, y: i32) -> Vec<u8> {
        record(EmrType::OffsetClipRgn, &point(x, y))
    }

    /// A solid black brush selected into every DC of the record list that
    /// follows, and a polygon covering far more than the canvas.
    fn black_filled_canvas() -> Vec<Vec<u8>> {
        let mut brush = Vec::new();
        dword(&mut brush, 1); // explicit handle
        dword(&mut brush, 0); // BS_SOLID
        dword(&mut brush, 0); // black COLORREF
        dword(&mut brush, 0);
        let mut select = Vec::new();
        dword(&mut select, 1);
        let mut polygon = Vec::new();
        for value in [-50, -50, 150, 150] {
            sdword(&mut polygon, value);
        }
        dword(&mut polygon, 4);
        for (x, y) in [(-50, -50), (150, -50), (150, 150), (-50, 150)] {
            polygon.extend_from_slice(&point(x, y));
        }
        vec![
            record(EmrType::CreateBrushIndirect, &brush),
            record(EmrType::SelectObject, &select),
            record(EmrType::Polygon, &polygon),
        ]
    }

    /// Converts a classic EMF and counts the non-white pixels its own SVG
    /// produces at canvas size. The synthetic canvas is 100 by 100 units, so
    /// one unit is one pixel.
    fn painted_pixels(records: Vec<Vec<u8>>) -> usize {
        let bytes = classic_emf(records);
        let parser = EmfParser::new(&bytes, &crate::Limits::default()).unwrap();
        let svg = EmfSvgConverter::new(&parser).convert().unwrap();
        let png = crate::raster::rasterize_svg(
            &svg,
            100,
            100,
            image::Rgba([255, 255, 255, 255]),
            image::ImageFormat::Png,
            &crate::raster::RasterLimits::default(),
        )
        .expect("converted EMF SVG must rasterize");
        image::load_from_memory_with_format(&png, image::ImageFormat::Png)
            .expect("rasterized PNG must decode")
            .to_rgba8()
            .pixels()
            .filter(|pixel| pixel.0 != [255; 4])
            .count()
    }

    /// The stroked elements of a converted classic EMF (every `<line` and
    /// every `<path` with a stroke, so not the clip definitions), in document
    /// order.
    fn line_elements(records: Vec<Vec<u8>>) -> Vec<String> {
        let bytes = classic_emf(records);
        let parser = EmfParser::new(&bytes, &crate::Limits::default()).unwrap();
        let svg = EmfSvgConverter::new(&parser).convert().unwrap();
        let mut found = Vec::new();
        let mut rest = svg.as_str();
        while let Some(start) = rest.find('<') {
            let tail = &rest[start..];
            let end = tail.find('>').unwrap() + 1;
            let element = &tail[..end];
            if element.starts_with("<line ")
                || (element.starts_with("<path ") && element.contains(" stroke=\""))
            {
                found.push(element.to_string());
            }
            rest = &tail[end..];
        }
        found
    }

    #[test]
    fn unclipped_line_runs_keep_their_serialized_form() {
        let elements = line_elements(vec![
            record(EmrType::MoveToEx, &point(10, 20)),
            line_to(30, 20),
            line_to(30, 40),
            line_to(10, 40),
            record(EmrType::MoveToEx, &point(50, 50)),
            line_to(60, 60),
            record(EmrType::MoveToEx, &point(70, 70)),
            line_to(80, 70),
            line_to(80, 80),
        ]);
        assert_eq!(
            elements,
            [
                r#"<path d="M10 20L30 20L30 40L10 40" fill="none" stroke="black" stroke-linecap="round" stroke-linejoin="round" stroke-miterlimit="10"/>"#,
                r#"<line x1="50" y1="50" x2="60" y2="60" stroke="black" stroke-linecap="round" stroke-linejoin="round" stroke-miterlimit="10"/>"#,
                r#"<line x1="70" y1="70" x2="80" y2="70" stroke="black" stroke-linecap="round" stroke-linejoin="round" stroke-miterlimit="10"/>"#,
                r#"<line x1="80" y1="70" x2="80" y2="80" stroke="black" stroke-linecap="round" stroke-linejoin="round" stroke-miterlimit="10"/>"#,
            ]
        );
    }

    #[test]
    fn unclipped_line_run_longer_than_the_buffer_threshold_keeps_its_split() {
        let mut records = vec![record(EmrType::MoveToEx, &point(0, 0))];
        for step in 1..=12 {
            records.push(line_to(step * 5, (step % 2) * 5));
        }
        let elements = line_elements(records);
        assert_eq!(
            elements,
            [
                r#"<path d="M0 0L5 5L10 0L15 5L20 0L25 5L30 0L35 5L40 0L45 5L50 0" fill="none" stroke="black" stroke-linecap="round" stroke-linejoin="round" stroke-miterlimit="10"/>"#,
                r#"<line x1="50" y1="0" x2="55" y2="5" stroke="black" stroke-linecap="round" stroke-linejoin="round" stroke-miterlimit="10"/>"#,
                r#"<line x1="55" y1="5" x2="60" y2="0" stroke="black" stroke-linecap="round" stroke-linejoin="round" stroke-miterlimit="10"/>"#,
            ]
        );
    }

    #[test]
    fn clipped_line_keeps_its_clip() {
        let elements = line_elements(vec![
            clip_rect(0, 0, 50, 50),
            record(EmrType::MoveToEx, &point(10, 20)),
            line_to(80, 90),
        ]);
        assert_eq!(
            elements,
            [
                r#"<line x1="10" y1="20" x2="80" y2="90" stroke="black" stroke-linecap="round" stroke-linejoin="round" stroke-miterlimit="10" clip-path="url(#emf-clip-0)"/>"#
            ]
        );
    }

    #[test]
    fn line_runs_do_not_merge_across_clip_regions() {
        let elements = line_elements(vec![
            clip_rect(0, 0, 50, 50),
            record(EmrType::MoveToEx, &point(10, 20)),
            line_to(30, 20),
            line_to(30, 40),
            line_to(10, 40),
            clear_clip(),
            line_to(10, 60),
            line_to(30, 60),
            line_to(30, 80),
        ]);
        assert_eq!(
            elements,
            [
                r#"<path d="M10 20L30 20L30 40L10 40" fill="none" stroke="black" stroke-linecap="round" stroke-linejoin="round" stroke-miterlimit="10" clip-path="url(#emf-clip-0)"/>"#,
                r#"<path d="M10 40L10 60L30 60L30 80" fill="none" stroke="black" stroke-linecap="round" stroke-linejoin="round" stroke-miterlimit="10"/>"#,
            ]
        );
    }

    #[test]
    fn excluded_clip_rect_paints_everything_but_its_rect() {
        let mut payload = Vec::new();
        for value in [20, 20, 80, 80] {
            sdword(&mut payload, value);
        }
        let mut records = vec![record(EmrType::ExcludeClipRect, &payload)];
        records.extend(black_filled_canvas());

        // 100 by 100 canvas minus the excluded 60 by 60 rectangle.
        assert_eq!(painted_pixels(records), 10_000 - 3_600);
    }

    #[test]
    fn offset_clip_moves_the_installed_region() {
        let mut records = vec![clip_rect(0, 0, 50, 50), offset_clip(80, 80)];
        records.extend(black_filled_canvas());

        // The 50 by 50 region keeps only its 20 by 20 corner on the canvas.
        assert_eq!(painted_pixels(records), 400);
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
        let parser = EmfParser::new(&bytes, &crate::Limits::default()).unwrap();
        let svg = EmfSvgConverter::new(&parser).convert().unwrap();
        assert!(svg.starts_with("<svg "));
        assert!(svg.contains("stroke=\"red\""));
        assert!(svg.contains("stroke-width=\"2\""));
        assert!(svg.contains("x1=\"10\" y1=\"20\""));
    }

    fn set_rop2(mode: u32) -> Vec<u8> {
        let mut out = Vec::new();
        dword(&mut out, mode);
        record(EmrType::SetRop2, &out)
    }

    /// A BitBlt record with no embedded bitmap, as producers write for fills.
    fn brush_bitblt(rop3: u32, left: i32, top: i32, width: i32, height: i32) -> Vec<u8> {
        let mut payload = vec![0u8; 92];
        for (offset, value) in [(16, left), (20, top), (24, width), (28, height)] {
            put_i32(&mut payload, offset, value);
        }
        put_u32(&mut payload, 32, rop3);
        put_f32(&mut payload, 44, 1.0);
        put_f32(&mut payload, 56, 1.0);
        record(EmrType::BitBlt, &payload)
    }

    fn create_solid_brush(handle: u32, color: u32) -> Vec<u8> {
        let mut brush = Vec::new();
        dword(&mut brush, handle);
        dword(&mut brush, brush_style::SOLID);
        dword(&mut brush, color);
        dword(&mut brush, 0);
        record(EmrType::CreateBrushIndirect, &brush)
    }

    fn select_object(handle: u32) -> Vec<u8> {
        let mut out = Vec::new();
        dword(&mut out, handle);
        record(EmrType::SelectObject, &out)
    }

    fn convert_with_report(
        records: Vec<Vec<u8>>,
    ) -> (String, Vec<crate::emfplus::renderer::RendererDiagnostic>) {
        let bytes = classic_emf(records);
        let parser = EmfParser::new(&bytes, &crate::Limits::default()).unwrap();
        EmfSvgConverter::new(&parser)
            .convert_with_diagnostics()
            .unwrap()
    }

    /// A StretchDIBits record copying a 1x`colors.len()` 24-bpp DIB to the
    /// logical rectangle (10, 20, 20, 30) with `rop3`.
    fn stretch_dib_record(colors: &[u32], rop3: u32) -> Vec<u8> {
        const FIXED: usize = 72;
        let mut info = vec![0u8; 40];
        put_u32(&mut info, 0, 40);
        put_i32(&mut info, 4, i32::try_from(colors.len()).unwrap());
        put_i32(&mut info, 8, 1);
        info[12..14].copy_from_slice(&1u16.to_le_bytes());
        info[14..16].copy_from_slice(&24u16.to_le_bytes());
        let mut bits = Vec::new();
        for color in colors {
            let [red, green, blue, _] = color.to_le_bytes();
            bits.extend_from_slice(&[blue, green, red]);
        }
        bits.resize(bits.len().div_ceil(4) * 4, 0);
        let mut payload = vec![0u8; FIXED];
        put_i32(&mut payload, 16, 10);
        put_i32(&mut payload, 20, 20);
        put_i32(&mut payload, 32, i32::try_from(colors.len()).unwrap());
        put_i32(&mut payload, 36, 1);
        put_u32(&mut payload, 40, (8 + FIXED) as u32);
        put_u32(&mut payload, 44, info.len() as u32);
        put_u32(&mut payload, 48, (8 + FIXED + info.len()) as u32);
        put_u32(&mut payload, 52, bits.len() as u32);
        put_u32(&mut payload, 60, rop3);
        put_i32(&mut payload, 64, 20);
        put_i32(&mut payload, 68, 30);
        payload.extend_from_slice(&info);
        payload.extend_from_slice(&bits);
        record(EmrType::StretchDIBits, &payload)
    }

    fn rectangle(left: i32, top: i32, right: i32, bottom: i32) -> Vec<u8> {
        record(EmrType::Rectangle, &rect_payload(left, top, right, bottom))
    }

    #[test]
    fn set_rop2_only_fails_when_a_record_paints_with_it() {
        // Setting and resetting a mix mode does not matter while nothing is
        // painted under it.
        let (svg, _) = convert_with_report(vec![set_rop2(6), set_rop2(13), rectangle(1, 2, 3, 4)]);
        assert!(svg.contains("<path "));

        // R2_NOT inverts the destination under any pen.
        let bytes = classic_emf(vec![set_rop2(6), rectangle(1, 2, 3, 4)]);
        let parser = EmfParser::new(&bytes, &crate::Limits::default()).unwrap();
        assert!(matches!(
            EmfSvgConverter::new(&parser).convert(),
            Err(Error::Unsupported(_))
        ));
    }

    #[test]
    fn rop2_nop_and_xor_with_black_paint_nothing() {
        for mode in [11, 7] {
            let (svg, diagnostics) = convert_with_report(vec![
                set_rop2(mode),
                create_solid_brush(1, 0),
                select_object(1),
                record(EmrType::MoveToEx, &point(10, 20)),
                line_to(30, 40),
                rectangle(1, 2, 3, 4),
            ]);
            assert!(!svg.contains("<line ") && !svg.contains("<path "), "{svg}");
            assert!(diagnostics.is_empty());
        }
    }

    #[test]
    fn rop2_constant_modes_paint_without_destination_pixels() {
        let (svg, diagnostics) = convert_with_report(vec![set_rop2(16), rectangle(10, 20, 30, 40)]);
        assert!(svg.contains("fill=\"white\" stroke=\"white\""));
        assert!(diagnostics.is_empty());
        // R2_NOTCOPYPEN inverts the default white brush and black pen.
        let (svg, _) = convert_with_report(vec![set_rop2(4), rectangle(10, 20, 30, 40)]);
        assert!(svg.contains("fill=\"black\" stroke=\"white\""));
    }

    #[test]
    fn undefined_rop2_mode_keeps_the_current_mode() {
        let (svg, diagnostics) = convert_with_report(vec![set_rop2(0), rectangle(10, 20, 30, 40)]);
        assert!(svg.contains("fill=\"white\" stroke=\"black\""));
        assert!(
            diagnostics
                .iter()
                .any(|diagnostic| diagnostic.code == "noncanonical-emf-invalid-rop2")
        );
    }

    #[test]
    fn patcopy_blit_fills_the_device_rectangle_with_the_brush() {
        let (svg, diagnostics) = convert_with_report(vec![
            create_solid_brush(1, 0x0000_00ff),
            select_object(1),
            brush_bitblt(0x00f0_0021, 10, 20, 30, 40),
        ]);
        assert!(svg.contains("<path d=\"M10 20L40 20L40 60L10 60Z\" fill=\"red\""));
        assert!(diagnostics.is_empty());
    }

    #[test]
    fn zero_size_blit_rectangles_draw_nothing() {
        let (svg, diagnostics) =
            convert_with_report(vec![brush_bitblt(0x00f0_0021, 10, 20, 0, 40)]);
        assert!(!svg.contains("<path "));
        assert!(
            diagnostics
                .iter()
                .any(|diagnostic| diagnostic.code == "noncanonical-emf-empty-bitmap-record")
        );
    }

    #[test]
    fn mask_and_image_blits_compose_one_exact_image() {
        // SRCAND with a black and white mask, then SRCPAINT with an image
        // that is black where the mask is white.
        let (svg, diagnostics) = convert_with_report(vec![
            stretch_dib_record(&[0, 0x00ff_ffff], 0x0088_00c6),
            stretch_dib_record(&[0x0000_00ff, 0], 0x00ee_0086),
            rectangle(0, 0, 5, 5),
        ]);
        assert_eq!(svg.matches("<image").count(), 1);
        assert!(svg.find("<image").unwrap() < svg.find("<path").unwrap());
        assert!(diagnostics.is_empty(), "{diagnostics:?}");

        // A lone SRCPAINT of a coloured image depends on the destination.
        let (svg, diagnostics) =
            convert_with_report(vec![stretch_dib_record(&[0x0000_00ff], 0x00ee_0086)]);
        assert!(svg.contains("<image"));
        assert!(
            diagnostics
                .iter()
                .any(|diagnostic| diagnostic.code == "emf-raster-operation-approximated")
        );
    }

    #[test]
    fn clipped_bitmaps_keep_the_clip_outside_the_transform() {
        let (svg, _) = convert_with_report(vec![
            record(EmrType::IntersectClipRect, &rect_payload(0, 0, 15, 25)),
            stretch_dib_record(&[0x0000_00ff], 0x00cc_0020),
        ]);
        assert!(svg.contains("<g clip-path=\"url(#"));
        assert!(svg.contains("><g transform=\"matrix(1 0 0 1 0 0)\""));
    }

    #[test]
    fn dib_pattern_brush_tiles_its_dib_in_device_pixels() {
        let (info, bits) = one_pixel_dib_parts();
        let mut payload = Vec::new();
        for value in [1u32, 0, 32, u32::try_from(info.len()).unwrap()] {
            dword(&mut payload, value);
        }
        dword(&mut payload, u32::try_from(32 + info.len()).unwrap());
        dword(&mut payload, u32::try_from(bits.len()).unwrap());
        payload.extend_from_slice(&info);
        payload.extend_from_slice(&bits);
        let (svg, diagnostics) = convert_with_report(vec![
            record(EmrType::CreateDIBPatternBrushPt, &payload),
            select_object(1),
            rectangle(10, 20, 30, 40),
        ]);
        assert!(svg.contains("<pattern id=\"emf-brush-"));
        assert!(svg.contains("fill=\"url(#emf-brush-"));
        assert!(diagnostics.is_empty(), "{diagnostics:?}");
    }

    #[test]
    fn patinvert_bracket_paints_the_middle_fill_in_the_brush_colour() {
        let patinvert = || brush_bitblt(0x005a_0049, 0, 0, 40, 40);
        let (svg, diagnostics) = convert_with_report(vec![
            create_solid_brush(1, 0x0000_00ff),
            create_solid_brush(2, 0),
            select_object(stock_objects::NULL_PEN),
            select_object(1),
            patinvert(),
            // The middle fill: a black rectangle, painted with R2_COPYPEN.
            select_object(2),
            rectangle(10, 10, 30, 30),
            select_object(1),
            patinvert(),
        ]);
        assert!(diagnostics.is_empty(), "{diagnostics:?}");
        // The XOR rectangle holds the middle rectangle, so only the XOR
        // rectangle paints, in red, clipped to the middle rectangle.
        assert!(!svg.contains("mask="));
        assert!(svg.contains("<g clip-path=\"url(#emf-xor-clip-"));
        assert!(svg.contains("<path d=\"M0 0L40 0L40 40L0 40Z\" fill=\"red\""));
        assert!(!svg.contains(XOR_FILL));

        // Without the second pass the XOR fill depends on the destination.
        let bytes = classic_emf(vec![
            create_solid_brush(1, 0x0000_00ff),
            select_object(1),
            patinvert(),
        ]);
        let parser = EmfParser::new(&bytes, &crate::Limits::default()).unwrap();
        assert!(matches!(
            EmfSvgConverter::new(&parser).convert(),
            Err(Error::Unsupported(_))
        ));
    }

    /// A Polygon record through `points`, bounded by their extremes.
    fn polygon_record(points: &[(i32, i32)]) -> Vec<u8> {
        let mut payload = Vec::new();
        let xs = points.iter().map(|point| point.0);
        let ys = points.iter().map(|point| point.1);
        for value in [
            xs.clone().min().unwrap(),
            ys.clone().min().unwrap(),
            xs.max().unwrap(),
            ys.max().unwrap(),
        ] {
            sdword(&mut payload, value);
        }
        dword(&mut payload, u32::try_from(points.len()).unwrap());
        for (x, y) in points {
            payload.extend_from_slice(&point(*x, *y));
        }
        record(EmrType::Polygon, &payload)
    }

    /// Two XOR strips with `vertices` vertices each, a black middle fill
    /// and the second XOR pass.
    fn polygon_xor_bracket(vertices: i32) -> Vec<Vec<u8>> {
        // A rectangle strip whose top edge holds collinear vertices.
        let strip = |left: i32| {
            let mut points = (0..vertices - 2)
                .map(|index| (left + index * 20 / (vertices - 3), 0))
                .collect::<Vec<_>>();
            points.extend([(left + 20, 50), (left, 50)]);
            polygon_record(&points)
        };
        let pass = || {
            vec![
                set_rop2(7),
                select_object(1),
                strip(0),
                select_object(2),
                strip(20),
            ]
        };
        let mut records = vec![
            create_solid_brush(1, 0x0000_00ff),
            create_solid_brush(2, 0x0000_ff00),
            create_solid_brush(3, 0),
            select_object(stock_objects::NULL_PEN),
        ];
        records.extend(pass());
        records.extend([
            set_rop2(13),
            select_object(3),
            polygon_record(&[(0, 0), (40, 0), (0, 40)]),
        ]);
        records.extend(pass());
        records
    }

    #[test]
    fn xor_bracket_outlines_with_many_vertices_are_not_overlap_tested() {
        let (svg, diagnostics) = convert_with_report(polygon_xor_bracket(
            i32::try_from(MAX_BRACKET_OUTLINE_VERTICES).unwrap(),
        ));
        assert!(diagnostics.is_empty(), "{diagnostics:?}");
        assert!(svg.contains("<g clip-path=\"url(#emf-xor-clip-"), "{svg}");

        let bytes = classic_emf(polygon_xor_bracket(
            i32::try_from(MAX_BRACKET_OUTLINE_VERTICES + 1).unwrap(),
        ));
        let parser = EmfParser::new(&bytes, &crate::Limits::default()).unwrap();
        assert!(matches!(
            EmfSvgConverter::new(&parser).convert(),
            Err(Error::Unsupported(_))
        ));
    }

    #[test]
    fn xor_fills_cover_stays_exact_at_i32_extremes() {
        let (low, high) = (i64::from(i32::MIN), i64::from(i32::MAX));
        let fill = |left: i64, right: i64| XorShape {
            element: String::new(),
            color: 0x0000_00ff,
            outline: Some(vec![(left, low), (right, low), (right, high), (left, high)]),
        };
        let middle = [(-5, -5), (5, -5), (0, 5)];
        assert!(xor_fills_cover(
            &[fill(low, 0), fill(0, high)],
            Some(&middle)
        ));
        assert!(!xor_fills_cover(
            &[fill(low, -1), fill(0, high)],
            Some(&middle)
        ));
    }

    #[test]
    fn xor_bracket_fails_once_its_waiting_fills_pass_a_limit() {
        let fills = |count: usize, limits: crate::Limits| {
            let patinvert = brush_bitblt(0x005a_0049, 0, 0, 40, 40);
            let mut records = vec![create_solid_brush(1, 0x0000_00ff), select_object(1)];
            records.extend(std::iter::repeat_n(patinvert, count));
            let bytes = classic_emf(records);
            let parser = EmfParser::new(&bytes, &limits).unwrap();
            let converter = EmfSvgConverter::with_limits(&parser, limits);
            let mut state = RenderState::new();
            // Every fill waits in the bracket until one fails it.
            let played = parser
                .records
                .iter()
                .take_while(|record| converter.play_record(record, &mut state).is_ok())
                .count();
            (played, state.xor_bracket.map(|bracket| bracket.first.len()))
        };
        let (played, waiting) = fills(MAX_BRACKET_SHAPES + 1, crate::Limits::default());
        assert_eq!(played, MAX_BRACKET_SHAPES + 2);
        assert_eq!(waiting, None);

        let limits = crate::Limits {
            max_output_bytes: 1000,
            ..crate::Limits::default()
        };
        let (played, waiting) = fills(100, limits);
        assert!(played < 20, "{played}");
        assert_eq!(waiting, None);
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
        let parser = EmfParser::new(&bytes, &crate::Limits::default()).unwrap();
        let svg = EmfSvgConverter::new(&parser).convert().unwrap();
        assert!(svg.contains("data:image/png;base64,"));
        assert!(svg.contains("transform=\"matrix(0 1 -1 0 7 9)\""));
        assert!(svg.contains("x=\"10\" y=\"20\" width=\"20\" height=\"30\""));
    }

    #[test]
    fn clipped_bitmap_reads_its_clip_in_device_space() {
        let bytes = classic_emf(vec![
            clip_rect(0, 0, 50, 50),
            record(EmrType::StretchDIBits, &stretch_dib_payload()),
        ]);
        let parser = EmfParser::new(&bytes, &crate::Limits::default()).unwrap();
        let svg = EmfSvgConverter::new(&parser).convert().unwrap();

        // The clip sits outside the group that carries the bitmap transform.
        assert!(svg.contains("<g clip-path=\"url(#emf-clip-0)\"><g transform=\"matrix("));
    }

    #[test]
    fn plg_blt_keeps_affine_destination_parallelogram() {
        let bytes = classic_emf(vec![record(EmrType::PlgBlt, &plg_blt_payload())]);
        let parser = EmfParser::new(&bytes, &crate::Limits::default()).unwrap();
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
        let parser = EmfParser::new(bytes, &crate::Limits::default()).unwrap();
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
        let font = Font {
            charset: 0,
            face_name: "Arial".to_owned(),
            ..Font::default()
        };
        let decoded = decode_text(
            &[0x80, 0x93, 0x94],
            0,
            3,
            false,
            &font,
            EmrType::ExtTextOutA,
        )
        .expect("single-byte text");
        assert_eq!(decoded.text, "€“”");
        assert!(decoded.exact);
    }

    #[test]
    fn symbol_face_text_decodes_to_greek_without_the_symbol_face() {
        let font = Font {
            charset: 2,
            face_name: "Symbol".to_owned(),
            ..Font::default()
        };
        let decoded =
            decode_text(b"ab", 0, 2, false, &font, EmrType::ExtTextOutA).expect("symbol text");
        assert_eq!(decoded.text, "αβ");
        assert!(decoded.exact);
        assert!(decoded.unicode_face);
        assert!(
            !font
                .svg_attrs(12.0, !decoded.unicode_face)
                .contains("Symbol")
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
        let parser = EmfParser::new(&bytes, &crate::Limits::default()).unwrap();
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
        let parser = EmfParser::new(&bytes, &crate::Limits::default()).unwrap();
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
        let dc = DeviceContext {
            world_transform: XForm {
                m11: 0.0,
                m12: 1.0,
                m21: -1.0,
                m22: 0.0,
                dx: 0.0,
                dy: 0.0,
            },
            ..DeviceContext::default()
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

    /// Converts a classic EMF built from `records` and returns its SVG and the
    /// diagnostic codes of playback.
    fn convert_codes(records: Vec<Vec<u8>>) -> (String, Vec<&'static str>) {
        let bytes = classic_emf(records);
        let parser = EmfParser::new(&bytes, &crate::Limits::default()).unwrap();
        let (svg, diagnostics) = EmfSvgConverter::new(&parser)
            .convert_with_diagnostics()
            .unwrap();
        (svg, diagnostics.into_iter().map(|d| d.code).collect())
    }

    fn handle_record(kind: EmrType, handle: u32) -> Vec<u8> {
        let mut out = Vec::new();
        dword(&mut out, handle);
        record(kind, &out)
    }

    fn brush_record(handle: u32) -> Vec<u8> {
        let mut out = Vec::new();
        for value in [handle, 0, 0x0000_00ff, 0] {
            dword(&mut out, value);
        }
        record(EmrType::CreateBrushIndirect, &out)
    }

    #[test]
    fn a_handle_deleted_while_selected_can_be_created_again() {
        let (svg, codes) = convert_codes(vec![
            brush_record(1),
            handle_record(EmrType::SelectObject, 1),
            handle_record(EmrType::DeleteObject, 1),
            brush_record(1),
            handle_record(EmrType::SelectObject, 1),
        ]);
        assert!(svg.starts_with("<svg "));
        assert!(codes.is_empty(), "{codes:?}");
    }

    #[test]
    fn object_table_misuse_that_gdi_ignores_reports_noncanonical_codes() {
        let (_, codes) = convert_codes(vec![
            brush_record(1),
            brush_record(1),
            brush_record(0x8000_0005),
            handle_record(EmrType::SelectObject, 9),
            record(EmrType::RestoreDc, &(-1_i32).to_le_bytes()),
            record(EmrType::CloseFigure, &[]),
        ]);
        assert_eq!(
            codes,
            [
                "noncanonical-emf-object-handle-reused",
                "noncanonical-emf-stock-object-handle",
                "noncanonical-emf-select-unknown-object",
                "noncanonical-emf-restore-dc-level",
                "noncanonical-emf-close-figure",
            ]
        );
    }

    fn set_world_transform(m11: f32, m12: f32, m21: f32, m22: f32, dx: f32, dy: f32) -> Vec<u8> {
        let mut out = Vec::new();
        for value in [m11, m12, m21, m22, dx, dy] {
            out.extend_from_slice(&value.to_le_bytes());
        }
        dword(&mut out, 4);
        record(EmrType::ModifyWorldTransform, &out)
    }

    fn log_font(height: i32, face: &str, charset: u8, pitch_and_family: u8) -> Vec<u8> {
        let mut out = Vec::new();
        for value in [height, 0, 0, 0, 400] {
            sdword(&mut out, value);
        }
        out.extend_from_slice(&[0, 0, 0, charset, 0, 0, 0, pitch_and_family]);
        let mut name: Vec<u8> = face.encode_utf16().flat_map(u16::to_le_bytes).collect();
        name.resize(64, 0);
        out.extend_from_slice(&name);
        out
    }

    fn font_record(handle: u32, height: i32, face: &str) -> Vec<u8> {
        font_record_with(handle, height, face, 0, 0)
    }

    fn font_record_with(
        handle: u32,
        height: i32,
        face: &str,
        charset: u8,
        pitch_and_family: u8,
    ) -> Vec<u8> {
        let mut out = Vec::new();
        dword(&mut out, handle);
        out.extend_from_slice(&log_font(height, face, charset, pitch_and_family));
        record(EmrType::ExtCreateFontIndirectW, &out)
    }

    fn extended_pen_record(handle: u32, style: u32, width: u32) -> Vec<u8> {
        let mut out = Vec::new();
        for value in [handle, 0, 0, 0, 0, style, width, 0, 0, 0, 0] {
            dword(&mut out, value);
        }
        record(EmrType::ExtCreatePen, &out)
    }

    fn text_out_w(x: i32, y: i32, text: &str, dx: Option<&[i32]>) -> Vec<u8> {
        let units: Vec<u16> = text.encode_utf16().collect();
        let string: Vec<u8> = units.iter().flat_map(|unit| unit.to_le_bytes()).collect();
        text_out(EmrType::ExtTextOutW, (x, y), &string, units.len(), 0, dx)
    }

    /// An ExtTextOutA or ExtTextOutW record: its header, the string at offset
    /// 76 padded to four bytes, then the advances.
    fn text_out(
        kind: EmrType,
        (x, y): (i32, i32),
        string: &[u8],
        count: usize,
        options: u32,
        dx: Option<&[i32]>,
    ) -> Vec<u8> {
        let mut string = string.to_vec();
        while (76 + string.len()) % 4 != 0 {
            string.push(0);
        }
        let mut out = Vec::new();
        out.extend_from_slice(&[0u8; 16]); // bounds
        dword(&mut out, 1); // GM_COMPATIBLE
        out.extend_from_slice(&1.0_f32.to_le_bytes());
        out.extend_from_slice(&1.0_f32.to_le_bytes());
        sdword(&mut out, x);
        sdword(&mut out, y);
        dword(&mut out, u32::try_from(count).unwrap());
        dword(&mut out, 76); // string follows the 68-byte header
        dword(&mut out, options);
        out.extend_from_slice(&[0u8; 16]); // rectangle
        let dx_offset = 76 + u32::try_from(string.len()).unwrap();
        dword(&mut out, if dx.is_some() { dx_offset } else { 0 });
        out.extend_from_slice(&string);
        if let Some(values) = dx {
            for value in values {
                sdword(&mut out, *value);
            }
        }
        record(kind, &out)
    }

    fn modify_world_transform(dx: f32, mode: u32) -> Vec<u8> {
        let mut out = Vec::new();
        for value in [1.0_f32, 0.0, 0.0, 1.0, dx, 0.0] {
            out.extend_from_slice(&value.to_le_bytes());
        }
        dword(&mut out, mode);
        record(EmrType::ModifyWorldTransform, &out)
    }

    #[test]
    fn modify_world_transform_set_replaces_the_transform() {
        let (svg, codes) = convert_codes(vec![
            modify_world_transform(7.0, 4),
            modify_world_transform(5.0, 4),
            record(EmrType::MoveToEx, &point(0, 0)),
            line_to(10, 0),
        ]);
        assert!(codes.is_empty(), "{codes:?}");
        assert!(
            svg.contains(r#"x1="5""#) && svg.contains(r#"x2="15""#),
            "{svg}"
        );

        let (_, codes) = convert_codes(vec![modify_world_transform(5.0, 9)]);
        assert_eq!(codes, ["noncanonical-emf-modify-world-transform-mode"]);
    }

    #[test]
    fn invalid_state_records_that_gdi_refuses_are_ignored() {
        let mut zero_extent = Vec::new();
        sdword(&mut zero_extent, 0);
        sdword(&mut zero_extent, 5);
        let (_, codes) = convert_codes(vec![
            record(EmrType::SetWindowExtEx, &zero_extent),
            record(EmrType::SetBkMode, &0_u32.to_le_bytes()),
            record(EmrType::SetPolyFillMode, &7_u32.to_le_bytes()),
        ]);
        assert_eq!(
            codes,
            [
                "noncanonical-emf-zero-mapping-extent",
                "noncanonical-emf-background-mode",
                "noncanonical-emf-polygon-fill-mode",
            ]
        );
    }

    #[test]
    fn shapes_without_area_and_empty_poly_records_draw_nothing() {
        let mut rect = Vec::new();
        for value in [3, 3, 3, 9] {
            sdword(&mut rect, value);
        }
        let mut poly = vec![0u8; 16];
        poly.extend_from_slice(&0_u32.to_le_bytes());
        let (svg, codes) = convert_codes(vec![
            record(EmrType::Rectangle, &rect),
            record(EmrType::Ellipse, &rect),
            record(EmrType::Polygon, &poly),
        ]);
        assert!(!svg.contains("<path"));
        assert_eq!(
            codes,
            [
                "noncanonical-emf-empty-shape",
                "noncanonical-emf-empty-shape",
                "noncanonical-emf-empty-poly-record"
            ]
        );
    }

    #[test]
    fn multibyte_ansi_text_uses_the_font_character_set() {
        let font = Font {
            charset: 128,
            face_name: "MS Gothic".to_owned(),
            ..Font::default()
        };
        let decoded = decode_text(
            &[0x82, 0xa0, b'a'],
            0,
            3,
            false,
            &font,
            EmrType::ExtTextOutA,
        )
        .expect("shift-jis text");
        assert_eq!(decoded.text, "あa");
        assert!(decoded.exact);

        let broken = decode_text(&[0x82], 0, 1, false, &font, EmrType::ExtTextOutA).unwrap();
        assert!(!broken.exact);
    }

    #[test]
    fn unicode_text_in_the_symbol_face_becomes_real_characters() {
        let font = Font {
            charset: 2,
            face_name: "Symbol".to_owned(),
            ..Font::default()
        };
        let bytes: Vec<u8> = [0xf070_u16, 0xf065]
            .iter()
            .flat_map(|u| u.to_le_bytes())
            .collect();
        let decoded = decode_text(&bytes, 0, 2, true, &font, EmrType::ExtTextOutW).unwrap();
        assert_eq!(decoded.text, "πε");
        assert!(decoded.unicode_face);
    }

    #[test]
    fn an_empty_clip_rectangle_hides_later_drawing() {
        let (svg, codes) = convert_codes(vec![
            clip_rect(5, 5, 5, 20),
            record(EmrType::MoveToEx, &point(0, 0)),
            line_to(30, 30),
        ]);
        assert!(codes.is_empty(), "{codes:?}");
        assert!(svg.contains("clip-path"), "{svg}");
    }

    fn emfplus_record(kind: u16, flags: u16, data: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&kind.to_le_bytes());
        out.extend_from_slice(&flags.to_le_bytes());
        dword(&mut out, u32::try_from(12 + data.len()).unwrap());
        dword(&mut out, u32::try_from(data.len()).unwrap());
        out.extend_from_slice(data);
        out
    }

    fn emfplus_comment(records: &[Vec<u8>]) -> Vec<u8> {
        let payload = records.concat();
        let mut body = Vec::new();
        dword(&mut body, u32::try_from(4 + payload.len()).unwrap());
        dword(&mut body, crate::emfplus::EMFPLUS_COMMENT_IDENTIFIER);
        body.extend_from_slice(&payload);
        record(EmrType::Comment, &body)
    }

    /// An `EmfPlusHeader` of a dual metafile recorded for a 96 DPI display.
    fn dual_header() -> Vec<u8> {
        let mut data = Vec::new();
        for value in [0xdbc0_1002_u32, 1, 96, 96] {
            dword(&mut data, value);
        }
        emfplus_record(0x4001, 1, &data)
    }

    fn emfplus_eof() -> Vec<u8> {
        emfplus_record(0x4002, 0, &[])
    }

    /// An `EmfPlusFillRects` record with one solid rectangle.
    fn emfplus_fill() -> Vec<u8> {
        let mut data = Vec::new();
        dword(&mut data, 0xff12_3456);
        dword(&mut data, 1);
        for value in [0.0_f32, 0.0, 50.0, 50.0] {
            data.extend_from_slice(&value.to_le_bytes());
        }
        emfplus_record(0x400a, 0x8000, &data)
    }

    fn classic_rectangle() -> Vec<u8> {
        let mut data = Vec::new();
        for value in [10, 10, 60, 60] {
            sdword(&mut data, value);
        }
        record(EmrType::Rectangle, &data)
    }

    fn convert(emf: &[u8]) -> (String, Vec<&'static str>) {
        let parser = EmfParser::new(emf, &crate::Limits::default()).unwrap();
        let (svg, diagnostics) = EmfSvgConverter::new(&parser)
            .convert_with_diagnostics()
            .unwrap();
        (
            svg,
            diagnostics
                .iter()
                .map(|diagnostic| diagnostic.code)
                .collect(),
        )
    }

    #[test]
    fn dual_metafile_plays_its_exact_emfplus_track_once() {
        let emf = classic_emf(vec![
            emfplus_comment(&[dual_header(), emfplus_fill(), emfplus_eof()]),
            classic_rectangle(),
        ]);
        let (svg, codes) = convert(&emf);
        assert!(svg.contains("#123456"), "{svg}");
        assert!(!svg.contains("<rect x=\"10\""), "{svg}");
        assert!(codes.is_empty(), "{codes:?}");
    }

    #[test]
    fn dual_metafile_with_get_dc_plays_its_classic_track() {
        let emf = classic_emf(vec![
            emfplus_comment(&[
                dual_header(),
                emfplus_fill(),
                emfplus_record(0x4004, 0, &[]),
            ]),
            classic_rectangle(),
            emfplus_comment(&[emfplus_eof()]),
        ]);
        let (svg, codes) = convert(&emf);
        assert!(!svg.contains("#123456") && svg.contains("M10 10"), "{svg}");
        assert_eq!(codes, vec!["noncanonical-emfplus-dual-getdc"]);
        assert!(
            EmfSvgConverter::new(&EmfParser::new(&emf, &crate::Limits::default()).unwrap())
                .convert()
                .is_ok()
        );
    }

    #[test]
    fn approximated_emfplus_track_falls_back_to_the_classic_track() {
        // A hatch brush is only approximated by EMF+ playback.
        let mut brush = Vec::new();
        for value in [0xdbc0_1002_u32, 1, 0, 0xff00_0000, 0xffff_ffff] {
            dword(&mut brush, value);
        }
        let mut fill = Vec::new();
        dword(&mut fill, 0);
        dword(&mut fill, 1);
        for value in [0.0_f32, 0.0, 50.0, 50.0] {
            fill.extend_from_slice(&value.to_le_bytes());
        }
        let emf = classic_emf(vec![
            emfplus_comment(&[
                dual_header(),
                emfplus_record(0x4008, 0x0100, &brush),
                emfplus_record(0x400a, 0, &fill),
                emfplus_eof(),
            ]),
            classic_rectangle(),
        ]);
        let (svg, codes) = convert(&emf);
        assert!(svg.contains("M10 10"), "{svg}");
        assert_eq!(codes, vec!["noncanonical-emfplus-approximated"]);
    }

    #[test]
    fn unreadable_emfplus_track_falls_back_to_the_classic_track() {
        // The EMF+ stream never ends.
        let emf = classic_emf(vec![
            emfplus_comment(&[dual_header(), emfplus_fill()]),
            classic_rectangle(),
        ]);
        let (svg, codes) = convert(&emf);
        assert!(!svg.contains("#123456") && svg.contains("M10 10"), "{svg}");
        assert_eq!(codes, vec!["noncanonical-emfplus-unreadable"]);
    }

    #[test]
    fn emfplus_records_after_classic_records_are_ignored() {
        let emf = classic_emf(vec![
            classic_rectangle(),
            emfplus_comment(&[dual_header(), emfplus_fill(), emfplus_eof()]),
        ]);
        let (svg, codes) = convert(&emf);
        assert!(!svg.contains("#123456") && svg.contains("M10 10"), "{svg}");
        assert_eq!(codes, vec!["noncanonical-emfplus-not-leading"]);
    }

    #[test]
    fn emfplus_track_shares_the_classic_view_box() {
        let mut emf = classic_emf(vec![emfplus_comment(&[
            dual_header(),
            emfplus_fill(),
            emfplus_eof(),
        ])]);
        // Bounds from (5, 7) to (100, 100).
        emf[8..12].copy_from_slice(&5_i32.to_le_bytes());
        emf[12..16].copy_from_slice(&7_i32.to_le_bytes());
        let (svg, _) = convert(&emf);
        assert!(svg.contains("viewBox=\"5 7 95 93\""), "{svg}");
    }

    #[test]
    fn angle_arc_sweeps_beyond_one_turn_stay_bounded() {
        let mut arc = point(50, 50);
        arc.extend_from_slice(&20_u32.to_le_bytes());
        arc.extend_from_slice(&0.0_f32.to_le_bytes());
        arc.extend_from_slice(&3.9e11_f32.to_le_bytes());
        let (svg, codes) = convert_codes(vec![
            record(EmrType::MoveToEx, &point(70, 50)),
            record(EmrType::AngleArc, &arc),
        ]);
        assert!(codes.is_empty(), "{codes:?}");
        // One turn plus the remainder: at most 64 segments.
        let segments = svg.matches(' ').count();
        assert!(segments < 400, "{} bytes of SVG", svg.len());
    }

    #[test]
    fn a_negative_logpen_width_draws_a_cosmetic_pen() {
        let mut pen = Vec::new();
        pen.extend_from_slice(&1_u32.to_le_bytes());
        pen.extend_from_slice(&0_u32.to_le_bytes());
        pen.extend_from_slice(&(-1_i32).to_le_bytes());
        pen.extend_from_slice(&0_u32.to_le_bytes());
        pen.extend_from_slice(&0_u32.to_le_bytes());
        let mut select = Vec::new();
        select.extend_from_slice(&1_u32.to_le_bytes());
        let (svg, _) = convert_codes(vec![
            record(EmrType::CreatePen, &pen),
            record(EmrType::SelectObject, &select),
            record(EmrType::MoveToEx, &point(0, 0)),
            line_to(30, 30),
        ]);
        assert!(!svg.contains("stroke-width=\"4294967295\""), "{svg}");
        assert!(!svg.contains("e9"), "{svg}");
    }

    #[test]
    fn classic_text_height_follows_the_world_transform() {
        let (svg, codes) = convert_codes(vec![
            set_world_transform(0.25, 0.0, 0.0, 0.25, 0.0, 0.0),
            font_record(1, -100, "Arial"),
            handle_record(EmrType::SelectObject, 1),
            text_out_w(0, 0, "A", None),
        ]);
        assert!(codes.is_empty(), "{codes:?}");
        assert!(svg.contains("font-size=\"25\""), "{svg}");
    }

    #[test]
    fn symbol_face_text_keeps_the_face_and_maps_its_positions() {
        let (svg, codes) = convert_codes(vec![
            font_record(1, -100, "MT Extra"),
            handle_record(EmrType::SelectObject, 1),
            text_out_w(0, 0, "h", None),
        ]);
        assert!(codes.is_empty(), "{codes:?}");
        // MT Extra position 0x68 is Planck's constant over two pi, and the
        // face stays selectable for positions without a Unicode equivalent.
        assert!(svg.contains('\u{210f}'), "{svg}");
        assert!(svg.contains(r#"font-family="MT Extra""#), "{svg}");
    }

    #[test]
    fn a_zero_text_height_uses_the_default_font_size() {
        let (svg, codes) = convert_codes(vec![
            font_record(1, 0, "Arial"),
            handle_record(EmrType::SelectObject, 1),
            text_out_w(0, 0, "A", None),
        ]);
        assert!(codes.is_empty(), "{codes:?}");
        assert!(svg.contains("font-size=\"12\""), "{svg}");
    }

    #[test]
    fn geometric_pen_width_follows_the_world_transform() {
        let (svg, codes) = convert_codes(vec![
            set_world_transform(0.25, 0.0, 0.0, 0.25, 0.0, 0.0),
            extended_pen_record(1, pen_style::GEOMETRIC | pen_style::SOLID, 6),
            handle_record(EmrType::SelectObject, 1),
            record(EmrType::MoveToEx, &point(0, 0)),
            line_to(200, 200),
        ]);
        assert!(codes.is_empty(), "{codes:?}");
        assert!(svg.contains("stroke-width=\"1.5\""), "{svg}");
    }

    #[test]
    fn cosmetic_pen_width_is_a_device_width() {
        let (svg, codes) = convert_codes(vec![
            set_world_transform(0.25, 0.0, 0.0, 0.25, 0.0, 0.0),
            extended_pen_record(1, pen_style::COSMETIC | pen_style::SOLID, 3),
            handle_record(EmrType::SelectObject, 1),
            record(EmrType::MoveToEx, &point(0, 0)),
            line_to(200, 200),
        ]);
        assert!(codes.is_empty(), "{codes:?}");
        assert!(svg.contains("stroke-width=\"3\""), "{svg}");
    }

    fn create_pen_record(handle: u32, style: u32, width: i32) -> Vec<u8> {
        let mut out = Vec::new();
        dword(&mut out, handle);
        dword(&mut out, style);
        sdword(&mut out, width);
        sdword(&mut out, 0);
        dword(&mut out, 0);
        record(EmrType::CreatePen, &out)
    }

    #[test]
    fn logpen_width_follows_the_world_transform() {
        // GDI creates a LogPen with its width in logical units even though
        // the style carries no PS_GEOMETRIC bit.
        let (svg, codes) = convert_codes(vec![
            set_world_transform(0.25, 0.0, 0.0, 0.25, 0.0, 0.0),
            create_pen_record(1, pen_style::SOLID, 8),
            handle_record(EmrType::SelectObject, 1),
            record(EmrType::MoveToEx, &point(0, 0)),
            line_to(200, 200),
        ]);
        assert!(codes.is_empty(), "{codes:?}");
        assert!(svg.contains("stroke-width=\"2\""), "{svg}");
        // Dashes are multiples of the drawn width.
        let (svg, _) = convert_codes(vec![
            set_world_transform(0.25, 0.0, 0.0, 0.25, 0.0, 0.0),
            create_pen_record(1, pen_style::DASH, 8),
            handle_record(EmrType::SelectObject, 1),
            record(EmrType::MoveToEx, &point(0, 0)),
            line_to(200, 200),
        ]);
        assert!(svg.contains("stroke-dasharray=\"6 2\""), "{svg}");
    }

    #[test]
    fn zero_width_pens_stay_one_device_pixel_wide() {
        for pen in [
            create_pen_record(1, pen_style::SOLID, 0),
            extended_pen_record(1, pen_style::GEOMETRIC | pen_style::SOLID, 0),
        ] {
            let (svg, codes) = convert_codes(vec![
                set_world_transform(16.0, 0.0, 0.0, 16.0, 0.0, 0.0),
                pen,
                handle_record(EmrType::SelectObject, 1),
                record(EmrType::MoveToEx, &point(0, 0)),
                line_to(20, 20),
            ]);
            assert!(codes.is_empty(), "{codes:?}");
            assert!(svg.contains("<line"), "{svg}");
            assert!(!svg.contains("stroke-width"), "{svg}");
        }
        // So does the stock black pen.
        let (svg, _) = convert_codes(vec![
            set_world_transform(16.0, 0.0, 0.0, 16.0, 0.0, 0.0),
            record(EmrType::MoveToEx, &point(0, 0)),
            line_to(20, 20),
        ]);
        assert!(!svg.contains("stroke-width"), "{svg}");
    }

    #[test]
    fn text_anchoring_uses_the_recorded_glyph_positions() {
        let (svg, codes) = convert_codes(vec![
            record(EmrType::SetTextAlign, &6_u32.to_le_bytes()),
            font_record(1, -10, "Arial"),
            handle_record(EmrType::SelectObject, 1),
            text_out_w(50, 20, "ab", Some(&[10, 10])),
        ]);
        assert!(codes.is_empty(), "{codes:?}");
        // The 20 unit string is centered on 50, so the glyphs start at 40 and
        // 50; SVG's own text-anchor would center each glyph on its position.
        assert!(svg.contains("<text x=\"40 50\""), "{svg}");
        assert!(!svg.contains("text-anchor=\"middle\""), "{svg}");
    }

    #[test]
    fn advances_place_characters_after_every_code_unit_they_cover() {
        // A Shift_JIS character takes two bytes and two advances.
        let (svg, codes) = convert_codes(vec![
            font_record_with(1, -10, "MS Mincho", 128, 0),
            handle_record(EmrType::SelectObject, 1),
            text_out(
                EmrType::ExtTextOutA,
                (0, 0),
                b"a\x82\xa0b",
                4,
                0,
                Some(&[10, 6, 6, 8]),
            ),
        ]);
        assert!(codes.is_empty(), "{codes:?}");
        assert!(svg.contains("<text x=\"0 10 22\" y=\"0 0 0\""), "{svg}");
        // So does a surrogate pair in UTF-16, and every space of a run keeps
        // its own position.
        let (svg, _) = convert_codes(vec![text_out_w(
            0,
            0,
            "a\u{1d400}  b",
            Some(&[10, 6, 6, 5, 5, 8]),
        )]);
        assert!(svg.contains("<text x=\"0 10 22 27 32\""), "{svg}");
        assert!(
            svg.contains("xml:space=\"preserve\">a\u{1d400}  b<"),
            "{svg}"
        );
    }

    #[test]
    fn vertical_advances_move_the_characters_down_the_page() {
        let (svg, codes) = convert_codes(vec![text_out(
            EmrType::ExtTextOutW,
            (5, 5),
            &[b'a', 0, b'b', 0],
            2,
            text_options::PDY,
            Some(&[10, 3, 10, 3]),
        )]);
        assert!(codes.is_empty(), "{codes:?}");
        assert!(svg.contains("<text x=\"5 15\" y=\"5 8\""), "{svg}");
    }

    #[test]
    fn the_font_family_names_the_recorded_face_then_its_class() {
        let (svg, codes) = convert_codes(vec![
            font_record_with(1, -10, "Garamond Premier", 0, 0x12),
            handle_record(EmrType::SelectObject, 1),
            text_out_w(0, 0, "a", None),
        ]);
        assert!(codes.is_empty(), "{codes:?}");
        assert!(
            svg.contains("font-family=\"Garamond Premier, serif\""),
            "{svg}"
        );
    }
}
