//! Safe WMF playback to a compact, self-contained SVG document.

mod bounds;
mod renderer;
mod simd;
mod state;
mod style;
mod transform;

use std::sync::Arc;

use super::parser::WmfParser;
use crate::svg_utils::{ClipGrouper, write_num};
use bounds::BoundsCalculator;
use litchi_core::error::{Error, Result};
use renderer::SvgRenderer;
use state::MappingState;
use transform::CoordinateTransform;

pub(super) type BitmapHook =
    dyn Fn(u16, &[u8], [f64; 4]) -> std::result::Result<Option<String>, String> + Send + Sync;

/// Stateful WMF-to-SVG converter.
pub struct WmfConverter {
    parser: WmfParser,
    bitmap_hook: Option<Arc<BitmapHook>>,
    limits: crate::Limits,
}

impl WmfConverter {
    pub fn new(parser: WmfParser) -> Self {
        Self {
            parser,
            bitmap_hook: None,
            limits: crate::Limits::default(),
        }
    }

    pub(crate) fn with_limits(parser: WmfParser, limits: crate::Limits) -> Self {
        Self {
            parser,
            bitmap_hook: None,
            limits,
        }
    }

    /// Installs an override for bitmap record rendering.
    ///
    /// The callback receives the canonical function, raw record parameters,
    /// and the transformed SVG destination `[x, y, width, height]`. Returning
    /// `None` explicitly declines the record and is fatal in strict playback.
    /// Without an override, the crate's bounded shared DIB renderer is used.
    #[must_use]
    pub fn with_bitmap_hook<F>(mut self, hook: F) -> Self
    where
        F: Fn(u16, &[u8], [f64; 4]) -> std::result::Result<Option<String>, String>
            + Send
            + Sync
            + 'static,
    {
        self.bitmap_hook = Some(Arc::new(hook));
        self
    }

    /// Strict conversion. Unsupported or malformed output-affecting records
    /// fail rather than silently disappearing from the document.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Unsupported`] when strict playback encounters a
    /// malformed or unsupported output-affecting record.
    pub fn to_svg(&self) -> Result<String> {
        let (svg, issues) = self.render();
        let inexact = issues
            .iter()
            .filter(|issue| issue.fatal || crate::codec::is_lossy_diagnostic(issue.code))
            .map(|issue| issue.message.as_str())
            .collect::<Vec<_>>();
        if inexact.is_empty() {
            Ok(svg)
        } else {
            Err(Error::Unsupported(format!(
                "WMF SVG playback is not exact: {}",
                inexact.join("; ")
            )))
        }
    }

    /// Converts and returns every non-fatal issue with its diagnostic code.
    pub(crate) fn to_svg_reported(&self) -> Result<(String, Vec<(&'static str, String)>)> {
        let (svg, issues) = self.render();
        let fatal = issues
            .iter()
            .filter(|issue| issue.fatal)
            .map(|issue| issue.message.as_str())
            .collect::<Vec<_>>();
        if !fatal.is_empty() {
            return Err(Error::Unsupported(format!(
                "WMF SVG playback failed: {}",
                fatal.join("; ")
            )));
        }
        Ok((
            svg,
            issues
                .into_iter()
                .map(|issue| (issue.code, issue.message))
                .collect(),
        ))
    }

    /// Lenient conversion that returns the SVG plus all playback diagnostics.
    /// Fatal diagnostics are included in the report but do not discard output.
    ///
    /// # Errors
    ///
    /// Reserved for conversion resource failures; record playback problems are
    /// returned in the diagnostic vector.
    pub fn to_svg_with_diagnostics(&self) -> Result<(String, Vec<String>)> {
        let (svg, issues) = self.render();
        Ok((
            svg,
            issues
                .into_iter()
                .map(|issue| {
                    if issue.fatal {
                        format!("error: {}", issue.message)
                    } else {
                        format!("warning: {}", issue.message)
                    }
                })
                .collect(),
        ))
    }

    /// Frame rectangle and playback mapping of the picture.
    ///
    /// A placeable header declares the frame in logical units. Without one,
    /// the metafile declares nothing: the first window extent that can frame
    /// a picture is the frame players give it, and the records' drawn bounds
    /// say how much of it they use.
    fn frame(&self) -> ((f64, f64, f64, f64), MappingState) {
        if let Some(placeable) = &self.parser.placeable {
            let bbox = normalized_bbox((
                f64::from(placeable.left),
                f64::from(placeable.top),
                f64::from(placeable.right),
                f64::from(placeable.bottom),
            ));
            // The placeable rectangle is the frame, in logical units of the
            // window the metafile is designed for.
            let mapping =
                MappingState::playback((bbox.0, bbox.1), (bbox.2 - bbox.0, bbox.3 - bbox.1));
            (bbox, mapping)
        } else {
            let extent =
                BoundsCalculator::first_window_extent(&self.parser.records).unwrap_or((1.0, 1.0));
            let mapping = MappingState::playback((0.0, 0.0), extent);
            let bbox = BoundsCalculator::scan_records(&self.parser.records, mapping).as_tuple();
            (bbox, mapping)
        }
    }

    /// Size a raster converter samples the picture at.
    ///
    /// A placeable header declares the display size of the picture in logical
    /// units. A metafile without one declares no size, so it gets the size of
    /// its SVG canvas, which keeps the aspect ratio of the frame its own
    /// records describe.
    pub(crate) fn natural_size(&self) -> (u32, u32) {
        if let Some(placeable) = &self.parser.placeable {
            return (
                u32::from(placeable.width().unsigned_abs()).max(1),
                u32::from(placeable.height().unsigned_abs()).max(1),
            );
        }
        let (bbox, _) = self.frame();
        let (width, height) = Self::calculate_dimensions(bbox);
        (rounded_dimension(width), rounded_dimension(height))
    }

    fn render(&self) -> (String, Vec<renderer::RenderIssue>) {
        let (bbox, mapping) = self.frame();
        let (svg_width, svg_height) = Self::calculate_dimensions(bbox);
        let transform = CoordinateTransform::new(bbox, svg_width, svg_height);
        let mut renderer = SvgRenderer::new(
            transform,
            self.bitmap_hook.as_deref(),
            dib_limits(self.limits),
            self.limits.max_state_depth,
            self.limits.max_objects,
            self.limits.max_path_points,
            self.limits.max_output_bytes,
            self.limits.max_uncompressed_bytes,
        );
        renderer.set_initial_mapping(mapping);
        let mut elements = String::with_capacity(4096);
        let mut groups = ClipGrouper::new();
        let mut element_count = 0usize;
        let mut output_limit_hit = false;
        let mut records = self.parser.records.iter();
        loop {
            let element = match records.next() {
                Some(record) => renderer.render_record(record),
                // Deferred bitmap records paint at the end of playback.
                None => match renderer.finish() {
                    Some(element) => Some(element),
                    None => break,
                },
            };
            if let Some(element) = element {
                element_count = element_count.saturating_add(svg_element_count(&element));
                // Grouping only removes bytes, so the ungrouped size bounds the
                // output.
                if elements
                    .len()
                    .checked_add(element.len())
                    .is_none_or(|size| size > self.limits.max_output_bytes)
                {
                    output_limit_hit = true;
                    break;
                }
                groups.push(&mut elements, &element);
            }
        }
        groups.finish(&mut elements);
        let (definitions, mut issues) = renderer.into_parts();
        element_count = element_count
            .saturating_add(svg_element_count(&definitions))
            .saturating_add(groups.groups());
        if element_count > self.limits.max_svg_elements {
            issues.push(renderer::RenderIssue {
                fatal: true,
                code: renderer::PLAYBACK_WARNING,
                message: format!(
                    "WMF SVG element count exceeds limit {}",
                    self.limits.max_svg_elements
                ),
            });
        }
        if output_limit_hit {
            issues.push(renderer::RenderIssue {
                fatal: true,
                code: renderer::PLAYBACK_WARNING,
                message: format!(
                    "WMF SVG output exceeds limit {} bytes",
                    self.limits.max_output_bytes
                ),
            });
        }

        let estimated = elements
            .len()
            .checked_add(definitions.len())
            .and_then(|size| size.checked_add(256));
        if estimated.is_none_or(|size| size > self.limits.max_output_bytes) {
            issues.push(renderer::RenderIssue {
                fatal: true,
                code: renderer::PLAYBACK_WARNING,
                message: format!(
                    "WMF SVG output exceeds limit {} bytes",
                    self.limits.max_output_bytes
                ),
            });
            return (String::new(), issues);
        }
        let mut svg = String::with_capacity(estimated.unwrap_or(256));
        svg.push_str(r#"<svg width=""#);
        write_num(&mut svg, svg_width);
        svg.push_str(r#"" height=""#);
        write_num(&mut svg, svg_height);
        svg.push_str(r#"" viewBox="0 0 "#);
        write_num(&mut svg, svg_width);
        svg.push(' ');
        write_num(&mut svg, svg_height);
        svg.push_str(r#"" xmlns="http://www.w3.org/2000/svg">"#);
        if !definitions.is_empty() {
            svg.push_str("<defs>");
            svg.push_str(&definitions);
            svg.push_str("</defs>");
        }
        svg.push_str(&elements);
        svg.push_str("</svg>");
        if svg.len() > self.limits.max_output_bytes {
            issues.push(renderer::RenderIssue {
                fatal: true,
                code: renderer::PLAYBACK_WARNING,
                message: format!(
                    "WMF SVG output exceeds limit {} bytes",
                    self.limits.max_output_bytes
                ),
            });
        }
        (svg, issues)
    }

    fn calculate_dimensions(bbox: (f64, f64, f64, f64)) -> (f64, f64) {
        const MAX_WIDTH: f64 = 768.0;
        const MAX_HEIGHT: f64 = 512.0;

        let width = (bbox.2 - bbox.0).abs();
        let height = (bbox.3 - bbox.1).abs();
        if !width.is_finite() || !height.is_finite() || width <= 0.0 || height <= 0.0 {
            return (768.0, 512.0);
        }
        let scale = (MAX_WIDTH / width).min(MAX_HEIGHT / height);
        ((width * scale).max(1.0), (height * scale).max(1.0))
    }
}

fn dib_limits(limits: crate::Limits) -> crate::dib::DibLimits {
    crate::dib::DibLimits {
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

fn svg_element_count(fragment: &str) -> usize {
    fragment
        .as_bytes()
        .windows(2)
        .filter(|pair| pair[0] == b'<' && !matches!(pair[1], b'/' | b'!' | b'?'))
        .count()
}

/// Rounds a canvas dimension to the pixel size a raster samples it at.
fn rounded_dimension(value: f64) -> u32 {
    if value.is_finite() {
        value.round().clamp(1.0, f64::from(u16::MAX)) as u32
    } else {
        1
    }
}

fn normalized_bbox(bbox: (f64, f64, f64, f64)) -> (f64, f64, f64, f64) {
    let left = bbox.0.min(bbox.2);
    let top = bbox.1.min(bbox.3);
    let right = bbox.0.max(bbox.2);
    let bottom = bbox.1.max(bbox.3);
    if [left, top, right, bottom]
        .iter()
        .all(|value| value.is_finite())
    {
        (left, top, right, bottom)
    } else {
        (0.0, 0.0, 1000.0, 1000.0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn dimensions_handle_reversed_i16_extrema_without_overflow() {
        let dimensions = WmfConverter::calculate_dimensions((
            f64::from(i16::MAX),
            f64::from(i16::MAX),
            f64::from(i16::MIN),
            f64::from(i16::MIN),
        ));
        assert_eq!(dimensions, (512.0, 512.0));
        assert!(dimensions.0.is_finite() && dimensions.1.is_finite());
    }

    #[test]
    fn santa_wmf_converts_end_to_end_without_non_finite_svg() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/images/wmf/santa.wmf"
        );
        let data = std::fs::read(path).expect("santa test WMF");
        let parser = WmfParser::new(&data, &crate::Limits::default()).expect("parse santa WMF");
        let svg = WmfConverter::new(parser)
            .to_svg()
            .expect("render santa WMF");
        assert!(svg.starts_with("<svg"));
        assert!(svg.contains("<polygon"));
        assert!(svg.ends_with("</svg>"));
        assert!(!svg.contains("NaN"));
        assert!(!svg.contains("inf"));
    }

    /// A memory metafile made of `(function, parameters)` records, optionally
    /// behind a placeable header.
    fn wmf(records: &[(u16, Vec<u8>)], placeable: Option<(i16, i16, i16, i16)>) -> Vec<u8> {
        let mut actions = Vec::new();
        let mut max_record = 0_u32;
        for (function, params) in records {
            let mut body = function.to_le_bytes().to_vec();
            body.extend_from_slice(params);
            if body.len() % 2 != 0 {
                body.push(0);
            }
            let words = u32::try_from((body.len() + 4) / 2).unwrap();
            max_record = max_record.max(words);
            actions.extend_from_slice(&words.to_le_bytes());
            actions.extend_from_slice(&body);
        }
        let mut out = Vec::new();
        if let Some((left, top, right, bottom)) = placeable {
            let mut header = vec![0_u8; 22];
            header[0..4].copy_from_slice(&0x9ac6_cdd7_u32.to_le_bytes());
            header[6..8].copy_from_slice(&left.to_le_bytes());
            header[8..10].copy_from_slice(&top.to_le_bytes());
            header[10..12].copy_from_slice(&right.to_le_bytes());
            header[12..14].copy_from_slice(&bottom.to_le_bytes());
            header[14..16].copy_from_slice(&1440_u16.to_le_bytes());
            let checksum = header[..20]
                .as_chunks::<2>()
                .0
                .iter()
                .map(|word| u16::from_le_bytes([word[0], word[1]]))
                .fold(0_u16, |value, word| value ^ word);
            header[20..22].copy_from_slice(&checksum.to_le_bytes());
            out.extend_from_slice(&header);
        }
        let mut header = vec![0_u8; 18];
        header[0..2].copy_from_slice(&1_u16.to_le_bytes());
        header[2..4].copy_from_slice(&9_u16.to_le_bytes());
        header[4..6].copy_from_slice(&0x0300_u16.to_le_bytes());
        // The declared size covers the standard header and the records of the
        // metafile itself, not the placeable header in front of them.
        let words = u32::try_from((18 + actions.len()) / 2).unwrap();
        header[6..10].copy_from_slice(&words.to_le_bytes());
        header[12..16].copy_from_slice(&max_record.to_le_bytes());
        out.extend_from_slice(&header);
        out.extend_from_slice(&actions);
        out
    }

    fn window_extent(width: i16, height: i16) -> (u16, Vec<u8>) {
        let mut params = Vec::new();
        params.extend_from_slice(&height.to_le_bytes());
        params.extend_from_slice(&width.to_le_bytes());
        (0x000b, params)
    }

    fn line(from: (i16, i16), to: (i16, i16)) -> Vec<(u16, Vec<u8>)> {
        let mut move_params = Vec::new();
        move_params.extend_from_slice(&from.1.to_le_bytes());
        move_params.extend_from_slice(&from.0.to_le_bytes());
        let mut line_params = Vec::new();
        line_params.extend_from_slice(&to.1.to_le_bytes());
        line_params.extend_from_slice(&to.0.to_le_bytes());
        vec![(0x0214, move_params), (0x0213, line_params)]
    }

    #[test]
    fn a_metafile_without_a_placeable_header_keeps_its_own_aspect_ratio() {
        let mut records = vec![window_extent(768, 34)];
        records.extend(line((0, 0), (768, 34)));
        records.push((0x0000, Vec::new()));
        let data = wmf(&records, None);
        let parser = WmfParser::new(&data, &crate::Limits::default()).expect("parse WMF");
        let converter = WmfConverter::new(parser);
        assert_eq!(converter.natural_size(), (768, 34));
        let svg = converter.to_svg().expect("render WMF");
        assert!(svg.contains(r#"width="768" height="34""#), "{svg}");
    }

    #[test]
    fn a_placeable_header_declares_the_natural_size() {
        let mut records = vec![window_extent(768, 34)];
        records.extend(line((0, 0), (768, 34)));
        records.push((0x0000, Vec::new()));
        let data = wmf(&records, Some((10, 20, 110, 220)));
        let parser = WmfParser::new(&data, &crate::Limits::default()).expect("parse WMF");
        assert_eq!(WmfConverter::new(parser).natural_size(), (100, 200));
        // A rectangle recorded right to left and bottom to top has the same
        // size.
        let data = wmf(&records, Some((110, 220, 10, 20)));
        let parser = WmfParser::new(&data, &crate::Limits::default()).expect("parse WMF");
        assert_eq!(WmfConverter::new(parser).natural_size(), (100, 200));
    }
}
