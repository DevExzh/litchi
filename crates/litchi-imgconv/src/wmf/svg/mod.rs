//! Safe WMF playback to a compact, self-contained SVG document.

mod bounds;
mod renderer;
mod simd;
mod state;
mod style;
mod transform;

use std::sync::Arc;

use super::parser::WmfParser;
use crate::svg_utils::write_num;
use bounds::BoundsCalculator;
use litchi_core::error::{Error, Result};
use renderer::SvgRenderer;
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
        if issues.is_empty() {
            Ok(svg)
        } else {
            Err(Error::Unsupported(format!(
                "WMF SVG playback is not exact: {}",
                issues
                    .iter()
                    .map(|issue| issue.message.as_str())
                    .collect::<Vec<_>>()
                    .join("; ")
            )))
        }
    }

    pub(crate) fn to_svg_reported(&self) -> Result<(String, Vec<String>)> {
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
        Ok((svg, issues.into_iter().map(|issue| issue.message).collect()))
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

    fn render(&self) -> (String, Vec<renderer::RenderIssue>) {
        let bbox = if let Some(placeable) = &self.parser.placeable {
            normalized_bbox((
                f64::from(placeable.left),
                f64::from(placeable.top),
                f64::from(placeable.right),
                f64::from(placeable.bottom),
            ))
        } else {
            BoundsCalculator::scan_records(&self.parser.records).as_tuple()
        };
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
        let mut elements = String::with_capacity(4096);
        let mut element_count = 0usize;
        let mut output_limit_hit = false;
        for record in &self.parser.records {
            if let Some(element) = renderer.render_record(record) {
                element_count = element_count.saturating_add(svg_element_count(&element));
                if elements
                    .len()
                    .checked_add(element.len())
                    .is_none_or(|size| size > self.limits.max_output_bytes)
                {
                    output_limit_hit = true;
                    break;
                }
                elements.push_str(&element);
            }
        }
        let (definitions, mut issues) = renderer.into_parts();
        element_count = element_count.saturating_add(svg_element_count(&definitions));
        if element_count > self.limits.max_svg_elements {
            issues.push(renderer::RenderIssue {
                fatal: true,
                message: format!(
                    "WMF SVG element count exceeds limit {}",
                    self.limits.max_svg_elements
                ),
            });
        }
        if output_limit_hit {
            issues.push(renderer::RenderIssue {
                fatal: true,
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
        let parser = WmfParser::new(&data).expect("parse santa WMF");
        let svg = WmfConverter::new(parser)
            .to_svg()
            .expect("render santa WMF");
        assert!(svg.starts_with("<svg"));
        assert!(svg.contains("<polygon"));
        assert!(svg.ends_with("</svg>"));
        assert!(!svg.contains("NaN"));
        assert!(!svg.contains("inf"));
    }
}
