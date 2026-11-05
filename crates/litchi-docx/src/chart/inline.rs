//! Charts drawn inline (or anchored) in a paragraph: the `w:drawing`
//! elements whose graphic is a `c:chart` or an extended `cx:chart`.

use std::borrow::Cow;
use std::ops::Range;

use litchi_ooxml_common::xml::attributes::first_wins;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use smallvec::SmallVec;

use crate::error::{Error, Result};

/// The `a:graphicData/@uri` of an ordinary chart.
pub(crate) const CHART_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
/// The Strict `a:graphicData/@uri` of an ordinary chart.
pub(crate) const STRICT_CHART_URI: &str = "http://purl.oclc.org/ooxml/drawingml/chart";

/// The content type of an ordinary chart part.
pub(crate) const CHART_CONTENT_TYPE: &str = super::model::CHART_CT;

/// The content type a main-document chart relationship's target must have,
/// or `None` for a relationship that names no chart.
pub(crate) fn chart_content_type(relationship_type: &str) -> Option<&'static str> {
    use litchi_opc::constants::relationship_type as rt;
    if relationship_type == litchi_drawingml::chart::chartex::RELATIONSHIP_TYPE {
        Some(litchi_drawingml::chart::chartex::CONTENT_TYPE)
    } else if matches!(relationship_type, rt::CHART | rt::STRICT_CHART) {
        Some(super::model::CHART_CT)
    } else {
        None
    }
}

/// One chart of a paragraph: the relationship of its chart part, its frame
/// name and description, and its extent.
///
/// ```
/// use litchi_docx::chart::InlineChart;
///
/// fn describe(chart: &InlineChart) -> String {
///     format!("{} ({} x {} EMU)", chart.name(), chart.width_emu(), chart.height_emu())
/// }
/// # let _ = describe;
/// ```
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct InlineChart {
    relationship_id: String,
    name: String,
    description: String,
    width_emu: i64,
    height_emu: i64,
    extended: bool,
    inline: bool,
}

impl InlineChart {
    /// The main-document relationship of the chart part (`r:id`).
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// The drawing object's name (`wp:docPr/@name`).
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// The alternative text (`wp:docPr/@descr`).
    #[must_use]
    pub fn description(&self) -> &str {
        &self.description
    }

    /// The width in EMUs (`wp:extent/@cx`).
    #[must_use]
    pub const fn width_emu(&self) -> i64 {
        self.width_emu
    }

    /// The height in EMUs (`wp:extent/@cy`).
    #[must_use]
    pub const fn height_emu(&self) -> i64 {
        self.height_emu
    }

    /// Whether the chart is an extended (`cx:chart`) one.
    #[must_use]
    pub const fn is_extended(&self) -> bool {
        self.extended
    }

    /// Whether the drawing is inline (`wp:inline`) rather than floating
    /// (`wp:anchor`).
    #[must_use]
    pub const fn is_inline(&self) -> bool {
        self.inline
    }
}

/// One chart drawing of a paragraph with the byte spans an edit rewrites,
/// relative to the paragraph's XML.
#[derive(Debug, Clone)]
pub(crate) struct ChartSpan {
    pub(crate) chart: InlineChart,
    /// The `w:drawing` element.
    pub(crate) drawing: Range<usize>,
    /// The enclosing `w:r`, when the drawing is its only content besides
    /// its properties.
    pub(crate) sole_run: Option<Range<usize>>,
    /// The `wp:extent` start tag.
    pub(crate) extent: Option<Range<usize>>,
}

fn attribute(element: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    first_wins(element)
        .flatten()
        .find(|attribute| attribute.key.local_name().as_ref() == name)
        .and_then(|attribute| {
            let text = std::str::from_utf8(&attribute.value).ok()?.to_owned();
            Some(quick_xml::escape::unescape(&text).map_or_else(|_| text.clone(), Cow::into_owned))
        })
}

#[derive(Default)]
struct Drawing {
    run: Option<(usize, usize)>,
    width: i64,
    height: i64,
    name: String,
    description: String,
    relationship_id: Option<String>,
    extended: bool,
    inline: bool,
    uri: Option<String>,
    extent: Option<Range<usize>>,
}

/// Every chart drawing of a paragraph's XML, in document order.
pub(crate) fn chart_spans(xml: &[u8]) -> Result<SmallVec<[ChartSpan; 2]>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut found = SmallVec::new();
    // Open elements: (local name, start, non-property child count).
    let mut stack: Vec<(Vec<u8>, usize, usize)> = Vec::new();
    let mut drawing: Option<Drawing> = None;
    // Charts waiting for their run to close: (index in `found`, run start).
    let mut pending: Vec<(usize, usize)> = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("paragraph position overflows".into()))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("paragraph XML: {error}")))?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("paragraph position overflows".into()))?;
        let (element, empty) = match &event {
            Event::Start(element) => (Some(element.clone()), false),
            Event::Empty(element) => (Some(element.clone()), true),
            Event::End(_) => {
                if let Some((local, opened, children)) = stack.pop() {
                    if local == b"drawing"
                        && let Some(open) = drawing.take()
                    {
                        finish(&mut found, &mut pending, open, opened..end);
                    } else if local == b"r" {
                        // A run's drawings are its sole content when it has
                        // no other child but its properties.
                        pending.retain(|(index, run)| {
                            if *run != opened {
                                return true;
                            }
                            if children == 1
                                && let Some(span) = found.get_mut(*index)
                            {
                                span.sole_run = Some(opened..end);
                            }
                            false
                        });
                    }
                }
                continue;
            },
            Event::Eof => break,
            _ => continue,
        };
        let Some(element) = element else { continue };
        let local = element.local_name().as_ref().to_vec();
        if let Some((_, _, children)) = stack.last_mut()
            && local != b"rPr"
        {
            *children += 1;
        }
        match local.as_slice() {
            b"drawing" => {
                let run = stack
                    .iter()
                    .rev()
                    .find(|(name, _, _)| name.as_slice() == b"r")
                    .map(|(_, start, _)| (*start, 0));
                drawing = Some(Drawing {
                    run,
                    width: 0,
                    height: 0,
                    ..Drawing::default()
                });
            },
            b"inline" | b"anchor" => {
                if let Some(open) = drawing.as_mut() {
                    open.inline = local == b"inline";
                }
            },
            b"extent" => {
                if let Some(open) = drawing.as_mut()
                    && open.extent.is_none()
                {
                    open.width = attribute(&element, b"cx")
                        .and_then(|value| value.parse().ok())
                        .unwrap_or(0);
                    open.height = attribute(&element, b"cy")
                        .and_then(|value| value.parse().ok())
                        .unwrap_or(0);
                    open.extent = Some(start..end);
                }
            },
            b"docPr" => {
                if let Some(open) = drawing.as_mut() {
                    open.name = attribute(&element, b"name").unwrap_or_default();
                    open.description = attribute(&element, b"descr").unwrap_or_default();
                }
            },
            b"graphicData" => {
                if let Some(open) = drawing.as_mut() {
                    open.uri = attribute(&element, b"uri");
                }
            },
            b"chart" => {
                if let Some(open) = drawing.as_mut() {
                    let uri = open.uri.as_deref().unwrap_or_default();
                    let extended = uri == litchi_drawingml::chart::chartex::GRAPHIC_DATA_URI;
                    if extended || uri == CHART_URI || uri == STRICT_CHART_URI {
                        open.extended = extended;
                        open.relationship_id = attribute(&element, b"id");
                    }
                }
            },
            _ => {},
        }
        if empty {
            if local == b"drawing"
                && let Some(open) = drawing.take()
            {
                finish(&mut found, &mut pending, open, start..end);
            }
        } else {
            stack.push((local, start, 0));
        }
    }
    Ok(found)
}

/// Record a closed drawing that holds a chart; its run, if any, is
/// resolved when the run closes.
fn finish(
    found: &mut SmallVec<[ChartSpan; 2]>,
    pending: &mut Vec<(usize, usize)>,
    open: Drawing,
    span: Range<usize>,
) {
    let Some(relationship_id) = open.relationship_id else {
        return;
    };
    if let Some((run, _)) = open.run {
        pending.push((found.len(), run));
    }
    found.push(ChartSpan {
        chart: InlineChart {
            relationship_id,
            name: open.name,
            description: open.description,
            width_emu: open.width,
            height_emu: open.height,
            extended: open.extended,
            inline: open.inline,
        },
        drawing: span,
        sole_run: None,
        extent: open.extent,
    });
}

/// Every chart of a paragraph's XML, in document order.
pub(crate) fn parse_inline_charts(xml: &[u8]) -> Result<SmallVec<[InlineChart; 2]>> {
    Ok(chart_spans(xml)?
        .into_iter()
        .map(|span| span.chart)
        .collect())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn chart_drawings_are_found_with_their_spans() {
        let xml = concat!(
            r#"<w:p xmlns:w="w" xmlns:wp="wp" xmlns:a="a" xmlns:c="c" xmlns:r="r">"#,
            r#"<w:r><w:t>x</w:t></w:r>"#,
            r#"<w:r><w:rPr/><w:drawing><wp:inline><wp:extent cx="100" cy="200"/><wp:docPr id="1" name="Sales &amp; costs" descr="d"/>"#,
            r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId7"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#,
            r#"<w:r><w:t>y</w:t><w:drawing><wp:anchor><wp:extent cx="5" cy="6"/><wp:docPr id="2" name="Ex"/>"#,
            r#"<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="cx" r:id="rId8"/></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#,
            r#"<w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="3" name="Pic"/><a:graphic><a:graphicData uri="pic"/></a:graphic></wp:inline></w:drawing></w:r>"#,
            "</w:p>"
        )
        .as_bytes();
        let spans = chart_spans(xml).unwrap();
        assert_eq!(spans.len(), 2);
        let first = &spans[0].chart;
        assert_eq!(
            (
                first.relationship_id(),
                first.name(),
                first.width_emu(),
                first.height_emu()
            ),
            ("rId7", "Sales & costs", 100, 200)
        );
        assert!(first.is_inline() && !first.is_extended());
        assert!(spans[0].sole_run.is_some());
        assert!(spans[1].chart.is_extended() && !spans[1].chart.is_inline());
        assert!(spans[1].sole_run.is_none());
        let run = spans[0].sole_run.clone().unwrap();
        assert!(xml[run.clone()].starts_with(b"<w:r><w:rPr/><w:drawing>"));
        assert!(xml[run].ends_with(b"</w:drawing></w:r>"));
        let extent = spans[0].extent.clone().unwrap();
        assert_eq!(&xml[extent], br#"<wp:extent cx="100" cy="200"/>"#);
        assert!(xml[spans[1].drawing.clone()].starts_with(b"<w:drawing>"));
    }
}
