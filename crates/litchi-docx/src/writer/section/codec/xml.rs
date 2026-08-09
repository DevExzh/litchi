#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::expect_used,
    reason = "the invariant is established immediately before extraction"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
use crate::error::{Error, Result};
use crate::header_footer::Kind;
use crate::section::Start;
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};
use std::fmt::Write;

use super::super::borders;
use super::super::model::{
    ChapterSep, Color, Display, Endnotes, Footnotes, GridType, LineNumberRestart,
    NoteNumberRestart, NotePos, OffsetFrom, PageNumberFormat, PageOrientation, SectionColumn,
    SectionColumns, SectionDocumentGrid, SectionHeaderFooterReference, SectionLineNumbering,
    SectionPageNumbering, SectionPaperSource, SectionProperties, SectionTextDirection,
    SectionVerticalAlignment, Style, ZOrder,
};
use super::package::write_references;
impl SectionProperties {
    pub(crate) fn from_xml(xml: &str) -> Result<Self> {
        let children = direct_children(xml)?;
        let mut properties = Self::default();
        let mut seen = std::collections::HashSet::new();
        let mut last_rank = 0u8;
        for (name, raw) in children {
            if let Some(rank) = section_child_rank(&name) {
                if rank < last_rank {
                    return Err(Error::InvalidFormat(format!(
                        "section property '{name}' is out of schema order"
                    )));
                }
                last_rank = rank;
            }
            if !seen.insert(name.clone())
                && !matches!(name.as_str(), "headerReference" | "footerReference")
            {
                return Err(Error::InvalidFormat(format!(
                    "section properties contain duplicate '{name}'"
                )));
            }
            match name.as_str() {
                "headerReference" => properties.headers.push(parse_header_footer(&raw)?),
                "footerReference" => properties.footers.push(parse_header_footer(&raw)?),
                "footnotePr" => properties.footnotes = Some(parse_footnotes(&raw)?),
                "endnotePr" => properties.endnotes = Some(parse_endnotes(&raw)?),
                "type" => {
                    let value = required_attr(&raw, b"val")?;
                    properties.start_type = Some(Start::from_xml(&value).ok_or_else(|| {
                        Error::InvalidFormat(format!("invalid section type '{value}'"))
                    })?);
                },
                "pgSz" => {
                    let attrs = attributes(&raw)?;
                    if let Some(value) = attr(&attrs, "w") {
                        properties.page_width = parse_u32(value, "page width")?;
                    }
                    if let Some(value) = attr(&attrs, "h") {
                        properties.page_height = parse_u32(value, "page height")?;
                    }
                    if let Some(value) = attr(&attrs, "orient") {
                        properties.orientation = PageOrientation::parse(value)?;
                    }
                },
                "pgMar" => {
                    let attrs = attributes(&raw)?;
                    assign_u32(&attrs, "top", &mut properties.margin_top)?;
                    assign_u32(&attrs, "bottom", &mut properties.margin_bottom)?;
                    assign_u32(&attrs, "left", &mut properties.margin_left)?;
                    assign_u32(&attrs, "right", &mut properties.margin_right)?;
                    assign_u32(&attrs, "header", &mut properties.header_distance)?;
                    assign_u32(&attrs, "footer", &mut properties.footer_distance)?;
                    assign_u32(&attrs, "gutter", &mut properties.gutter)?;
                },
                "pgNumType" => properties.page_numbering = Some(parse_page_numbering(&raw)?),
                "paperSrc" => {
                    let attrs = attributes(&raw)?;
                    properties.paper_source = Some(SectionPaperSource {
                        first: attr(&attrs, "first")
                            .map(|value| parse_u32(value, "first paper source"))
                            .transpose()?,
                        other: attr(&attrs, "other")
                            .map(|value| parse_u32(value, "other paper source"))
                            .transpose()?,
                    });
                },
                "pgBorders" => properties.page_borders = Some(parse_page_borders(&raw)?),
                "lnNumType" => {
                    properties.line_numbering = Some(parse_line_numbering(&raw)?);
                },
                "cols" => properties.columns = Some(parse_columns(&raw)?),
                "formProt" => properties.form_protection = parse_on_off(&raw)?,
                "vAlign" => {
                    properties.vertical_alignment = Some(SectionVerticalAlignment::parse(
                        &required_attr(&raw, b"val")?,
                    )?);
                },
                "titlePg" => properties.title_page = parse_on_off(&raw)?,
                "textDirection" => {
                    properties.text_direction =
                        Some(SectionTextDirection::parse(&required_attr(&raw, b"val")?)?);
                },
                "bidi" => properties.bidirectional = parse_on_off(&raw)?,
                "rtlGutter" => properties.rtl_gutter = parse_on_off(&raw)?,
                "docGrid" => properties.document_grid = Some(parse_grid(&raw)?),
                "printerSettings" => {
                    properties.printer_settings_relationship_id = Some(required_attr(&raw, b"id")?);
                },
                _ => properties.preserved_unknown_children.push(raw),
            }
        }
        properties.validate()?;
        Ok(properties)
    }

    pub(crate) fn write_xml(
        &self,
        xml: &mut String,
        rels: Option<&super::super::super::relmap::RelationshipMapper>,
    ) -> Result<()> {
        self.validate()?;
        xml.push_str("<w:sectPr>");
        write_references(xml, "headerReference", &self.headers, rels, true)?;
        write_references(xml, "footerReference", &self.footers, rels, false)?;
        if let Some(note) = &self.footnotes {
            write_footnotes(xml, note)?;
        } else if rels.is_some_and(|rels| rels.get_footnotes_id().is_some()) {
            xml.push_str("<w:footnotePr><w:numFmt w:val=\"decimal\"/></w:footnotePr>");
        }
        if let Some(note) = &self.endnotes {
            write_endnotes(xml, note)?;
        } else if rels.is_some_and(|rels| rels.get_endnotes_id().is_some()) {
            xml.push_str("<w:endnotePr><w:numFmt w:val=\"decimal\"/></w:endnotePr>");
        }
        if let Some(start_type) = self.start_type {
            write!(xml, "<w:type w:val=\"{}\"/>", start_type.to_xml())
                .map_err(|error| Error::Xml(error.to_string()))?;
        }
        write!(
            xml,
            "<w:pgSz w:w=\"{}\" w:h=\"{}\" w:orient=\"{}\"/>",
            self.page_width,
            self.page_height,
            self.orientation.as_str()
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
        write!(
            xml,
            "<w:pgMar w:top=\"{}\" w:right=\"{}\" w:bottom=\"{}\" w:left=\"{}\" w:header=\"{}\" w:footer=\"{}\" w:gutter=\"{}\"/>",
            self.margin_top,
            self.margin_right,
            self.margin_bottom,
            self.margin_left,
            self.header_distance,
            self.footer_distance,
            self.gutter
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
        if let Some(paper_source) = &self.paper_source {
            xml.push_str("<w:paperSrc");
            if let Some(first) = paper_source.first {
                write!(xml, " w:first=\"{first}\"")
                    .map_err(|error| Error::Xml(error.to_string()))?;
            }
            if let Some(other) = paper_source.other {
                write!(xml, " w:other=\"{other}\"")
                    .map_err(|error| Error::Xml(error.to_string()))?;
            }
            xml.push_str("/>");
        }
        if let Some(borders) = &self.page_borders {
            write_page_borders(xml, borders)?;
        }
        if let Some(numbering) = &self.line_numbering {
            write_line_numbering(xml, numbering)?;
        }
        if let Some(numbering) = &self.page_numbering {
            write_page_numbering(xml, numbering)?;
        }
        if let Some(columns) = &self.columns {
            write_columns(xml, columns)?;
        }
        if self.form_protection {
            xml.push_str("<w:formProt/>");
        }
        if let Some(alignment) = self.vertical_alignment {
            write!(xml, "<w:vAlign w:val=\"{}\"/>", alignment.as_str())
                .map_err(|error| Error::Xml(error.to_string()))?;
        }
        if self.title_page {
            xml.push_str("<w:titlePg/>");
        }
        if let Some(direction) = self.text_direction {
            write!(xml, "<w:textDirection w:val=\"{}\"/>", direction.as_str())
                .map_err(|error| Error::Xml(error.to_string()))?;
        }
        if self.bidirectional {
            xml.push_str("<w:bidi/>");
        }
        if self.rtl_gutter {
            xml.push_str("<w:rtlGutter/>");
        }
        if let Some(grid) = &self.document_grid {
            write_grid(xml, grid)?;
        }
        if let Some(id) = &self.printer_settings_relationship_id {
            write!(xml, "<w:printerSettings r:id=\"{}\"/>", escape(id))
                .map_err(|error| Error::Xml(error.to_string()))?;
        }
        for child in &self.preserved_unknown_children {
            xml.push_str(child);
        }
        xml.push_str("</w:sectPr>");
        Ok(())
    }
}

fn section_child_rank(name: &str) -> Option<u8> {
    match name {
        "headerReference" => Some(0),
        "footerReference" => Some(1),
        "footnotePr" => Some(2),
        "endnotePr" => Some(3),
        "type" => Some(4),
        "pgSz" => Some(5),
        "pgMar" => Some(6),
        "paperSrc" => Some(7),
        "pgBorders" => Some(8),
        "lnNumType" => Some(9),
        "pgNumType" => Some(10),
        "cols" => Some(11),
        "formProt" => Some(12),
        "vAlign" => Some(13),
        "titlePg" => Some(14),
        "textDirection" => Some(15),
        "bidi" => Some(16),
        "rtlGutter" => Some(17),
        "docGrid" => Some(18),
        "printerSettings" => Some(19),
        _ => None,
    }
}

pub(super) fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
}

fn direct_children(xml: &str) -> Result<Vec<(String, String)>> {
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut child: Option<(String, usize, usize)> = None;
    let mut children = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_source_error| Error::InvalidFormat("section XML offset overflow".into()))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_source_error| Error::InvalidFormat("section XML offset overflow".into()))?;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if root_seen || element.local_name().as_ref() != b"sectPr" {
                        return Err(Error::InvalidFormat(
                            "section properties have an invalid root".into(),
                        ));
                    }
                    root_seen = true;
                } else if depth == 1 {
                    child = Some((
                        String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
                        start,
                        1,
                    ));
                } else if let Some((_, _, child_depth)) = child.as_mut() {
                    *child_depth += 1;
                }
                depth += 1;
            },
            Event::Empty(element) if depth == 1 => {
                let name = String::from_utf8_lossy(element.local_name().as_ref()).into_owned();
                children.push((name, xml[start..end].to_string()));
            },
            Event::End(_) => {
                if let Some((_, _, child_depth)) = child.as_mut() {
                    *child_depth -= 1;
                    if *child_depth == 0 {
                        let (name, child_start, _) = child.take().expect("present");
                        children.push((name, xml[child_start..end].to_string()));
                    }
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid section XML nesting".into()))?;
            },
            Event::Eof => break,
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !root_seen || depth != 0 {
        return Err(Error::InvalidFormat(
            "unterminated section properties".into(),
        ));
    }
    Ok(children)
}

fn attributes(xml: &str) -> Result<Vec<(String, String)>> {
    let mut reader = Reader::from_str(xml);
    let element = loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(element) | Event::Empty(element) => break element,
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "section property has no element".into(),
                ));
            },
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    };
    let mut result = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
        if result.iter().any(|(candidate, _)| candidate == &name) {
            return Err(Error::InvalidFormat(format!(
                "duplicate section property attribute '{name}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        result.push((name, value));
    }
    Ok(result)
}

fn attr<'a>(attrs: &'a [(String, String)], name: &str) -> Option<&'a str> {
    attrs
        .iter()
        .find_map(|(candidate, value)| (candidate == name).then_some(value.as_str()))
}

fn required_attr(xml: &str, name: &[u8]) -> Result<String> {
    let name = String::from_utf8_lossy(name);
    let attrs = attributes(xml)?;
    attr(&attrs, &name)
        .map(ToOwned::to_owned)
        .ok_or_else(|| Error::InvalidFormat(format!("missing section attribute '{name}'")))
}

fn parse_u32(value: &str, description: &str) -> Result<u32> {
    value.parse().map_err(|_source_error| {
        Error::InvalidFormat(format!("invalid {description} value '{value}'"))
    })
}

fn assign_u32(attrs: &[(String, String)], name: &str, slot: &mut u32) -> Result<()> {
    if let Some(value) = attr(attrs, name) {
        *slot = parse_u32(value, name)?;
    }
    Ok(())
}

fn parse_header_footer(xml: &str) -> Result<SectionHeaderFooterReference> {
    let kind = Kind::from_xml(&required_attr(xml, b"type")?)
        .ok_or_else(|| Error::InvalidFormat("invalid section header/footer type".to_string()))?;
    Ok(SectionHeaderFooterReference {
        kind,
        relationship_id: Some(required_attr(xml, b"id")?),
        part: None,
    })
}

fn parse_page_numbering(xml: &str) -> Result<SectionPageNumbering> {
    let attrs = attributes(xml)?;
    Ok(SectionPageNumbering {
        format: attr(&attrs, "fmt")
            .map(PageNumberFormat::parse)
            .transpose()?
            .unwrap_or(PageNumberFormat::Decimal),
        start: attr(&attrs, "start")
            .map(|value| parse_u32(value, "page number start"))
            .transpose()?,
        chapter_style: attr(&attrs, "chapStyle")
            .map(|value| {
                value
                    .parse::<u8>()
                    .map_err(|_source_error| Error::InvalidFormat("invalid chapter style".into()))
            })
            .transpose()?,
        chapter_separator: attr(&attrs, "chapSep").map(ChapterSep::parse).transpose()?,
    })
}

fn parse_columns(xml: &str) -> Result<SectionColumns> {
    let attrs = attributes(xml)?;
    let mut columns = SectionColumns {
        equal_width: attr(&attrs, "equalWidth")
            .is_none_or(|value| value != "0" && value != "false"),
        count: attr(&attrs, "num")
            .map(|value| {
                value.parse::<u16>().map_err(|_source_error| {
                    Error::InvalidFormat("invalid section column count".into())
                })
            })
            .transpose()?
            .unwrap_or(1),
        space: attr(&attrs, "space")
            .map(|value| parse_u32(value, "column space"))
            .transpose()?,
        separator: attr(&attrs, "sep").is_some_and(|value| value == "1" || value == "true"),
        columns: Vec::new(),
    };
    for (name, raw) in direct_nested_children(xml)? {
        if name != "col" {
            return Err(Error::InvalidFormat(format!(
                "invalid child '{name}' in section columns"
            )));
        }
        let attrs = attributes(&raw)?;
        columns.columns.push(SectionColumn {
            width: parse_u32(
                attr(&attrs, "w")
                    .ok_or_else(|| Error::InvalidFormat("section column omits width".into()))?,
                "column width",
            )?,
            space: attr(&attrs, "space")
                .map(|value| parse_u32(value, "column space"))
                .transpose()?,
        });
    }
    Ok(columns)
}

fn direct_nested_children(xml: &str) -> Result<Vec<(String, String)>> {
    let open_end = xml
        .find('>')
        .ok_or_else(|| Error::InvalidFormat("invalid section property".into()))?;
    let close = xml.rfind("</").unwrap_or(xml.len());
    if close <= open_end + 1 {
        return Ok(Vec::new());
    }
    direct_children(&format!(
        "<w:sectPr>{}</w:sectPr>",
        &xml[open_end + 1..close]
    ))
}

#[derive(Debug, Clone, Copy)]
struct ParsedNote<P> {
    format: PageNumberFormat,
    start: Option<u32>,
    restart: Option<NoteNumberRestart>,
    position: Option<P>,
}

fn parse_note_properties<P: NotePos>(xml: &str) -> Result<ParsedNote<P>> {
    let mut result = ParsedNote {
        format: PageNumberFormat::Decimal,
        start: None,
        restart: None,
        position: None,
    };
    for (name, raw) in direct_nested_children(xml)? {
        let value = required_attr(&raw, b"val")?;
        match name.as_str() {
            "numFmt" => result.format = PageNumberFormat::parse(&value)?,
            "numStart" => result.start = Some(parse_u32(&value, "note number start")?),
            "numRestart" => result.restart = Some(NoteNumberRestart::parse(&value)?),
            "pos" => result.position = Some(P::parse(&value)?),
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "invalid note property '{name}'"
                )));
            },
        }
    }
    Ok(result)
}

fn parse_footnotes(xml: &str) -> Result<Footnotes> {
    let parsed = parse_note_properties(xml)?;
    Ok(Footnotes {
        format: parsed.format,
        start: parsed.start,
        restart: parsed.restart,
        position: parsed.position,
    })
}

fn parse_endnotes(xml: &str) -> Result<Endnotes> {
    let parsed = parse_note_properties(xml)?;
    Ok(Endnotes {
        format: parsed.format,
        start: parsed.start,
        restart: parsed.restart,
        position: parsed.position,
    })
}

fn parse_grid(xml: &str) -> Result<SectionDocumentGrid> {
    let attrs = attributes(xml)?;
    Ok(SectionDocumentGrid {
        grid_type: attr(&attrs, "type")
            .map(GridType::parse)
            .transpose()?
            .unwrap_or(GridType::Default),
        line_pitch: attr(&attrs, "linePitch")
            .map(|value| parse_u32(value, "grid line pitch"))
            .transpose()?,
        char_space: attr(&attrs, "charSpace")
            .map(|value| {
                value.parse::<i32>().map_err(|_source_error| {
                    Error::InvalidFormat("invalid grid character space".into())
                })
            })
            .transpose()?,
    })
}

fn parse_page_borders(xml: &str) -> Result<borders::Borders> {
    let attrs = attributes(xml)?;
    let mut borders = borders::Borders {
        offset_from: attr(&attrs, "offsetFrom")
            .map(OffsetFrom::parse)
            .transpose()?
            .unwrap_or(OffsetFrom::Page),
        z_order: attr(&attrs, "zOrder")
            .map(ZOrder::parse)
            .transpose()?
            .unwrap_or(ZOrder::Back),
        display: attr(&attrs, "display")
            .map(Display::parse)
            .transpose()?
            .unwrap_or(Display::AllPages),
        ..borders::Borders::default()
    };
    for (name, raw) in direct_nested_children(xml)? {
        let edge = match name.as_str() {
            "top" => &mut borders.top,
            "left" => &mut borders.left,
            "bottom" => &mut borders.bottom,
            "right" => &mut borders.right,
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "invalid child '{name}' in section page borders"
                )));
            },
        };
        if edge.is_some() {
            return Err(Error::InvalidFormat(format!(
                "duplicate '{name}' page border edge"
            )));
        }
        *edge = Some(parse_page_border(&raw)?);
    }
    Ok(borders)
}

fn parse_page_border(xml: &str) -> Result<borders::Border> {
    let attrs = attributes(xml)?;
    let on_off =
        |name: &str| attr(&attrs, name).is_some_and(|value| matches!(value, "1" | "true" | "on"));
    Ok(borders::Border {
        style: Style::parse(
            attr(&attrs, "val")
                .ok_or_else(|| Error::InvalidFormat("page border omits style".into()))?,
        )?,
        size: attr(&attrs, "sz")
            .map(|value| parse_u32(value, "page border size"))
            .transpose()?,
        space: attr(&attrs, "space")
            .map(|value| parse_u32(value, "page border space"))
            .transpose()?,
        color: attr(&attrs, "color").map(Color::parse).transpose()?,
        shadow: on_off("shadow"),
        frame: on_off("frame"),
    })
}

fn parse_line_numbering(xml: &str) -> Result<SectionLineNumbering> {
    let attrs = attributes(xml)?;
    Ok(SectionLineNumbering {
        count_by: attr(&attrs, "countBy")
            .map(|value| parse_u32(value, "line-number increment"))
            .transpose()?,
        start: attr(&attrs, "start")
            .map(|value| parse_u32(value, "line-number start"))
            .transpose()?,
        distance: attr(&attrs, "distance")
            .map(|value| parse_u32(value, "line-number distance"))
            .transpose()?,
        restart: attr(&attrs, "restart")
            .map(LineNumberRestart::parse)
            .transpose()?,
    })
}

fn parse_on_off(xml: &str) -> Result<bool> {
    let attrs = attributes(xml)?;
    Ok(attr(&attrs, "val").is_none_or(|value| matches!(value, "1" | "true" | "on")))
}

fn write_note_properties<P: NotePos>(
    xml: &mut String,
    element: &str,
    format: PageNumberFormat,
    start: Option<u32>,
    restart: Option<NoteNumberRestart>,
    position: Option<P>,
) -> Result<()> {
    write!(
        xml,
        "<w:{element}><w:numFmt w:val=\"{}\"/>",
        format.as_str()
    )
    .map_err(|error| Error::Xml(error.to_string()))?;
    if let Some(start) = start {
        write!(xml, "<w:numStart w:val=\"{start}\"/>")
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(restart) = restart {
        write!(xml, "<w:numRestart w:val=\"{}\"/>", restart.as_str())
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(position) = position {
        write!(xml, "<w:pos w:val=\"{}\"/>", position.as_str())
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    write!(xml, "</w:{element}>").map_err(|error| Error::Xml(error.to_string()))
}

fn write_footnotes(xml: &mut String, note: &Footnotes) -> Result<()> {
    write_note_properties(
        xml,
        "footnotePr",
        note.format,
        note.start,
        note.restart,
        note.position,
    )
}

fn write_endnotes(xml: &mut String, note: &Endnotes) -> Result<()> {
    write_note_properties(
        xml,
        "endnotePr",
        note.format,
        note.start,
        note.restart,
        note.position,
    )
}

fn write_page_numbering(xml: &mut String, numbering: &SectionPageNumbering) -> Result<()> {
    write!(xml, "<w:pgNumType w:fmt=\"{}\"", numbering.format.as_str())
        .map_err(|error| Error::Xml(error.to_string()))?;
    if let Some(start) = numbering.start {
        write!(xml, " w:start=\"{start}\"").map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(style) = numbering.chapter_style {
        write!(xml, " w:chapStyle=\"{style}\"").map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(separator) = numbering.chapter_separator {
        write!(xml, " w:chapSep=\"{}\"", separator.as_str())
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    xml.push_str("/>");
    Ok(())
}

fn write_columns(xml: &mut String, columns: &SectionColumns) -> Result<()> {
    write!(
        xml,
        "<w:cols w:equalWidth=\"{}\" w:num=\"{}\"",
        i32::from(columns.equal_width),
        columns.count
    )
    .map_err(|error| Error::Xml(error.to_string()))?;
    if let Some(space) = columns.space {
        write!(xml, " w:space=\"{space}\"").map_err(|error| Error::Xml(error.to_string()))?;
    }
    if columns.separator {
        xml.push_str(" w:sep=\"1\"");
    }
    if columns.columns.is_empty() {
        xml.push_str("/>");
    } else {
        xml.push('>');
        for column in &columns.columns {
            write!(xml, "<w:col w:w=\"{}\"", column.width)
                .map_err(|error| Error::Xml(error.to_string()))?;
            if let Some(space) = column.space {
                write!(xml, " w:space=\"{space}\"")
                    .map_err(|error| Error::Xml(error.to_string()))?;
            }
            xml.push_str("/>");
        }
        xml.push_str("</w:cols>");
    }
    Ok(())
}

fn write_grid(xml: &mut String, grid: &SectionDocumentGrid) -> Result<()> {
    write!(xml, "<w:docGrid w:type=\"{}\"", grid.grid_type.as_str())
        .map_err(|error| Error::Xml(error.to_string()))?;
    if let Some(pitch) = grid.line_pitch {
        write!(xml, " w:linePitch=\"{pitch}\"").map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(space) = grid.char_space {
        write!(xml, " w:charSpace=\"{space}\"").map_err(|error| Error::Xml(error.to_string()))?;
    }
    xml.push_str("/>");
    Ok(())
}

fn write_page_borders(xml: &mut String, borders: &borders::Borders) -> Result<()> {
    write!(
        xml,
        "<w:pgBorders w:offsetFrom=\"{}\" w:zOrder=\"{}\" w:display=\"{}\"",
        borders.offset_from.as_str(),
        borders.z_order.as_str(),
        borders.display.as_str()
    )
    .map_err(|error| Error::Xml(error.to_string()))?;
    let edges = [
        ("top", &borders.top),
        ("left", &borders.left),
        ("bottom", &borders.bottom),
        ("right", &borders.right),
    ];
    if edges.iter().all(|(_, edge)| edge.is_none()) {
        xml.push_str("/>");
        return Ok(());
    }
    xml.push('>');
    for (name, edge) in edges {
        if let Some(border) = edge {
            write_page_border(xml, name, border)?;
        }
    }
    xml.push_str("</w:pgBorders>");
    Ok(())
}

fn write_page_border(xml: &mut String, name: &str, border: &borders::Border) -> Result<()> {
    write!(xml, "<w:{name} w:val=\"{}\"", escape(border.style.as_str()))
        .map_err(|error| Error::Xml(error.to_string()))?;
    if let Some(size) = border.size {
        write!(xml, " w:sz=\"{size}\"").map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(space) = border.space {
        write!(xml, " w:space=\"{space}\"").map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(color) = border.color {
        match color {
            Color::Auto => xml.push_str(" w:color=\"auto\""),
            Color::Rgb([red, green, blue]) => {
                write!(xml, " w:color=\"{red:02X}{green:02X}{blue:02X}\"")
                    .map_err(|error| Error::Xml(error.to_string()))?;
            },
        }
    }
    if border.shadow {
        xml.push_str(" w:shadow=\"1\"");
    }
    if border.frame {
        xml.push_str(" w:frame=\"1\"");
    }
    xml.push_str("/>");
    Ok(())
}

fn write_line_numbering(xml: &mut String, numbering: &SectionLineNumbering) -> Result<()> {
    xml.push_str("<w:lnNumType");
    if let Some(count_by) = numbering.count_by {
        write!(xml, " w:countBy=\"{count_by}\"").map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(start) = numbering.start {
        write!(xml, " w:start=\"{start}\"").map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(distance) = numbering.distance {
        write!(xml, " w:distance=\"{distance}\"").map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(restart) = numbering.restart {
        write!(xml, " w:restart=\"{}\"", restart.as_str())
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    xml.push_str("/>");
    Ok(())
}
