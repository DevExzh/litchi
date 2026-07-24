//! Authoring model and serializer for XLSX DrawingML shapes and text boxes.
//!
//! [`XlsxShapeSpec`] describes one text-box shape to embed in a worksheet
//! drawing part: a preset geometry, an anchor on the sheet grid, text-body
//! properties, and rich-text runs. It deliberately reuses the typed read
//! model from [`crate::xlsx::shapes`] (`XlsxShapePreset`, `XlsxShapeAnchor`,
//! `XlsxShapeBodyProperties`, paragraphs and runs) so anything authored here
//! round-trips through the shape inventory with identical semantics.
//!
//! [`write_shape_anchor_xml`] serializes one spec as a single
//! `xdr:twoCellAnchor`/`xdr:oneCellAnchor`/`xdr:absoluteAnchor` element for
//! the worksheet drawing part emitted by
//! `MutableWorksheet::generate_drawing_xml`. Everything is inert: no
//! rendering, no layout computation, and all inputs are bounded and
//! validated at authoring time.

use std::fmt::Write as _;

use litchi_core::xml::escape::escape_xml;

use crate::xlsx::shapes::{
    XlsxEmuExtent, XlsxEmuOffset, XlsxEditAs, XlsxShapeAnchor, XlsxShapeBodyProperties,
    XlsxShapeParagraph, XlsxShapePreset, XlsxTextAutofit, XlsxTextDirection,
    XlsxTextVerticalAnchor, XlsxTextWrap,
};

/// Maximum number of authored shapes per worksheet.
const MAX_SHAPES_PER_WORKSHEET: usize = 4096;
/// Maximum aggregate run text bytes across one authored shape.
const MAX_SHAPE_TEXT_BYTES: usize = 1024 * 1024;
/// Maximum length of the shape name or description.
const MAX_SHAPE_NAME_BYTES: usize = 4096;
/// Maximum paragraphs in one authored shape.
const MAX_SHAPE_PARAGRAPHS: usize = 16_384;
/// Maximum runs in one authored paragraph.
const MAX_RUNS_PER_PARAGRAPH: usize = 4096;

/// One authored DrawingML text-box shape for a worksheet drawing part.
///
/// Construct with [`XlsxShapeSpec::text_box`] (or [`XlsxShapeSpec::shape`]
/// for a non-text-box shape), then adjust `body_properties`, `paragraphs`,
/// or flags before handing it to `MutableWorksheet::add_shape`.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxShapeSpec {
    /// Shape name written to `xdr:cNvPr@name`.
    pub name: String,
    /// Alternative text (`xdr:cNvPr@descr`), when set.
    pub description: Option<String>,
    /// Whether the shape is hidden (`xdr:cNvPr@hidden`).
    pub hidden: bool,
    /// Whether the shape is marked as a text box (`xdr:cNvSpPr@txBox`).
    pub is_text_box: bool,
    /// How the shape is anchored on the worksheet grid.
    pub anchor: XlsxShapeAnchor,
    /// Preset geometry (`a:prstGeom@prst`).
    pub preset: XlsxShapePreset,
    /// Text-body properties (`a:bodyPr`).
    pub body_properties: XlsxShapeBodyProperties,
    /// The text story as paragraphs with runs.
    pub paragraphs: Vec<XlsxShapeParagraph>,
}

impl XlsxShapeSpec {
    /// A text-box shape with the given preset, anchor, and plain-text story.
    ///
    /// `text` is split into paragraphs on `\n`; each paragraph becomes one
    /// unformatted run. Body properties start at the ECMA-376 defaults.
    pub fn text_box(
        name: impl Into<String>,
        anchor: XlsxShapeAnchor,
        preset: XlsxShapePreset,
        text: &str,
    ) -> Self {
        Self {
            is_text_box: true,
            ..Self::shape(name, anchor, preset, text)
        }
    }

    /// A plain (non-text-box) shape with the given preset, anchor, and text.
    pub fn shape(
        name: impl Into<String>,
        anchor: XlsxShapeAnchor,
        preset: XlsxShapePreset,
        text: &str,
    ) -> Self {
        let paragraphs = text
            .split('\n')
            .map(|line| XlsxShapeParagraph {
                runs: vec![crate::xlsx::shapes::XlsxShapeRun {
                    text: line.to_string(),
                    ..crate::xlsx::shapes::XlsxShapeRun::default()
                }],
            })
            .collect();
        Self {
            name: name.into(),
            description: None,
            hidden: false,
            is_text_box: false,
            anchor,
            preset,
            body_properties: XlsxShapeBodyProperties::default(),
            paragraphs,
        }
    }

    /// Validate the spec against worksheet bounds and the module limits.
    pub(crate) fn validate(&self, existing: usize) -> Result<(), String> {
        if existing >= MAX_SHAPES_PER_WORKSHEET {
            return Err("worksheet shape count limit exceeded".to_string());
        }
        if self.name.is_empty() {
            return Err("shape name cannot be empty".to_string());
        }
        if self.name.len() > MAX_SHAPE_NAME_BYTES
            || self
                .description
                .as_ref()
                .is_some_and(|d| d.len() > MAX_SHAPE_NAME_BYTES)
        {
            return Err("shape name/description is too long".to_string());
        }
        if self.paragraphs.len() > MAX_SHAPE_PARAGRAPHS {
            return Err("shape paragraph count limit exceeded".to_string());
        }
        let mut text_bytes = 0usize;
        for paragraph in &self.paragraphs {
            if paragraph.runs.len() > MAX_RUNS_PER_PARAGRAPH {
                return Err("shape run count limit exceeded".to_string());
            }
            text_bytes += paragraph.runs.iter().map(|run| run.text.len()).sum::<usize>();
        }
        if text_bytes > MAX_SHAPE_TEXT_BYTES {
            return Err("shape text bytes limit exceeded".to_string());
        }
        validate_anchor(&self.anchor)
    }
}

/// Validate anchor markers against worksheet bounds (and ordering for
/// two-cell anchors), mirroring the checks applied to images and charts.
fn validate_anchor(anchor: &XlsxShapeAnchor) -> Result<(), String> {
    const MAX_COLUMNS: u32 = 16_384;
    const MAX_ROWS: u32 = 1_048_576;
    let markers: &[crate::xlsx::shapes::XlsxCellMarker] = match anchor {
        XlsxShapeAnchor::TwoCell { from, to, .. } => {
            if to.row < from.row || to.column < from.column {
                return Err("shape anchor cannot be descending".to_string());
            }
            &[*from, *to]
        },
        XlsxShapeAnchor::OneCell { from, .. } => &[*from],
        XlsxShapeAnchor::Absolute { .. } => &[],
    };
    for marker in markers {
        if marker.column >= MAX_COLUMNS || marker.row >= MAX_ROWS {
            return Err("shape anchor exceeds worksheet bounds".to_string());
        }
    }
    Ok(())
}

/// Serialize one authored shape as a complete anchor element of a worksheet
/// drawing part.
///
/// `id` is the drawing-wide unique object ID written to `xdr:cNvPr@id`; the
/// caller is responsible for keeping it unique across pictures, charts, and
/// shapes in the same drawing part.
pub(crate) fn write_shape_anchor_xml(
    xml: &mut String,
    spec: &XlsxShapeSpec,
    id: u32,
) -> Result<(), String> {
    write_anchor_open(xml, &spec.anchor);
    write_shape_xml(xml, spec, id)?;
    xml.push_str("<xdr:clientData/>");
    match spec.anchor {
        XlsxShapeAnchor::TwoCell { .. } => xml.push_str("</xdr:twoCellAnchor>"),
        XlsxShapeAnchor::OneCell { .. } => xml.push_str("</xdr:oneCellAnchor>"),
        XlsxShapeAnchor::Absolute { .. } => xml.push_str("</xdr:absoluteAnchor>"),
    }
    Ok(())
}

fn write_anchor_open(xml: &mut String, anchor: &XlsxShapeAnchor) {
    match anchor {
        XlsxShapeAnchor::TwoCell { from, to, edit_as } => {
            match edit_as {
                // The ECMA-376 default; omitted to keep output canonical.
                XlsxEditAs::TwoCell => xml.push_str("<xdr:twoCellAnchor>"),
                XlsxEditAs::OneCell => xml.push_str(r#"<xdr:twoCellAnchor editAs="oneCell">"#),
                XlsxEditAs::Absolute => xml.push_str(r#"<xdr:twoCellAnchor editAs="absolute">"#),
            }
            write_marker(xml, "from", from);
            write_marker(xml, "to", to);
        },
        XlsxShapeAnchor::OneCell { from, extent } => {
            xml.push_str("<xdr:oneCellAnchor>");
            write_marker(xml, "from", from);
            write_extent(xml, extent);
        },
        XlsxShapeAnchor::Absolute { position, extent } => {
            xml.push_str("<xdr:absoluteAnchor>");
            write_position(xml, position);
            write_extent(xml, extent);
        },
    }
}

fn write_marker(xml: &mut String, name: &str, marker: &crate::xlsx::shapes::XlsxCellMarker) {
    let _ = write!(
        xml,
        "<xdr:{name}><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff>\
         <xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:{name}>",
        marker.column,
        marker.column_offset.emu(),
        marker.row,
        marker.row_offset.emu()
    );
}

fn write_extent(xml: &mut String, extent: &XlsxEmuExtent) {
    let _ = write!(
        xml,
        r#"<xdr:ext cx="{}" cy="{}"/>"#,
        extent.width.emu(),
        extent.height.emu()
    );
}

fn write_position(xml: &mut String, position: &XlsxEmuOffset) {
    let _ = write!(
        xml,
        r#"<xdr:pos x="{}" y="{}"/>"#,
        position.x.emu(),
        position.y.emu()
    );
}

fn write_shape_xml(xml: &mut String, spec: &XlsxShapeSpec, id: u32) -> Result<(), String> {
    let _ = write!(
        xml,
        r#"<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{id}" name="{}""#,
        escape_xml(&spec.name)
    );
    if let Some(description) = &spec.description {
        let _ = write!(xml, r#" descr="{}""#, escape_xml(description));
    }
    if spec.hidden {
        xml.push_str(r#" hidden="1""#);
    }
    xml.push_str("/>");
    if spec.is_text_box {
        xml.push_str(r#"<xdr:cNvSpPr txBox="1"/>"#);
    } else {
        xml.push_str("<xdr:cNvSpPr/>");
    }
    xml.push_str("</xdr:nvSpPr><xdr:spPr>");
    xml.push_str(r#"<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>"#);
    let _ = write!(
        xml,
        r#"<a:prstGeom prst="{}"><a:avLst/></a:prstGeom>"#,
        escape_xml(spec.preset.as_str())
    );
    xml.push_str("</xdr:spPr><xdr:txBody>");
    write_body_properties(xml, &spec.body_properties);
    xml.push_str("<a:lstStyle/>");
    for paragraph in &spec.paragraphs {
        xml.push_str("<a:p>");
        for run in &paragraph.runs {
            xml.push_str("<a:r>");
            write_run_properties(xml, run);
            let _ = write!(xml, "<a:t>{}</a:t>", escape_xml(&run.text));
            xml.push_str("</a:r>");
        }
        xml.push_str("</a:p>");
    }
    xml.push_str("</xdr:txBody></xdr:sp>");
    Ok(())
}

fn write_body_properties(xml: &mut String, body: &XlsxShapeBodyProperties) {
    xml.push_str("<a:bodyPr");
    let _ = write!(
        xml,
        r#" lIns="{}" tIns="{}" rIns="{}" bIns="{}""#,
        body.insets.left.emu(),
        body.insets.top.emu(),
        body.insets.right.emu(),
        body.insets.bottom.emu()
    );
    if body.vertical_anchor != XlsxTextVerticalAnchor::Top {
        let token = match body.vertical_anchor {
            XlsxTextVerticalAnchor::Top => unreachable!(),
            XlsxTextVerticalAnchor::Center => "ctr",
            XlsxTextVerticalAnchor::Bottom => "b",
            XlsxTextVerticalAnchor::Justified => "just",
            XlsxTextVerticalAnchor::Distributed => "dist",
        };
        let _ = write!(xml, r#" anchor="{token}""#);
    }
    if body.anchor_center {
        xml.push_str(r#" anchorCtr="1""#);
    }
    if body.direction != XlsxTextDirection::Horizontal {
        let token = match body.direction {
            XlsxTextDirection::Horizontal => unreachable!(),
            XlsxTextDirection::Vertical => "vert",
            XlsxTextDirection::Vertical270 => "vert270",
            XlsxTextDirection::WordArtVertical => "wordArtVert",
            XlsxTextDirection::EastAsianVertical => "eaVert",
            XlsxTextDirection::MongolianVertical => "mongolianVert",
            XlsxTextDirection::WordArtVerticalRtl => "wordArtVertRtl",
        };
        let _ = write!(xml, r#" vert="{token}""#);
    }
    if body.wrap == XlsxTextWrap::None {
        xml.push_str(r#" wrap="none""#);
    }
    if body.column_count != 1 {
        let _ = write!(xml, r#" numCol="{}""#, body.column_count);
    }
    if body.space_first_last_paragraph {
        xml.push_str(r#" spcFirstLastPara="1""#);
    }
    match body.autofit {
        XlsxTextAutofit::NoAutofit => xml.push_str("><a:noAutofit/></a:bodyPr>"),
        XlsxTextAutofit::ShapeAutofit => xml.push_str("><a:spAutoFit/></a:bodyPr>"),
        XlsxTextAutofit::NormalAutofit => xml.push_str("><a:normAutofit/></a:bodyPr>"),
    }
}

fn write_run_properties(xml: &mut String, run: &crate::xlsx::shapes::XlsxShapeRun) {
    if run.bold.is_none()
        && run.italic.is_none()
        && run.underline.is_none()
        && run.font_size_hundredths.is_none()
    {
        xml.push_str("<a:rPr/>");
        return;
    }
    xml.push_str("<a:rPr");
    if let Some(size) = run.font_size_hundredths {
        let _ = write!(xml, r#" sz="{size}""#);
    }
    if let Some(bold) = run.bold {
        xml.push_str(if bold { r#" b="1""# } else { r#" b="0""# });
    }
    if let Some(italic) = run.italic {
        xml.push_str(if italic { r#" i="1""# } else { r#" i="0""# });
    }
    if let Some(underline) = run.underline {
        xml.push_str(if underline { r#" u="sng""# } else { r#" u="none""# });
    }
    xml.push_str("/>");
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::shapes::{
        XlsxCellMarker, XlsxDrawingObject, XlsxEmu, XlsxShapeRun, XlsxTextInsets,
        parse_drawing_shapes,
    };
    use crate::xlsx::writer::sheet::MutableWorksheet;

    fn marker(column: u32, row: u32) -> XlsxCellMarker {
        XlsxCellMarker {
            column,
            column_offset: XlsxEmu(100),
            row,
            row_offset: XlsxEmu(200),
        }
    }

    fn two_cell() -> XlsxShapeAnchor {
        XlsxShapeAnchor::TwoCell {
            from: marker(1, 2),
            to: marker(5, 9),
            edit_as: XlsxEditAs::OneCell,
        }
    }

    fn drawing_wrap(body: &str) -> String {
        format!(
            "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" \
             xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">{body}</xdr:wsDr>"
        )
    }

    fn parse_single(xml: &str) -> crate::xlsx::shapes::XlsxAnchoredObject {
        let objects = parse_drawing_shapes(&drawing_wrap(xml)).unwrap().unwrap();
        assert_eq!(objects.len(), 1);
        objects.into_iter().next().unwrap()
    }

    #[test]
    fn two_cell_text_box_round_trips_through_reader() {
        let mut spec = XlsxShapeSpec::text_box("Box 1", two_cell(), XlsxShapePreset::RoundRectangle, "Hello");
        spec.description = Some("alt <text>".to_string());
        spec.hidden = true;
        spec.body_properties = XlsxShapeBodyProperties {
            insets: XlsxTextInsets {
                left: XlsxEmu(182880),
                top: XlsxEmu(91440),
                right: XlsxEmu(182880),
                bottom: XlsxEmu(91440),
            },
            vertical_anchor: XlsxTextVerticalAnchor::Center,
            anchor_center: true,
            direction: XlsxTextDirection::Vertical270,
            wrap: XlsxTextWrap::None,
            autofit: XlsxTextAutofit::ShapeAutofit,
            column_count: 2,
            space_first_last_paragraph: true,
        };
        spec.paragraphs = vec![
            XlsxShapeParagraph {
                runs: vec![
                    XlsxShapeRun {
                        text: "Bold &".to_string(),
                        bold: Some(true),
                        italic: Some(false),
                        underline: Some(true),
                        font_size_hundredths: Some(1200),
                    },
                    XlsxShapeRun {
                        text: " plain".to_string(),
                        ..XlsxShapeRun::default()
                    },
                ],
            },
            XlsxShapeParagraph {
                runs: vec![XlsxShapeRun {
                    text: "second".to_string(),
                    ..XlsxShapeRun::default()
                }],
            },
        ];

        let mut xml = String::new();
        write_shape_anchor_xml(&mut xml, &spec, 7).unwrap();
        let anchored = parse_single(&xml);
        assert_eq!(anchored.anchor, spec.anchor);
        let XlsxDrawingObject::Shape(shape) = &anchored.object else {
            panic!("expected a shape");
        };
        assert_eq!(shape.non_visual.id, Some(7));
        assert_eq!(shape.non_visual.name.as_deref(), Some("Box 1"));
        assert_eq!(shape.non_visual.description.as_deref(), Some("alt <text>"));
        assert!(shape.non_visual.hidden);
        assert!(!shape.non_visual.locked);
        assert!(shape.is_text_box);
        assert_eq!(shape.preset, Some(XlsxShapePreset::RoundRectangle));
        let body = shape.text_body.as_ref().unwrap();
        assert_eq!(body.body_properties, spec.body_properties);
        assert_eq!(body.paragraphs, spec.paragraphs);
        assert_eq!(body.text(), "Bold & plain\nsecond");
    }

    #[test]
    fn one_cell_and_absolute_anchors_round_trip() {
        let one_cell = XlsxShapeAnchor::OneCell {
            from: marker(3, 4),
            extent: XlsxEmuExtent {
                width: XlsxEmu(914400),
                height: XlsxEmu(457200),
            },
        };
        let absolute = XlsxShapeAnchor::Absolute {
            position: XlsxEmuOffset {
                x: XlsxEmu(123),
                y: XlsxEmu(456),
            },
            extent: XlsxEmuExtent {
                width: XlsxEmu(789),
                height: XlsxEmu(101),
            },
        };
        for anchor in [one_cell, absolute] {
            let spec = XlsxShapeSpec::shape("S", anchor, XlsxShapePreset::Other("custGeom".into()), "");
            let mut xml = String::new();
            write_shape_anchor_xml(&mut xml, &spec, 1).unwrap();
            let anchored = parse_single(&xml);
            assert_eq!(anchored.anchor, anchor);
            let XlsxDrawingObject::Shape(shape) = &anchored.object else {
                panic!("expected a shape");
            };
            assert!(!shape.is_text_box);
            assert_eq!(
                shape.preset,
                Some(XlsxShapePreset::Other("custGeom".to_string()))
            );
        }
    }

    #[test]
    fn default_body_properties_round_trip() {
        let spec = XlsxShapeSpec::text_box("Defaults", two_cell(), XlsxShapePreset::Rectangle, "x");
        let mut xml = String::new();
        write_shape_anchor_xml(&mut xml, &spec, 2).unwrap();
        let anchored = parse_single(&xml);
        let XlsxDrawingObject::Shape(shape) = &anchored.object else {
            panic!("expected a shape");
        };
        let body = shape.text_body.as_ref().unwrap();
        assert_eq!(body.body_properties, XlsxShapeBodyProperties::default());
        // The default edit-as token is omitted from the output.
        let default_edit = XlsxShapeAnchor::TwoCell {
            from: marker(0, 0),
            to: marker(1, 1),
            edit_as: XlsxEditAs::TwoCell,
        };
        let spec = XlsxShapeSpec::text_box("E", default_edit, XlsxShapePreset::Rectangle, "x");
        let mut xml = String::new();
        write_shape_anchor_xml(&mut xml, &spec, 3).unwrap();
        assert!(!xml.contains("editAs"));
        assert_eq!(parse_single(&xml).anchor, default_edit);
    }

    #[test]
    fn validation_rejects_invalid_specs() {
        let mut spec = XlsxShapeSpec::text_box("", two_cell(), XlsxShapePreset::Rectangle, "x");
        assert!(spec.validate(0).is_err());
        spec.name = "ok".to_string();
        spec.anchor = XlsxShapeAnchor::TwoCell {
            from: marker(5, 9),
            to: marker(1, 2),
            edit_as: XlsxEditAs::TwoCell,
        };
        assert!(spec.validate(0).is_err());
        spec.anchor = XlsxShapeAnchor::TwoCell {
            from: marker(16_384, 0),
            to: marker(16_385, 1),
            edit_as: XlsxEditAs::TwoCell,
        };
        assert!(spec.validate(0).is_err());
        spec.anchor = two_cell();
        assert!(spec.validate(MAX_SHAPES_PER_WORKSHEET).is_err());
        assert!(spec.validate(0).is_ok());
    }

    #[test]
    fn worksheet_api_adds_removes_and_serializes_shapes() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.add_text_box("First", two_cell(), XlsxShapePreset::Rectangle, "hello")
            .unwrap();
        ws.add_shape(XlsxShapeSpec::text_box(
            "Second",
            two_cell(),
            XlsxShapePreset::Ellipse,
            "world",
        ))
        .unwrap();
        assert_eq!(ws.shapes().len(), 2);
        assert!(ws.remove_shape(5).is_err());
        let removed = ws.remove_shape(0).unwrap();
        assert_eq!(removed.name, "First");
        assert_eq!(ws.shapes().len(), 1);

        let xml = ws.generate_drawing_xml().unwrap().unwrap();
        assert!(xml.contains("<xdr:sp"));
        assert!(xml.contains(r#"prst="ellipse""#));
        let objects = parse_drawing_shapes(&xml).unwrap().unwrap();
        assert_eq!(objects.len(), 1);
        let XlsxDrawingObject::Shape(shape) = &objects[0].object else {
            panic!("expected a shape");
        };
        assert_eq!(shape.non_visual.name.as_deref(), Some("Second"));
        assert_eq!(shape.text_body.as_ref().unwrap().text(), "world");
    }

    #[test]
    fn shapes_coexist_with_images_in_drawing_xml() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.add_image(vec![1, 2, 3], "png", 1, 1, 2, 2, Some("Logo"))
            .unwrap();
        ws.add_text_box("Note", two_cell(), XlsxShapePreset::Rectangle, "note text")
            .unwrap();
        let xml = ws.generate_drawing_xml().unwrap().unwrap();
        // The image keeps rId1; the shape follows with the next object ID.
        assert!(xml.contains(r#"r:embed="rId1""#));
        assert!(xml.contains(r#"<xdr:cNvPr id="2" name="Note""#));
        let objects = parse_drawing_shapes(&xml).unwrap().unwrap();
        assert_eq!(objects.len(), 1, "pictures stay with the image pipeline");
    }

    #[test]
    fn authored_shapes_round_trip_through_a_saved_package() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("shapes.xlsx");
        let sheet_name;
        {
            let mut workbook = crate::xlsx::workbook::Workbook::create().unwrap();
            let ws = workbook.worksheet_mut(0).unwrap();
            sheet_name = ws.name().to_string();
            ws.add_image(vec![9, 9, 9], "png", 1, 1, 2, 2, None).unwrap();
            ws.add_text_box(
                "Greeting",
                XlsxShapeAnchor::TwoCell {
                    from: XlsxCellMarker {
                        column: 2,
                        column_offset: XlsxEmu(57150),
                        row: 1,
                        row_offset: XlsxEmu(47625),
                    },
                    to: XlsxCellMarker {
                        column: 6,
                        column_offset: XlsxEmu(0),
                        row: 8,
                        row_offset: XlsxEmu(0),
                    },
                    edit_as: XlsxEditAs::TwoCell,
                },
                XlsxShapePreset::RoundRectangle,
                "line one\nline two",
            )
            .unwrap();
            let mut fancy = XlsxShapeSpec::text_box(
                "Fancy",
                XlsxShapeAnchor::OneCell {
                    from: XlsxCellMarker {
                        column: 8,
                        column_offset: XlsxEmu(0),
                        row: 10,
                        row_offset: XlsxEmu(0),
                    },
                    extent: XlsxEmuExtent {
                        width: XlsxEmu(1_828_800),
                        height: XlsxEmu(914_400),
                    },
                },
                XlsxShapePreset::Ellipse,
                "fancy",
            );
            fancy.body_properties.vertical_anchor = XlsxTextVerticalAnchor::Bottom;
            fancy.body_properties.autofit = XlsxTextAutofit::NormalAutofit;
            fancy.body_properties.wrap = XlsxTextWrap::None;
            fancy.paragraphs[0].runs[0].bold = Some(true);
            fancy.paragraphs[0].runs[0].font_size_hundredths = Some(1400);
            ws.add_shape(fancy).unwrap();
            drop(ws);
            workbook.save(&path).unwrap();
        }

        let workbook = crate::xlsx::workbook::Workbook::open(&path).unwrap();
        let inventory = workbook.shapes_on_sheet(&sheet_name).unwrap();
        assert_eq!(inventory.objects.len(), 2);

        let XlsxDrawingObject::Shape(greeting) = &inventory.objects[0].object else {
            panic!("expected a shape");
        };
        assert_eq!(greeting.non_visual.name.as_deref(), Some("Greeting"));
        assert!(greeting.is_text_box);
        assert_eq!(greeting.preset, Some(XlsxShapePreset::RoundRectangle));
        assert_eq!(
            greeting.text_body.as_ref().unwrap().text(),
            "line one\nline two"
        );
        let XlsxShapeAnchor::TwoCell { from, .. } = inventory.objects[0].anchor else {
            panic!("expected a two-cell anchor");
        };
        assert_eq!(from.column, 2);
        assert_eq!(from.column_offset, XlsxEmu(57150));

        let XlsxDrawingObject::Shape(fancy) = &inventory.objects[1].object else {
            panic!("expected a shape");
        };
        assert!(matches!(
            inventory.objects[1].anchor,
            XlsxShapeAnchor::OneCell { .. }
        ));
        let body = fancy.text_body.as_ref().unwrap();
        assert_eq!(
            body.body_properties.vertical_anchor,
            XlsxTextVerticalAnchor::Bottom
        );
        assert_eq!(body.body_properties.autofit, XlsxTextAutofit::NormalAutofit);
        assert_eq!(body.body_properties.wrap, XlsxTextWrap::None);
        assert_eq!(body.paragraphs[0].runs[0].bold, Some(true));
        assert_eq!(body.paragraphs[0].runs[0].font_size_hundredths, Some(1400));

        // The saved package stays valid for the crate's own readers, and the
        // image pipeline still sees its picture.
        let package = litchi_opc::OpcPackage::open(&path).unwrap();
        let drawing_part = package
            .get_part(&litchi_opc::PackURI::new("/xl/drawings/drawing1.xml").unwrap())
            .unwrap();
        assert_eq!(
            drawing_part.content_type(),
            litchi_opc::constants::content_type::OFC_DRAWING
        );
        let drawing_xml = std::str::from_utf8(drawing_part.blob()).unwrap();
        assert!(drawing_xml.contains("<xdr:pic>"));
        assert!(drawing_xml.contains("<xdr:sp "));
    }
}
