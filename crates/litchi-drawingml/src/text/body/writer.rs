//! Canonical writer for `DrawingML` text-body content.
//!
//! The `write_contents` function emits the neutral `a:CT_TextBody` children
//! shared by `SpreadsheetDrawing`, `PresentationML`, and other `DrawingML` hosts:
//! a:bodyPr, a:lstStyle, paragraphs, runs, and text. Host crates retain the
//! surrounding txBody element and their anchor, relationship, and package
//! state. The emitted order follows the `CT_TextBody` sequence documented by
//! [MS-ODRAWXML] and its `SpreadsheetDrawing` example.
//!
//! The writer authors fresh bodies: text runs are written as `a:r`, line
//! breaks as `a:br`, and both retain the supported `sz`, `b`, `i` and `u`
//! properties. Fields, paragraph properties and the other character
//! properties the reader models are refused before any output is written. A
//! body that came from a reader with skipped list-style, effect, extension,
//! or other unmodeled markup is also refused, because this writer cannot
//! reproduce that projection losslessly.

use litchi_core::xml::escape::escape_xml;

use super::{Body, Paragraph, ParagraphProperties, Properties, Run, RunKind, RunProperties};
use crate::text::{Anchor, Autofit, Columns, Direction, Wrap};

/// Write the neutral content of one `DrawingML` text body.
///
/// The caller owns the host-specific wrapper, for example
/// <xdr:txBody>...</xdr:txBody> or <p:txBody>...</p:txBody>.
///
/// # Errors
///
/// Returns [`crate::Error::Invalid`] when the body contains content this
/// writer cannot represent without loss, including a reader projection marked
/// as incomplete. The output is unchanged on error.
pub fn write(xml: &mut String, body: &Body) -> crate::Result<()> {
    write_contents(xml, &body.properties, &body.paragraphs)
}

/// Write the neutral content of one `DrawingML` text body without requiring an
/// allocation when a host stores properties and paragraphs as separate fields.
///
/// # Errors
///
/// Returns [`crate::Error::Invalid`] when the body properties or paragraphs
/// contain content this writer cannot represent without loss. The output is
/// unchanged on error.
pub fn write_contents(
    xml: &mut String,
    properties: &Properties,
    paragraphs: &[Paragraph],
) -> crate::Result<()> {
    validate(properties, paragraphs)?;
    write_properties(xml, properties);
    xml.push_str("<a:lstStyle/>");
    for paragraph in paragraphs {
        xml.push_str("<a:p>");
        for run in &paragraph.runs {
            match &run.kind {
                RunKind::Text => {
                    xml.push_str("<a:r>");
                    write_run_properties(xml, run);
                    xml.push_str(&format!("<a:t>{}</a:t>", escape_xml(&run.text)));
                    xml.push_str("</a:r>");
                },
                RunKind::Break => {
                    xml.push_str("<a:br>");
                    write_run_properties(xml, run);
                    xml.push_str("</a:br>");
                },
                RunKind::Field(_) => {
                    // `validate` above rejects fields before the first byte is
                    // appended. Keep this arm defensive if that policy ever
                    // changes.
                    return Err(crate::Error::Invalid(
                        "text body fields cannot be written".to_owned(),
                    ));
                },
            }
        }
        xml.push_str("</a:p>");
    }
    Ok(())
}

fn write_properties(xml: &mut String, properties: &Properties) {
    xml.push_str("<a:bodyPr");
    xml.push_str(&format!(
        r#" lIns="{}" tIns="{}" rIns="{}" bIns="{}""#,
        properties.insets.left,
        properties.insets.top,
        properties.insets.right,
        properties.insets.bottom
    ));
    if properties.vertical_anchor != Anchor::Top {
        xml.push_str(&format!(
            r#" anchor="{}""#,
            properties.vertical_anchor.token()
        ));
    }
    if properties.anchor_center {
        xml.push_str(r#" anchorCtr="1""#);
    }
    if properties.direction != Direction::Horizontal {
        xml.push_str(&format!(r#" vert="{}""#, properties.direction.token()));
    }
    if properties.wrap == Wrap::None {
        xml.push_str(r#" wrap="none""#);
    }
    if properties.column_count != Columns::ONE {
        xml.push_str(&format!(r#" numCol="{}""#, properties.column_count));
    }
    if properties.space_first_last_paragraph {
        xml.push_str(r#" spcFirstLastPara="1""#);
    }
    match properties.autofit {
        Autofit::None => xml.push_str("><a:noAutofit/></a:bodyPr>"),
        Autofit::Shape => xml.push_str("><a:spAutoFit/></a:bodyPr>"),
        Autofit::Normal => xml.push_str("><a:normAutofit/></a:bodyPr>"),
    }
}

fn write_run_properties(xml: &mut String, run: &Run) {
    let properties = &run.properties;
    if properties.bold.is_none()
        && properties.italic.is_none()
        && properties.underline.is_none()
        && properties.size.is_none()
    {
        xml.push_str("<a:rPr/>");
        return;
    }
    xml.push_str("<a:rPr");
    if let Some(size) = properties.size {
        xml.push_str(&format!(r#" sz="{size}""#));
    }
    if let Some(bold) = properties.bold {
        xml.push_str(if bold { r#" b="1""# } else { r#" b="0""# });
    }
    if let Some(italic) = properties.italic {
        xml.push_str(if italic { r#" i="1""# } else { r#" i="0""# });
    }
    if let Some(underline) = properties.underline {
        xml.push_str(&format!(r#" u="{}""#, underline.dml()));
    }
    xml.push_str("/>");
}

/// The first part of `paragraphs` the writer would not write, as a short
/// description, or `None` when it writes every value they hold.
///
/// This is retained for hosts that validate their complete authoring graph
/// before serializing it. [`write()`] and [`write_contents`] enforce the same
/// policy themselves.
#[must_use]
pub fn unsupported(paragraphs: &[Paragraph]) -> Option<&'static str> {
    for paragraph in paragraphs {
        if paragraph.properties != ParagraphProperties::default() {
            return Some("paragraph properties");
        }
        if paragraph.end_properties.is_some() {
            return Some("paragraph end properties");
        }
        for run in &paragraph.runs {
            if matches!(&run.kind, RunKind::Field(_)) {
                return Some("fields");
            }
            if matches!(&run.kind, RunKind::Break) && run.text != "\n" {
                return Some("line breaks whose text is not a single newline");
            }
            let written = RunProperties {
                bold: run.properties.bold,
                italic: run.properties.italic,
                underline: run.properties.underline,
                size: run.properties.size,
                ..RunProperties::default()
            };
            if run.properties != written {
                return Some("character properties other than sz, b, i and u");
            }
        }
    }
    None
}

/// Refuse paragraph content the serializer cannot reproduce exactly.
fn validate(properties: &Properties, paragraphs: &[Paragraph]) -> crate::Result<()> {
    if properties.has_unmodeled_content {
        return Err(crate::Error::Invalid(
            "text body contains unmodeled content".to_owned(),
        ));
    }
    if let Some(content) = unsupported(paragraphs) {
        return Err(crate::Error::Invalid(format!(
            "text body {content} cannot be written without loss"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{unsupported, write, write_contents};
    use crate::text::body::{
        Body, Field, Insets, Paragraph, ParagraphProperties, Properties, Run, RunKind,
        RunProperties,
    };
    use crate::text::{
        Anchor, Autofit, Columns, Coordinate32, Direction, TextSize, Underline, Wrap,
    };

    #[test]
    fn writes_defaults_and_escapes_text() {
        let mut xml = String::new();
        let body = Body {
            properties: Properties::default(),
            paragraphs: vec![Paragraph {
                runs: vec![Run::plain("A & < B")],
                ..Paragraph::default()
            }],
        };
        write(&mut xml, &body).unwrap();

        assert_eq!(
            xml,
            "<a:bodyPr lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr/><a:t>A &amp; &lt; B</a:t></a:r></a:p>"
        );
    }

    #[test]
    fn writes_non_default_body_and_run_properties() {
        let properties = Properties {
            insets: Insets {
                left: Coordinate32::from(1),
                top: Coordinate32::from(2),
                right: Coordinate32::from(3),
                bottom: Coordinate32::from(4),
            },
            vertical_anchor: Anchor::Center,
            anchor_center: true,
            direction: Direction::Vertical270,
            wrap: Wrap::None,
            autofit: Autofit::Shape,
            column_count: Columns::new(2).unwrap(),
            space_first_last_paragraph: true,
            has_unmodeled_content: false,
        };
        let paragraphs = [Paragraph {
            runs: vec![Run {
                text: "styled".to_owned(),
                properties: RunProperties {
                    bold: Some(true),
                    italic: Some(false),
                    underline: Some(Underline::DotDashHeavy),
                    size: Some(TextSize::new(1_200).unwrap()),
                    ..RunProperties::default()
                },
                ..Run::default()
            }],
            ..Paragraph::default()
        }];
        assert_eq!(unsupported(&paragraphs), None);
        let mut xml = String::new();

        write_contents(&mut xml, &properties, &paragraphs).unwrap();

        assert_eq!(
            xml,
            "<a:bodyPr lIns=\"1\" tIns=\"2\" rIns=\"3\" bIns=\"4\" anchor=\"ctr\" anchorCtr=\"1\" vert=\"vert270\" wrap=\"none\" numCol=\"2\" spcFirstLastPara=\"1\"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr sz=\"1200\" b=\"1\" i=\"0\" u=\"dotDashHeavy\"/><a:t>styled</a:t></a:r></a:p>"
        );
    }

    #[test]
    fn reports_content_the_writer_would_drop() {
        let plain = Paragraph {
            runs: vec![Run::plain("x")],
            ..Paragraph::default()
        };
        assert_eq!(unsupported(std::slice::from_ref(&plain)), None);
        let aligned = Paragraph {
            properties: ParagraphProperties {
                level: Some(1),
                ..ParagraphProperties::default()
            },
            ..plain.clone()
        };
        assert_eq!(unsupported(&[aligned]), Some("paragraph properties"));
        let broken = Paragraph {
            runs: vec![Run {
                kind: RunKind::Break,
                text: "\n".to_owned(),
                ..Run::default()
            }],
            ..Paragraph::default()
        };
        assert_eq!(unsupported(&[broken]), None);
        let field = Paragraph {
            runs: vec![Run {
                kind: RunKind::Field(Field::default()),
                ..Run::default()
            }],
            ..Paragraph::default()
        };
        assert_eq!(unsupported(&[field]), Some("fields"));
        let colored = Paragraph {
            runs: vec![Run {
                text: "x".to_owned(),
                properties: RunProperties {
                    language: Some("en-US".to_owned()),
                    ..RunProperties::default()
                },
                ..Run::default()
            }],
            ..Paragraph::default()
        };
        assert_eq!(
            unsupported(&[colored]),
            Some("character properties other than sz, b, i and u")
        );
    }

    #[test]
    fn writes_breaks_as_break_elements() {
        let paragraphs = [Paragraph {
            runs: vec![Run {
                kind: RunKind::Break,
                text: "\n".to_owned(),
                properties: RunProperties {
                    size: Some(TextSize::new(1_000).unwrap()),
                    ..RunProperties::default()
                },
            }],
            ..Paragraph::default()
        }];
        let mut xml = String::new();

        write_contents(&mut xml, &Properties::default(), &paragraphs).unwrap();

        assert!(xml.contains("<a:br><a:rPr sz=\"1000\"/></a:br>"));
        assert!(!xml.contains("<a:t>"));
    }

    #[test]
    fn refuses_lossy_body_without_touching_output() {
        let paragraphs = [Paragraph {
            properties: ParagraphProperties {
                level: Some(1),
                ..ParagraphProperties::default()
            },
            runs: vec![Run::plain("authored")],
            ..Paragraph::default()
        }];
        let mut xml = "prefix".to_owned();

        let error = write_contents(&mut xml, &Properties::default(), &paragraphs).unwrap_err();

        assert_eq!(xml, "prefix");
        assert!(
            matches!(error, crate::Error::Invalid(message) if message == "text body paragraph properties cannot be written without loss")
        );
    }

    #[test]
    fn refuses_break_with_non_newline_text_without_touching_output() {
        let paragraphs = [Paragraph {
            runs: vec![Run {
                kind: RunKind::Break,
                text: "wrong".to_owned(),
                ..Run::default()
            }],
            ..Paragraph::default()
        }];
        let mut xml = "prefix".to_owned();

        let error = write(
            &mut xml,
            &Body {
                properties: Properties::default(),
                paragraphs: paragraphs.to_vec(),
            },
        )
        .unwrap_err();

        assert_eq!(xml, "prefix");
        assert!(
            matches!(error, crate::Error::Invalid(message) if message == "text body line breaks whose text is not a single newline cannot be written without loss")
        );
    }

    #[test]
    fn refuses_incomplete_projection_without_touching_output() {
        let properties = Properties {
            has_unmodeled_content: true,
            ..Properties::default()
        };
        let mut xml = "prefix".to_owned();

        let error = write_contents(&mut xml, &properties, &[]).unwrap_err();

        assert_eq!(xml, "prefix");
        assert!(
            matches!(error, crate::Error::Invalid(message) if message == "text body contains unmodeled content")
        );
    }

    #[test]
    fn refuses_reader_projection_with_unmodeled_markup() {
        let body = crate::text::body::reader::read(
            br#"<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:effectLst><a:outerShdw/></a:effectLst></a:defRPr></a:lvl1pPr></a:lstStyle>
                <a:p><a:r><a:t>text</a:t></a:r></a:p>
            </a:txBody>"#,
        )
        .unwrap();
        assert!(body.properties.has_unmodeled_content);

        let mut xml = "prefix".to_owned();
        let error = write(&mut xml, &body).unwrap_err();

        assert_eq!(xml, "prefix");
        assert!(
            matches!(error, crate::Error::Invalid(message) if message == "text body contains unmodeled content")
        );
    }
}
