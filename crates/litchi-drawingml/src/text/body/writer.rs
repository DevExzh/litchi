//! Canonical writer for DrawingML text-body content.
//!
//! The write_contents function emits the neutral a:CT_TextBody children
//! shared by SpreadsheetDrawing, PresentationML, and other DrawingML hosts:
//! a:bodyPr, a:lstStyle, paragraphs, runs, and text. Host crates retain the
//! surrounding txBody element and their anchor, relationship, and package
//! state. The emitted order follows the CT_TextBody sequence documented by
//! [MS-ODRAWXML] and the SpreadsheetDrawing example in 3rdparty/specs/.

use std::fmt::Write as _;

use litchi_core::xml::escape::escape_xml;

use super::{Body, Paragraph, Properties, Run};
use crate::text::{Anchor, Autofit, Columns, Direction, Wrap};

/// Write the neutral content of one DrawingML text body.
///
/// The caller owns the host-specific wrapper, for example
/// <xdr:txBody>...</xdr:txBody> or <p:txBody>...</p:txBody>.
pub fn write(xml: &mut String, body: &Body) {
    write_contents(xml, &body.properties, &body.paragraphs);
}

/// Write the neutral content of one DrawingML text body without requiring an
/// allocation when a host stores properties and paragraphs as separate fields.
pub fn write_contents(xml: &mut String, properties: &Properties, paragraphs: &[Paragraph]) {
    write_properties(xml, properties);
    xml.push_str("<a:lstStyle/>");
    for paragraph in paragraphs {
        xml.push_str("<a:p>");
        for run in &paragraph.runs {
            xml.push_str("<a:r>");
            write_run_properties(xml, run);
            let _ = write!(xml, "<a:t>{}</a:t>", escape_xml(&run.text));
            xml.push_str("</a:r>");
        }
        xml.push_str("</a:p>");
    }
}

fn write_properties(xml: &mut String, properties: &Properties) {
    xml.push_str("<a:bodyPr");
    let _ = write!(
        xml,
        r#" lIns="{}" tIns="{}" rIns="{}" bIns="{}""#,
        properties.insets.left,
        properties.insets.top,
        properties.insets.right,
        properties.insets.bottom
    );
    if properties.vertical_anchor != Anchor::Top {
        let _ = write!(xml, r#" anchor="{}""#, properties.vertical_anchor.token());
    }
    if properties.anchor_center {
        xml.push_str(r#" anchorCtr="1""#);
    }
    if properties.direction != Direction::Horizontal {
        let _ = write!(xml, r#" vert="{}""#, properties.direction.token());
    }
    if properties.wrap == Wrap::None {
        xml.push_str(r#" wrap="none""#);
    }
    if properties.column_count != Columns::ONE {
        let _ = write!(xml, r#" numCol="{}""#, properties.column_count);
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
    if run.bold.is_none()
        && run.italic.is_none()
        && run.underline.is_none()
        && run.font_size.is_none()
    {
        xml.push_str("<a:rPr/>");
        return;
    }
    xml.push_str("<a:rPr");
    if let Some(size) = run.font_size {
        let _ = write!(xml, r#" sz="{size}""#);
    }
    if let Some(bold) = run.bold {
        xml.push_str(if bold { r#" b="1""# } else { r#" b="0""# });
    }
    if let Some(italic) = run.italic {
        xml.push_str(if italic { r#" i="1""# } else { r#" i="0""# });
    }
    if let Some(underline) = run.underline {
        let _ = write!(xml, r#" u="{}""#, underline.dml());
    }
    xml.push_str("/>");
}

#[cfg(test)]
mod tests {
    use super::{write, write_contents};
    use crate::text::body::{Body, Insets, Paragraph, Properties, Run};
    use crate::text::{
        Anchor, Autofit, Columns, Coordinate32, Direction, TextSize, Underline, Wrap,
    };

    #[test]
    fn writes_defaults_and_escapes_text() {
        let mut xml = String::new();
        let body = Body {
            properties: Properties::default(),
            paragraphs: vec![Paragraph {
                runs: vec![Run {
                    text: "A & < B".to_owned(),
                    ..Run::default()
                }],
            }],
        };
        write(&mut xml, &body);

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
        };
        let paragraphs = [Paragraph {
            runs: vec![Run {
                text: "styled".to_owned(),
                bold: Some(true),
                italic: Some(false),
                underline: Some(Underline::DotDashHeavy),
                font_size: Some(TextSize::new(1_200).unwrap()),
            }],
        }];
        let mut xml = String::new();

        write_contents(&mut xml, &properties, &paragraphs);

        assert_eq!(
            xml,
            "<a:bodyPr lIns=\"1\" tIns=\"2\" rIns=\"3\" bIns=\"4\" anchor=\"ctr\" anchorCtr=\"1\" vert=\"vert270\" wrap=\"none\" numCol=\"2\" spcFirstLastPara=\"1\"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr sz=\"1200\" b=\"1\" i=\"0\" u=\"dotDashHeavy\"/><a:t>styled</a:t></a:r></a:p>"
        );
    }
}
