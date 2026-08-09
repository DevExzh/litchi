//! Small XML codecs used by the semantic authoring layer.

use litchi_core::xml::escape_xml;
use std::fmt::Write as _;

fn encode_text_content(text: &str) -> String {
    fn flush_plain(output: &mut String, plain: &mut String) {
        if !plain.is_empty() {
            output.push_str(&escape_xml(plain));
            plain.clear();
        }
    }

    let mut output = String::with_capacity(text.len());
    let mut plain = String::new();
    let mut characters = text.chars().peekable();
    while let Some(character) = characters.next() {
        match character {
            ' ' => {
                flush_plain(&mut output, &mut plain);
                let mut count = 1usize;
                while characters.next_if_eq(&' ').is_some() {
                    count += 1;
                }
                if count == 1 && !output.is_empty() && characters.peek().is_some() {
                    output.push(' ');
                } else if count == 1 {
                    output.push_str("<text:s/>");
                } else {
                    let _ = write!(output, r#"<text:s text:c="{count}"/>"#);
                }
            },
            '\t' => {
                flush_plain(&mut output, &mut plain);
                output.push_str("<text:tab/>");
            },
            '\r' => {
                flush_plain(&mut output, &mut plain);
                output.push_str("<text:line-break/>");
            },
            _ => plain.push(character),
        }
    }
    flush_plain(&mut output, &mut plain);
    output
}

pub(crate) fn generate_text_paragraphs(text: &str, style_name: Option<&str>) -> String {
    let escaped_style = style_name.map(escape_xml);
    let mut output = String::with_capacity(text.len() + 32);
    for paragraph in text.split('\n') {
        output.push_str("<text:p");
        if let Some(style) = escaped_style.as_deref() {
            output.push_str(r" text:style-name=");
            output.push('"');
            output.push_str(style);
            output.push('"');
        }
        output.push('>');
        output.push_str(&encode_text_content(paragraph));
        output.push_str("</text:p>");
    }
    output
}

pub(super) fn push_optional_attribute(output: &mut String, name: &str, value: Option<&str>) {
    if let Some(text) = value {
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(&escape_xml(text));
        output.push('"');
    }
}
