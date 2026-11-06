//! Direct handout placeholder projection with namespace and owner boundaries.

use super::codec::{check_element, check_empty_depth};
use super::{Limits, Master};
use crate::presentation_properties::metadata::process_str;
use crate::presentation_properties::{A_NS, A_STRICT, P_NS, P_STRICT};
use crate::{CodecLimitKind, Error, LimitKind, Result};
use litchi_ooxml_common::xml::attributes::first_wins;
use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Root,
    Common,
    Tree,
    Shape,
    NonVisual,
    Application,
    TextBody,
    Paragraph,
    Run,
    Field,
    Text,
    Background,
    BackgroundProperties,
    SolidFill,
    Other,
}

#[derive(Default)]
struct Placeholder {
    slot: Option<usize>,
    text: Option<String>,
    paragraphs: usize,
    auto_date: bool,
    saw_placeholder: bool,
    saw_body: bool,
}

struct Parser {
    master: Master,
    placeholders: [bool; 4],
    explicit: Option<[bool; 4]>,
    current: Option<Placeholder>,
    p: &'static str,
    a: &'static str,
    common_seen: bool,
}

fn invalid(message: &str) -> Error {
    Error::Invalid(message.into())
}

fn attribute(element: &BytesStart<'_>, key: &[u8], decoder: Decoder) -> Result<Option<String>> {
    for attr in first_wins(element) {
        let attr = match attr {
            Ok(attr) => attr,
            Err(quick_xml::events::attributes::AttrError::Duplicated(..)) => continue,
            Err(error) => return Err(Error::Xml(error.to_string())),
        };
        if attr.key.as_ref() == key {
            return attr
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map(|value| Some(value.into_owned()))
                .map_err(|e| Error::Xml(e.to_string()));
        }
    }
    Ok(None)
}

fn boolean(value: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid("invalid handout header/footer Boolean")),
    }
}

impl Parser {
    fn start(
        &mut self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        limits: Limits,
    ) -> Result<Context> {
        let name = element.local_name();
        let p =
            matches!(namespace, ResolveResult::Bound(Namespace(ns)) if *ns == self.p.as_bytes());
        let a =
            matches!(namespace, ResolveResult::Bound(Namespace(ns)) if *ns == self.a.as_bytes());
        if p {
            match (parent, name.as_ref()) {
                (Context::Root, b"cSld") => {
                    if self.common_seen {
                        return Err(invalid("duplicate handout common slide data"));
                    }
                    self.common_seen = true;
                    return Ok(Context::Common);
                },
                (Context::Root, b"hf") => {
                    if self.explicit.is_some() {
                        return Err(invalid("duplicate handout header/footer settings"));
                    }
                    // CT_HeaderFooter defaults each omitted attribute to true.
                    let mut flags = [true; 4];
                    for (index, key) in [b"hdr".as_slice(), b"ftr", b"dt", b"sldNum"]
                        .into_iter()
                        .enumerate()
                    {
                        if let Some(value) = attribute(element, key, decoder)? {
                            flags[index] = boolean(&value)?;
                        }
                    }
                    self.explicit = Some(flags);
                },
                (Context::Common, b"spTree") => return Ok(Context::Tree),
                (Context::Tree, b"sp") => {
                    self.current = Some(Placeholder::default());
                    return Ok(Context::Shape);
                },
                (Context::Shape, b"nvSpPr") => return Ok(Context::NonVisual),
                (Context::NonVisual, b"nvPr") => return Ok(Context::Application),
                (Context::Application, b"ph") => {
                    let current = self
                        .current
                        .as_mut()
                        .ok_or_else(|| invalid("handout placeholder has no shape owner"))?;
                    if current.saw_placeholder || current.saw_body {
                        return Err(invalid("duplicate or misplaced handout placeholder"));
                    }
                    current.saw_placeholder = true;
                    current.slot = match attribute(element, b"type", decoder)?.as_deref() {
                        Some("hdr") => Some(0),
                        Some("ftr") => Some(1),
                        Some("dt") => Some(2),
                        Some("sldNum") => Some(3),
                        _ => None,
                    };
                },
                (Context::Shape, b"txBody") => {
                    let current = self
                        .current
                        .as_mut()
                        .ok_or_else(|| invalid("handout text has no shape owner"))?;
                    if current.saw_body {
                        return Err(invalid("duplicate handout placeholder text body"));
                    }
                    current.saw_body = true;
                    return Ok(Context::TextBody);
                },
                (Context::Common, b"bg") => return Ok(Context::Background),
                (Context::Background, b"bgPr") => return Ok(Context::BackgroundProperties),
                _ => {},
            }
        }
        if a {
            match (parent, name.as_ref()) {
                (Context::BackgroundProperties, b"solidFill") => return Ok(Context::SolidFill),
                (Context::SolidFill, b"srgbClr") => {
                    if self.master.background_color.is_some() {
                        return Err(invalid("ambiguous handout background RGB color"));
                    }
                    let color = attribute(element, b"val", decoder)?
                        .ok_or_else(|| invalid("handout RGB color has no value"))?;
                    super::codec::validate_color(&color)?;
                    self.master.background_color = Some(color);
                },
                (Context::TextBody, b"p") => {
                    if let Some(current) = &mut self.current {
                        if current.paragraphs != 0 {
                            append(current, "\n", limits)?;
                        }
                        current.paragraphs += 1;
                    }
                    return Ok(Context::Paragraph);
                },
                (Context::Paragraph, b"r") => return Ok(Context::Run),
                (Context::Paragraph, b"fld") => {
                    if let Some(current) = &mut self.current {
                        if current.slot == Some(2)
                            && attribute(element, b"type", decoder)?
                                .is_some_and(|s| s.starts_with("datetime"))
                        {
                            current.auto_date = true;
                        }
                    }
                    return Ok(Context::Field);
                },
                (Context::Run | Context::Field, b"t") => {
                    if let Some(current) = &mut self.current {
                        append(current, "", limits)?;
                    }
                    return Ok(Context::Text);
                },
                (Context::Paragraph, b"br") => {
                    if let Some(current) = &mut self.current {
                        append(current, "\n", limits)?;
                    }
                },
                _ => {},
            }
        }
        Ok(Context::Other)
    }

    fn finish(&mut self, context: Context) -> Result<()> {
        if context == Context::Shape {
            let current = self
                .current
                .take()
                .ok_or_else(|| invalid("handout shape owner is missing"))?;
            if let Some(slot) = current.slot {
                if self.placeholders[slot] {
                    return Err(Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedContent,
                        "parse_handout_master",
                        "multiple handout placeholders have the same header/footer role",
                    ));
                }
                self.placeholders[slot] = true;
                match slot {
                    0 => self.master.header_footer.header_text = current.text,
                    1 => self.master.header_footer.footer_text = current.text,
                    2 => {
                        self.master.header_footer.date_time_text = current.text;
                        self.master.header_footer.auto_date = current.auto_date;
                    },
                    _ => {},
                }
            }
        }
        Ok(())
    }
}

fn append(shape: &mut Placeholder, text: &str, limits: Limits) -> Result<()> {
    if shape.slot.is_none() {
        return Ok(());
    }
    super::codec::validate_text(text, limits)?;
    let value = shape.text.get_or_insert_with(String::new);
    let size = value
        .len()
        .checked_add(text.len())
        .ok_or(Error::CapacityOverflow {
            resource: "handout placeholder text",
        })?;
    if size > limits.max_string_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::Codec(CodecLimitKind::StringBytes),
            observed: size,
            maximum: limits.max_string_bytes(),
        });
    }
    value
        .try_reserve(text.len())
        .map_err(|source| Error::Allocation {
            resource: "handout placeholder text",
            source,
        })?;
    value.push_str(text);
    Ok(())
}

pub(super) fn parse(xml: &str, limits: Limits) -> Result<Master> {
    if xml.len() > limits.max_input_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::Codec(CodecLimitKind::InputBytes),
            observed: xml.len(),
            maximum: limits.max_input_bytes(),
        });
    }
    let xml = process_str(xml, limits)?;
    if xml.len() > limits.max_output_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::Codec(CodecLimitKind::OutputBytes),
            observed: xml.len(),
            maximum: limits.max_output_bytes(),
        });
    }
    let mut reader = NsReader::from_str(xml.as_ref());
    let mut parser = Parser {
        master: Master::default(),
        placeholders: [false; 4],
        explicit: None,
        current: None,
        p: P_NS,
        a: A_NS,
        common_seen: false,
    };
    let mut stack = Vec::new();
    let mut nodes = 0;
    let mut closed = false;
    loop {
        let event = reader.read_event().map_err(|e| Error::Xml(e.to_string()))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                check_empty_depth(stack.len(), limits)?;
                check_element(element, reader.decoder(), &mut nodes, limits)?;
                let namespace = reader.resolver().resolve_element(element.name()).0;
                if matches!(namespace, ResolveResult::Unknown(_)) {
                    return Err(invalid("unbound namespace prefix in handout master"));
                }
                let context = if stack.is_empty() {
                    if closed || element.local_name().as_ref() != b"handoutMaster" {
                        return Err(invalid("expected one handoutMaster root"));
                    }
                    match namespace {
                        ResolveResult::Bound(Namespace(ns)) if ns == P_NS.as_bytes() => {},
                        ResolveResult::Bound(Namespace(ns)) if ns == P_STRICT.as_bytes() => {
                            parser.p = P_STRICT;
                            parser.a = A_STRICT;
                        },
                        _ => return Err(invalid("handoutMaster root has a foreign namespace")),
                    }
                    Context::Root
                } else {
                    parser.start(
                        *stack
                            .last()
                            .ok_or_else(|| invalid("missing handout context"))?,
                        &namespace,
                        element,
                        reader.decoder(),
                        limits,
                    )?
                };
                if matches!(event, Event::Start(_)) {
                    stack.push(context);
                } else {
                    parser.finish(context)?;
                    if context == Context::Root {
                        closed = true;
                    }
                }
            },
            Event::End(_) => {
                let context = stack
                    .pop()
                    .ok_or_else(|| invalid("handout-master XML has an unmatched end tag"))?;
                parser.finish(context)?;
                if context == Context::Root {
                    closed = true;
                }
            },
            Event::Text(text) => {
                let text = text.decode().map_err(|e| Error::Xml(e.to_string()))?;
                if stack.last() == Some(&Context::Text) {
                    if let Some(current) = &mut parser.current {
                        append(current, &text, limits)?;
                    }
                } else if stack.is_empty() && !text.trim().is_empty() {
                    return Err(invalid("text outside handout master root"));
                }
            },
            Event::CData(text) => {
                if stack.is_empty() {
                    return Err(invalid("CDATA outside handout master root"));
                }
                if stack.last() == Some(&Context::Text) {
                    if let Some(current) = &mut parser.current {
                        append(
                            current,
                            &text.decode().map_err(|e| Error::Xml(e.to_string()))?,
                            limits,
                        )?;
                    }
                }
            },
            Event::GeneralRef(reference) => {
                if stack.is_empty() {
                    return Err(invalid("entity reference outside handout master root"));
                }
                if stack.last() == Some(&Context::Text) {
                    if let Some(current) = &mut parser.current {
                        append(current, &decode_xml_reference(&reference)?, limits)?;
                    }
                }
            },
            Event::DocType(_) => return Err(invalid("handout master cannot contain a DTD")),
            Event::Eof => break,
            _ => {},
        }
    }
    if !closed || !stack.is_empty() {
        return Err(invalid(
            "handout-master XML has a missing or unterminated root",
        ));
    }
    let flags = parser.explicit.unwrap_or(parser.placeholders);
    let footer = &mut parser.master.header_footer;
    footer.show_header = flags[0];
    footer.show_footer = flags[1];
    footer.show_date_time = flags[2];
    footer.show_slide_number = flags[3];
    Ok(parser.master)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "codec boundary assertions"
)]
mod tests {
    use super::*;

    fn shape(kind: &str, content: &str) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="field"/><p:cNvSpPr/><p:nvPr><p:ph type="{kind}"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:lstStyle/>{content}</p:txBody></p:sp>"#
        )
    }
    fn document(shapes: &str, tail: &str) -> String {
        format!(
            r#"<p:handoutMaster xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:f="urn:foreign"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></p:bgPr></p:bg><p:spTree>{shapes}</p:spTree></p:cSld>{tail}</p:handoutMaster>"#
        )
    }

    #[test]
    fn placeholder_text_fields_and_flags_follow_owned_namespaces() {
        let header = shape(
            "hdr",
            "<a:p><a:r><a:t>  H &amp; </a:t></a:r><a:br/><a:r><a:t><![CDATA[<tail>]]></a:t></a:r></a:p><a:p><a:r><a:t>Next</a:t></a:r></a:p>",
        );
        let date = shape(
            "dt",
            r#"<a:p><a:fld type="datetime1"><a:t>2026-11-06</a:t></a:fld></a:p>"#,
        );
        let fake = shape("ftr", "<a:p><a:r><a:t>foreign</a:t></a:r></a:p>").replace("p:", "f:");
        let nested = format!(
            "<p:grpSp>{}</p:grpSp>",
            shape("ftr", "<a:p><a:r><a:t>nested</a:t></a:r></a:p>")
        );
        let xml = document(
            &(header + &date + &fake + &nested),
            "<p:extLst><p:hf hdr=\"0\"/><a:srgbClr val=\"FFFFFF\"/></p:extLst>",
        );
        for xml in [
            xml.clone(),
            xml.replace("<p:", "<q:")
                .replace("</p:", "</q:")
                .replace("xmlns:p=", "xmlns:q=")
                .replace("<a:", "<d:")
                .replace("</a:", "</d:")
                .replace("xmlns:a=", "xmlns:d="),
            xml.replace(P_NS, P_STRICT).replace(A_NS, A_STRICT),
        ] {
            let master = parse(&xml, Limits::DEFAULT).unwrap();
            assert_eq!(
                master.header_footer.header_text.as_deref(),
                Some("  H & \n<tail>\nNext")
            );
            assert!(master.header_footer.show_header);
            assert!(!master.header_footer.show_footer);
            assert_eq!(master.header_footer.footer_text, None);
            assert!(master.header_footer.auto_date);
            assert_eq!(
                master.header_footer.date_time_text.as_deref(),
                Some("2026-11-06")
            );
            assert_eq!(master.background_color.as_deref(), Some("112233"));
        }
    }

    #[test]
    fn explicit_booleans_override_placeholders_and_omitted_attributes_default_true() {
        let header = shape("hdr", "<a:p><a:r><a:t>hidden text</a:t></a:r></a:p>");
        let master = parse(
            &document(&header, r#"<p:hf hdr="false" dt="0"/>"#),
            Limits::DEFAULT,
        )
        .unwrap();
        assert!(!master.header_footer.show_header);
        assert_eq!(
            master.header_footer.header_text.as_deref(),
            Some("hidden text")
        );
        assert!(!master.header_footer.show_date_time);
        assert!(master.header_footer.show_footer);
        assert!(master.header_footer.show_slide_number);
        assert!(parse(&document("", r#"<p:hf hdr="yes"/>"#), Limits::DEFAULT).is_err());
        assert!(parse(&document("", "<p:hf/><p:hf/>"), Limits::DEFAULT).is_err());
    }

    #[test]
    fn compatibility_projection_selects_only_the_active_direct_placeholder() {
        let selected = shape("hdr", "<a:p><a:r><a:t>active</a:t></a:r></a:p>");
        let skipped = shape("hdr", "<a:p><a:r><a:t>inactive</a:t></a:r></a:p>");
        let branches = format!(
            r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:future="urn:future"><mc:Choice Requires="future">{skipped}</mc:Choice><mc:Fallback>{selected}</mc:Fallback></mc:AlternateContent>"#
        );
        assert_eq!(
            parse(&document(&branches, ""), Limits::DEFAULT)
                .unwrap()
                .header_footer
                .header_text
                .as_deref(),
            Some("active")
        );
        assert!(matches!(
            parse(
                &document(&(selected.clone() + &selected), ""),
                Limits::DEFAULT
            ),
            Err(Error::Refused(refusal))
                if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
                    && refusal.operation() == "parse_handout_master"
        ));
    }

    #[test]
    fn text_limits_cover_combined_runs_and_writer_escaped_output() {
        let runs = format!(
            "<a:p><a:r><a:t>{}</a:t></a:r><a:r><a:t>{}</a:t></a:r></a:p>",
            "x".repeat(64),
            "y".repeat(64)
        );
        let xml = document(&shape("hdr", &runs), "");
        parse(&xml, Limits::DEFAULT.with_string_bytes(128).unwrap()).unwrap();
        assert!(matches!(
            parse(&xml, Limits::DEFAULT.with_string_bytes(127).unwrap()),
            Err(Error::Limit {
                kind: LimitKind::Codec(CodecLimitKind::StringBytes),
                observed: 128,
                maximum: 127
            })
        ));
        assert!(parse(&xml, Limits::DEFAULT.with_depth(7).unwrap()).is_err());
        parse(&xml, Limits::DEFAULT.with_depth(8).unwrap()).unwrap();
        let master = Master::new().with_header("<&>".repeat(100));
        assert!(matches!(
            master.to_xml(Limits::DEFAULT.with_output_bytes(100).unwrap()),
            Err(Error::Limit { .. })
        ));
        assert!(
            Master::new()
                .with_header("bad\0text")
                .to_xml(Limits::DEFAULT)
                .is_err()
        );
        assert!(
            Master::new()
                .with_background_color("not-rgb")
                .to_xml(Limits::DEFAULT)
                .is_err()
        );
    }

    #[test]
    fn fresh_writer_keeps_text_whitespace_dates_flags_and_rgb_color() {
        for master in [
            Master::new(),
            Master::new()
                .with_header("  A & <B> 😀\nnext  ")
                .with_footer("")
                .with_fixed_date("2026-11-06")
                .with_background_color("ABCDEF"),
            Master::new().with_date_time().with_slide_numbers(),
        ] {
            let parsed = parse(&master.to_xml(Limits::DEFAULT).unwrap(), Limits::DEFAULT).unwrap();
            assert_eq!(
                parsed.header_footer.header_text,
                master.header_footer.header_text
            );
            assert_eq!(
                parsed.header_footer.footer_text,
                master.header_footer.footer_text
            );
            assert_eq!(
                parsed.header_footer.date_time_text,
                master.header_footer.date_time_text
            );
            assert_eq!(
                parsed.header_footer.auto_date,
                master.header_footer.auto_date
            );
            assert_eq!(
                [
                    parsed.header_footer.show_header,
                    parsed.header_footer.show_footer,
                    parsed.header_footer.show_date_time,
                    parsed.header_footer.show_slide_number
                ],
                [
                    master.header_footer.show_header,
                    master.header_footer.show_footer,
                    master.header_footer.show_date_time,
                    master.header_footer.show_slide_number
                ]
            );
            assert_eq!(parsed.background_color, master.background_color);
        }
    }

    #[test]
    fn rejects_invalid_roots_and_content_outside_the_document() {
        let root = format!(r#"<p:handoutMaster xmlns:p="{P_NS}"/>"#);
        for xml in [
            String::new(),
            "<p:handoutMaster/>".into(),
            root.replace(P_NS, "urn:foreign"),
            format!("{root}{root}"),
            format!("{root}text"),
            format!("{root}<![CDATA[outside]]>"),
            format!("&amp;{root}"),
            format!("<!DOCTYPE handoutMaster>{root}"),
        ] {
            assert!(parse(&xml, Limits::DEFAULT).is_err(), "{xml}");
        }
    }

    #[test]
    fn generated_field_ids_have_the_ooxml_guid_shape() {
        let xml = Master::new()
            .with_date_time()
            .to_xml(Limits::DEFAULT)
            .unwrap();
        let mut reader = NsReader::from_str(&xml);
        let mut ids = Vec::new();
        loop {
            match reader.read_event().unwrap() {
                Event::Start(element) if element.local_name().as_ref() == b"fld" => {
                    let id = attribute(&element, b"id", reader.decoder())
                        .unwrap()
                        .unwrap();
                    assert_eq!(id.len(), 38, "{id}");
                    assert!(id.starts_with('{') && id.ends_with('}'));
                    assert_eq!(
                        id[1..37].split('-').map(str::len).collect::<Vec<_>>(),
                        [8, 4, 4, 4, 12]
                    );
                    assert!(
                        id[1..37]
                            .bytes()
                            .all(|b| b.is_ascii_hexdigit() || b == b'-')
                    );
                    ids.push(id);
                },
                Event::Eof => break,
                _ => {},
            }
        }
        assert_eq!(ids.len(), 2);
        assert_ne!(ids[0], ids[1]);
    }
}
