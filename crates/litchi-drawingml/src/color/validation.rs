//! Structural and scalar validation for `DrawingML` color fragments.

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::{Error, Result};

/// Maximum accepted color-fragment size.
pub const MAX_XML_BYTES: usize = 64 * 1024;
/// Maximum accepted nesting depth for an opaque fragment.
pub const MAX_DEPTH: usize = 32;
/// Maximum accepted element count for an opaque fragment.
pub const MAX_NODES: usize = 256;
/// Maximum number of typed color transforms in one color choice.
pub const MAX_TRANSFORMS: usize = 64;

/// Caller-selected bounds for one raw color fragment.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[must_use]
pub struct Limits {
    pub max_xml_bytes: usize,
    pub max_depth: usize,
    pub max_nodes: usize,
    pub max_transforms: usize,
}

impl Limits {
    pub const DEFAULT: Self = Self {
        max_xml_bytes: MAX_XML_BYTES,
        max_depth: MAX_DEPTH,
        max_nodes: MAX_NODES,
        max_transforms: MAX_TRANSFORMS,
    };

    #[must_use]
    pub const fn new(
        max_xml_bytes: usize,
        max_depth: usize,
        max_nodes: usize,
        max_transforms: usize,
    ) -> Self {
        Self {
            max_xml_bytes,
            max_depth,
            max_nodes,
            max_transforms,
        }
    }

    pub(crate) fn validate(self) -> Result<Self> {
        for (kind, value, maximum) in [
            (
                crate::LimitKind::XmlBytes,
                self.max_xml_bytes,
                MAX_XML_BYTES,
            ),
            (crate::LimitKind::XmlDepth, self.max_depth, MAX_DEPTH),
            (crate::LimitKind::XmlNodes, self.max_nodes, MAX_NODES),
            (
                crate::LimitKind::ColorTransforms,
                self.max_transforms,
                MAX_TRANSFORMS,
            ),
        ] {
            if value == 0 || value > maximum {
                return Err(Error::invalid_limit(kind, value, maximum));
            }
        }
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

const MIN_PERCENTAGE: i32 = -100_000;
const MAX_PERCENTAGE: i32 = 100_000;
const MAX_POSITIVE_ANGLE: u32 = 21_600_000;

pub(super) fn percentage(value: i32, kind: &'static str) -> Result<i32> {
    if !(MIN_PERCENTAGE..=MAX_PERCENTAGE).contains(&value) {
        return Err(Error::Invalid(format!(
            "DrawingML {kind} must be between -100000 and 100000"
        )));
    }
    Ok(value)
}

pub(super) fn positive_percentage(value: u32, kind: &'static str) -> Result<u32> {
    if value > MAX_PERCENTAGE as u32 {
        return Err(Error::Invalid(format!(
            "DrawingML {kind} must be at most 100000"
        )));
    }
    Ok(value)
}

pub(super) fn positive_angle(value: u32) -> Result<u32> {
    if value > MAX_POSITIVE_ANGLE {
        return Err(Error::Invalid(
            "DrawingML positive fixed angles must be at most 21600000".into(),
        ));
    }
    Ok(value)
}

pub(super) fn parse_percentage(value: &str, kind: &'static str) -> Result<i32> {
    let value = parse_scaled(value, kind)?;
    percentage(value, kind)
}

pub(super) fn parse_positive_percentage(value: &str, kind: &'static str) -> Result<u32> {
    let value = parse_scaled(value, kind)?;
    if value < 0 {
        return Err(Error::Invalid(format!(
            "DrawingML {kind} cannot be negative"
        )));
    }
    let value = u32::try_from(value).map_err(|error| {
        Error::Invalid(format!(
            "DrawingML {kind} cannot be represented as a positive percentage: {error}"
        ))
    })?;
    positive_percentage(value, kind)
}

pub(super) fn parse_angle(value: &str) -> Result<i32> {
    value.trim().parse::<i32>().map_err(|_error| {
        Error::Invalid(format!(
            "DrawingML angles must be signed 60000ths of a degree: {value:?}"
        ))
    })
}

pub(super) fn parse_positive_angle(value: &str) -> Result<u32> {
    let value = value.trim().parse::<u32>().map_err(|_error| {
        Error::Invalid(format!(
            "DrawingML positive angles must be unsigned 60000ths of a degree: {value:?}"
        ))
    })?;
    positive_angle(value)
}

/// Parse the `DrawingML` percentage lexical union.
///
/// The standard stores values as thousandths of a percent. Office also reads
/// a human-readable percent sign form; accepting both keeps the shared owner
/// useful for real-world files while always writing the canonical integer form.
pub(crate) fn parse_scaled(value: &str, kind: &'static str) -> Result<i32> {
    let value = value.trim();
    if let Some(value) = value.strip_suffix('%') {
        let value = value.trim();
        if value.is_empty() {
            return Err(Error::Invalid(format!(
                "DrawingML {kind} has an empty percent value"
            )));
        }

        let (negative, value) = match value.as_bytes().first() {
            Some(b'-') => (true, &value[1..]),
            Some(b'+') => (false, &value[1..]),
            _ => (false, value),
        };
        if value.is_empty() {
            return Err(Error::Invalid(format!(
                "DrawingML {kind} has an empty percent value"
            )));
        }

        let (whole, fraction) = value.split_once('.').unwrap_or((value, ""));
        if fraction.len() > 3 || !fraction.bytes().all(|byte| byte.is_ascii_digit()) {
            return Err(Error::Invalid(format!(
                "DrawingML {kind} percent values have at most three fractional digits"
            )));
        }
        let whole = if whole.is_empty() {
            0
        } else {
            whole.parse::<u64>().map_err(|_error| {
                Error::Invalid(format!("invalid DrawingML {kind} percent: {value:?}%"))
            })?
        };
        let fraction = if fraction.is_empty() {
            0
        } else {
            let parsed = fraction.parse::<u32>().map_err(|_error| {
                Error::Invalid(format!("invalid DrawingML {kind} percent: {value:?}%"))
            })?;
            let digits = u32::try_from(fraction.len()).map_err(|error| {
                Error::Invalid(format!(
                    "DrawingML {kind} percent precision is invalid: {error}"
                ))
            })?;
            parsed * 10_u32.pow(3 - digits)
        };
        let scaled = whole
            .checked_mul(1000)
            .and_then(|whole| whole.checked_add(u64::from(fraction)))
            .ok_or_else(|| Error::Invalid(format!("DrawingML {kind} percent overflows")))?;
        let signed = i64::try_from(scaled)
            .ok()
            .and_then(|scaled| {
                if negative {
                    scaled.checked_neg()
                } else {
                    Some(scaled)
                }
            })
            .and_then(|scaled| i32::try_from(scaled).ok())
            .ok_or_else(|| Error::Invalid(format!("DrawingML {kind} percent overflows")))?;
        return Ok(signed);
    }

    value.parse::<i32>().map_err(|_error| {
        Error::Invalid(format!(
            "DrawingML {kind} must be an integer thousandth of a percent: {value:?}"
        ))
    })
}

/// Validate and return the original fragment without allocating a copy.
pub(crate) fn validated_fragment(xml: &[u8], limits: Limits) -> Result<&[u8]> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(Error::limit(
            crate::LimitKind::XmlBytes,
            xml.len(),
            limits.max_xml_bytes,
        ));
    }

    let mut reader = Reader::from_reader(xml);
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut nodes = 0usize;
    let mut transforms = 0usize;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("DrawingML color node count overflow".into()))?;
                if nodes > limits.max_nodes {
                    return Err(Error::limit(
                        crate::LimitKind::XmlNodes,
                        nodes,
                        limits.max_nodes,
                    ));
                }
                if is_transform_name(element.local_name().as_ref()) {
                    transforms = transforms.checked_add(1).ok_or_else(|| {
                        Error::Invalid("DrawingML color transform count overflow".into())
                    })?;
                    if transforms > limits.max_transforms {
                        return Err(Error::limit(
                            crate::LimitKind::ColorTransforms,
                            transforms,
                            limits.max_transforms,
                        ));
                    }
                }
                if stack.is_empty() {
                    if root_seen {
                        return Err(Error::Invalid(
                            "DrawingML color fragment contains multiple roots".into(),
                        ));
                    }
                    root_seen = true;
                }
                if stack.len() >= limits.max_depth {
                    return Err(Error::limit(
                        crate::LimitKind::XmlDepth,
                        stack.len().saturating_add(1),
                        limits.max_depth,
                    ));
                }
                stack.push(element.name().as_ref().to_vec());
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("DrawingML color node count overflow".into()))?;
                if nodes > limits.max_nodes {
                    return Err(Error::limit(
                        crate::LimitKind::XmlNodes,
                        nodes,
                        limits.max_nodes,
                    ));
                }
                if is_transform_name(element.local_name().as_ref()) {
                    transforms = transforms.checked_add(1).ok_or_else(|| {
                        Error::Invalid("DrawingML color transform count overflow".into())
                    })?;
                    if transforms > limits.max_transforms {
                        return Err(Error::limit(
                            crate::LimitKind::ColorTransforms,
                            transforms,
                            limits.max_transforms,
                        ));
                    }
                }
                if stack.is_empty() {
                    if root_seen {
                        return Err(Error::Invalid(
                            "DrawingML color fragment contains multiple roots".into(),
                        ));
                    }
                    root_seen = true;
                    root_closed = true;
                }
                if stack.len() >= limits.max_depth {
                    return Err(Error::limit(
                        crate::LimitKind::XmlDepth,
                        stack.len().saturating_add(1),
                        limits.max_depth,
                    ));
                }
            },
            Event::End(element) => {
                let Some(expected) = stack.pop() else {
                    return Err(Error::Invalid(
                        "DrawingML color fragment has an unmatched closing element".into(),
                    ));
                };
                if expected.as_slice() != element.name().as_ref() {
                    return Err(Error::Invalid(
                        "DrawingML color fragment has mismatched closing elements".into(),
                    ));
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text) if stack.is_empty() => {
                if !text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    return Err(Error::Invalid(
                        "DrawingML color fragment contains text outside its root".into(),
                    ));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if stack.is_empty() => {
                return Err(Error::Invalid(
                    "DrawingML color fragment contains data outside its root".into(),
                ));
            },
            Event::Decl(_) | Event::DocType(_) => {
                return Err(Error::Invalid(
                    "DrawingML color fragment cannot contain an XML declaration or doctype".into(),
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(Error::Invalid(
            "DrawingML color fragment must contain one complete root".into(),
        ));
    }
    Ok(xml)
}

fn is_transform_name(name: &[u8]) -> bool {
    matches!(
        name,
        b"alpha"
            | b"alphaMod"
            | b"alphaOff"
            | b"blue"
            | b"blueMod"
            | b"blueOff"
            | b"comp"
            | b"gamma"
            | b"gray"
            | b"green"
            | b"greenMod"
            | b"greenOff"
            | b"hue"
            | b"hueMod"
            | b"hueOff"
            | b"inv"
            | b"invGamma"
            | b"lum"
            | b"lumMod"
            | b"lumOff"
            | b"red"
            | b"redMod"
            | b"redOff"
            | b"sat"
            | b"satMod"
            | b"satOff"
            | b"shade"
            | b"tint"
    )
}

#[cfg(test)]
mod envelope_tests {
    //! What the color fragment check accepts before, around and after its one element, and the exact error reported otherwise. Each
    //! expected line is the row label followed by the result.

    const DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;

    /// One reader's root element in the forms the rows are built from.
    struct Root<'a> {
        /// The start tag.
        open: &'a str,
        /// Content that makes the element valid for its reader.
        body: &'a str,
        /// The end tag.
        close: &'a str,
        /// The element as one empty-element tag.
        empty: &'a str,
    }

    fn rows(root: &Root<'_>) -> Vec<(&'static str, String)> {
        let Root {
            open,
            body,
            close,
            empty,
        } = *root;
        let document = format!("{open}{body}{close}");
        let d = document.as_str();
        vec![
            ("empty input", String::new()),
            ("whitespace only", " \n".to_owned()),
            ("declaration only", DECLARATION.to_owned()),
            ("empty-element root", empty.to_owned()),
            ("start and end root", format!("{open}{close}")),
            ("root then whitespace", format!("{d}\n ")),
            ("root then comment", format!("{d}<!--c-->")),
            ("root then PI", format!("{d}<?pi x?>")),
            ("root then text", format!("{d}text")),
            ("root then second root", format!("{d}{d}")),
            ("root then foreign start", format!("{d}<foreign></foreign>")),
            ("root then foreign empty", format!("{d}<foreign/>")),
            ("declaration after root", format!("{d}{DECLARATION}")),
            ("two declarations", format!("{DECLARATION}{DECLARATION}{d}")),
            (
                "comment before declaration",
                format!("<!--c-->{DECLARATION}{d}"),
            ),
            ("text before root", format!("text{d}")),
            ("unterminated root", format!("{open}{body}")),
            ("stray end tag", format!("{d}</foreign>")),
            ("DOCTYPE", format!("<!DOCTYPE x>{d}")),
            ("top-level CDATA", format!("<![CDATA[x]]>{d}")),
            ("declaration then root", format!("{DECLARATION}{d}")),
            (
                "whitespace before declaration",
                format!(" {DECLARATION}{d}"),
            ),
            ("PI before declaration", format!("<?pi x?>{DECLARATION}{d}")),
            ("comment then root", format!("<!--c-->{d}")),
            ("PI then root", format!("<?pi x?>{d}")),
            ("root then text then second root", format!("{d}text{d}")),
            ("empty-element root then root", format!("{empty}{d}")),
            (
                "declaration after empty-element root",
                format!("{empty}{DECLARATION}"),
            ),
            ("reference before root", format!("&amp;{d}")),
            ("reference after root", format!("{d}&amp;")),
            ("CDATA after root", format!("{d}<![CDATA[x]]>")),
            ("end tag before root", format!("</foreign>{d}")),
            (
                "declaration inside root",
                format!("{open}{DECLARATION}{body}{close}"),
            ),
            (
                "text, CDATA and reference inside root",
                format!("{open}{body}text<![CDATA[x]]>&amp;{close}"),
            ),
        ]
    }

    fn assert_matrix(root: &Root<'_>, expected: &[&str], read: impl Fn(&str) -> String) {
        let actual = rows(root)
            .into_iter()
            .map(|(label, xml)| format!("{label}: {}", read(&xml)))
            .collect::<Vec<_>>();
        assert_eq!(actual, expected, "actual rows:\n{actual:#?}");
    }

    fn outcome<T, E: std::fmt::Debug>(
        result: Result<T, E>,
        summary: impl Fn(T) -> String,
    ) -> String {
        match result {
            Ok(value) => format!("Ok({})", summary(value)),
            Err(error) => format!("{error:?}"),
        }
    }

    /// The color fragment check.
    #[test]
    fn color_fragment_envelope() {
        let root = Root {
            open: r#"<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="FF0000">"#,
            body: r#"<a:alpha val="50000"/>"#,
            close: r"</a:srgbClr>",
            empty: r#"<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="FF0000"/>"#,
        };
        assert_matrix(
            &root,
            &[
                "empty input: Invalid(\"DrawingML color fragment must contain one complete root\")",
                "whitespace only: Invalid(\"DrawingML color fragment must contain one complete root\")",
                "declaration only: Invalid(\"DrawingML color fragment cannot contain an XML declaration or doctype\")",
                "empty-element root: Ok(89)",
                "start and end root: Ok(100)",
                "root then whitespace: Ok(124)",
                "root then comment: Ok(130)",
                "root then PI: Ok(130)",
                "root then text: Invalid(\"DrawingML color fragment contains text outside its root\")",
                "root then second root: Invalid(\"DrawingML color fragment contains multiple roots\")",
                "root then foreign start: Invalid(\"DrawingML color fragment contains multiple roots\")",
                "root then foreign empty: Invalid(\"DrawingML color fragment contains multiple roots\")",
                "declaration after root: Invalid(\"DrawingML color fragment cannot contain an XML declaration or doctype\")",
                "two declarations: Invalid(\"DrawingML color fragment cannot contain an XML declaration or doctype\")",
                "comment before declaration: Invalid(\"DrawingML color fragment cannot contain an XML declaration or doctype\")",
                "text before root: Invalid(\"DrawingML color fragment contains text outside its root\")",
                "unterminated root: Invalid(\"DrawingML color fragment must contain one complete root\")",
                "stray end tag: Xml(\"ill-formed document: close tag `</foreign>` does not match any open tag\")",
                "DOCTYPE: Invalid(\"DrawingML color fragment cannot contain an XML declaration or doctype\")",
                "top-level CDATA: Invalid(\"DrawingML color fragment contains data outside its root\")",
                "declaration then root: Invalid(\"DrawingML color fragment cannot contain an XML declaration or doctype\")",
                "whitespace before declaration: Invalid(\"DrawingML color fragment cannot contain an XML declaration or doctype\")",
                "PI before declaration: Invalid(\"DrawingML color fragment cannot contain an XML declaration or doctype\")",
                "comment then root: Ok(130)",
                "PI then root: Ok(130)",
                "root then text then second root: Invalid(\"DrawingML color fragment contains text outside its root\")",
                "empty-element root then root: Invalid(\"DrawingML color fragment contains multiple roots\")",
                "declaration after empty-element root: Invalid(\"DrawingML color fragment cannot contain an XML declaration or doctype\")",
                "reference before root: Invalid(\"DrawingML color fragment contains data outside its root\")",
                "reference after root: Invalid(\"DrawingML color fragment contains data outside its root\")",
                "CDATA after root: Invalid(\"DrawingML color fragment contains data outside its root\")",
                "end tag before root: Xml(\"ill-formed document: close tag `</foreign>` does not match any open tag\")",
                "declaration inside root: Invalid(\"DrawingML color fragment cannot contain an XML declaration or doctype\")",
                "text, CDATA and reference inside root: Ok(144)",
            ],
            |xml| {
                outcome(
                    super::validated_fragment(xml.as_bytes(), super::Limits::DEFAULT),
                    |bytes| bytes.len().to_string(),
                )
            },
        );
    }
}
