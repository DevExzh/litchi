//! Office Math equations in `DrawingML` text bodies.
//!
//! PowerPoint and Excel write an equation as `a14:m` inside a paragraph of a
//! shape's text body ([MS-ODRAWXML] §2.3.3.41; §3.5 for the example), in the
//! `mc:Choice Requires="a14"` branch of a markup-compatibility block
//! (ECMA-376 Part 3) whose `mc:Fallback` shows readers without Office Math
//! a picture of the equation or its linear text. The hosts differ in where
//! the block sits (around a slide shape, around a worksheet drawing anchor)
//! and share everything this module holds: the OMML a new or replaced
//! equation is written as ([`prepare`]), with the run properties the Office
//! applications put on every math run ([`HostRun`]); the picture of the
//! equation ([`render`], [`fallback_picture`]); the linear-text fallback run
//! ([`linear_run`]); and small checks of the XML an edit stages.
//!
//! Conversions go through `litchi-equation` and its default
//! [`litchi_equation::Limits`]. Requires the `equation` feature.

use std::fmt;

use litchi_equation::{Justification, Limits, OmmlHost, PngOptions};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::Reader;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

/// The markup-compatibility namespace (ECMA-376 Part 3).
pub const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
/// The Office 2010 drawing namespace that holds `a14:m`.
pub const A14: &str = "http://schemas.microsoft.com/office/drawing/2010/main";
/// The Office Math namespace.
pub const MATH: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
/// The transitional `DrawingML` main namespace.
pub const TRANSITIONAL_DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
/// The relationship namespaces (transitional and strict).
const RELATIONSHIPS: [&[u8]; 2] = [
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships",
];

/// Pixels per CSS pixel of the fallback picture (384 pixels per inch).
const PICTURE_SCALE: f64 = 4.0;
/// The most pixels of a fallback picture; a larger frame is drawn at a
/// lower resolution.
const MAX_FALLBACK_PIXELS: f64 = 4_194_304.0;
/// EMU per point (ECMA-376 Part 1, §20.1.10.16).
pub const EMU_PER_POINT: f64 = 12_700.0;
/// The font the Office applications name on every math run.
pub const DEFAULT_RUN_FONTS: &str = r#"<a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/>"#;
/// The picture-fill locks of a fallback shape, as PowerPoint writes them.
pub const FALLBACK_LOCKS: &str = r#"<a:spLocks noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noTextEdit="1"/>"#;
/// The fill of a fallback picture: the whole frame, as PowerPoint writes it
/// for a picture of the shape's text body.
pub const FALLBACK_FILL: &str = "<a:stretch><a:fillRect/></a:stretch>";
/// The default text-body insets in EMU: left, top, right and bottom
/// (ECMA-376 Part 1, §21.1.2.1.1 `bodyPr`).
pub const DEFAULT_INSETS: [i64; 4] = [91_440, 45_720, 91_440, 45_720];

/// What an equation is written from.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum EquationSource<'a> {
    /// LaTeX math, converted to OMML through `litchi-equation`.
    Latex(&'a str),
    /// An OMML `m:oMath` or `m:oMathPara` element, written as given. The
    /// conventional `m` (math) and `a` (DrawingML) prefixes may be left
    /// undeclared.
    Omml(&'a str),
}

impl<'a> EquationSource<'a> {
    /// The caller's text: the LaTeX or the OMML.
    #[must_use]
    pub const fn text(self) -> &'a str {
        match self {
            Self::Latex(text) | Self::Omml(text) => text,
        }
    }
}

/// Why an equation could not be prepared or checked.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// Malformed XML, with the reader's message.
    Xml(String),
    /// Markup or input this module refuses, with the reason.
    Invalid(&'static str),
    /// Caller OMML refers to a relationship (`r:id`, `r:embed` and the
    /// like), which the host part has none of for it.
    RelationshipReference,
    /// The equation does not convert or render.
    Equation(litchi_equation::Error),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Xml(message) => formatter.write_str(message),
            Self::Invalid(message) => formatter.write_str(message),
            Self::RelationshipReference => {
                formatter.write_str("equation OMML cannot refer to relationships")
            },
            Self::Equation(error) => fmt::Display::fmt(error, formatter),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Equation(error) => Some(error),
            Self::Xml(_) | Self::Invalid(_) | Self::RelationshipReference => None,
        }
    }
}

impl From<litchi_equation::Error> for Error {
    fn from(error: litchi_equation::Error) -> Self {
        Self::Equation(error)
    }
}

/// Result type of this module.
pub type Result<T> = std::result::Result<T, Error>;

const fn invalid(message: &'static str) -> Error {
    Error::Invalid(message)
}

fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

/// The host run properties every new math run gets: the attributes and the
/// child elements of an `a:rPr` (ECMA-376 Part 1, §21.1.2.3.9
/// `CT_TextCharacterProperties`), as raw XML with the `a` prefix.
#[derive(Debug, Clone)]
pub struct HostRun {
    attributes: Vec<(String, String)>,
    /// Each child's local name and raw XML, in order.
    children: Vec<(String, String)>,
}

/// The fill choices of `CT_TextCharacterProperties`.
const FILLS: [&str; 6] = [
    "noFill",
    "solidFill",
    "gradFill",
    "blipFill",
    "pattFill",
    "grpFill",
];

impl HostRun {
    /// What PowerPoint and Excel write, with `size` in hundredths of a
    /// point.
    #[must_use]
    pub fn native(size: Option<u32>) -> Self {
        let mut attributes = vec![
            ("lang".to_owned(), "en-US".to_owned()),
            ("i".to_owned(), "1".to_owned()),
        ];
        if let Some(size) = size {
            attributes.insert(1, ("sz".to_owned(), size.to_string()));
        }
        Self {
            attributes,
            children: vec![("latin".to_owned(), DEFAULT_RUN_FONTS.to_owned())],
        }
    }

    /// The text size, in points.
    #[must_use]
    pub fn font_size_pt(&self) -> Option<f64> {
        self.attributes
            .iter()
            .find(|(key, _)| key == "sz")
            .and_then(|(_, value)| value.parse::<u32>().ok())
            .filter(|size| (100..=400_000).contains(size))
            .map(|size| f64::from(size) / 100.0)
    }

    /// The first host `a:rPr` of a math run in `omml`; `None` when there is
    /// none or it, or one of its children, uses another prefix than `a`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Xml`] for malformed XML.
    pub fn of(omml: &str) -> Result<Option<Self>> {
        let mut reader = Reader::from_reader(omml.as_bytes());
        let mut math_prefix: Option<Vec<u8>> = None;
        let mut path: Vec<Vec<u8>> = Vec::new();
        // The template found, the path depth of its element while it is
        // open, and the start of its open child.
        let mut found: Option<(Self, usize)> = None;
        let mut child: Option<(usize, String)> = None;
        loop {
            let start = position(&reader)?;
            let event = reader.read_event().map_err(xml_error)?;
            let end = position(&reader)?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    let name = element.name();
                    let prefix = name.prefix().map(|prefix| prefix.as_ref().to_vec());
                    let local = String::from_utf8_lossy(name.local_name().as_ref()).into_owned();
                    let empty = matches!(event, Event::Empty(_));
                    if math_prefix.is_none() {
                        math_prefix = Some(prefix.clone().unwrap_or_default());
                    }
                    if let Some((run, level)) = found.as_mut() {
                        if path.len() == *level + 1 {
                            if prefix.as_deref() != Some(b"a") {
                                return Ok(None);
                            }
                            if empty {
                                run.children.push((local, omml[start..end].to_owned()));
                            } else {
                                child = Some((start, local));
                            }
                        }
                    } else if local == "rPr"
                        && prefix.as_deref() == Some(b"a")
                        && path.last().is_some_and(|parent| {
                            math_prefix
                                .as_ref()
                                .is_some_and(|math| parent.as_slice() == qualified(math, b"r"))
                        })
                    {
                        let mut attributes = Vec::new();
                        for attribute in element.checked_attributes() {
                            let attribute = attribute.map_err(xml_error)?;
                            let key =
                                std::str::from_utf8(attribute.key.as_ref()).map_err(xml_error)?;
                            let value = std::str::from_utf8(&attribute.value).map_err(xml_error)?;
                            attributes.push((key.to_owned(), value.to_owned()));
                        }
                        let run = Self {
                            attributes,
                            children: Vec::new(),
                        };
                        if empty {
                            return Ok(Some(run));
                        }
                        found = Some((run, path.len()));
                    }
                    if !empty {
                        path.push(name.as_ref().to_vec());
                    }
                },
                Event::End(_) => {
                    path.pop();
                    if let Some((run, level)) = found.as_mut() {
                        if path.len() == *level + 1
                            && let Some((child_start, local)) = child.take()
                        {
                            run.children
                                .push((local, omml[child_start..end].to_owned()));
                        } else if path.len() == *level {
                            return Ok(found.map(|(run, _)| run));
                        }
                    }
                },
                Event::Eof => return Ok(None),
                _ => {},
            }
        }
    }

    /// An `a:rPr` element: these properties with `overrides` (attributes
    /// the converter wrote) on top, and `fills` (the converter's fill) in
    /// place of this template's fill. `upright` drops the italic switch.
    #[must_use]
    pub fn element(&self, overrides: &[(String, String)], fills: &str, upright: bool) -> String {
        let mut out = String::from("<a:rPr");
        for (key, value) in &self.attributes {
            if overrides.iter().any(|(name, _)| name == key) || (upright && key == "i") {
                continue;
            }
            push_attribute(&mut out, key, value);
        }
        for (key, value) in overrides {
            push_attribute(&mut out, key, value);
        }
        let mut children = String::new();
        let mut filled = fills.is_empty();
        for (local, raw) in &self.children {
            let is_fill = FILLS.contains(&local.as_str());
            // The fill follows only the outline (`a:ln`).
            if !filled && local != "ln" {
                children.push_str(fills);
                filled = true;
            }
            if is_fill && !fills.is_empty() {
                continue;
            }
            children.push_str(raw);
        }
        if !filled {
            children.push_str(fills);
        }
        if children.is_empty() {
            out.push_str("/>");
        } else {
            out.push('>');
            out.push_str(&children);
            out.push_str("</a:rPr>");
        }
        out
    }
}

/// Append ` key="value"` to `out`. Values come from parsed XML and are
/// still escaped; only a quote of the other kind can need escaping here.
pub fn push_attribute(out: &mut String, key: &str, value: &str) {
    out.push(' ');
    out.push_str(key);
    out.push_str("=\"");
    out.push_str(&value.replace('"', "&quot;"));
    out.push('"');
}

fn qualified(prefix: &[u8], local: &[u8]) -> Vec<u8> {
    if prefix.is_empty() {
        local.to_vec()
    } else {
        let mut out = prefix.to_vec();
        out.push(b':');
        out.extend_from_slice(local);
        out
    }
}

fn position<R>(reader: &Reader<R>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_error| invalid("the equation is too large"))
}

/// Give every run of converter-written OMML the host run properties the
/// Office applications write: `run` for runs without any, merged under the
/// converter's own `a:rPr` (colour, bold, italic) where it wrote one.
///
/// # Errors
///
/// Returns [`Error::Xml`] for malformed XML.
pub fn apply_host_runs(omml: &str, run: &HostRun) -> Result<String> {
    let mut reader = Reader::from_reader(omml.as_bytes());
    let mut out = String::with_capacity(omml.len() * 2);
    let mut copied = 0usize;
    // For each open element: its name, and for an `m:r`, whether its host
    // properties are written and whether its `m:rPr` makes it upright.
    let mut stack: Vec<(Vec<u8>, bool, bool)> = Vec::new();
    // An open converter `a:rPr`: its start, its attributes and where its
    // children start.
    let mut pending: Option<(usize, Vec<(String, String)>, usize, bool)> = None;
    loop {
        let start = position(&reader)?;
        let event = reader.read_event().map_err(xml_error)?;
        let end = position(&reader)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let name = element.name().as_ref().to_vec();
                let empty = matches!(event, Event::Empty(_));
                if let Some((parent, written, upright)) = stack.last_mut()
                    && parent.as_slice() == b"m:r"
                    && !*written
                {
                    match name.as_slice() {
                        b"m:rPr" => {
                            let tail = &omml[start..];
                            let close = if empty {
                                end - start
                            } else {
                                tail.find("</m:rPr>").map_or(0, |at| at + "</m:rPr>".len())
                            };
                            let properties = &tail[..close];
                            *upright = properties.contains("<m:nor/>")
                                || properties.contains(r#"m:val="p""#)
                                || properties.contains(r#"m:val="b""#);
                        },
                        b"a:rPr" => {
                            *written = true;
                            let mut attributes = Vec::new();
                            for attribute in element.checked_attributes() {
                                let attribute = attribute.map_err(xml_error)?;
                                attributes.push((
                                    String::from_utf8_lossy(attribute.key.as_ref()).into_owned(),
                                    String::from_utf8_lossy(&attribute.value).into_owned(),
                                ));
                            }
                            if empty {
                                out.push_str(&omml[copied..start]);
                                out.push_str(&run.element(&attributes, "", *upright));
                                copied = end;
                            } else {
                                pending = Some((start, attributes, end, *upright));
                            }
                        },
                        _ => {
                            *written = true;
                            out.push_str(&omml[copied..start]);
                            out.push_str(&run.element(&[], "", *upright));
                            copied = start;
                        },
                    }
                }
                if !empty {
                    stack.push((name, false, false));
                }
            },
            Event::End(ref element) => {
                stack.pop();
                if element.name().as_ref() == b"a:rPr"
                    && let Some((open, attributes, children, upright)) = pending.take()
                {
                    out.push_str(&omml[copied..open]);
                    out.push_str(&run.element(&attributes, &omml[children..start], upright));
                    copied = end;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    out.push_str(&omml[copied..]);
    Ok(out)
}

/// Namespace prefixes in scope and their URIs.
pub type Bindings = Vec<(Vec<u8>, Vec<u8>)>;

/// The namespace prefixes the root element of `xml` binds, and their URIs.
///
/// # Errors
///
/// Returns [`Error::Xml`] for malformed XML.
pub fn root_bindings(xml: &[u8]) -> Result<Bindings> {
    let mut reader = Reader::from_reader(xml);
    loop {
        match reader.read_event().map_err(xml_error)? {
            Event::Start(ref element) | Event::Empty(ref element) => {
                return Ok(element
                    .checked_attributes()
                    .flatten()
                    .filter_map(|attribute| {
                        attribute
                            .key
                            .as_ref()
                            .strip_prefix(b"xmlns:")
                            .map(|prefix| (prefix.to_vec(), attribute.value.into_owned()))
                    })
                    .collect());
            },
            Event::Eof => return Ok(Vec::new()),
            _ => {},
        }
    }
}

/// Check caller OMML and declare the `m` and `a` prefixes it leaves
/// undeclared and the host part does not bind (`bound`) to the same
/// namespace, so it can stand inside the part. `dml` is the part's
/// `DrawingML` namespace.
///
/// # Errors
///
/// Returns [`Error::Xml`] for malformed OMML and [`Error::Invalid`] for
/// OMML that is not one `m:oMath` or `m:oMathPara` element or uses an
/// undeclared prefix, and [`Error::RelationshipReference`] for OMML that
/// refers to relationships.
pub fn admit_omml(omml: &str, dml: &str, bound: &Bindings) -> Result<String> {
    let omml = omml.trim();
    let mut reader = NsReader::from_reader(omml.as_bytes());
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut root_name_end = None;
    let mut undeclared: Vec<Vec<u8>> = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("the equation is too large"))?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(format!("malformed OMML: {error}")))?;
        let note = |undeclared: &mut Vec<Vec<u8>>, namespace: &ResolveResult<'_>| {
            if let ResolveResult::Unknown(prefix) = namespace
                && !undeclared.contains(prefix)
            {
                undeclared.push(prefix.clone());
            }
        };
        let math = matches!(&namespace, ResolveResult::Bound(Namespace(value)) if *value == MATH.as_bytes())
            || matches!(&namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"m");
        note(&mut undeclared, &namespace);
        drop(namespace);
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                for attribute in element.checked_attributes() {
                    let attribute = attribute
                        .map_err(|error| Error::Xml(format!("malformed OMML: {error}")))?;
                    let key = attribute.key;
                    if key.as_ref() == b"xmlns" || key.as_ref().starts_with(b"xmlns:") {
                        continue;
                    }
                    let (namespace, _) = reader.resolver().resolve_attribute(key);
                    let uri = match &namespace {
                        ResolveResult::Bound(Namespace(uri)) => Some(uri.to_vec()),
                        ResolveResult::Unknown(prefix) => bound
                            .iter()
                            .find(|(name, _)| name == prefix)
                            .map(|(_, uri)| uri.clone()),
                        ResolveResult::Unbound => None,
                    };
                    if uri.is_some_and(|uri| RELATIONSHIPS.contains(&uri.as_slice())) {
                        return Err(Error::RelationshipReference);
                    }
                    note(&mut undeclared, &namespace);
                }
                if depth == 0 {
                    roots += 1;
                    let local = element.local_name();
                    if !matches!(local.as_ref(), b"oMath" | b"oMathPara") {
                        return Err(invalid(
                            "equation OMML must be one m:oMath or m:oMathPara element",
                        ));
                    }
                    if !math {
                        return Err(invalid("equation OMML must use the Office Math namespace"));
                    }
                    root_name_end = Some(start + 1 + element.name().as_ref().len());
                }
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                }
            },
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Text(ref text) if depth == 0 => {
                if text.iter().any(|byte| !byte.is_ascii_whitespace()) {
                    return Err(invalid("equation OMML has text outside its root"));
                }
            },
            Event::DocType(_) | Event::PI(_) | Event::Decl(_) => {
                return Err(invalid(
                    "equation OMML cannot hold declarations or processing instructions",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let at = root_name_end.ok_or_else(|| invalid("equation OMML is empty"))?;
    if depth != 0 {
        return Err(Error::Xml(
            "malformed OMML: an element is not closed".to_owned(),
        ));
    }
    if roots != 1 {
        return Err(invalid(
            "equation OMML must be one m:oMath or m:oMathPara element",
        ));
    }
    let mut declarations = String::new();
    for prefix in undeclared {
        let inherited = bound
            .iter()
            .find(|(name, _)| *name == prefix)
            .map(|(_, uri)| uri.as_slice());
        match (prefix.as_slice(), inherited) {
            (b"m", Some(uri)) if uri == MATH.as_bytes() => {},
            (b"a", Some(uri)) if uri == dml.as_bytes() => {},
            (b"m", _) => declarations.push_str(&format!(r#" xmlns:m="{MATH}""#)),
            (b"a", _) => declarations.push_str(&format!(r#" xmlns:a="{dml}""#)),
            (_, Some(_)) => {},
            (_, None) => {
                return Err(invalid(
                    "equation OMML uses a namespace prefix it does not declare",
                ));
            },
        }
    }
    let mut out = String::with_capacity(omml.len() + declarations.len());
    out.push_str(&omml[..at]);
    out.push_str(&declarations);
    out.push_str(&omml[at..]);
    Ok(out)
}

/// The OMML form of an equation: a display paragraph with its
/// justification, or an inline `m:oMath`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Form {
    /// An `m:oMath`.
    Inline,
    /// An `m:oMathPara`, with the justification it names.
    Display(Option<Justification>),
}

impl Form {
    /// The form of `omml`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Xml`] for malformed XML.
    pub fn of(omml: &str) -> Result<Self> {
        let mut reader = Reader::from_reader(omml.as_bytes());
        let mut root = true;
        loop {
            match reader.read_event().map_err(xml_error)? {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    let local = element.local_name();
                    if root {
                        if local.as_ref() != b"oMathPara" {
                            return Ok(Self::Inline);
                        }
                        root = false;
                    } else if local.as_ref() == b"jc" {
                        let value = element
                            .checked_attributes()
                            .flatten()
                            .find(|attribute| attribute.key.local_name().as_ref() == b"val")
                            .map(|attribute| attribute.value.into_owned());
                        return Ok(Self::Display(match value.as_deref() {
                            Some(b"left") => Some(Justification::Left),
                            Some(b"right") => Some(Justification::Right),
                            Some(b"center") => Some(Justification::Center),
                            Some(b"centerGroup") => Some(Justification::CenterGroup),
                            _ => None,
                        }));
                    } else if local.as_ref() == b"oMath" {
                        return Ok(Self::Display(None));
                    }
                },
                Event::Eof => {
                    return Ok(if root {
                        Self::Inline
                    } else {
                        Self::Display(None)
                    });
                },
                _ => {},
            }
        }
    }
}

/// OMML ready for a host part, with what its conversion lost.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub struct Prepared {
    /// The OMML to write inside `a14:m`.
    pub omml: String,
    /// What the LaTeX conversion could not carry into OMML, one message per
    /// loss; empty for OMML input.
    pub losses: Vec<String>,
}

/// The OMML of `source` for a part whose `DrawingML` namespace is `dml` and
/// whose root binds `bound`: LaTeX converted in `form` with `run` on every
/// math run, or caller OMML admitted with [`admit_omml`]. The result reads
/// back through the converter.
///
/// # Errors
///
/// Returns [`Error::Equation`] for input that does not convert, and the
/// errors of [`admit_omml`].
pub fn prepare(
    source: EquationSource<'_>,
    form: Form,
    run: &HostRun,
    dml: &str,
    bound: &Bindings,
) -> Result<Prepared> {
    let limits = Limits::DEFAULT;
    let (omml, losses) = match source {
        EquationSource::Latex(latex) => {
            let converted = match form {
                Form::Inline => {
                    litchi_equation::latex_to_omml(latex, OmmlHost::DrawingMl, &limits)?
                },
                Form::Display(justification) => litchi_equation::latex_to_omml_paragraph(
                    latex,
                    OmmlHost::DrawingMl,
                    justification,
                    &limits,
                )?,
            };
            let losses = converted.losses().to_vec();
            let mut omml = apply_host_runs(converted.xml(), run)?;
            if dml != TRANSITIONAL_DML {
                omml = omml.replacen(
                    &format!(r#" xmlns:a="{TRANSITIONAL_DML}""#),
                    &format!(r#" xmlns:a="{dml}""#),
                    1,
                );
            }
            (omml, losses)
        },
        EquationSource::Omml(omml) => (admit_omml(omml, dml, bound)?, Vec::new()),
    };
    // The converter must read back what is written, so the reader of the
    // part and every later edit can.
    litchi_equation::omml_to_latex(&omml, &limits)?;
    Ok(Prepared { omml, losses })
}

/// A picture of an equation, the size it covers in the host, in EMU, and
/// the scale it was drawn at.
#[derive(Debug)]
pub struct Picture {
    png: litchi_equation::Png,
    width: i64,
    height: i64,
    scale: f64,
}

impl Picture {
    /// The width the equation covers, in EMU.
    #[must_use]
    pub const fn width(&self) -> i64 {
        self.width
    }

    /// The height the equation covers, in EMU.
    #[must_use]
    pub const fn height(&self) -> i64 {
        self.height
    }
}

/// The pixels of a rendered equation, at most.
const MAX_PICTURE_PIXELS: f64 = 16_000_000.0;
/// The widest or tallest rendered equation, in pixels (the renderer's own
/// bound is 8192).
const MAX_PICTURE_SIDE: f64 = 8_000.0;

/// Draw `omml` at `font_size_pt`.
///
/// # Errors
///
/// Returns [`Error::Equation`] for OMML that does not convert or render.
pub fn render(omml: &str, font_size_pt: f64) -> Result<Picture> {
    let options = PngOptions::new(font_size_pt)?;
    // Draw at PICTURE_SCALE, or lower when that would exceed the pixel
    // bounds: a long equation is drawn coarser instead of refused.
    let (width_pt, height_pt) = litchi_equation::omml_size_pt(omml, &options, &Limits::DEFAULT)?;
    let pixels_per_point = 96.0 / 72.0;
    let (width_px, height_px) = (
        (width_pt * pixels_per_point).max(1.0),
        (height_pt * pixels_per_point).max(1.0),
    );
    let scale = PICTURE_SCALE
        .min(MAX_PICTURE_SIDE / width_px)
        .min(MAX_PICTURE_SIDE / height_px)
        .min((MAX_PICTURE_PIXELS / (width_px * height_px)).sqrt());
    let options = options
        .with_scale(scale)?
        .with_max_pixels(MAX_PICTURE_PIXELS as u64 + 1_000_000)?;
    let png = litchi_equation::omml_to_png(omml, &options, &Limits::DEFAULT)?;
    let (width, height) = png.size_pt();
    let emu = |points: f64| -> i64 {
        // Bounded by the renderer's pixel limit, far below i64::MAX.
        #[allow(
            clippy::cast_possible_truncation,
            reason = "the picture is at most a few thousand points wide"
        )]
        let value = (points * EMU_PER_POINT).ceil() as i64;
        value.max(1)
    };
    Ok(Picture {
        width: emu(width),
        height: emu(height),
        png,
        scale,
    })
}

/// The frame size of a shape that fits a picture inside its text insets.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the frame overflows.
pub fn fitted(picture: &Picture, insets: [i64; 4]) -> Result<(i64, i64)> {
    let [left, top, right, bottom] = insets.map(|inset| inset.max(0));
    let width = picture
        .width
        .checked_add(left)
        .and_then(|value| value.checked_add(right))
        .ok_or_else(|| invalid("the equation frame is too large"))?;
    let height = picture
        .height
        .checked_add(top)
        .and_then(|value| value.checked_add(bottom))
        .ok_or_else(|| invalid("the equation frame is too large"))?;
    Ok((width, height))
}

/// The fallback picture of a shape: its whole frame, transparent, with the
/// equation at its own size inside the text insets, placed by `place`
/// (horizontal and vertical, 0 at the start, 1 at the end) and scaled down
/// only to fit, as PowerPoint draws its picture of the text body. Readers
/// stretch the picture over the frame, so the equation keeps its
/// proportions only when the picture has the frame's.
///
/// # Errors
///
/// Returns [`Error::Equation`] when the picture cannot be composed.
pub fn fallback_picture(
    picture: &Picture,
    (width, height): (i64, i64),
    insets: [i64; 4],
    place: (f64, f64),
) -> Result<Vec<u8>> {
    #![allow(
        clippy::cast_precision_loss,
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "frames and insets are far below 2^52 EMU, canvases a few thousand pixels"
    )]
    let [left, top, right, bottom] = insets.map(|inset| inset.max(0) as f64);
    let (width, height) = (width.max(1) as f64, height.max(1) as f64);
    // Pixels per EMU at the rendering scale (96 pixels per inch at 1x).
    let natural = picture.scale * 96.0 / 914_400.0;
    let mut density = natural;
    let pixels = (width * density).ceil() * (height * density).ceil();
    if pixels > MAX_FALLBACK_PIXELS {
        density *= (MAX_FALLBACK_PIXELS / pixels).sqrt();
    }
    let room = (
        (width - left - right).max(1.0),
        (height - top - bottom).max(1.0),
    );
    let fit = (room.0 / picture.width as f64)
        .min(room.1 / picture.height as f64)
        .min(1.0);
    let zoom = (density / natural * fit).clamp(f64::MIN_POSITIVE, 1.0);
    let shown = (picture.width as f64 * fit, picture.height as f64 * fit);
    let x = left + (room.0 - shown.0).max(0.0) * place.0.clamp(0.0, 1.0);
    let y = top + (room.1 - shown.1).max(0.0) * place.1.clamp(0.0, 1.0);
    let canvas = |extent: f64| ((extent * density).ceil().max(1.0)) as u32;
    Ok(picture
        .png
        .placed(
            canvas(width),
            canvas(height),
            x * density,
            y * density,
            zoom,
        )?
        .into_bytes())
}

/// The horizontal position of an equation in its room, from 0 (left) to 1
/// (right): a display paragraph's justification (centred as a group when it
/// names none), otherwise its paragraph's alignment (`a:pPr@algn`).
///
/// # Errors
///
/// Returns [`Error::Xml`] for malformed OMML.
pub fn horizontal(omml: &str, align: Option<&str>) -> Result<f64> {
    Ok(match Form::of(omml)? {
        Form::Display(Some(Justification::Left)) => 0.0,
        Form::Display(Some(Justification::Right)) => 1.0,
        Form::Display(_) => 0.5,
        Form::Inline => match align {
            Some("ctr") => 0.5,
            Some("r") => 1.0,
            _ => 0.0,
        },
    })
}

/// The vertical position of the text in its room, from 0 (top) to 1
/// (bottom), from the text anchor (`a:bodyPr@anchor`, top when absent).
#[must_use]
pub fn vertical(anchor: Option<&str>) -> f64 {
    match anchor {
        Some("ctr") => 0.5,
        Some("b") => 1.0,
        _ => 0.0,
    }
}

/// The linear form of an equation, as a text fallback shows it:
/// UnicodeMath, else LaTeX, else nothing.
#[must_use]
pub fn linear_text(omml: &str) -> String {
    let limits = Limits::DEFAULT;
    litchi_equation::omml_to_unicode_math(omml, &limits)
        .or_else(|_error| litchi_equation::omml_to_latex(omml, &limits))
        .unwrap_or_default()
}

/// The text of a run that stands for an equation in a text fallback: its
/// linear form ([`linear_text`]), escaped for character data.
#[must_use]
pub fn linear_run(omml: &str) -> String {
    let linear = linear_text(omml);
    let properties = HostRun::of(omml).ok().flatten().map_or_else(
        || r#"<a:rPr lang="en-US"/>"#.to_owned(),
        |run| run.element(&[], "", true),
    );
    format!(
        "<a:r>{properties}<a:t>{}</a:t></a:r>",
        quick_xml::escape::escape(linear.as_str())
    )
}

/// The qualified name of the element whose start tag begins `tag`.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when `tag` is no start tag.
pub fn tag_name(tag: &[u8]) -> Result<&str> {
    let tag = tag
        .strip_prefix(b"<")
        .ok_or_else(|| invalid("the equation block is malformed"))?;
    let end = tag
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || *byte == b'/' || *byte == b'>')
        .unwrap_or(tag.len());
    std::str::from_utf8(&tag[..end]).map_err(|_error| invalid("the equation block is malformed"))
}

/// An element standing outside its markup-compatibility wrappers: `element`
/// (its bytes) with the namespace declarations of the start tags before it
/// in its block (`wrappers`: the `mc:AlternateContent` and `mc:Choice` start
/// tags) that it does not make itself, so every prefix it uses stays
/// declared.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when `element` is no start tag and
/// [`Error::Xml`] for text that is not UTF-8.
pub fn carry_declarations(wrappers: &[u8], element: &[u8]) -> Result<String> {
    let declarations = |tags: &[u8], stop_after_first: bool| -> Result<Vec<(Vec<u8>, Vec<u8>)>> {
        let mut reader = Reader::from_reader(tags);
        let mut out = Vec::new();
        loop {
            match reader.read_event() {
                Ok(Event::Start(ref element) | Event::Empty(ref element)) => {
                    for attribute in element.checked_attributes().flatten() {
                        let key = attribute.key.as_ref();
                        if key == b"xmlns" || key.starts_with(b"xmlns:") {
                            out.push((key.to_vec(), attribute.value.into_owned()));
                        }
                    }
                    if stop_after_first {
                        return Ok(out);
                    }
                },
                // The wrappers are start tags only; their end is not here.
                Ok(Event::Eof) | Err(_) => return Ok(out),
                Ok(_) => {},
            }
        }
    };
    let own = declarations(element, true)?;
    let name = tag_name(element)?;
    let mut carried = String::new();
    let mut seen: Vec<Vec<u8>> = own.iter().map(|(key, _)| key.clone()).collect();
    // The innermost declaration of a prefix wins: walk from the Choice out.
    for (key, value) in declarations(wrappers, false)?.into_iter().rev() {
        if seen.contains(&key) {
            continue;
        }
        let key_text = String::from_utf8(key.clone()).map_err(xml_error)?;
        let value_text = String::from_utf8(value).map_err(xml_error)?;
        push_attribute(&mut carried, &key_text, &value_text);
        seen.push(key);
    }
    let text = std::str::from_utf8(element).map_err(xml_error)?;
    let at = 1 + name.len();
    Ok(format!("{}{carried}{}", &text[..at], &text[at..]))
}

/// Refuse XML that uses a namespace prefix it does not declare, before it
/// is staged: a commit refuses such a part, with every other edit.
///
/// # Errors
///
/// Returns [`Error::Xml`] for malformed XML and [`Error::Invalid`] for an
/// undeclared prefix.
pub fn require_declared_prefixes(xml: &[u8]) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let unknown = matches!(namespace, ResolveResult::Unknown(_));
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                if unknown {
                    return Err(invalid(
                        "the equation edit leaves a namespace prefix undeclared",
                    ));
                }
                for attribute in element.checked_attributes().flatten() {
                    let key = attribute.key;
                    if key.as_ref().starts_with(b"xmlns") || key.prefix().is_none() {
                        continue;
                    }
                    if key.as_ref().starts_with(b"xml:") {
                        continue;
                    }
                    if matches!(
                        reader.resolver().resolve_attribute(key).0,
                        ResolveResult::Unknown(_)
                    ) {
                        return Err(invalid(
                            "the equation edit leaves a namespace prefix undeclared",
                        ));
                    }
                }
            },
            Event::Eof => return Ok(()),
            _ => {},
        }
    }
}

/// The largest `cNvPr` ID anywhere in `xml`, both branches of every
/// markup-compatibility block included; 1 when there is none.
///
/// # Errors
///
/// Returns [`Error::Xml`] for malformed XML.
pub fn max_raw_shape_id(xml: &[u8]) -> Result<u32> {
    let mut reader = Reader::from_reader(xml);
    let mut maximum = 1u32;
    loop {
        match reader.read_event().map_err(xml_error)? {
            Event::Start(ref element) | Event::Empty(ref element)
                if element.local_name().as_ref() == b"cNvPr" =>
            {
                for attribute in element.checked_attributes().flatten() {
                    if attribute.key.as_ref() == b"id"
                        && let Some(id) = std::str::from_utf8(&attribute.value)
                            .ok()
                            .and_then(|value| value.trim().parse::<u32>().ok())
                    {
                        maximum = maximum.max(id);
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(maximum)
}

/// The ID of the first `cNvPr` in `xml`.
///
/// # Errors
///
/// Returns [`Error::Xml`] for malformed XML.
pub fn cnvpr_id(xml: &[u8]) -> Result<Option<u32>> {
    let mut reader = Reader::from_reader(xml);
    loop {
        match reader.read_event().map_err(xml_error)? {
            Event::Start(ref element) | Event::Empty(ref element)
                if element.local_name().as_ref() == b"cNvPr" =>
            {
                return Ok(element
                    .checked_attributes()
                    .flatten()
                    .find(|attribute| attribute.key.as_ref() == b"id")
                    .and_then(|attribute| {
                        std::str::from_utf8(&attribute.value)
                            .ok()
                            .and_then(|value| value.trim().parse().ok())
                    }));
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const DML_BINDING: (&[u8], &[u8]) = (b"a", TRANSITIONAL_DML.as_bytes());

    fn bound() -> Bindings {
        vec![(DML_BINDING.0.to_vec(), DML_BINDING.1.to_vec())]
    }

    #[test]
    fn converter_runs_get_the_host_properties() {
        let run = HostRun::native(Some(2400));
        let omml = r#"<m:oMath xmlns:m="m" xmlns:a="a"><m:r><m:t>x</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>sin</m:t></m:r><m:r><a:rPr b="1"/><m:t>B</m:t></m:r><m:r><a:rPr i="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><m:t>c</m:t></m:r></m:oMath>"#;
        let written = apply_host_runs(omml, &run).unwrap();
        let fonts = DEFAULT_RUN_FONTS;
        assert_eq!(
            written,
            format!(
                concat!(
                    r#"<m:oMath xmlns:m="m" xmlns:a="a">"#,
                    r#"<m:r><a:rPr lang="en-US" sz="2400" i="1">{fonts}</a:rPr><m:t>x</m:t></m:r>"#,
                    r#"<m:r><m:rPr><m:sty m:val="p"/></m:rPr><a:rPr lang="en-US" sz="2400">{fonts}</a:rPr><m:t>sin</m:t></m:r>"#,
                    r#"<m:r><a:rPr lang="en-US" sz="2400" i="1" b="1">{fonts}</a:rPr><m:t>B</m:t></m:r>"#,
                    r#"<m:r><a:rPr lang="en-US" sz="2400" i="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>{fonts}</a:rPr><m:t>c</m:t></m:r>"#,
                    "</m:oMath>"
                ),
                fonts = fonts
            )
        );
    }

    #[test]
    fn the_first_math_run_properties_are_the_template() {
        let omml = r#"<m:oMath><m:r><a:rPr lang="de-DE" sz="3200" i="1"><a:latin typeface="Cambria Math"/></a:rPr><m:t>x</m:t></m:r></m:oMath>"#;
        let run = HostRun::of(omml).unwrap().unwrap();
        assert_eq!(run.font_size_pt(), Some(32.0));
        assert_eq!(
            run.children,
            [(
                "latin".to_owned(),
                r#"<a:latin typeface="Cambria Math"/>"#.to_owned()
            )]
        );
        assert_eq!(
            run.element(&[], "", false),
            r#"<a:rPr lang="de-DE" sz="3200" i="1"><a:latin typeface="Cambria Math"/></a:rPr>"#
        );
        // Fonts after a fill are still found.
        let filled = r#"<m:oMath><m:r><a:rPr i="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>x</m:t></m:r></m:oMath>"#;
        let template = HostRun::of(filled).unwrap().unwrap();
        assert_eq!(template.children.len(), 2);
        // The template's fill is kept, or replaced by the converter's.
        assert_eq!(
            template.element(&[], "", false),
            r#"<a:rPr i="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr>"#
        );
        assert_eq!(
            template.element(&[], "<a:noFill/>", true),
            r#"<a:rPr><a:noFill/><a:latin typeface="Cambria Math"/></a:rPr>"#
        );
        // A run without host properties, or another prefix, is no template.
        assert!(
            HostRun::of("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>")
                .unwrap()
                .is_none()
        );
        assert!(
            HostRun::of(r#"<m:oMath><m:r><w:rPr/><m:t>x</m:t></m:r></m:oMath>"#)
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn caller_omml_is_admitted_with_its_missing_prefixes() {
        let dml = TRANSITIONAL_DML;
        assert_eq!(
            admit_omml(
                "  <m:oMath><m:r><a:rPr/><m:t>x</m:t></m:r></m:oMath> ",
                dml,
                &bound()
            )
            .unwrap(),
            format!(r#"<m:oMath xmlns:m="{MATH}"><m:r><a:rPr/><m:t>x</m:t></m:r></m:oMath>"#)
        );
        assert_eq!(
            admit_omml("<m:oMath><a:rPr/></m:oMath>", dml, &Vec::new()).unwrap(),
            format!(r#"<m:oMath xmlns:m="{MATH}" xmlns:a="{dml}"><a:rPr/></m:oMath>"#)
        );
        let declared = format!(r#"<q:oMathPara xmlns:q="{MATH}"/>"#);
        assert_eq!(admit_omml(&declared, dml, &bound()).unwrap(), declared);
        for bad in [
            "",
            "x",
            "<m:r/>",
            "<m:oMath/><m:oMath/>",
            "<m:oMath><z:r/></m:oMath>",
            "<m:oMath>",
            r#"<?xml version="1.0"?><m:oMath/>"#,
            r#"<m:oMath xmlns:m="urn:other"/>"#,
        ] {
            assert!(admit_omml(bad, dml, &bound()).is_err(), "{bad}");
        }
    }

    #[test]
    fn the_form_of_an_equation_is_kept() {
        assert!(matches!(Form::of("<m:oMath/>").unwrap(), Form::Inline));
        assert!(matches!(
            Form::of(r#"<m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath/></m:oMathPara>"#).unwrap(),
            Form::Display(Some(Justification::Left))
        ));
        assert!(matches!(
            Form::of("<m:oMathPara><m:oMath/></m:oMathPara>").unwrap(),
            Form::Display(None)
        ));
    }

    #[test]
    fn shape_ids_are_read_from_every_branch() {
        assert_eq!(
            max_raw_shape_id(br#"<r><p:cNvPr id="7"/><x><a:cNvPr id="12"/></x></r>"#).unwrap(),
            12
        );
        assert_eq!(max_raw_shape_id(b"<r/>").unwrap(), 1);
        assert_eq!(
            cnvpr_id(br#"<r><x:cNvPr id=" 4 "/><x:cNvPr id="9"/></r>"#).unwrap(),
            Some(4)
        );
        assert_eq!(cnvpr_id(b"<r/>").unwrap(), None);
    }

    #[test]
    fn an_unwrapped_element_keeps_the_declarations_it_uses() {
        let wrappers = br#"<mc:AlternateContent xmlns:mc="m" xmlns:a="outer"><mc:Choice xmlns:a14="x" xmlns:a="inner" Requires="a14">"#;
        assert_eq!(
            carry_declarations(wrappers, br#"<p:sp xmlns:mc="own"><a:t/></p:sp>"#).unwrap(),
            r#"<p:sp xmlns:a="inner" xmlns:a14="x" xmlns:mc="own"><a:t/></p:sp>"#
        );
        assert!(carry_declarations(wrappers, b"text").is_err());
    }

    #[test]
    fn tags_and_prefixes_are_checked() {
        assert_eq!(tag_name(b"<mc:Fallback>").unwrap(), "mc:Fallback");
        assert_eq!(tag_name(b"<a/>").unwrap(), "a");
        assert!(tag_name(b"x").is_err());
        assert!(require_declared_prefixes(br#"<a:r xmlns:a="u"><a:t/></a:r>"#).is_ok());
        assert!(require_declared_prefixes(b"<a:r/>").is_err());
        assert!(require_declared_prefixes(br#"<r b:x="1"/>"#).is_err());
    }

    #[test]
    fn linear_runs_are_escaped_upright_text() {
        let run = linear_run(
            r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><a:rPr lang="en-US" i="1"/><m:t>a&lt;b</m:t></m:r></m:oMath>"#,
        );
        assert!(
            run.starts_with(r#"<a:r><a:rPr lang="en-US"/><a:t>"#),
            "{run}"
        );
        assert!(run.contains("&lt;"), "{run}");
    }
}
