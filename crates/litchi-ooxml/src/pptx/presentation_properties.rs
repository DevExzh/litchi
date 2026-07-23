//! Typed PresentationML presentation properties with inert extension payloads.

use litchi_core::sheet::Result;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::{
    Reader, XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
};

const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const P_STRICT: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const A_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const R_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const P14_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P15_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2012/main";
const DISCARD_IMAGE_EDIT_DATA_URI: &str = "{E76CE94A-603C-4142-B9EB-6D1370010A27}";
const DEFAULT_IMAGE_DPI_URI: &str = "{D31A062A-798A-4329-ABDD-BBA856620510}";
const CHART_TRACKING_REF_BASED_URI: &str = "{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}";
const BROWSE_MODE_URI: &str = "{F99C55AA-B7CB-42B0-86F8-08522FDF87E8}";
const LASER_COLOR_URI: &str = "{EC167BDD-8182-4AB7-AECC-EB403E3ABB37}";
const SHOW_MEDIA_CONTROLS_URI: &str = "{2FDB2607-1784-4EEB-B798-7EB5836EED8A}";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/presProps";
const CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml";
const MAX_BYTES: usize = 8 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 100_000;
const MAX_STRING: usize = 1024 * 1024;
const MAX_EXTENSIONS: usize = 1024;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct InertHtmlTarget {
    pub relationship_id: String,
    pub target: Option<String>,
    pub relationship_type: Option<String>,
    pub external: Option<bool>,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BrowserSupport {
    V3,
    V4,
    V3V4,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum WebScreenSize {
    S544x376,
    S640x480,
    S720x512,
    S800x600,
    S1024x768,
    S1152x882,
    S1152x900,
    S1280x1024,
    S1600x1200,
    S1800x1400,
    S1920x1200,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum WebColor {
    None,
    Browser,
    PresentationText,
    PresentationAccent,
    WhiteTextOnBlack,
    BlackTextOnWhite,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PrintOutput {
    Slides,
    Handouts1,
    Handouts2,
    Handouts3,
    Handouts4,
    Handouts6,
    Handouts9,
    Notes,
    Outline,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PrintColorMode {
    BlackWhite,
    Gray,
    Color,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SlideSelection {
    All,
    Range { start: u32, end: u32 },
    CustomShow(u32),
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ShowMode {
    Present,
    Browse { show_scrollbar: Option<bool> },
    Kiosk { restart: Option<u32> },
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ColorKind {
    ScRgb,
    Srgb,
    Hsl,
    System,
    Scheme,
    Preset,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PresentationColor {
    pub kind: ColorKind,
    pub attributes: Vec<(String, String)>,
    pub xml: Vec<u8>,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OpaquePresentationExtension {
    pub uri: String,
    pub xml: Vec<u8>,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PresentationPropertyExtension {
    DiscardImageEditData(bool),
    DefaultImageDpi(u32),
    ChartTrackingReferenceBased(bool),
    Unknown(OpaquePresentationExtension),
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SlideShowExtension {
    BrowseMode { show_status: Option<bool> },
    LaserColor(PresentationColor),
    ShowMediaControls(bool),
    Unknown(OpaquePresentationExtension),
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HtmlPublishProperties {
    pub show_speaker_notes: Option<bool>,
    pub browser: Option<BrowserSupport>,
    pub target: InertHtmlTarget,
    pub slides: SlideSelection,
    pub extension_xml: Option<Vec<u8>>,
}
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct WebProperties {
    pub show_animation: Option<bool>,
    pub resize_graphics: Option<bool>,
    pub allow_png: Option<bool>,
    pub rely_on_vml: Option<bool>,
    pub organize_in_folders: Option<bool>,
    pub use_long_filenames: Option<bool>,
    pub image_size: Option<WebScreenSize>,
    pub encoding: Option<String>,
    pub color: Option<WebColor>,
    pub extension_xml: Option<Vec<u8>>,
}
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct PrintProperties {
    pub output: Option<PrintOutput>,
    pub color_mode: Option<PrintColorMode>,
    pub hidden_slides: Option<bool>,
    pub scale_to_fit_paper: Option<bool>,
    pub frame_slides: Option<bool>,
    pub extension_xml: Option<Vec<u8>>,
}
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct ShowProperties {
    pub loop_show: Option<bool>,
    pub show_narration: Option<bool>,
    pub show_animation: Option<bool>,
    pub use_timings: Option<bool>,
    pub mode: Option<ShowMode>,
    pub slides: Option<SlideSelection>,
    pub pen_color: Option<PresentationColor>,
    pub extensions: Vec<SlideShowExtension>,
}
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct PresentationProperties {
    pub html_publish: Option<HtmlPublishProperties>,
    pub web: Option<WebProperties>,
    pub print: Option<PrintProperties>,
    pub show: Option<ShowProperties>,
    pub recent_colors: Vec<PresentationColor>,
    pub extensions: Vec<PresentationPropertyExtension>,
}

impl PresentationProperties {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_BYTES {
            return Err(invalid("presentation properties exceed 8 MiB"));
        }
        let processed = crate::common::mce::process_ooxml(xml)?;
        if processed.len() > MAX_BYTES {
            return Err(invalid("processed presentation properties exceed 8 MiB"));
        }
        let root = parse_dom(processed.as_ref())?;
        project(&root)
    }
    pub fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        validate(self)?;
        let p = if strict { P_STRICT } else { P_NS };
        let a = if strict { A_STRICT } else { A_NS };
        let r = if strict { R_STRICT } else { R_NS };
        let mut x = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentationPr xmlns:p="{p}" xmlns:a="{a}" xmlns:r="{r}">"#
        );
        if let Some(v) = &self.html_publish {
            write_html(&mut x, v, strict)?;
        }
        if let Some(v) = &self.web {
            write_web(&mut x, v, strict)?;
        }
        if let Some(v) = &self.print {
            write_print(&mut x, v, strict)?;
        }
        if let Some(v) = &self.show {
            write_show(&mut x, v, strict)?;
        }
        if !self.recent_colors.is_empty() {
            x.push_str("<p:clrMru>");
            for c in &self.recent_colors {
                write_opaque(&mut x, &c.xml, strict)?;
            }
            x.push_str("</p:clrMru>");
        }
        write_presentation_extensions(&mut x, &self.extensions, strict)?;
        x.push_str("</p:presentationPr>");
        if x.len() > MAX_BYTES {
            return Err(invalid("serialized presentation properties exceed 8 MiB"));
        }
        Ok(x.into_bytes())
    }
}

pub fn load_from_package(package: &OpcPackage) -> Result<Option<PresentationProperties>> {
    let presentation = package.main_document_part()?;
    let mut found = presentation
        .rels()
        .iter()
        .filter(|r| matches!(r.reltype(), REL | STRICT_REL));
    let Some(rel) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(invalid(
            "presentation has multiple presentation-properties relationships",
        ));
    }
    if rel.is_external() {
        return Err(invalid(
            "presentation-properties relationship cannot be external",
        ));
    }
    let uri: PackURI = rel.target_partname()?;
    let part = package.get_part(&uri)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(invalid(format!(
            "presentation-properties part '{uri}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    let mut value = PresentationProperties::parse(part.blob())?;
    if let Some(html) = value.html_publish.as_mut() {
        let target = part
            .rels()
            .get(&html.target.relationship_id)
            .ok_or_else(|| {
                invalid(format!(
                    "missing HTML publish relationship '{}'",
                    html.target.relationship_id
                ))
            })?;
        html.target.target = Some(target.target_ref().to_string());
        html.target.relationship_type = Some(target.reltype().to_string());
        html.target.external = Some(target.is_external());
    }
    Ok(Some(value))
}

#[derive(Clone, Debug)]
struct Attr {
    qname: String,
    ns: String,
    local: String,
    value: String,
}
#[derive(Clone, Debug)]
enum Content {
    Node(Node),
    Text(String),
    CData(String),
    Comment(String),
}
#[derive(Clone, Debug)]
struct Node {
    qname: String,
    ns: String,
    local: String,
    attrs: Vec<Attr>,
    bindings: Vec<(String, String)>,
    content: Vec<Content>,
}
fn parse_dom(xml: &[u8]) -> Result<Node> {
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = Reader::from_reader(xml);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    loop {
        let d = reader.decoder();
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(invalid("presentation-properties node limit exceeded"));
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("presentation-properties depth limit exceeded"));
                }
                let node = make_node(&e, d, &stack)?;
                stack.push(node);
            },
            Ok(Event::Empty(e)) => {
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(invalid("presentation-properties node limit exceeded"));
                }
                let node = make_node(&e, d, &stack)?;
                attach(&mut stack, &mut root, node)?;
            },
            Ok(Event::End(_)) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;
                attach(&mut stack, &mut root, node)?;
            },
            Ok(Event::Text(t)) => {
                let text = t.decode().map_err(xml_error)?.into_owned();
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Text(text))
                } else if !text.trim().is_empty() {
                    return Err(invalid("text outside presentationPr"));
                }
            },
            Ok(Event::CData(t)) => {
                let text = t.decode().map_err(xml_error)?.into_owned();
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::CData(text))
                } else {
                    return Err(invalid("CDATA outside presentationPr"));
                }
            },
            Ok(Event::GeneralRef(r)) => {
                let text = crate::common::xml::decode_xml_reference(&r)?;
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Text(text))
                } else {
                    return Err(invalid("entity outside presentationPr"));
                }
            },
            Ok(Event::Comment(c)) => {
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Comment(
                        c.decode().map_err(xml_error)?.into_owned(),
                    ))
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Ok(Event::Decl(_)) => {},
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(e)),
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated presentation-properties XML"));
    }
    root.ok_or_else(|| invalid("missing presentationPr root"))
}
fn make_node(e: &BytesStart<'_>, d: Decoder, stack: &[Node]) -> Result<Node> {
    let q = std::str::from_utf8(e.name().as_ref())
        .map_err(xml_error)?
        .to_string();
    let mut bindings = stack.last().map(|n| n.bindings.clone()).unwrap_or_default();
    let mut raw = Vec::new();
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        raw.push((
            std::str::from_utf8(a.key.as_ref())
                .map_err(xml_error)?
                .to_string(),
            a.decoded_and_normalized_value(XmlVersion::Implicit1_0, d)
                .map_err(xml_error)?
                .into_owned(),
        ));
    }
    for (k, v) in &raw {
        if k == "xmlns" || k.starts_with("xmlns:") {
            let key = k.strip_prefix("xmlns:").unwrap_or("").to_string();
            if let Some(old) = bindings.iter_mut().find(|x| x.0 == key) {
                old.1 = v.clone()
            } else {
                bindings.push((key, v.clone()))
            }
        }
    }
    let (p, l) = split(&q)?;
    let local = l.to_string();
    let ns = resolve(&bindings, p)?;
    let mut attrs = Vec::new();
    for (k, v) in raw {
        if k == "xmlns" || k.starts_with("xmlns:") {
            continue;
        }
        let (ap, al) = split(&k)?;
        let attr_ns = if ap.is_empty() {
            String::new()
        } else {
            resolve(&bindings, ap)?
        };
        let attr_local = al.to_string();
        attrs.push(Attr {
            qname: k,
            ns: attr_ns,
            local: attr_local,
            value: v,
        });
    }
    Ok(Node {
        qname: q,
        ns,
        local,
        attrs,
        bindings,
        content: Vec::new(),
    })
}
fn attach(stack: &mut [Node], root: &mut Option<Node>, node: Node) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.content.push(Content::Node(node));
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn project(root: &Node) -> Result<PresentationProperties> {
    expect(root, P_NS, P_STRICT, "presentationPr")?;
    no_attrs(root)?;
    let mut out = PresentationProperties::default();
    let mut order = 0;
    let mut has_extensions = false;
    for child in children(root)? {
        let index = match child.local.as_str() {
            "htmlPubPr" => 0,
            "webPr" => 1,
            "prnPr" => 2,
            "showPr" => 3,
            "clrMru" => 4,
            "extLst" => 5,
            _ => return Err(invalid("unexpected presentationPr child")),
        };
        expect_p(child)?;
        if index < order {
            return Err(invalid("presentationPr children are out of order"));
        }
        order = index;
        match index {
            0 => {
                if out.html_publish.is_some() {
                    return Err(invalid("duplicate htmlPubPr"));
                }
                out.html_publish = Some(parse_html(child)?);
            },
            1 => {
                if out.web.is_some() {
                    return Err(invalid("duplicate webPr"));
                }
                out.web = Some(parse_web(child)?);
            },
            2 => {
                if out.print.is_some() {
                    return Err(invalid("duplicate prnPr"));
                }
                out.print = Some(parse_print(child)?);
            },
            3 => {
                if out.show.is_some() {
                    return Err(invalid("duplicate showPr"));
                }
                out.show = Some(parse_show(child)?);
            },
            4 => {
                if !out.recent_colors.is_empty() {
                    return Err(invalid("duplicate clrMru"));
                }
                no_attrs(child)?;
                for color in children(child)? {
                    out.recent_colors.push(parse_color(color)?);
                }
                if out.recent_colors.len() > 10 {
                    return Err(invalid("clrMru permits at most ten colors"));
                }
            },
            5 => {
                if has_extensions {
                    return Err(invalid("duplicate presentation extLst"));
                }
                has_extensions = true;
                out.extensions = parse_presentation_extensions(child)?;
            },
            _ => unreachable!(),
        }
    }
    validate(&out)?;
    Ok(out)
}
fn parse_html(n: &Node) -> Result<HtmlPublishProperties> {
    let show_speaker_notes = bool_opt(n, "showSpeakerNotes")?;
    let browser = attr_opt(n, "", "pubBrowser")?
        .map(parse_browser)
        .transpose()?;
    let id = attr_req(n, R_NS, R_STRICT, "id")?;
    only_attrs(
        n,
        &[
            ("", "showSpeakerNotes"),
            ("", "pubBrowser"),
            (R_NS, "id"),
            (R_STRICT, "id"),
        ],
    )?;
    let mut slides = None;
    let mut ext = None;
    for c in children(n)? {
        expect_p(c)?;
        match c.local.as_str() {
            "extLst" => {
                if ext.is_some() {
                    return Err(invalid("duplicate htmlPubPr extLst"));
                }
                ext = Some(node_xml(c, false)?);
            },
            _ => {
                if slides.is_some() || ext.is_some() {
                    return Err(invalid("invalid HTML slide selection order"));
                }
                slides = Some(parse_selection(c)?);
            },
        }
    }
    Ok(HtmlPublishProperties {
        show_speaker_notes,
        browser,
        target: InertHtmlTarget {
            relationship_id: id,
            target: None,
            relationship_type: None,
            external: None,
        },
        slides: slides.ok_or_else(|| invalid("htmlPubPr requires a slide selection"))?,
        extension_xml: ext,
    })
}
fn parse_web(n: &Node) -> Result<WebProperties> {
    let mut v = WebProperties {
        show_animation: bool_opt(n, "showAnimation")?,
        resize_graphics: bool_opt(n, "resizeGraphics")?,
        allow_png: bool_opt(n, "allowPng")?,
        rely_on_vml: bool_opt(n, "relyOnVml")?,
        organize_in_folders: bool_opt(n, "organizeInFolders")?,
        use_long_filenames: bool_opt(n, "useLongFilenames")?,
        image_size: attr_opt(n, "", "imgSz")?.map(parse_screen).transpose()?,
        encoding: attr_opt(n, "", "encoding")?,
        color: attr_opt(n, "", "clr")?.map(parse_web_color).transpose()?,
        extension_xml: None,
    };
    only_attrs(
        n,
        &[
            ("", "showAnimation"),
            ("", "resizeGraphics"),
            ("", "allowPng"),
            ("", "relyOnVml"),
            ("", "organizeInFolders"),
            ("", "useLongFilenames"),
            ("", "imgSz"),
            ("", "encoding"),
            ("", "clr"),
        ],
    )?;
    v.extension_xml = single_ext(n)?;
    Ok(v)
}
fn parse_print(n: &Node) -> Result<PrintProperties> {
    let mut v = PrintProperties {
        output: attr_opt(n, "", "prnWhat")?.map(parse_output).transpose()?,
        color_mode: attr_opt(n, "", "clrMode")?
            .map(parse_print_color)
            .transpose()?,
        hidden_slides: bool_opt(n, "hiddenSlides")?,
        scale_to_fit_paper: bool_opt(n, "scaleToFitPaper")?,
        frame_slides: bool_opt(n, "frameSlides")?,
        extension_xml: None,
    };
    only_attrs(
        n,
        &[
            ("", "prnWhat"),
            ("", "clrMode"),
            ("", "hiddenSlides"),
            ("", "scaleToFitPaper"),
            ("", "frameSlides"),
        ],
    )?;
    v.extension_xml = single_ext(n)?;
    Ok(v)
}
fn parse_show(n: &Node) -> Result<ShowProperties> {
    let mut v = ShowProperties {
        loop_show: bool_opt(n, "loop")?,
        show_narration: bool_opt(n, "showNarration")?,
        show_animation: bool_opt(n, "showAnimation")?,
        use_timings: bool_opt(n, "useTimings")?,
        ..Default::default()
    };
    only_attrs(
        n,
        &[
            ("", "loop"),
            ("", "showNarration"),
            ("", "showAnimation"),
            ("", "useTimings"),
        ],
    )?;
    let mut stage = 0;
    let mut has_extensions = false;
    for c in children(n)? {
        expect_p(c)?;
        match c.local.as_str() {
            "present" | "browse" | "kiosk" => {
                if stage > 0 || v.mode.is_some() {
                    return Err(invalid("invalid show mode order"));
                }
                v.mode = Some(parse_mode(c)?);
            },
            "sldAll" | "sldRg" | "custShow" => {
                if stage > 1 || v.slides.is_some() {
                    return Err(invalid("invalid show selection order"));
                }
                stage = 1;
                v.slides = Some(parse_selection(c)?);
            },
            "penClr" => {
                if stage > 2 || v.pen_color.is_some() {
                    return Err(invalid("invalid penClr order"));
                }
                stage = 2;
                no_attrs(c)?;
                let colors = children(c)?;
                if colors.len() != 1 {
                    return Err(invalid("penClr requires one color"));
                }
                v.pen_color = Some(parse_color(colors[0])?);
            },
            "extLst" => {
                if has_extensions {
                    return Err(invalid("duplicate show extLst"));
                }
                has_extensions = true;
                stage = 3;
                v.extensions = parse_show_extensions(c)?;
            },
            _ => return Err(invalid("unexpected showPr child")),
        }
    }
    Ok(v)
}
fn parse_mode(n: &Node) -> Result<ShowMode> {
    match n.local.as_str() {
        "present" => {
            no_attrs(n)?;
            empty(n)?;
            Ok(ShowMode::Present)
        },
        "browse" => {
            let x = bool_opt(n, "showScrollbar")?;
            only_attrs(n, &[("", "showScrollbar")])?;
            empty(n)?;
            Ok(ShowMode::Browse { show_scrollbar: x })
        },
        "kiosk" => {
            let x = u32_opt(n, "restart")?;
            only_attrs(n, &[("", "restart")])?;
            empty(n)?;
            Ok(ShowMode::Kiosk { restart: x })
        },
        _ => Err(invalid("invalid show mode")),
    }
}
fn parse_selection(n: &Node) -> Result<SlideSelection> {
    match n.local.as_str() {
        "sldAll" => {
            no_attrs(n)?;
            empty(n)?;
            Ok(SlideSelection::All)
        },
        "sldRg" => {
            let start = u32_req(n, "st")?;
            let end = u32_req(n, "end")?;
            only_attrs(n, &[("", "st"), ("", "end")])?;
            empty(n)?;
            if start > end {
                return Err(invalid("slide range start exceeds end"));
            }
            Ok(SlideSelection::Range { start, end })
        },
        "custShow" => {
            let id = u32_req(n, "id")?;
            only_attrs(n, &[("", "id")])?;
            empty(n)?;
            Ok(SlideSelection::CustomShow(id))
        },
        _ => Err(invalid("invalid slide selection")),
    }
}
fn parse_color(n: &Node) -> Result<PresentationColor> {
    let kind = match (n.ns.as_str(), n.local.as_str()) {
        (A_NS | A_STRICT, "scrgbClr") => ColorKind::ScRgb,
        (A_NS | A_STRICT, "srgbClr") => ColorKind::Srgb,
        (A_NS | A_STRICT, "hslClr") => ColorKind::Hsl,
        (A_NS | A_STRICT, "sysClr") => ColorKind::System,
        (A_NS | A_STRICT, "schemeClr") => ColorKind::Scheme,
        (A_NS | A_STRICT, "prstClr") => ColorKind::Preset,
        _ => return Err(invalid("invalid DrawingML color")),
    };
    let attributes = n
        .attrs
        .iter()
        .map(|a| (a.qname.clone(), a.value.clone()))
        .collect();
    Ok(PresentationColor {
        kind,
        attributes,
        xml: node_xml(n, false)?,
    })
}

fn parse_presentation_extensions(n: &Node) -> Result<Vec<PresentationPropertyExtension>> {
    no_attrs(n)?;
    let extensions = children(n)?;
    if extensions.len() > MAX_EXTENSIONS {
        return Err(invalid("presentation extension count exceeds limit"));
    }
    let mut out = Vec::with_capacity(extensions.len());
    let mut discard = false;
    let mut dpi = false;
    let mut tracking = false;
    for ext in extensions {
        let uri = extension_uri(ext)?;
        let value =
            match uri.as_str() {
                DISCARD_IMAGE_EDIT_DATA_URI => {
                    if discard {
                        return Err(invalid("duplicate discardImageEditData extension"));
                    }
                    discard = true;
                    PresentationPropertyExtension::DiscardImageEditData(parse_extension_bool(
                        ext,
                        P14_NS,
                        "discardImageEditData",
                    )?)
                },
                DEFAULT_IMAGE_DPI_URI => {
                    if dpi {
                        return Err(invalid("duplicate defaultImageDpi extension"));
                    }
                    dpi = true;
                    PresentationPropertyExtension::DefaultImageDpi(parse_extension_u32(
                        ext,
                        P14_NS,
                        "defaultImageDpi",
                    )?)
                },
                CHART_TRACKING_REF_BASED_URI => {
                    if tracking {
                        return Err(invalid("duplicate chartTrackingRefBased extension"));
                    }
                    tracking = true;
                    PresentationPropertyExtension::ChartTrackingReferenceBased(
                        parse_extension_bool(ext, P15_NS, "chartTrackingRefBased")?,
                    )
                },
                _ => PresentationPropertyExtension::Unknown(OpaquePresentationExtension {
                    uri,
                    xml: node_xml(ext, false)?,
                }),
            };
        out.push(value);
    }
    Ok(out)
}
fn parse_show_extensions(n: &Node) -> Result<Vec<SlideShowExtension>> {
    no_attrs(n)?;
    let extensions = children(n)?;
    if extensions.len() > MAX_EXTENSIONS {
        return Err(invalid("slide-show extension count exceeds limit"));
    }
    let mut out = Vec::with_capacity(extensions.len());
    let mut browse_mode = false;
    let mut laser = false;
    let mut controls = false;
    for ext in extensions {
        let uri = extension_uri(ext)?;
        let value = match uri.as_str() {
            BROWSE_MODE_URI => {
                if browse_mode {
                    return Err(invalid("duplicate browseMode extension"));
                }
                browse_mode = true;
                let payload = extension_payload(ext, P14_NS, "browseMode")?;
                let show_status = bool_opt(payload, "showStatus")?;
                only_attrs(payload, &[("", "showStatus")])?;
                empty(payload)?;
                SlideShowExtension::BrowseMode { show_status }
            },
            LASER_COLOR_URI => {
                if laser {
                    return Err(invalid("duplicate laserClr extension"));
                }
                laser = true;
                let payload = extension_payload(ext, P14_NS, "laserClr")?;
                no_attrs(payload)?;
                let colors = children(payload)?;
                if colors.len() != 1 {
                    return Err(invalid("laserClr requires exactly one DrawingML color"));
                }
                SlideShowExtension::LaserColor(parse_color(colors[0])?)
            },
            SHOW_MEDIA_CONTROLS_URI => {
                if controls {
                    return Err(invalid("duplicate showMediaCtrls extension"));
                }
                controls = true;
                SlideShowExtension::ShowMediaControls(parse_extension_bool(
                    ext,
                    P14_NS,
                    "showMediaCtrls",
                )?)
            },
            _ => SlideShowExtension::Unknown(OpaquePresentationExtension {
                uri,
                xml: node_xml(ext, false)?,
            }),
        };
        out.push(value);
    }
    Ok(out)
}
fn extension_uri(n: &Node) -> Result<String> {
    expect_p(n)?;
    if n.local != "ext" {
        return Err(invalid("extLst contains a non-ext child"));
    }
    let uri =
        attr_opt(n, "", "uri")?.ok_or_else(|| invalid("presentation extension requires uri"))?;
    if uri.is_empty() {
        return Err(invalid("presentation extension uri is empty"));
    }
    only_attrs(n, &[("", "uri")])?;
    Ok(uri)
}
fn extension_payload<'a>(n: &'a Node, namespace: &str, local: &str) -> Result<&'a Node> {
    let payload = children(n)?;
    if payload.len() != 1 {
        return Err(invalid(format!(
            "{local} extension requires exactly one payload"
        )));
    }
    let payload = payload[0];
    if payload.ns != namespace || payload.local != local {
        return Err(invalid(format!("extension uri does not contain {local}")));
    }
    Ok(payload)
}
fn parse_extension_bool(n: &Node, namespace: &str, local: &str) -> Result<bool> {
    let payload = extension_payload(n, namespace, local)?;
    let value = bool_req(payload, "val")?;
    only_attrs(payload, &[("", "val")])?;
    empty(payload)?;
    Ok(value)
}
fn parse_extension_u32(n: &Node, namespace: &str, local: &str) -> Result<u32> {
    let payload = extension_payload(n, namespace, local)?;
    let value = u32_req(payload, "val")?;
    only_attrs(payload, &[("", "val")])?;
    empty(payload)?;
    Ok(value)
}

fn write_html(x: &mut String, v: &HtmlPublishProperties, s: bool) -> Result<()> {
    x.push_str("<p:htmlPubPr");
    bool_opt_write(x, "showSpeakerNotes", v.show_speaker_notes);
    if let Some(b) = v.browser {
        attr_write(x, "pubBrowser", browser_str(b));
    }
    x.push_str(" r:id=\"");
    esc_attr(x, &v.target.relationship_id);
    x.push_str("\">");
    write_selection(x, &v.slides);
    if let Some(e) = &v.extension_xml {
        write_opaque(x, e, s)?;
    }
    x.push_str("</p:htmlPubPr>");
    Ok(())
}
fn write_web(x: &mut String, v: &WebProperties, s: bool) -> Result<()> {
    x.push_str("<p:webPr");
    for (n, b) in [
        ("showAnimation", v.show_animation),
        ("resizeGraphics", v.resize_graphics),
        ("allowPng", v.allow_png),
        ("relyOnVml", v.rely_on_vml),
        ("organizeInFolders", v.organize_in_folders),
        ("useLongFilenames", v.use_long_filenames),
    ] {
        bool_opt_write(x, n, b)
    }
    if let Some(z) = v.image_size {
        attr_write(x, "imgSz", screen_str(z));
    }
    if let Some(z) = &v.encoding {
        attr_write(x, "encoding", z);
    }
    if let Some(z) = v.color {
        attr_write(x, "clr", web_color_str(z));
    }
    if let Some(e) = &v.extension_xml {
        x.push('>');
        write_opaque(x, e, s)?;
        x.push_str("</p:webPr>");
    } else {
        x.push_str("/>");
    }
    Ok(())
}
fn write_print(x: &mut String, v: &PrintProperties, s: bool) -> Result<()> {
    x.push_str("<p:prnPr");
    if let Some(z) = v.output {
        attr_write(x, "prnWhat", output_str(z));
    }
    if let Some(z) = v.color_mode {
        attr_write(x, "clrMode", print_color_str(z));
    }
    for (n, b) in [
        ("hiddenSlides", v.hidden_slides),
        ("scaleToFitPaper", v.scale_to_fit_paper),
        ("frameSlides", v.frame_slides),
    ] {
        bool_opt_write(x, n, b)
    }
    if let Some(e) = &v.extension_xml {
        x.push('>');
        write_opaque(x, e, s)?;
        x.push_str("</p:prnPr>");
    } else {
        x.push_str("/>");
    }
    Ok(())
}
fn write_show(x: &mut String, v: &ShowProperties, s: bool) -> Result<()> {
    x.push_str("<p:showPr");
    for (n, b) in [
        ("loop", v.loop_show),
        ("showNarration", v.show_narration),
        ("showAnimation", v.show_animation),
        ("useTimings", v.use_timings),
    ] {
        bool_opt_write(x, n, b)
    }
    x.push('>');
    if let Some(m) = &v.mode {
        match m {
            ShowMode::Present => x.push_str("<p:present/>"),
            ShowMode::Browse { show_scrollbar } => {
                x.push_str("<p:browse");
                bool_opt_write(x, "showScrollbar", *show_scrollbar);
                x.push_str("/>");
            },
            ShowMode::Kiosk { restart } => {
                x.push_str("<p:kiosk");
                if let Some(v) = restart {
                    attr_write(x, "restart", &v.to_string());
                }
                x.push_str("/>");
            },
        }
    }
    if let Some(z) = &v.slides {
        write_selection(x, z);
    }
    if let Some(c) = &v.pen_color {
        x.push_str("<p:penClr>");
        write_opaque(x, &c.xml, s)?;
        x.push_str("</p:penClr>");
    }
    write_show_extensions(x, &v.extensions, s)?;
    x.push_str("</p:showPr>");
    Ok(())
}
fn write_selection(x: &mut String, v: &SlideSelection) {
    match v {
        SlideSelection::All => x.push_str("<p:sldAll/>"),
        SlideSelection::Range { start, end } => {
            x.push_str(&format!("<p:sldRg st=\"{start}\" end=\"{end}\"/>"))
        },
        SlideSelection::CustomShow(id) => x.push_str(&format!("<p:custShow id=\"{id}\"/>")),
    }
}
fn write_presentation_extensions(
    x: &mut String,
    v: &[PresentationPropertyExtension],
    strict: bool,
) -> Result<()> {
    if v.is_empty() {
        return Ok(());
    }
    x.push_str("<p:extLst>");
    for extension in v {
        match extension {
            PresentationPropertyExtension::DiscardImageEditData(value) => write_bool_extension(
                x,
                DISCARD_IMAGE_EDIT_DATA_URI,
                "p14",
                "discardImageEditData",
                *value,
            ),
            PresentationPropertyExtension::DefaultImageDpi(value) => {
                x.push_str("<p:ext uri=\"");
                x.push_str(DEFAULT_IMAGE_DPI_URI);
                x.push_str("\"><p14:defaultImageDpi xmlns:p14=\"");
                x.push_str(P14_NS);
                x.push_str("\" val=\"");
                x.push_str(&value.to_string());
                x.push_str("\"/></p:ext>");
            },
            PresentationPropertyExtension::ChartTrackingReferenceBased(value) => {
                write_bool_extension(
                    x,
                    CHART_TRACKING_REF_BASED_URI,
                    "p15",
                    "chartTrackingRefBased",
                    *value,
                )
            },
            PresentationPropertyExtension::Unknown(value) => {
                write_unknown_extension(x, value, strict)?
            },
        }
    }
    x.push_str("</p:extLst>");
    Ok(())
}
fn write_show_extensions(x: &mut String, v: &[SlideShowExtension], strict: bool) -> Result<()> {
    if v.is_empty() {
        return Ok(());
    }
    x.push_str("<p:extLst>");
    for extension in v {
        match extension {
            SlideShowExtension::BrowseMode { show_status } => {
                x.push_str("<p:ext uri=\"");
                x.push_str(BROWSE_MODE_URI);
                x.push_str("\"><p14:browseMode xmlns:p14=\"");
                x.push_str(P14_NS);
                x.push('"');
                bool_opt_write(x, "showStatus", *show_status);
                x.push_str("/></p:ext>");
            },
            SlideShowExtension::LaserColor(color) => {
                x.push_str("<p:ext uri=\"");
                x.push_str(LASER_COLOR_URI);
                x.push_str("\"><p14:laserClr xmlns:p14=\"");
                x.push_str(P14_NS);
                x.push_str("\">");
                write_opaque(x, &color.xml, strict)?;
                x.push_str("</p14:laserClr></p:ext>");
            },
            SlideShowExtension::ShowMediaControls(value) => {
                write_bool_extension(x, SHOW_MEDIA_CONTROLS_URI, "p14", "showMediaCtrls", *value)
            },
            SlideShowExtension::Unknown(value) => write_unknown_extension(x, value, strict)?,
        }
    }
    x.push_str("</p:extLst>");
    Ok(())
}
fn write_bool_extension(x: &mut String, uri: &str, prefix: &str, local: &str, value: bool) {
    x.push_str("<p:ext uri=\"");
    x.push_str(uri);
    x.push_str("\"><");
    x.push_str(prefix);
    x.push(':');
    x.push_str(local);
    x.push_str(" xmlns:");
    x.push_str(prefix);
    x.push_str("=\"");
    x.push_str(if prefix == "p15" { P15_NS } else { P14_NS });
    x.push_str("\" val=\"");
    x.push_str(if value { "1" } else { "0" });
    x.push_str("\"/></p:ext>");
}
fn write_unknown_extension(
    x: &mut String,
    v: &OpaquePresentationExtension,
    strict: bool,
) -> Result<()> {
    bounded(&v.uri)?;
    let node = parse_dom(&v.xml)?;
    let uri = extension_uri(&node)?;
    if uri != v.uri {
        return Err(invalid("opaque extension uri does not match its XML"));
    }
    if known_extension_uri(&uri) {
        return Err(invalid(
            "known presentation extension cannot be stored as opaque",
        ));
    }
    write_opaque(x, &v.xml, strict)
}
fn known_extension_uri(uri: &str) -> bool {
    matches!(
        uri,
        DISCARD_IMAGE_EDIT_DATA_URI
            | DEFAULT_IMAGE_DPI_URI
            | CHART_TRACKING_REF_BASED_URI
            | BROWSE_MODE_URI
            | LASER_COLOR_URI
            | SHOW_MEDIA_CONTROLS_URI
    )
}
fn write_opaque(x: &mut String, v: &[u8], strict: bool) -> Result<()> {
    validate_fragment(v)?;
    let mut text = std::str::from_utf8(v).map_err(xml_error)?.to_string();
    if strict {
        text = text
            .replace(P_NS, P_STRICT)
            .replace(A_NS, A_STRICT)
            .replace(R_NS, R_STRICT)
    } else {
        text = text
            .replace(P_STRICT, P_NS)
            .replace(A_STRICT, A_NS)
            .replace(R_STRICT, R_NS)
    }
    x.push_str(&text);
    Ok(())
}

fn validate(v: &PresentationProperties) -> Result<()> {
    if v.recent_colors.len() > 10 {
        return Err(invalid("clrMru permits at most ten colors"));
    }
    if v.extensions.len() > MAX_EXTENSIONS {
        return Err(invalid("presentation extension count exceeds limit"));
    }
    if v.show
        .as_ref()
        .is_some_and(|show| show.extensions.len() > MAX_EXTENSIONS)
    {
        return Err(invalid("slide-show extension count exceeds limit"));
    }
    for c in v
        .recent_colors
        .iter()
        .chain(v.show.iter().filter_map(|s| s.pen_color.as_ref()))
        .chain(
            v.show
                .iter()
                .flat_map(|s| s.extensions.iter())
                .filter_map(|e| match e {
                    SlideShowExtension::LaserColor(c) => Some(c),
                    _ => None,
                }),
        )
    {
        validate_fragment(&c.xml)?;
    }
    for e in [
        v.html_publish
            .as_ref()
            .and_then(|x| x.extension_xml.as_deref()),
        v.web.as_ref().and_then(|x| x.extension_xml.as_deref()),
        v.print.as_ref().and_then(|x| x.extension_xml.as_deref()),
    ]
    .into_iter()
    .flatten()
    {
        validate_fragment(e)?;
    }
    for e in v
        .extensions
        .iter()
        .filter_map(|e| match e {
            PresentationPropertyExtension::Unknown(v) => Some(v),
            _ => None,
        })
        .chain(
            v.show
                .iter()
                .flat_map(|s| s.extensions.iter())
                .filter_map(|e| match e {
                    SlideShowExtension::Unknown(v) => Some(v),
                    _ => None,
                }),
        )
    {
        let mut sink = String::new();
        write_unknown_extension(&mut sink, e, false)?;
    }
    if let Some(h) = &v.html_publish {
        bounded(&h.target.relationship_id)?;
    }
    if let Some(w) = &v.web {
        if let Some(e) = &w.encoding {
            bounded(e)?;
        }
    }
    Ok(())
}
fn validate_fragment(v: &[u8]) -> Result<()> {
    if v.len() > MAX_BYTES {
        return Err(invalid("opaque presentation-properties XML exceeds limit"));
    }
    parse_dom(v).map(|_| ())
}
fn node_xml(n: &Node, strict: bool) -> Result<Vec<u8>> {
    let mut x = String::new();
    write_node(&mut x, n, strict)?;
    Ok(x.into_bytes())
}
fn write_node(x: &mut String, n: &Node, strict: bool) -> Result<()> {
    x.push('<');
    x.push_str(&n.qname);
    for (p, u) in &n.bindings {
        x.push_str(if p.is_empty() { " xmlns=\"" } else { " xmlns:" });
        if !p.is_empty() {
            x.push_str(p);
            x.push_str("=\"");
        }
        let value = if strict {
            map_ns(u, true)
        } else {
            map_ns(u, false)
        };
        esc_attr(x, value);
        x.push('"');
    }
    for a in &n.attrs {
        x.push(' ');
        x.push_str(&a.qname);
        x.push_str("=\"");
        esc_attr(x, &a.value);
        x.push('"');
    }
    if n.content.is_empty() {
        x.push_str("/>");
        return Ok(());
    }
    x.push('>');
    for c in &n.content {
        match c {
            Content::Node(v) => write_node(x, v, strict)?,
            Content::Text(v) => esc_text(x, v),
            Content::CData(v) => {
                x.push_str("<![CDATA[");
                x.push_str(v);
                x.push_str("]]>");
            },
            Content::Comment(v) => {
                x.push_str("<!--");
                x.push_str(v);
                x.push_str("-->");
            },
        }
    }
    x.push_str("</");
    x.push_str(&n.qname);
    x.push('>');
    Ok(())
}
fn map_ns<'a>(v: &'a str, s: bool) -> &'a str {
    if s {
        match v {
            P_NS => P_STRICT,
            A_NS => A_STRICT,
            R_NS => R_STRICT,
            _ => v,
        }
    } else {
        match v {
            P_STRICT => P_NS,
            A_STRICT => A_NS,
            R_STRICT => R_NS,
            _ => v,
        }
    }
}

fn children(n: &Node) -> Result<Vec<&Node>> {
    let mut v = Vec::new();
    for c in &n.content {
        match c {
            Content::Node(x) => v.push(x),
            Content::Text(x) if x.trim().is_empty() => {},
            Content::Comment(_) => {},
            _ => return Err(invalid("unexpected text in typed presentation properties")),
        }
    }
    Ok(v)
}
fn empty(n: &Node) -> Result<()> {
    if children(n)?.is_empty() {
        Ok(())
    } else {
        Err(invalid("leaf presentation property has children"))
    }
}
fn expect(n: &Node, a: &str, b: &str, l: &str) -> Result<()> {
    if (n.ns == a || n.ns == b) && n.local == l {
        Ok(())
    } else {
        Err(invalid(format!("expected {l}")))
    }
}
fn expect_p(n: &Node) -> Result<()> {
    if n.ns == P_NS || n.ns == P_STRICT {
        Ok(())
    } else {
        Err(invalid("expected PresentationML namespace"))
    }
}
fn attr_opt(n: &Node, ns: &str, l: &str) -> Result<Option<String>> {
    let mut v = None;
    for a in &n.attrs {
        if a.ns == ns && a.local == l {
            if v.is_some() {
                return Err(invalid("duplicate attribute"));
            }
            bounded(&a.value)?;
            v = Some(a.value.clone());
        }
    }
    Ok(v)
}
fn attr_req(n: &Node, a: &str, b: &str, l: &str) -> Result<String> {
    let mut v = None;
    for x in &n.attrs {
        if (x.ns == a || x.ns == b) && x.local == l {
            if v.is_some() {
                return Err(invalid("duplicate attribute"));
            }
            v = Some(x.value.clone());
        }
    }
    v.ok_or_else(|| invalid(format!("missing required attribute '{l}'")))
}
fn no_attrs(n: &Node) -> Result<()> {
    if n.attrs.is_empty() {
        Ok(())
    } else {
        Err(invalid("unexpected attributes"))
    }
}
fn only_attrs(n: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    for a in &n.attrs {
        if !allowed
            .iter()
            .any(|(ns, l)| (*ns == a.ns || (*ns == R_NS && a.ns == R_STRICT)) && *l == a.local)
        {
            return Err(invalid(format!("unexpected attribute '{}'", a.qname)));
        }
    }
    Ok(())
}
fn bool_opt(n: &Node, l: &str) -> Result<Option<bool>> {
    match attr_opt(n, "", l)?.as_deref() {
        None => Ok(None),
        Some("1" | "true") => Ok(Some(true)),
        Some("0" | "false") => Ok(Some(false)),
        _ => Err(invalid(format!("invalid boolean '{l}'"))),
    }
}
fn u32_opt(n: &Node, l: &str) -> Result<Option<u32>> {
    attr_opt(n, "", l)?
        .map(|v| {
            v.parse()
                .map_err(|_| invalid(format!("invalid integer '{l}'")))
        })
        .transpose()
}
fn u32_req(n: &Node, l: &str) -> Result<u32> {
    u32_opt(n, l)?.ok_or_else(|| invalid(format!("missing integer '{l}'")))
}
fn bool_req(n: &Node, l: &str) -> Result<bool> {
    bool_opt(n, l)?.ok_or_else(|| invalid(format!("missing boolean '{l}'")))
}
fn single_ext(n: &Node) -> Result<Option<Vec<u8>>> {
    let c = children(n)?;
    if c.len() > 1 {
        return Err(invalid("property permits only one extLst"));
    }
    if let Some(e) = c.first() {
        expect_p(e)?;
        if e.local != "extLst" {
            return Err(invalid("unexpected property child"));
        }
        Ok(Some(node_xml(e, false)?))
    } else {
        Ok(None)
    }
}
fn split(q: &str) -> Result<(&str, &str)> {
    if let Some((p, l)) = q.split_once(':') {
        if l.is_empty() || l.contains(':') {
            return Err(invalid("invalid QName"));
        }
        Ok((p, l))
    } else {
        Ok(("", q))
    }
}
fn resolve(bindings: &[(String, String)], p: &str) -> Result<String> {
    bindings
        .iter()
        .rev()
        .find(|x| x.0 == p)
        .map(|x| x.1.clone())
        .ok_or_else(|| invalid(format!("unbound namespace prefix '{p}'")))
}
fn bounded(v: &str) -> Result<()> {
    if v.len() > MAX_STRING {
        Err(invalid("presentation property string exceeds 1 MiB"))
    } else {
        Ok(())
    }
}
fn attr_write(x: &mut String, n: &str, v: &str) {
    x.push(' ');
    x.push_str(n);
    x.push_str("=\"");
    esc_attr(x, v);
    x.push('"')
}
fn bool_opt_write(x: &mut String, n: &str, v: Option<bool>) {
    if let Some(v) = v {
        attr_write(x, n, if v { "1" } else { "0" })
    }
}
fn esc_attr(x: &mut String, v: &str) {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;"),
            '<' => x.push_str("&lt;"),
            '"' => x.push_str("&quot;"),
            '\r' => x.push_str("&#xD;"),
            '\n' => x.push_str("&#xA;"),
            '\t' => x.push_str("&#x9;"),
            _ => x.push(c),
        }
    }
}
fn esc_text(x: &mut String, v: &str) {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;"),
            '<' => x.push_str("&lt;"),
            '>' => x.push_str("&gt;"),
            _ => x.push(c),
        }
    }
}
macro_rules! enums{($parse:ident,$write:ident,$ty:ty,$($s:literal=>$v:path),+)=>{fn $parse(v:String)->Result<$ty>{match v.as_str(){$($s=>Ok($v),)+_=>Err(invalid(format!("invalid enumeration '{v}'")))}}fn $write(v:$ty)->&'static str{match v{$($v=>$s,)+}}}}
enums!(parse_browser,browser_str,BrowserSupport,"v3"=>BrowserSupport::V3,"v4"=>BrowserSupport::V4,"v3v4"=>BrowserSupport::V3V4);
enums!(parse_screen,screen_str,WebScreenSize,"544x376"=>WebScreenSize::S544x376,"640x480"=>WebScreenSize::S640x480,"720x512"=>WebScreenSize::S720x512,"800x600"=>WebScreenSize::S800x600,"1024x768"=>WebScreenSize::S1024x768,"1152x882"=>WebScreenSize::S1152x882,"1152x900"=>WebScreenSize::S1152x900,"1280x1024"=>WebScreenSize::S1280x1024,"1600x1200"=>WebScreenSize::S1600x1200,"1800x1400"=>WebScreenSize::S1800x1400,"1920x1200"=>WebScreenSize::S1920x1200);
enums!(parse_web_color,web_color_str,WebColor,"none"=>WebColor::None,"browser"=>WebColor::Browser,"presentationText"=>WebColor::PresentationText,"presentationAccent"=>WebColor::PresentationAccent,"whiteTextOnBlack"=>WebColor::WhiteTextOnBlack,"blackTextOnWhite"=>WebColor::BlackTextOnWhite);
enums!(parse_output,output_str,PrintOutput,"slides"=>PrintOutput::Slides,"handouts1"=>PrintOutput::Handouts1,"handouts2"=>PrintOutput::Handouts2,"handouts3"=>PrintOutput::Handouts3,"handouts4"=>PrintOutput::Handouts4,"handouts6"=>PrintOutput::Handouts6,"handouts9"=>PrintOutput::Handouts9,"notes"=>PrintOutput::Notes,"outline"=>PrintOutput::Outline);
enums!(parse_print_color,print_color_str,PrintColorMode,"bw"=>PrintColorMode::BlackWhite,"gray"=>PrintColorMode::Gray,"clr"=>PrintColorMode::Color);
fn invalid(v: impl Into<String>) -> Box<dyn std::error::Error + Send + Sync> {
    std::io::Error::new(std::io::ErrorKind::InvalidData, v.into()).into()
}
fn xml_error(e: impl std::fmt::Display) -> Box<dyn std::error::Error + Send + Sync> {
    invalid(e.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_local_presentation_properties_fixture() {
        let value = PresentationProperties::parse(include_bytes!(
            "../../../../test-data/ooxml/pptx/presentation-properties/basic_presentation.xml"
        ))
        .unwrap();
        let show = value.show.as_ref().unwrap();
        assert_eq!(show.mode, Some(ShowMode::Kiosk { restart: Some(5) }));
        assert_eq!(
            show.extensions,
            vec![SlideShowExtension::BrowseMode {
                show_status: Some(false)
            }]
        );
        let strict = value.to_xml(true).unwrap();
        let again = PresentationProperties::parse(&strict).unwrap();
        assert_eq!(again.show, value.show);
    }
    #[test]
    fn strict_mce_and_typed_roundtrip() {
        let xml = format!(
            r#"<p:presentationPr xmlns:p="{P_STRICT}" xmlns:r="{R_STRICT}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:u" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><u:x/></mc:Choice><mc:Fallback><p:webPr showAnimation="1" imgSz="1920x1200" clr="browser"/></mc:Fallback></mc:AlternateContent><p:prnPr prnWhat="handouts6" clrMode="gray"/><p:showPr loop="1"><p:kiosk restart="5"/><p:sldRg st="2" end="4"/></p:showPr></p:presentationPr>"#
        );
        let v = PresentationProperties::parse(xml.as_bytes()).unwrap();
        assert_eq!(v.web.unwrap().image_size, Some(WebScreenSize::S1920x1200));
        assert_eq!(v.print.unwrap().output, Some(PrintOutput::Handouts6));
        assert_eq!(
            v.show.unwrap().slides,
            Some(SlideSelection::Range { start: 2, end: 4 })
        );
    }
    #[test]
    fn typed_extensions_preserve_unknown_entries_without_resolving_relationships() {
        let xml = format!(
            r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:r="{R_NS}" xmlns:p14="{P14_NS}" xmlns:p15="{P15_NS}" xmlns:a="{A_NS}" xmlns:v="urn:producer"><p:showPr><p:extLst><p:ext uri="{LASER_COLOR_URI}"><p14:laserClr><a:schemeClr val="accent1"/></p14:laserClr></p:ext><p:ext uri="urn:producer:show"><v:payload r:id="rIdNeverFetched"><v:nested value="opaque"/></v:payload></p:ext><p:ext uri="{SHOW_MEDIA_CONTROLS_URI}"><p14:showMediaCtrls val="0"/></p:ext></p:extLst></p:showPr><p:extLst><p:ext uri="{DISCARD_IMAGE_EDIT_DATA_URI}"><p14:discardImageEditData val="1"/></p:ext><p:ext uri="urn:producer:root"><v:data href="https://example.invalid/not-opened"/></p:ext><p:ext uri="{DEFAULT_IMAGE_DPI_URI}"><p14:defaultImageDpi val="4294967295"/></p:ext><p:ext uri="{CHART_TRACKING_REF_BASED_URI}"><p15:chartTrackingRefBased val="false"/></p:ext></p:extLst></p:presentationPr>"#
        );
        let value = PresentationProperties::parse(xml.as_bytes()).unwrap();
        let show = value.show.as_ref().unwrap();
        assert!(
            matches!(&show.extensions[..],[SlideShowExtension::LaserColor(_),SlideShowExtension::Unknown(OpaquePresentationExtension{uri,..}),SlideShowExtension::ShowMediaControls(false)] if uri=="urn:producer:show")
        );
        assert!(
            matches!(&value.extensions[..],[PresentationPropertyExtension::DiscardImageEditData(true),PresentationPropertyExtension::Unknown(OpaquePresentationExtension{uri,..}),PresentationPropertyExtension::DefaultImageDpi(u32::MAX),PresentationPropertyExtension::ChartTrackingReferenceBased(false)] if uri=="urn:producer:root")
        );
        for strict in [false, true] {
            let written = value.to_xml(strict).unwrap();
            let again = PresentationProperties::parse(&written).unwrap();
            assert_eq!(
                again
                    .extensions
                    .iter()
                    .filter(|e| matches!(e, PresentationPropertyExtension::Unknown(_)))
                    .count(),
                1
            );
            assert_eq!(
                again
                    .show
                    .as_ref()
                    .unwrap()
                    .extensions
                    .iter()
                    .filter(|e| matches!(e, SlideShowExtension::Unknown(_)))
                    .count(),
                1
            );
            let text = String::from_utf8(written).unwrap();
            assert!(text.contains("r:id=\"rIdNeverFetched\""));
            assert!(text.contains("https://example.invalid/not-opened"));
        }
    }
    #[test]
    fn browse_mode_extension_round_trips() {
        let xml = format!(
            r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:p14="{P14_NS}"><p:showPr><p:extLst><p:ext uri="{BROWSE_MODE_URI}"><p14:browseMode showStatus="0"/></p:ext></p:extLst></p:showPr></p:presentationPr>"#
        );
        let value = PresentationProperties::parse(xml.as_bytes()).unwrap();
        assert_eq!(
            value.show.as_ref().unwrap().extensions,
            vec![SlideShowExtension::BrowseMode {
                show_status: Some(false)
            }]
        );
        for strict in [false, true] {
            let written = value.to_xml(strict).unwrap();
            let again = PresentationProperties::parse(&written).unwrap();
            assert_eq!(again.show, value.show);
        }

        let no_status = format!(
            r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:p14="{P14_NS}"><p:showPr><p:extLst><p:ext uri="{BROWSE_MODE_URI}"><p14:browseMode/></p:ext></p:extLst></p:showPr></p:presentationPr>"#
        );
        assert!(matches!(
            PresentationProperties::parse(no_status.as_bytes()),
            Ok(PresentationProperties {
                show: Some(ShowProperties {
                    extensions,
                    ..
                }),
                ..
            }) if extensions == vec![SlideShowExtension::BrowseMode { show_status: None }]
        ));
    }
    #[test]
    fn rejects_hostile_typed_extension_grammar_and_bounds() {
        let root = |body: &str| {
            format!(
                r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:r="{R_NS}" xmlns:p14="{P14_NS}" xmlns:p15="{P15_NS}" xmlns:a="{A_NS}">{body}</p:presentationPr>"#
            )
        };
        let cases = [
            root(&format!(
                r#"<p:extLst><p:ext uri="{DISCARD_IMAGE_EDIT_DATA_URI}"><p14:defaultImageDpi val="1"/></p:ext></p:extLst>"#
            )),
            root(&format!(
                r#"<p:extLst><p:ext uri="{DISCARD_IMAGE_EDIT_DATA_URI}"><p14:discardImageEditData/></p:ext></p:extLst>"#
            )),
            root(&format!(
                r#"<p:extLst><p:ext uri="{DISCARD_IMAGE_EDIT_DATA_URI}"><p14:discardImageEditData val="yes"/></p:ext></p:extLst>"#
            )),
            root(&format!(
                r#"<p:extLst><p:ext uri="{DEFAULT_IMAGE_DPI_URI}"><p14:defaultImageDpi val="4294967296"/></p:ext></p:extLst>"#
            )),
            root(&format!(
                r#"<p:extLst><p:ext uri="{DISCARD_IMAGE_EDIT_DATA_URI}"><p14:discardImageEditData val="0"/></p:ext><p:ext uri="{DISCARD_IMAGE_EDIT_DATA_URI}"><p14:discardImageEditData val="1"/></p:ext></p:extLst>"#
            )),
            root(&format!(
                r#"<p:extLst><p:ext uri="{CHART_TRACKING_REF_BASED_URI}"><p15:chartTrackingRefBased val="1" r:id="rIdNo"/></p:ext></p:extLst>"#
            )),
            root(&format!(
                r#"<p:showPr><p:extLst><p:ext uri="{LASER_COLOR_URI}"><p14:laserClr><a:srgbClr val="FF0000"/><a:srgbClr val="00FF00"/></p14:laserClr></p:ext></p:extLst></p:showPr>"#
            )),
            root(&format!(
                r#"<p:showPr><p:extLst><p:ext uri="{SHOW_MEDIA_CONTROLS_URI}"><p14:showMediaCtrls val="1"><p14:x/></p14:showMediaCtrls></p:ext></p:extLst></p:showPr>"#
            )),
            root(
                r#"<p:extLst><p:ext uri=""><p14:discardImageEditData val="1"/></p:ext></p:extLst>"#,
            ),
        ];
        for xml in cases {
            assert!(
                PresentationProperties::parse(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
        let mut value = PresentationProperties::default();
        value.extensions = (0..=MAX_EXTENSIONS)
            .map(|i| {
                PresentationPropertyExtension::Unknown(OpaquePresentationExtension {
                    uri: format!("urn:vendor:{i}"),
                    xml: format!(r#"<p:ext xmlns:p="{P_NS}" uri="urn:vendor:{i}"/>"#).into_bytes(),
                })
            })
            .collect();
        assert!(value.to_xml(false).is_err());
    }
    #[test]
    fn rejects_malformed_and_unsafe() {
        for x in [
            format!(
                r#"<p:presentationPr xmlns:p="{P_NS}"><p:showPr loop="maybe"/></p:presentationPr>"#
            ),
            format!(
                r#"<p:presentationPr xmlns:p="{P_NS}"><p:showPr><p:sldRg st="4" end="2"/></p:showPr></p:presentationPr>"#
            ),
            format!(r#"<!DOCTYPE x><p:presentationPr xmlns:p="{P_NS}"/>"#),
        ] {
            assert!(
                PresentationProperties::parse(x.as_bytes()).is_err(),
                "accepted {x}"
            );
        }
        let mut v = PresentationProperties::default();
        v.extensions.push(PresentationPropertyExtension::Unknown(
            OpaquePresentationExtension {
                uri: "urn:bad".into(),
                xml: b"<?bad?><p:extLst/>".to_vec(),
            },
        ));
        assert!(v.to_xml(false).is_err());
    }
}
