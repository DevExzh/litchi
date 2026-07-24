//! Typed PresentationML view properties with bounded opaque extensions.

use litchi_core::sheet::Result;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::{
    Reader, XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
};

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const AS: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/viewProps";
const CT: &str = "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml";
const MAX: usize = 8 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 100_000;
const MAX_GUIDES: usize = 4096;
const MAX_SLIDES: usize = 100_000;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ViewKind {
    Slide,
    SlideMaster,
    Notes,
    Handout,
    NotesMaster,
    Outline,
    SlideSorter,
    SlideThumbnail,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SplitterState {
    Minimized,
    Restored,
    Maximized,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum GuideOrientation {
    Horizontal,
    Vertical,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Ratio {
    pub numerator: i64,
    pub denominator: i64,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Point {
    pub x: i64,
    pub y: i64,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CommonView {
    pub variable_scale: Option<bool>,
    pub scale_x: Ratio,
    pub scale_y: Ratio,
    pub origin: Point,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Guide {
    pub orientation: Option<GuideOrientation>,
    pub position: Option<i32>,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CommonSlideView {
    pub snap_to_grid: Option<bool>,
    pub snap_to_objects: Option<bool>,
    pub show_guides: Option<bool>,
    pub view: CommonView,
    pub guides: Vec<Guide>,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RestoredPane {
    pub size: u32,
    pub auto_adjust: Option<bool>,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NormalView {
    pub show_outline_icons: Option<bool>,
    pub snap_vertical_splitter: Option<bool>,
    pub vertical_bar_state: Option<SplitterState>,
    pub horizontal_bar_state: Option<SplitterState>,
    pub prefer_single_view: Option<bool>,
    pub restored_left: RestoredPane,
    pub restored_top: RestoredPane,
    pub extension_xml: Option<Vec<u8>>,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OutlineSlide {
    pub relationship_id: String,
    pub collapse: Option<bool>,
    pub target: Option<String>,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OutlineView {
    pub view: CommonView,
    pub slides: Vec<OutlineSlide>,
    pub extension_xml: Option<Vec<u8>>,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SimpleView {
    pub view: CommonView,
    pub extension_xml: Option<Vec<u8>>,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SorterView {
    pub show_formatting: Option<bool>,
    pub view: CommonView,
    pub extension_xml: Option<Vec<u8>>,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SlideLikeView {
    pub common: CommonSlideView,
    pub extension_xml: Option<Vec<u8>>,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct GridSpacing {
    pub cx: u32,
    pub cy: u32,
}
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct ViewProperties {
    pub last_view: Option<ViewKind>,
    pub show_comments: Option<bool>,
    pub normal: Option<NormalView>,
    pub slide: Option<SlideLikeView>,
    pub outline: Option<OutlineView>,
    pub notes_text: Option<SimpleView>,
    pub sorter: Option<SorterView>,
    pub notes: Option<SlideLikeView>,
    pub grid_spacing: Option<GridSpacing>,
    pub extension_xml: Option<Vec<u8>>,
}

impl ViewProperties {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX {
            return Err(invalid("view properties exceed 8 MiB"));
        }
        let x = crate::common::mce::process_ooxml(xml)?;
        if x.len() > MAX {
            return Err(invalid("processed view properties exceed 8 MiB"));
        }
        project(&parse_dom(x.as_ref())?)
    }
    pub fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        validate(self)?;
        let p = if strict { PS } else { P };
        let a = if strict { AS } else { A };
        let r = if strict { RS } else { R };
        let mut x = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewPr xmlns:p="{p}" xmlns:a="{a}" xmlns:r="{r}""#
        );
        if let Some(v) = self.last_view {
            attr(&mut x, "lastView", view_str(v));
        }
        bool_attr(&mut x, "showComments", self.show_comments);
        x.push('>');
        if let Some(v) = &self.normal {
            write_normal(&mut x, v, strict)?;
        }
        if let Some(v) = &self.slide {
            write_slide_like(&mut x, "slideViewPr", v, strict)?;
        }
        if let Some(v) = &self.outline {
            write_outline(&mut x, v, strict)?;
        }
        if let Some(v) = &self.notes_text {
            write_simple(&mut x, "notesTextViewPr", v, strict)?;
        }
        if let Some(v) = &self.sorter {
            write_sorter(&mut x, v, strict)?;
        }
        if let Some(v) = &self.notes {
            write_slide_like(&mut x, "notesViewPr", v, strict)?;
        }
        if let Some(v) = &self.grid_spacing {
            x.push_str(&format!("<p:gridSpacing cx=\"{}\" cy=\"{}\"/>", v.cx, v.cy));
        }
        if let Some(v) = &self.extension_xml {
            opaque(&mut x, v, strict)?;
        }
        x.push_str("</p:viewPr>");
        if x.len() > MAX {
            return Err(invalid("serialized view properties exceed 8 MiB"));
        }
        Ok(x.into_bytes())
    }
}

pub fn load_from_package(package: &OpcPackage) -> Result<Option<ViewProperties>> {
    let presentation = package.main_document_part()?;
    let mut found = presentation
        .rels()
        .iter()
        .filter(|x| matches!(x.reltype(), REL | STRICT_REL));
    let Some(rel) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(invalid(
            "presentation has multiple view-properties relationships",
        ));
    }
    if rel.is_external() {
        return Err(invalid("view-properties relationship cannot be external"));
    }
    let uri: PackURI = rel.target_partname()?;
    let part = package.get_part(&uri)?;
    if part.content_type() != CT {
        return Err(invalid(format!(
            "view-properties part '{uri}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    let mut value = ViewProperties::parse(part.blob())?;
    if let Some(outline) = value.outline.as_mut() {
        for slide in &mut outline.slides {
            let relationship = part.rels().get(&slide.relationship_id).ok_or_else(|| {
                invalid(format!(
                    "missing outline slide relationship '{}'",
                    slide.relationship_id
                ))
            })?;
            if relationship.is_external() {
                return Err(invalid("outline slide relationship cannot be external"));
            }
            slide.target = Some(relationship.target_ref().to_string());
        }
    }
    Ok(Some(value))
}

#[derive(Clone)]
struct Attr {
    q: String,
    ns: String,
    l: String,
    v: String,
}
#[derive(Clone)]
enum Content {
    Node(Node),
    Text(String),
    CData(String),
    Comment(String),
}
#[derive(Clone)]
struct Node {
    q: String,
    ns: String,
    l: String,
    attrs: Vec<Attr>,
    bindings: Vec<(String, String)>,
    content: Vec<Content>,
}
fn parse_dom(xml: &[u8]) -> Result<Node> {
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut rd = Reader::from_reader(xml);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut count = 0;
    loop {
        let d = rd.decoder();
        match rd.read_event() {
            Ok(Event::Start(e)) => {
                count += 1;
                if count > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(invalid("view-properties XML resource limit exceeded"));
                }
                stack.push(make(&e, d, &stack)?);
            },
            Ok(Event::Empty(e)) => {
                count += 1;
                if count > MAX_NODES {
                    return Err(invalid("view-properties node limit exceeded"));
                }
                let n = make(&e, d, &stack)?;
                attach(&mut stack, &mut root, n)?;
            },
            Ok(Event::End(_)) => {
                let n = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;
                attach(&mut stack, &mut root, n)?;
            },
            Ok(Event::Text(t)) => {
                let v = t.decode().map_err(xml_error)?.into_owned();
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Text(v))
                } else if !v.trim().is_empty() {
                    return Err(invalid("text outside viewPr"));
                }
            },
            Ok(Event::CData(t)) => {
                if let Some(n) = stack.last_mut() {
                    n.content
                        .push(Content::CData(t.decode().map_err(xml_error)?.into_owned()))
                } else {
                    return Err(invalid("CDATA outside viewPr"));
                }
            },
            Ok(Event::Comment(t)) => {
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Comment(
                        t.decode().map_err(xml_error)?.into_owned(),
                    ))
                }
            },
            Ok(Event::GeneralRef(t)) => {
                if let Some(n) = stack.last_mut() {
                    n.content
                        .push(Content::Text(crate::common::xml::decode_xml_reference(&t)?))
                } else {
                    return Err(invalid("entity outside viewPr"));
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
        return Err(invalid("unterminated view-properties XML"));
    }
    root.ok_or_else(|| invalid("missing viewPr root"))
}
fn make(e: &BytesStart<'_>, d: Decoder, stack: &[Node]) -> Result<Node> {
    let q = std::str::from_utf8(e.name().as_ref())
        .map_err(xml_error)?
        .to_string();
    let mut bindings = stack.last().map(|x| x.bindings.clone()).unwrap_or_default();
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
            if let Some(x) = bindings.iter_mut().find(|x| x.0 == key) {
                x.1 = v.clone()
            } else {
                bindings.push((key, v.clone()))
            }
        }
    }
    let (pr, lo) = split(&q)?;
    let local = lo.to_string();
    let ns = resolve(&bindings, pr)?;
    let mut attrs = Vec::new();
    for (q, v) in raw {
        if q == "xmlns" || q.starts_with("xmlns:") {
            continue;
        }
        let (pr, lo) = split(&q)?;
        let ans = if pr.is_empty() {
            String::new()
        } else {
            resolve(&bindings, pr)?
        };
        let local = lo.to_string();
        attrs.push(Attr {
            q,
            ns: ans,
            l: local,
            v,
        });
    }
    Ok(Node {
        q,
        ns,
        l: local,
        attrs,
        bindings,
        content: Vec::new(),
    })
}
fn attach(stack: &mut [Node], root: &mut Option<Node>, n: Node) -> Result<()> {
    if let Some(p) = stack.last_mut() {
        p.content.push(Content::Node(n))
    } else if root.replace(n).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn project(n: &Node) -> Result<ViewProperties> {
    expect_p(n, "viewPr")?;
    let mut v = ViewProperties {
        last_view: aopt(n, "lastView")?.map(parse_view).transpose()?,
        show_comments: bopt(n, "showComments")?,
        ..Default::default()
    };
    only(n, &["lastView", "showComments"])?;
    let mut order = 0;
    for c in kids(n)? {
        expect_p_any(c)?;
        let i = match c.l.as_str() {
            "normalViewPr" => 0,
            "slideViewPr" => 1,
            "outlineViewPr" => 2,
            "notesTextViewPr" => 3,
            "sorterViewPr" => 4,
            "notesViewPr" => 5,
            "gridSpacing" => 6,
            "extLst" => 7,
            _ => return Err(invalid("unexpected viewPr child")),
        };
        if i < order {
            return Err(invalid("viewPr children out of order"));
        }
        order = i;
        match i {
            0 => set(&mut v.normal, parse_normal(c)?)?,
            1 => set(&mut v.slide, parse_slide_like(c)?)?,
            2 => set(&mut v.outline, parse_outline(c)?)?,
            3 => set(&mut v.notes_text, parse_simple(c)?)?,
            4 => set(&mut v.sorter, parse_sorter(c)?)?,
            5 => set(&mut v.notes, parse_slide_like(c)?)?,
            6 => set(
                &mut v.grid_spacing,
                GridSpacing {
                    cx: u32req(c, "cx")?,
                    cy: u32req(c, "cy")?,
                },
            )?,
            7 => set(&mut v.extension_xml, node_xml(c, false)?)?,
            _ => unreachable!(),
        }
    }
    validate(&v)?;
    Ok(v)
}
fn parse_normal(n: &Node) -> Result<NormalView> {
    let show_outline_icons = bopt(n, "showOutlineIcons")?;
    let snap_vertical_splitter = bopt(n, "snapVertSplitter")?;
    let vertical_bar_state = aopt(n, "vertBarState")?.map(parse_splitter).transpose()?;
    let horizontal_bar_state = aopt(n, "horzBarState")?.map(parse_splitter).transpose()?;
    let prefer_single_view = bopt(n, "preferSingleView")?;
    only(
        n,
        &[
            "showOutlineIcons",
            "snapVertSplitter",
            "vertBarState",
            "horzBarState",
            "preferSingleView",
        ],
    )?;
    let c = kids(n)?;
    if c.len() < 2 || c.len() > 3 || c[0].l != "restoredLeft" || c[1].l != "restoredTop" {
        return Err(invalid(
            "normalViewPr requires restoredLeft then restoredTop",
        ));
    }
    let restored_left = parse_pane(c[0])?;
    let restored_top = parse_pane(c[1])?;
    let extension_xml = if c.len() == 3 {
        if c[2].l != "extLst" {
            return Err(invalid("invalid normalViewPr extension"));
        }
        Some(node_xml(c[2], false)?)
    } else {
        None
    };
    Ok(NormalView {
        show_outline_icons,
        snap_vertical_splitter,
        vertical_bar_state,
        horizontal_bar_state,
        prefer_single_view,
        restored_left,
        restored_top,
        extension_xml,
    })
}
fn parse_pane(n: &Node) -> Result<RestoredPane> {
    let size = u32req(n, "sz")?;
    if size > 100000 {
        return Err(invalid("restored pane size exceeds 100000"));
    }
    let auto_adjust = bopt(n, "autoAdjust")?;
    only(n, &["sz", "autoAdjust"])?;
    leaf(n)?;
    Ok(RestoredPane { size, auto_adjust })
}
fn parse_common(n: &Node) -> Result<CommonView> {
    expect_p(n, "cViewPr")?;
    let variable_scale = bopt(n, "varScale")?;
    only(n, &["varScale"])?;
    let c = kids(n)?;
    if c.len() != 2 || c[0].l != "scale" || c[1].l != "origin" {
        return Err(invalid("cViewPr requires scale then origin"));
    }
    let s = kids(c[0])?;
    if s.len() != 2 || s[0].l != "sx" || s[1].l != "sy" {
        return Err(invalid("scale requires DrawingML sx then sy"));
    }
    let scale_x = ratio(s[0])?;
    let scale_y = ratio(s[1])?;
    let origin = Point {
        x: i64req(c[1], "x")?,
        y: i64req(c[1], "y")?,
    };
    only(c[1], &["x", "y"])?;
    leaf(c[1])?;
    Ok(CommonView {
        variable_scale,
        scale_x,
        scale_y,
        origin,
    })
}
fn ratio(n: &Node) -> Result<Ratio> {
    if n.ns != A && n.ns != AS {
        return Err(invalid("ratio must use DrawingML namespace"));
    }
    let numerator = i64req(n, "n")?;
    let denominator = i64req(n, "d")?;
    if denominator == 0 {
        return Err(invalid("scale denominator cannot be zero"));
    }
    only(n, &["n", "d"])?;
    leaf(n)?;
    Ok(Ratio {
        numerator,
        denominator,
    })
}
fn parse_common_slide(n: &Node) -> Result<CommonSlideView> {
    expect_p(n, "cSldViewPr")?;
    let snap_to_grid = bopt(n, "snapToGrid")?;
    let snap_to_objects = bopt(n, "snapToObjects")?;
    let show_guides = bopt(n, "showGuides")?;
    only(n, &["snapToGrid", "snapToObjects", "showGuides"])?;
    let c = kids(n)?;
    if c.is_empty() || c.len() > 2 || c[0].l != "cViewPr" {
        return Err(invalid("cSldViewPr requires cViewPr"));
    }
    let view = parse_common(c[0])?;
    let mut guides = Vec::new();
    if c.len() == 2 {
        if c[1].l != "guideLst" {
            return Err(invalid("unexpected cSldViewPr child"));
        }
        for g in kids(c[1])? {
            if g.l != "guide" || guides.len() >= MAX_GUIDES {
                return Err(invalid("invalid guide list or limit"));
            }
            let orientation = aopt(g, "orient")?
                .map(|x| match x.as_str() {
                    "horz" => Ok(GuideOrientation::Horizontal),
                    "vert" => Ok(GuideOrientation::Vertical),
                    _ => Err(invalid("invalid guide orientation")),
                })
                .transpose()?;
            let position = i32opt(g, "pos")?;
            only(g, &["orient", "pos"])?;
            leaf(g)?;
            guides.push(Guide {
                orientation,
                position,
            });
        }
    }
    Ok(CommonSlideView {
        snap_to_grid,
        snap_to_objects,
        show_guides,
        view,
        guides,
    })
}
fn parse_slide_like(n: &Node) -> Result<SlideLikeView> {
    noattrs(n)?;
    let c = kids(n)?;
    if c.is_empty() || c.len() > 2 || c[0].l != "cSldViewPr" {
        return Err(invalid("slide-like view requires cSldViewPr"));
    }
    let common = parse_common_slide(c[0])?;
    let extension_xml = ext_second(&c)?;
    Ok(SlideLikeView {
        common,
        extension_xml,
    })
}
fn parse_simple(n: &Node) -> Result<SimpleView> {
    noattrs(n)?;
    let c = kids(n)?;
    if c.is_empty() || c.len() > 2 || c[0].l != "cViewPr" {
        return Err(invalid("view requires cViewPr"));
    }
    Ok(SimpleView {
        view: parse_common(c[0])?,
        extension_xml: ext_second(&c)?,
    })
}
fn parse_sorter(n: &Node) -> Result<SorterView> {
    let show_formatting = bopt(n, "showFormatting")?;
    only(n, &["showFormatting"])?;
    let c = kids(n)?;
    if c.is_empty() || c.len() > 2 || c[0].l != "cViewPr" {
        return Err(invalid("sorterViewPr requires cViewPr"));
    }
    Ok(SorterView {
        show_formatting,
        view: parse_common(c[0])?,
        extension_xml: ext_second(&c)?,
    })
}
fn parse_outline(n: &Node) -> Result<OutlineView> {
    noattrs(n)?;
    let c = kids(n)?;
    if c.is_empty() || c.len() > 3 || c[0].l != "cViewPr" {
        return Err(invalid("outlineViewPr requires cViewPr"));
    }
    let view = parse_common(c[0])?;
    let mut slides = Vec::new();
    let mut extension_xml = None;
    for child in &c[1..] {
        match child.l.as_str() {
            "sldLst" => {
                if !slides.is_empty() {
                    return Err(invalid("duplicate outline slide list"));
                }
                for s in kids(child)? {
                    if slides.len() >= MAX_SLIDES {
                        return Err(invalid("outline slide limit exceeded"));
                    }
                    expect_p(s, "sld")?;
                    let id = arel(s, "id")?;
                    let collapse = bopt(s, "collapse")?;
                    only_rel(s, &["collapse"], "id")?;
                    leaf(s)?;
                    slides.push(OutlineSlide {
                        relationship_id: id,
                        collapse,
                        target: None,
                    });
                }
            },
            "extLst" => {
                if extension_xml.is_some() {
                    return Err(invalid("duplicate outline extension"));
                }
                extension_xml = Some(node_xml(child, false)?);
            },
            _ => return Err(invalid("unexpected outline child")),
        }
    }
    Ok(OutlineView {
        view,
        slides,
        extension_xml,
    })
}
fn ext_second(c: &[&Node]) -> Result<Option<Vec<u8>>> {
    if c.len() == 2 {
        if c[1].l != "extLst" {
            return Err(invalid("expected extLst"));
        }
        Ok(Some(node_xml(c[1], false)?))
    } else {
        Ok(None)
    }
}

fn write_normal(x: &mut String, v: &NormalView, s: bool) -> Result<()> {
    x.push_str("<p:normalViewPr");
    for (n, b) in [
        ("showOutlineIcons", v.show_outline_icons),
        ("snapVertSplitter", v.snap_vertical_splitter),
        ("preferSingleView", v.prefer_single_view),
    ] {
        bool_attr(x, n, b)
    }
    if let Some(z) = v.vertical_bar_state {
        attr(x, "vertBarState", splitter_str(z));
    }
    if let Some(z) = v.horizontal_bar_state {
        attr(x, "horzBarState", splitter_str(z));
    }
    x.push('>');
    pane(x, "restoredLeft", &v.restored_left);
    pane(x, "restoredTop", &v.restored_top);
    if let Some(e) = &v.extension_xml {
        opaque(x, e, s)?;
    }
    x.push_str("</p:normalViewPr>");
    Ok(())
}
fn pane(x: &mut String, n: &str, v: &RestoredPane) {
    x.push_str(&format!("<p:{n} sz=\"{}\"", v.size));
    bool_attr(x, "autoAdjust", v.auto_adjust);
    x.push_str("/>")
}
fn write_common(x: &mut String, v: &CommonView) {
    x.push_str("<p:cViewPr");
    bool_attr(x, "varScale", v.variable_scale);
    x.push_str("><p:scale>");
    ratio_write(x, "a:sx", &v.scale_x);
    ratio_write(x, "a:sy", &v.scale_y);
    x.push_str(&format!(
        "</p:scale><p:origin x=\"{}\" y=\"{}\"/></p:cViewPr>",
        v.origin.x, v.origin.y
    ));
}
fn ratio_write(x: &mut String, n: &str, v: &Ratio) {
    x.push_str(&format!(
        "<{n} n=\"{}\" d=\"{}\"/>",
        v.numerator, v.denominator
    ));
}
fn write_common_slide(x: &mut String, v: &CommonSlideView) {
    x.push_str("<p:cSldViewPr");
    for (n, b) in [
        ("snapToGrid", v.snap_to_grid),
        ("snapToObjects", v.snap_to_objects),
        ("showGuides", v.show_guides),
    ] {
        bool_attr(x, n, b)
    }
    x.push('>');
    write_common(x, &v.view);
    if !v.guides.is_empty() {
        x.push_str("<p:guideLst>");
        for g in &v.guides {
            x.push_str("<p:guide");
            if let Some(o) = g.orientation {
                attr(
                    x,
                    "orient",
                    if o == GuideOrientation::Horizontal {
                        "horz"
                    } else {
                        "vert"
                    },
                );
            }
            if let Some(p) = g.position {
                attr(x, "pos", &p.to_string());
            }
            x.push_str("/>");
        }
        x.push_str("</p:guideLst>");
    }
    x.push_str("</p:cSldViewPr>");
}
fn write_slide_like(x: &mut String, n: &str, v: &SlideLikeView, s: bool) -> Result<()> {
    x.push_str(&format!("<p:{n}>"));
    write_common_slide(x, &v.common);
    if let Some(e) = &v.extension_xml {
        opaque(x, e, s)?;
    }
    x.push_str(&format!("</p:{n}>"));
    Ok(())
}
fn write_simple(x: &mut String, n: &str, v: &SimpleView, s: bool) -> Result<()> {
    x.push_str(&format!("<p:{n}>"));
    write_common(x, &v.view);
    if let Some(e) = &v.extension_xml {
        opaque(x, e, s)?;
    }
    x.push_str(&format!("</p:{n}>"));
    Ok(())
}
fn write_sorter(x: &mut String, v: &SorterView, s: bool) -> Result<()> {
    x.push_str("<p:sorterViewPr");
    bool_attr(x, "showFormatting", v.show_formatting);
    x.push('>');
    write_common(x, &v.view);
    if let Some(e) = &v.extension_xml {
        opaque(x, e, s)?;
    }
    x.push_str("</p:sorterViewPr>");
    Ok(())
}
fn write_outline(x: &mut String, v: &OutlineView, s: bool) -> Result<()> {
    x.push_str("<p:outlineViewPr>");
    write_common(x, &v.view);
    if !v.slides.is_empty() {
        x.push_str("<p:sldLst>");
        for slide in &v.slides {
            x.push_str("<p:sld r:id=\"");
            esc(x, &slide.relationship_id);
            x.push('"');
            bool_attr(x, "collapse", slide.collapse);
            x.push_str("/>");
        }
        x.push_str("</p:sldLst>");
    }
    if let Some(e) = &v.extension_xml {
        opaque(x, e, s)?;
    }
    x.push_str("</p:outlineViewPr>");
    Ok(())
}

fn validate(v: &ViewProperties) -> Result<()> {
    for ext in [
        v.extension_xml.as_deref(),
        v.normal.as_ref().and_then(|x| x.extension_xml.as_deref()),
        v.slide.as_ref().and_then(|x| x.extension_xml.as_deref()),
        v.outline.as_ref().and_then(|x| x.extension_xml.as_deref()),
        v.notes_text
            .as_ref()
            .and_then(|x| x.extension_xml.as_deref()),
        v.sorter.as_ref().and_then(|x| x.extension_xml.as_deref()),
        v.notes.as_ref().and_then(|x| x.extension_xml.as_deref()),
    ]
    .into_iter()
    .flatten()
    {
        parse_dom(ext)?;
    }
    for common in v
        .slide
        .iter()
        .map(|x| &x.common)
        .chain(v.notes.iter().map(|x| &x.common))
    {
        if common.guides.len() > MAX_GUIDES {
            return Err(invalid("guide limit exceeded"));
        }
    }
    if v.outline
        .as_ref()
        .is_some_and(|x| x.slides.len() > MAX_SLIDES)
    {
        return Err(invalid("outline slide limit exceeded"));
    }
    Ok(())
}
fn opaque(x: &mut String, b: &[u8], strict: bool) -> Result<()> {
    parse_dom(b)?;
    let mut s = std::str::from_utf8(b).map_err(xml_error)?.to_string();
    if strict {
        s = s.replace(P, PS).replace(A, AS).replace(R, RS)
    } else {
        s = s.replace(PS, P).replace(AS, A).replace(RS, R)
    }
    x.push_str(&s);
    Ok(())
}
fn node_xml(n: &Node, strict: bool) -> Result<Vec<u8>> {
    let mut x = String::new();
    node_write(&mut x, n, strict)?;
    Ok(x.into_bytes())
}
fn node_write(x: &mut String, n: &Node, s: bool) -> Result<()> {
    x.push('<');
    x.push_str(&n.q);
    for (p, u) in &n.bindings {
        if p.is_empty() {
            x.push_str(" xmlns=\"")
        } else {
            x.push_str(" xmlns:");
            x.push_str(p);
            x.push_str("=\"")
        }
        esc(x, mapns(u, s));
        x.push('"');
    }
    for a in &n.attrs {
        x.push(' ');
        x.push_str(&a.q);
        x.push_str("=\"");
        esc(x, &a.v);
        x.push('"');
    }
    if n.content.is_empty() {
        x.push_str("/>");
        return Ok(());
    }
    x.push('>');
    for c in &n.content {
        match c {
            Content::Node(n) => node_write(x, n, s)?,
            Content::Text(v) => text(x, v),
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
    x.push_str(&n.q);
    x.push('>');
    Ok(())
}
fn mapns<'a>(v: &'a str, s: bool) -> &'a str {
    if s {
        match v {
            P => PS,
            A => AS,
            R => RS,
            _ => v,
        }
    } else {
        match v {
            PS => P,
            AS => A,
            RS => R,
            _ => v,
        }
    }
}

fn kids(n: &Node) -> Result<Vec<&Node>> {
    let mut v = Vec::new();
    for c in &n.content {
        match c {
            Content::Node(x) => v.push(x),
            Content::Text(x) if x.trim().is_empty() => {},
            Content::Comment(_) => {},
            _ => return Err(invalid("unexpected text in typed view properties")),
        }
    }
    Ok(v)
}
fn leaf(n: &Node) -> Result<()> {
    if kids(n)?.is_empty() {
        Ok(())
    } else {
        Err(invalid("leaf view property has children"))
    }
}
fn expect_p(n: &Node, l: &str) -> Result<()> {
    if (n.ns == P || n.ns == PS) && n.l == l {
        Ok(())
    } else {
        Err(invalid(format!("expected {l}")))
    }
}
fn expect_p_any(n: &Node) -> Result<()> {
    if n.ns == P || n.ns == PS {
        Ok(())
    } else {
        Err(invalid("expected PresentationML child"))
    }
}
fn aopt(n: &Node, l: &str) -> Result<Option<String>> {
    let mut v = None;
    for a in &n.attrs {
        if a.ns.is_empty() && a.l == l {
            if v.is_some() {
                return Err(invalid("duplicate attribute"));
            }
            v = Some(a.v.clone());
        }
    }
    Ok(v)
}
fn arel(n: &Node, l: &str) -> Result<String> {
    let mut v = None;
    for a in &n.attrs {
        if (a.ns == R || a.ns == RS) && a.l == l {
            if v.is_some() {
                return Err(invalid("duplicate relationship attribute"));
            }
            v = Some(a.v.clone());
        }
    }
    v.ok_or_else(|| invalid(format!("missing relationship attribute '{l}'")))
}
fn bopt(n: &Node, l: &str) -> Result<Option<bool>> {
    match aopt(n, l)?.as_deref() {
        None => Ok(None),
        Some("1" | "true") => Ok(Some(true)),
        Some("0" | "false") => Ok(Some(false)),
        _ => Err(invalid(format!("invalid boolean '{l}'"))),
    }
}
fn i64req(n: &Node, l: &str) -> Result<i64> {
    aopt(n, l)?
        .ok_or_else(|| invalid(format!("missing '{l}'")))?
        .parse()
        .map_err(|_| invalid(format!("invalid integer '{l}'")))
}
fn u32req(n: &Node, l: &str) -> Result<u32> {
    aopt(n, l)?
        .ok_or_else(|| invalid(format!("missing '{l}'")))?
        .parse()
        .map_err(|_| invalid(format!("invalid integer '{l}'")))
}
fn i32opt(n: &Node, l: &str) -> Result<Option<i32>> {
    aopt(n, l)?
        .map(|x| {
            x.parse()
                .map_err(|_| invalid(format!("invalid integer '{l}'")))
        })
        .transpose()
}
fn only(n: &Node, allowed: &[&str]) -> Result<()> {
    for a in &n.attrs {
        if !a.ns.is_empty() || !allowed.contains(&a.l.as_str()) {
            return Err(invalid(format!("unexpected attribute '{}'", a.q)));
        }
    }
    Ok(())
}
fn only_rel(n: &Node, plain: &[&str], rel: &str) -> Result<()> {
    for a in &n.attrs {
        if !((a.ns.is_empty() && plain.contains(&a.l.as_str()))
            || ((a.ns == R || a.ns == RS) && a.l == rel))
        {
            return Err(invalid(format!("unexpected attribute '{}'", a.q)));
        }
    }
    Ok(())
}
fn noattrs(n: &Node) -> Result<()> {
    only(n, &[])
}
fn set<T>(slot: &mut Option<T>, v: T) -> Result<()> {
    if slot.replace(v).is_some() {
        Err(invalid("duplicate view property"))
    } else {
        Ok(())
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
fn resolve(b: &[(String, String)], p: &str) -> Result<String> {
    b.iter()
        .rev()
        .find(|x| x.0 == p)
        .map(|x| x.1.clone())
        .ok_or_else(|| invalid(format!("unbound prefix '{p}'")))
}
fn attr(x: &mut String, n: &str, v: &str) {
    x.push(' ');
    x.push_str(n);
    x.push_str("=\"");
    esc(x, v);
    x.push('"')
}
fn bool_attr(x: &mut String, n: &str, v: Option<bool>) {
    if let Some(v) = v {
        attr(x, n, if v { "1" } else { "0" })
    }
}
fn esc(x: &mut String, v: &str) {
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
fn text(x: &mut String, v: &str) {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;"),
            '<' => x.push_str("&lt;"),
            '>' => x.push_str("&gt;"),
            _ => x.push(c),
        }
    }
}
macro_rules! en{($p:ident,$w:ident,$t:ty,$($s:literal=>$v:path),+)=>{fn $p(s:String)->Result<$t>{match s.as_str(){$($s=>Ok($v),)+_=>Err(invalid(format!("invalid enumeration '{s}'")))}}fn $w(v:$t)->&'static str{match v{$($v=>$s,)+}}}}
en!(parse_view,view_str,ViewKind,"sldView"=>ViewKind::Slide,"sldMasterView"=>ViewKind::SlideMaster,"notesView"=>ViewKind::Notes,"handoutView"=>ViewKind::Handout,"notesMasterView"=>ViewKind::NotesMaster,"outlineView"=>ViewKind::Outline,"sldSorterView"=>ViewKind::SlideSorter,"sldThumbnailView"=>ViewKind::SlideThumbnail);
en!(parse_splitter,splitter_str,SplitterState,"minimized"=>SplitterState::Minimized,"restored"=>SplitterState::Restored,"maximized"=>SplitterState::Maximized);
fn invalid(v: impl Into<String>) -> Box<dyn std::error::Error + Send + Sync> {
    std::io::Error::new(std::io::ErrorKind::InvalidData, v.into()).into()
}
fn xml_error(e: impl std::fmt::Display) -> Box<dyn std::error::Error + Send + Sync> {
    invalid(e.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    fn f(b: &[u8]) -> ViewProperties {
        let p = OpcPackage::from_bytes(b).unwrap();
        load_from_package(&p).unwrap().unwrap()
    }
    #[test]
    fn poi_prprops_guides_grid_and_strict_roundtrip() {
        let v = f(include_bytes!(
            "../../../../test-data/poi/test-data/slideshow/prProps.pptx"
        ));
        assert_eq!(v.slide.as_ref().unwrap().common.guides.len(), 2);
        assert_eq!(
            v.grid_spacing,
            Some(GridSpacing {
                cx: 72008,
                cy: 72008
            })
        );
        let x = v.to_xml(true).unwrap();
        let r = ViewProperties::parse(&x).unwrap();
        assert_eq!(
            r.slide.unwrap().common.view.scale_x,
            Ratio {
                numerator: 66,
                denominator: 100
            }
        );
    }
    #[test]
    fn poi_outline_and_splitters() {
        let v = f(include_bytes!(
            "../../../../test-data/poi/test-data/slideshow/45545_Comment.pptx"
        ));
        let n = v.normal.unwrap();
        assert_eq!(n.vertical_bar_state, Some(SplitterState::Minimized));
        assert_eq!(n.horizontal_bar_state, Some(SplitterState::Maximized));
        assert!(v.outline.is_some());
        assert_eq!(v.sorter.unwrap().view.origin.y, 1026);
    }
    #[test]
    fn libreoffice_ratio_and_sparse_views() {
        let v = f(include_bytes!(
            "../../../../test-data/libreoffice-core/oox/qa/unit/data/shape-text-alignment.pptx"
        ));
        assert_eq!(
            v.notes_text.unwrap().view.scale_x,
            Ratio {
                numerator: 1,
                denominator: 1
            }
        );
        assert!(v.sorter.is_none());
    }
    #[test]
    fn strict_mce_and_malformed() {
        let x = format!(
            r#"<p:viewPr xmlns:p="{PS}" xmlns:a="{AS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:u" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><u:x/></mc:Choice><mc:Fallback><p:gridSpacing cx="10" cy="20"/></mc:Fallback></mc:AlternateContent></p:viewPr>"#
        );
        assert_eq!(
            ViewProperties::parse(x.as_bytes()).unwrap().grid_spacing,
            Some(GridSpacing { cx: 10, cy: 20 })
        );
        for bad in [
            format!(r#"<p:viewPr xmlns:p="{P}" showComments="maybe"/>"#),
            format!(
                r#"<p:viewPr xmlns:p="{P}" xmlns:a="{A}"><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="1" d="0"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr></p:viewPr>"#
            ),
            format!(r#"<!DOCTYPE x><p:viewPr xmlns:p="{P}"/>"#),
        ] {
            assert!(
                ViewProperties::parse(bad.as_bytes()).is_err(),
                "accepted {bad}"
            );
        }
    }
}
