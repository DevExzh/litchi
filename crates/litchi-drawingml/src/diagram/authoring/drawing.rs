//! Litchi's own deterministic layout of an authored `SmartArt` graphic,
//! written as the cached drawing (`dsp:drawing`).
//!
//! Every shape's `modelId` is the presentation point that shows it in the
//! data model (`40000 + node`, `70000 + node` for transitions), so a reader
//! can map a drawn shape back to its node. Integer EMU arithmetic keeps the
//! output identical on every platform.

use litchi_core::xml::escape_xml;

use super::{Flat, Layout, Node, Spec};

const DSP: &str = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const DGM: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

/// EMU per point.
const EMU_PER_POINT: i64 = 12_700;

/// A box in frame coordinates.
#[derive(Debug, Clone, Copy)]
struct Rect {
    x: i64,
    y: i64,
    w: i64,
    h: i64,
}

/// The style of a node shape.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Look {
    Solid,
    Translucent,
}

struct Writer {
    xml: String,
}

impl Writer {
    fn new(capacity: usize) -> Self {
        let mut xml = String::with_capacity(capacity);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(&format!(
            r#"<dsp:drawing xmlns:dgm="{DGM}" xmlns:dsp="{DSP}" xmlns:a="{A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>"#
        ));
        Self { xml }
    }

    fn finish(mut self) -> String {
        self.xml.push_str("</dsp:spTree></dsp:drawing>");
        self.xml
    }

    fn open(&mut self, model: u32, rect: Rect, flip_h: bool, geometry: &str, adjust: Option<i64>) {
        self.xml.push_str(&format!(
            r#"<dsp:sp modelId="{model}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm{}><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/></a:xfrm><a:prstGeom prst="{geometry}"><a:avLst>"#,
            if flip_h { r#" flipH="1""# } else { "" },
            rect.x,
            rect.y,
            rect.w,
            rect.h
        ));
        if let Some(adjust) = adjust {
            self.xml
                .push_str(&format!(r#"<a:gd name="adj" fmla="val {adjust}"/>"#));
        }
        self.xml.push_str("</a:avLst></a:prstGeom>");
    }

    /// A node shape with its text: the node's lines, then its descendants
    /// as bullet lines when `bullets` is set.
    fn node(
        &mut self,
        model: u32,
        rect: Rect,
        geometry: &str,
        adjust: Option<i64>,
        look: Look,
        node: &Node,
        bullets: bool,
    ) {
        self.open(model, rect, false, geometry, adjust);
        match look {
            Look::Solid => self.xml.push_str(r#"<a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></dsp:spPr><dsp:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style>"#),
            Look::Translucent => self.xml.push_str(r#"<a:solidFill><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:solidFill><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></dsp:spPr><dsp:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></dsp:style>"#),
        }
        let mut lines: Vec<(u32, &str)> = node.text.split('\n').map(|line| (0, line)).collect();
        if bullets {
            fn walk<'n>(out: &mut Vec<(u32, &'n str)>, nodes: &'n [Node], level: u32) {
                for child in nodes {
                    for line in child.text.split('\n') {
                        out.push((level, line));
                    }
                    walk(out, &child.children, level + 1);
                }
            }
            walk(&mut lines, &node.children, 1);
        }
        let longest = lines
            .iter()
            .map(|(level, line)| line.chars().count() + 2 * (*level as usize))
            .max()
            .unwrap_or(1)
            .max(1);
        let area = text_rect(rect, geometry, adjust);
        let size = font_size(area, lines.len().max(1), longest);
        let inset = (area.w.min(area.h) / 20).max(0);
        let text_color = if look == Look::Translucent {
            "tx1"
        } else {
            "lt1"
        };
        self.xml.push_str(&format!(
            r#"<dsp:txBody><a:bodyPr spcFirstLastPara="0" vert="horz" wrap="square" lIns="{inset}" tIns="{inset}" rIns="{inset}" bIns="{inset}" numCol="1" spcCol="1270" anchor="ctr" anchorCtr="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>"#
        ));
        for (level, line) in &lines {
            let line = line.strip_suffix('\r').unwrap_or(line);
            if *level == 0 {
                self.xml.push_str(r#"<a:p><a:pPr marL="0" lvl="0" indent="0" algn="ctr" defTabSz="914400"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:spcAft><a:spcPct val="35000"/></a:spcAft><a:buNone/></a:pPr>"#);
            } else {
                let margin = 114_300 * i64::from(*level);
                self.xml.push_str(&format!(
                    r#"<a:p><a:pPr marL="{margin}" lvl="{level}" indent="-114300" algn="l" defTabSz="914400"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:spcAft><a:spcPct val="15000"/></a:spcAft><a:buChar char="&#8226;"/></a:pPr>"#
                ));
            }
            let run_size = if *level == 0 { size } else { size * 4 / 5 };
            if line.is_empty() {
                self.xml.push_str(&format!(
                    r#"<a:endParaRPr lang="en-US" sz="{run_size}" kern="1200"/></a:p>"#
                ));
            } else {
                self.xml.push_str(&format!(
                    r#"<a:r><a:rPr lang="en-US" sz="{run_size}" kern="1200"><a:solidFill><a:schemeClr val="{text_color}"/></a:solidFill></a:rPr><a:t>{}</a:t></a:r></a:p>"#,
                    escape_xml(line)
                ));
            }
        }
        self.xml.push_str(&format!(
            r#"</dsp:txBody><dsp:txXfrm><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/></dsp:txXfrm></dsp:sp>"#,
            area.x, area.y, area.w, area.h
        ));
    }

    /// A text-less connecting shape (arrow or line).
    fn connector(&mut self, model: u32, rect: Rect, flip_h: bool, geometry: &str) {
        self.open(model, rect, flip_h, geometry, None);
        if geometry == "line" {
            self.xml.push_str(r#"<a:noFill/><a:ln w="12700"><a:solidFill><a:schemeClr val="accent1"><a:shade val="60000"/></a:schemeClr></a:solidFill></a:ln></dsp:spPr><dsp:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"/></dsp:style></dsp:sp>"#);
        } else {
            self.xml.push_str(r#"<a:solidFill><a:schemeClr val="accent1"><a:tint val="60000"/></a:schemeClr></a:solidFill><a:ln><a:noFill/></a:ln></dsp:spPr><dsp:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style></dsp:sp>"#);
        }
    }
}

/// The part of a shape its text uses: the preset geometry's text rectangle,
/// approximately.
fn text_rect(rect: Rect, geometry: &str, adjust: Option<i64>) -> Rect {
    match geometry {
        "triangle" => Rect {
            x: rect.x + rect.w / 4,
            y: rect.y + rect.h / 2,
            w: rect.w / 2,
            h: rect.h / 2,
        },
        "ellipse" => Rect {
            x: rect.x + rect.w * 15 / 100,
            y: rect.y + rect.h * 15 / 100,
            w: rect.w * 70 / 100,
            h: rect.h * 70 / 100,
        },
        "trapezoid" => {
            let inset = adjust.unwrap_or(25_000) * rect.w.min(rect.h) / 100_000;
            Rect {
                x: rect.x + inset,
                y: rect.y,
                w: (rect.w - 2 * inset).max(1),
                h: rect.h,
            }
        },
        _ => rect,
    }
}

/// A font size in hundredths of a point that fits `lines` lines of at most
/// `longest` characters in `rect`, between 8 pt and 40 pt.
fn font_size(rect: Rect, lines: usize, longest: usize) -> i64 {
    let lines = i64::try_from(lines).unwrap_or(i64::MAX).max(1);
    let longest = i64::try_from(longest).unwrap_or(i64::MAX).max(1);
    // Line height ≈ 1.25 × size; glyph width ≈ 0.55 × size; 80% usable.
    let by_height = rect.h.saturating_mul(80) / (lines.saturating_mul(125) * EMU_PER_POINT / 100);
    let by_width = rect.w.saturating_mul(80) / (longest.saturating_mul(55) * EMU_PER_POINT / 100);
    by_height.min(by_width).clamp(800, 4000) / 100 * 100
}

fn ceil_div(value: i64, divisor: i64) -> i64 {
    (value + divisor - 1) / divisor
}

fn count(value: usize) -> i64 {
    i64::try_from(value).unwrap_or(i64::MAX)
}

/// Write the drawing of `spec` for a `width` × `height` EMU frame.
pub(super) fn write(spec: &Spec, width: i64, height: i64, capacity: usize) -> String {
    let mut out = Writer::new(capacity);
    let flat = spec.flatten();
    let top: Vec<&Flat<'_>> = flat.iter().filter(|item| item.depth == 0).collect();
    if top.is_empty() {
        return out.finish();
    }
    let n = count(top.len());
    match spec.layout {
        Layout::List | Layout::Matrix => {
            let columns = if spec.layout == Layout::Matrix {
                n.min(2)
            } else {
                best_columns(n, width, height)
            };
            let rows = ceil_div(n, columns);
            // Cells with 10% gaps; boxes keep a 3:2 aspect at most.
            let cell_w = width * 10 / (columns * 11 - 1).max(1);
            let cell_h = height * 10 / (rows * 11 - 1).max(1);
            let (box_w, box_h) = if spec.layout == Layout::Matrix {
                (cell_w, cell_h)
            } else {
                let w = cell_w.min(cell_h * 3 / 2);
                (w, w * 2 / 3)
            };
            let grid_w = box_w * columns + box_w / 10 * (columns - 1);
            let grid_h = box_h * rows + box_h / 10 * (rows - 1);
            let x0 = (width - grid_w) / 2;
            let y0 = (height - grid_h) / 2;
            for (index, item) in top.iter().enumerate() {
                let index = count(index);
                let (row, column) = (index / columns, index % columns);
                let rect = Rect {
                    x: x0 + column * (box_w + box_w / 10),
                    y: y0 + row * (box_h + box_h / 10),
                    w: box_w,
                    h: box_h,
                };
                let geometry = if spec.layout == Layout::Matrix {
                    "roundRect"
                } else {
                    "rect"
                };
                out.node(
                    40_000 + item.id,
                    rect,
                    geometry,
                    None,
                    Look::Solid,
                    item.node,
                    true,
                );
            }
        },
        Layout::Process => {
            // n boxes and n - 1 arrows of 0.4 box width, with 0.1 gaps.
            let box_w = width * 10 / (n * 10 + (n - 1) * 6);
            let box_h = (box_w * 3 / 4).min(height);
            let arrow_w = box_w * 4 / 10;
            let arrow_h = box_h * 4 / 10;
            let y = (height - box_h) / 2;
            let total = box_w * n + (arrow_w + box_w / 5) * (n - 1);
            let mut x = (width - total) / 2;
            for (index, item) in top.iter().enumerate() {
                let rect = Rect {
                    x,
                    y,
                    w: box_w,
                    h: box_h,
                };
                out.node(
                    40_000 + item.id,
                    rect,
                    "roundRect",
                    None,
                    Look::Solid,
                    item.node,
                    true,
                );
                x += box_w + box_w / 10;
                if index + 1 < top.len() {
                    out.connector(
                        70_000 + item.id,
                        Rect {
                            x,
                            y: y + (box_h - arrow_h) / 2,
                            w: arrow_w,
                            h: arrow_h,
                        },
                        false,
                        "rightArrow",
                    );
                    x += arrow_w + box_w / 10;
                }
            }
        },
        Layout::Cycle => {
            let side = width.min(height);
            // sin(π/n) in parts per million, for n ≥ 2.
            let sine = if n <= 1 { 1_000_000 } else { sine_ppm(n) };
            let diameter = if n <= 1 {
                side / 2
            } else {
                (side * sine / (1_150_000 + sine)).min(side * 35 / 100)
            };
            let radius = (side - diameter) / 2;
            let (cx, cy) = (width / 2, height / 2);
            for (index, item) in top.iter().enumerate() {
                let (dx, dy) = if n <= 1 {
                    (0, 0)
                } else {
                    unit_ppm(count(index), n)
                };
                let rect = Rect {
                    x: cx + radius * dx / 1_000_000 - diameter / 2,
                    y: cy + radius * dy / 1_000_000 - diameter / 2,
                    w: diameter,
                    h: diameter,
                };
                out.node(
                    40_000 + item.id,
                    rect,
                    "ellipse",
                    None,
                    Look::Solid,
                    item.node,
                    true,
                );
            }
        },
        Layout::Venn => {
            let diameter = (height * 9 / 10).min(width * 100 / (100 + (n - 1) * 65));
            let step = diameter * 65 / 100;
            let total = diameter + step * (n - 1);
            let x0 = (width - total) / 2;
            let y = (height - diameter) / 2;
            for (index, item) in top.iter().enumerate() {
                let rect = Rect {
                    x: x0 + step * count(index),
                    y,
                    w: diameter,
                    h: diameter,
                };
                out.node(
                    40_000 + item.id,
                    rect,
                    "ellipse",
                    None,
                    Look::Translucent,
                    item.node,
                    true,
                );
            }
        },
        Layout::Pyramid => {
            let total_h = height * 9 / 10;
            let base = width.min(total_h * 6 / 5);
            let slice = total_h / n;
            let x_center = width / 2;
            let y0 = (height - total_h) / 2;
            for (index, item) in top.iter().enumerate() {
                let index = count(index);
                let top_w = base * index / n;
                let bottom_w = base * (index + 1) / n;
                let rect = Rect {
                    x: x_center - bottom_w / 2,
                    y: y0 + slice * index,
                    w: bottom_w,
                    h: slice,
                };
                if index == 0 {
                    out.node(
                        40_000 + item.id,
                        rect,
                        "triangle",
                        None,
                        Look::Solid,
                        item.node,
                        true,
                    );
                } else {
                    let inset = (bottom_w - top_w) / 2;
                    let adjust = inset * 100_000 / bottom_w.min(slice).max(1);
                    out.node(
                        40_000 + item.id,
                        rect,
                        "trapezoid",
                        Some(adjust),
                        Look::Solid,
                        item.node,
                        true,
                    );
                }
            }
        },
        Layout::Hierarchy => hierarchy(&mut out, &flat, width, height),
    }
    out.finish()
}

/// The column count whose 3:2 boxes are largest in the frame.
fn best_columns(n: i64, width: i64, height: i64) -> i64 {
    let mut best = (1, 0);
    for columns in 1..=n {
        let rows = ceil_div(n, columns);
        let cell_w = width * 10 / (columns * 11 - 1);
        let cell_h = height * 10 / (rows * 11 - 1);
        let w = cell_w.min(cell_h * 3 / 2);
        if w > best.1 {
            best = (columns, w);
        }
    }
    best.0
}

/// `sin(π / n)` in parts per million, by a table for small `n` and the
/// small-angle approximation beyond it.
fn sine_ppm(n: i64) -> i64 {
    const TABLE: [i64; 13] = [
        0, 0, 1_000_000, 866_025, 707_107, 587_785, 500_000, 433_884, 382_683, 342_020, 309_017,
        281_733, 258_819,
    ];
    usize::try_from(n)
        .ok()
        .and_then(|index| TABLE.get(index).copied())
        .unwrap_or(3_141_593 / n)
}

/// The unit vector at `index` of `n` positions around a ring, starting at
/// the top and going clockwise, in parts per million.
fn unit_ppm(index: i64, n: i64) -> (i64, i64) {
    // Angle in millionths of a full turn.
    let turn = index * 1_000_000 / n;
    let (sine, cosine) = sin_cos_turn(turn);
    (sine, -cosine)
}

/// `(sin, cos)` of `turn` millionths of a full turn, in parts per million.
fn sin_cos_turn(turn: i64) -> (i64, i64) {
    // Reduce to the first octant and use a Taylor series with integer
    // fixed point (1e6); accurate to a few ppm, ample for placement.
    let turn = turn.rem_euclid(1_000_000);
    let quadrant = turn / 250_000;
    let rest = turn % 250_000;
    // Angle in radians × 1e6.
    let angle = rest * 6_283_185 / 1_000_000;
    let a2 = angle * angle / 1_000_000;
    let sine = angle - angle * a2 / 6_000_000 + angle * a2 / 1_000_000 * a2 / 120_000_000
        - angle * a2 / 1_000_000 * a2 / 1_000_000 * a2 / 5_040_000_000;
    let cosine = 1_000_000 - a2 / 2 + a2 * a2 / 24_000_000 - a2 * a2 / 1_000_000 * a2 / 720_000_000;
    match quadrant {
        0 => (sine, cosine),
        1 => (cosine, -sine),
        2 => (-sine, -cosine),
        _ => (-cosine, sine),
    }
}

/// A layered tree: leaves share the width, each parent centred over its
/// leaves, one row per level, a line from each parent to each child.
fn hierarchy(out: &mut Writer, flat: &[Flat<'_>], width: i64, height: i64) {
    let mut leaves = vec![0_i64; flat.len()];
    for index in (0..flat.len()).rev() {
        let item = flat[index];
        if leaves[index] == 0 {
            leaves[index] = 1;
        }
        if item.parent != 0 {
            leaves[item.parent as usize - 1] += leaves[index];
        }
    }
    let total: i64 = flat
        .iter()
        .enumerate()
        .filter(|(_, item)| item.parent == 0)
        .map(|(index, _)| leaves[index])
        .sum();
    let levels = flat
        .iter()
        .map(|item| i64::from(item.depth))
        .max()
        .unwrap_or(0)
        + 1;
    let slot = width * 10 / (total * 12 - 2).max(1) * 12 / 10;
    let box_w = slot * 10 / 12;
    let row = height * 10 / (levels * 15 - 5).max(1);
    let box_h = row.min(box_w * 2 / 3);
    let pitch = box_h * 3 / 2;
    let grid_h = box_h + pitch * (levels - 1);
    let y0 = (height - grid_h) / 2;
    let grid_w = slot * total - (slot - box_w);
    // Left edge (in slots) of each node's leaf span.
    let mut start = vec![0_i64; flat.len()];
    let mut next_child = vec![0_i64; flat.len() + 1];
    for (index, item) in flat.iter().enumerate() {
        let parent = item.parent as usize;
        let base = if parent == 0 { 0 } else { start[parent - 1] };
        start[index] = base + next_child[parent];
        next_child[parent] += leaves[index];
    }
    let x0 = (width - grid_w) / 2;
    let rect_of = |index: usize| {
        let center = x0 + start[index] * slot + leaves[index] * slot / 2 - (slot - box_w) / 2;
        Rect {
            x: center - box_w / 2,
            y: y0 + i64::from(flat[index].depth) * pitch,
            w: box_w,
            h: box_h,
        }
    };
    for (index, item) in flat.iter().enumerate() {
        if item.parent == 0 {
            continue;
        }
        let parent = rect_of(item.parent as usize - 1);
        let child = rect_of(index);
        let (px, py) = (parent.x + parent.w / 2, parent.y + parent.h);
        let (cx, cy) = (child.x + child.w / 2, child.y);
        out.connector(
            70_000 + item.id,
            Rect {
                x: px.min(cx),
                y: py,
                w: (cx - px).abs(),
                h: cy - py,
            },
            cx < px,
            "line",
        );
    }
    for (index, item) in flat.iter().enumerate() {
        out.node(
            40_000 + item.id,
            rect_of(index),
            "roundRect",
            None,
            Look::Solid,
            item.node,
            false,
        );
    }
}
