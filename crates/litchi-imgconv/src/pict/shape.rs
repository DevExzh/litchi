//! QuickDraw shape geometry as SVG path data.
//!
//! QuickDraw coordinates lie on the grid lines between pixels. A shape's
//! interior covers the pixels inside its rectangle; a frame is drawn inside
//! the shape with the pen's width and height; a line is the area swept by
//! the rectangular pen, which hangs below and to the right of the pen
//! location (*Inside Macintosh: Imaging With QuickDraw*, chapter 3, "The
//! Graphics Pen" and "Drawing Lines and Shapes"). Arc angles are degrees
//! clockwise from twelve o'clock, scaled to the rectangle's proportions.

use std::f64::consts::PI;

use crate::svg_utils::write_num;

/// An output-space rectangle `(left, top, right, bottom)`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct Bounds {
    pub(super) left: f64,
    pub(super) top: f64,
    pub(super) right: f64,
    pub(super) bottom: f64,
}

impl Bounds {
    pub(super) fn is_empty(&self) -> bool {
        self.right <= self.left || self.bottom <= self.top
    }

    pub(super) fn width(&self) -> f64 {
        self.right - self.left
    }

    pub(super) fn height(&self) -> f64 {
        self.bottom - self.top
    }

    fn inset(&self, dx: f64, dy: f64) -> Self {
        Self {
            left: self.left + dx,
            top: self.top + dy,
            right: self.right - dx,
            bottom: self.bottom - dy,
        }
    }
}

fn point(path: &mut String, command: char, x: f64, y: f64) {
    path.push(command);
    write_num(path, x);
    path.push(' ');
    write_num(path, y);
}

pub(super) fn rect(path: &mut String, bounds: Bounds) {
    point(path, 'M', bounds.left, bounds.top);
    path.push('H');
    write_num(path, bounds.right);
    path.push('V');
    write_num(path, bounds.bottom);
    path.push('H');
    write_num(path, bounds.left);
    path.push('Z');
}

fn arc_to(path: &mut String, rx: f64, ry: f64, large: bool, clockwise: bool, x: f64, y: f64) {
    path.push('A');
    write_num(path, rx);
    path.push(' ');
    write_num(path, ry);
    path.push_str(if large { " 0 1 " } else { " 0 0 " });
    path.push(if clockwise { '1' } else { '0' });
    path.push(' ');
    write_num(path, x);
    path.push(' ');
    write_num(path, y);
}

pub(super) fn oval(path: &mut String, bounds: Bounds) {
    let rx = bounds.width() / 2.0;
    let ry = bounds.height() / 2.0;
    let cy = bounds.top + ry;
    point(path, 'M', bounds.left, cy);
    arc_to(path, rx, ry, true, true, bounds.right, cy);
    arc_to(path, rx, ry, true, true, bounds.left, cy);
    path.push('Z');
}

/// A rounded rectangle whose corner ovals are `oval_width` by
/// `oval_height`.
pub(super) fn round_rect(path: &mut String, bounds: Bounds, oval_width: f64, oval_height: f64) {
    let rx = (oval_width / 2.0).clamp(0.0, bounds.width() / 2.0);
    let ry = (oval_height / 2.0).clamp(0.0, bounds.height() / 2.0);
    if rx <= 0.0 || ry <= 0.0 {
        rect(path, bounds);
        return;
    }
    let Bounds {
        left,
        top,
        right,
        bottom,
    } = bounds;
    point(path, 'M', left + rx, top);
    path.push('H');
    write_num(path, right - rx);
    arc_to(path, rx, ry, false, true, right, top + ry);
    path.push('V');
    write_num(path, bottom - ry);
    arc_to(path, rx, ry, false, true, right - rx, bottom);
    path.push('H');
    write_num(path, left + rx);
    arc_to(path, rx, ry, false, true, left, bottom - ry);
    path.push('V');
    write_num(path, top + ry);
    arc_to(path, rx, ry, false, true, left + rx, top);
    path.push('Z');
}

/// Returns the output-space point at QuickDraw angle `degrees` on the oval
/// inscribed in `bounds`.
fn angle_point(bounds: Bounds, degrees: f64) -> (f64, f64) {
    let radians = degrees * PI / 180.0;
    let rx = bounds.width() / 2.0;
    let ry = bounds.height() / 2.0;
    (
        bounds.left + rx + rx * radians.sin(),
        bounds.top + ry - ry * radians.cos(),
    )
}

/// Normalizes an arc to a start angle and a non-negative clockwise sweep.
fn normalize_arc(start: f64, sweep: f64) -> (f64, f64) {
    if sweep < 0.0 {
        (start + sweep, (-sweep).min(360.0))
    } else {
        (start, sweep.min(360.0))
    }
}

/// A pie wedge from the oval's centre (`PaintArc`).
pub(super) fn wedge(path: &mut String, bounds: Bounds, start: f64, sweep: f64) {
    let (start, sweep) = normalize_arc(start, sweep);
    if sweep <= 0.0 {
        return;
    }
    if sweep >= 360.0 {
        oval(path, bounds);
        return;
    }
    let (x0, y0) = angle_point(bounds, start);
    let (x1, y1) = angle_point(bounds, start + sweep);
    let rx = bounds.width() / 2.0;
    let ry = bounds.height() / 2.0;
    point(path, 'M', bounds.left + rx, bounds.top + ry);
    point(path, 'L', x0, y0);
    arc_to(path, rx, ry, sweep > 180.0, true, x1, y1);
    path.push('Z');
}

/// The band between an oval and the same oval inset by the pen, limited to
/// the arc's angles (`FrameArc`).
pub(super) fn arc_band(path: &mut String, bounds: Bounds, start: f64, sweep: f64, pen: (f64, f64)) {
    let (start, sweep) = normalize_arc(start, sweep);
    if sweep <= 0.0 {
        return;
    }
    let inner = bounds.inset(pen.0, pen.1);
    if sweep >= 360.0 {
        frame_oval(path, bounds, pen);
        return;
    }
    let end = start + sweep;
    let large = sweep > 180.0;
    let (x0, y0) = angle_point(bounds, start);
    let (x1, y1) = angle_point(bounds, end);
    point(path, 'M', x0, y0);
    arc_to(
        path,
        bounds.width() / 2.0,
        bounds.height() / 2.0,
        large,
        true,
        x1,
        y1,
    );
    if inner.is_empty() {
        let center = (
            bounds.left + bounds.width() / 2.0,
            bounds.top + bounds.height() / 2.0,
        );
        point(path, 'L', center.0, center.1);
    } else {
        let (x2, y2) = angle_point(inner, end);
        let (x3, y3) = angle_point(inner, start);
        point(path, 'L', x2, y2);
        arc_to(
            path,
            inner.width() / 2.0,
            inner.height() / 2.0,
            large,
            false,
            x3,
            y3,
        );
    }
    path.push('Z');
}

/// Frames are filled with the even-odd rule: the shape and the shape inset
/// by the pen.
pub(super) fn frame_rect(path: &mut String, bounds: Bounds, pen: (f64, f64)) {
    rect(path, bounds);
    let inner = bounds.inset(pen.0, pen.1);
    if !inner.is_empty() {
        rect(path, inner);
    }
}

pub(super) fn frame_oval(path: &mut String, bounds: Bounds, pen: (f64, f64)) {
    oval(path, bounds);
    let inner = bounds.inset(pen.0, pen.1);
    if !inner.is_empty() {
        oval(path, inner);
    }
}

pub(super) fn frame_round_rect(
    path: &mut String,
    bounds: Bounds,
    oval: (f64, f64),
    pen: (f64, f64),
) {
    round_rect(path, bounds, oval.0, oval.1);
    let inner = bounds.inset(pen.0, pen.1);
    if !inner.is_empty() {
        round_rect(path, inner, oval.0 - 2.0 * pen.0, oval.1 - 2.0 * pen.1);
    }
}

pub(super) fn polygon(path: &mut String, points: &[(f64, f64)]) {
    for (index, &(x, y)) in points.iter().enumerate() {
        point(path, if index == 0 { 'M' } else { 'L' }, x, y);
    }
    if !points.is_empty() {
        path.push('Z');
    }
}

/// The area swept by a `pen.0` by `pen.1` pen moving from `from` to `to`:
/// the convex hull of the pen rectangle at both ends, wound clockwise so
/// that several sweeps in one path combine with the nonzero rule.
pub(super) fn pen_line(path: &mut String, from: (f64, f64), to: (f64, f64), pen: (f64, f64)) {
    let (width, height) = pen;
    let corners = |(x, y): (f64, f64)| {
        [
            (x, y),
            (x + width, y),
            (x + width, y + height),
            (x, y + height),
        ]
    };
    let mut points: Vec<(f64, f64)> = corners(from).into_iter().chain(corners(to)).collect();
    points.sort_by(|a, b| a.0.total_cmp(&b.0).then(a.1.total_cmp(&b.1)));
    points.dedup();
    let hull = convex_hull(&points);
    polygon(path, &hull);
}

/// Andrew's monotone chain over points sorted by x then y. In SVG's y-down
/// space the result runs clockwise.
fn convex_hull(points: &[(f64, f64)]) -> Vec<(f64, f64)> {
    if points.len() < 3 {
        return points.to_vec();
    }
    let cross = |o: (f64, f64), a: (f64, f64), b: (f64, f64)| {
        (a.0 - o.0) * (b.1 - o.1) - (a.1 - o.1) * (b.0 - o.0)
    };
    let mut lower: Vec<(f64, f64)> = Vec::new();
    for &p in points {
        while lower.len() >= 2 && cross(lower[lower.len() - 2], lower[lower.len() - 1], p) <= 0.0 {
            lower.pop();
        }
        lower.push(p);
    }
    let mut upper: Vec<(f64, f64)> = Vec::new();
    for &p in points.iter().rev() {
        while upper.len() >= 2 && cross(upper[upper.len() - 2], upper[upper.len() - 1], p) <= 0.0 {
            upper.pop();
        }
        upper.push(p);
    }
    lower.pop();
    upper.pop();
    lower.extend(upper);
    lower
}

#[cfg(test)]
mod tests {
    use super::*;

    const BOX: Bounds = Bounds {
        left: 0.0,
        top: 0.0,
        right: 10.0,
        bottom: 20.0,
    };

    #[test]
    fn horizontal_pen_line_covers_both_end_pixels() {
        let mut path = String::new();
        pen_line(&mut path, (0.0, 0.0), (10.0, 0.0), (1.0, 1.0));
        assert_eq!(path, "M0 0L11 0L11 1L0 1Z");
    }

    #[test]
    fn diagonal_pen_line_is_a_hexagon() {
        let mut path = String::new();
        pen_line(&mut path, (0.0, 0.0), (4.0, 4.0), (2.0, 2.0));
        assert_eq!(path, "M0 0L2 0L6 4L6 6L4 6L0 2Z");
    }

    #[test]
    fn frame_rect_insets_by_pen_and_degenerates_to_solid() {
        let mut path = String::new();
        frame_rect(&mut path, BOX, (2.0, 3.0));
        assert_eq!(path, "M0 0H10V20H0ZM2 3H8V17H2Z");
        let mut solid = String::new();
        frame_rect(&mut solid, BOX, (5.0, 1.0));
        assert_eq!(solid, "M0 0H10V20H0Z");
    }

    #[test]
    fn arc_angles_run_clockwise_from_twelve_oclock() {
        let mut path = String::new();
        wedge(&mut path, BOX, 0.0, 90.0);
        assert_eq!(path, "M5 10L5 0A5 10 0 0 1 10 10Z");
        let mut reversed = String::new();
        wedge(&mut reversed, BOX, 90.0, -90.0);
        assert_eq!(reversed, path);
    }

    #[test]
    fn round_rect_clamps_oval_to_the_rectangle() {
        let mut path = String::new();
        round_rect(&mut path, BOX, 40.0, 4.0);
        assert!(path.starts_with("M5 0H5A5 2 0 0 1 10 2V18"));
        let mut square = String::new();
        round_rect(&mut square, BOX, 0.0, 4.0);
        assert_eq!(square, "M0 0H10V20H0Z");
    }
}
