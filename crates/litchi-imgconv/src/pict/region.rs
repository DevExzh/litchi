//! QuickDraw regions.
//!
//! A region record is `rgnSize` (which counts the whole record), `rgnBBox`
//! and, for non-rectangular regions, a list of scan lines. Each scan line is
//! a vertical coordinate followed by horizontal inversion points and a
//! `0x7FFF` terminator; a final `0x7FFF` ends the list. An inversion point
//! toggles membership of every pixel at or right of it, from that line
//! downwards (*Inside Macintosh: Imaging With QuickDraw*, "Regions", and the
//! region data layout used by `ClipRgn` and the `*Rgn` opcodes).

use super::types::{PictRect, read_i16, read_u16};

const END: i16 = 0x7FFF;

/// A region decoded into non-overlapping rectangles.
#[derive(Debug, Clone, PartialEq)]
pub(super) struct Region {
    pub(super) bounds: PictRect,
    /// Rectangles as `(left, top, right, bottom)`.
    pub(super) rects: Vec<[i32; 4]>,
}

/// Why a region could not be decoded.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum RegionError {
    /// The record ends before its bounding box or scan-line list does.
    Truncated,
    /// The region needs more rectangles than the caller allows.
    TooComplex,
}

impl Region {
    /// Returns whether the region is exactly its bounding rectangle.
    pub(super) fn is_rectangular(&self) -> bool {
        let bounds = [
            i32::from(self.bounds.left),
            i32::from(self.bounds.top),
            i32::from(self.bounds.right),
            i32::from(self.bounds.bottom),
        ];
        self.rects.len() == 1 && self.rects[0] == bounds
    }

    /// Writes the region as SVG path data, mapping each coordinate through
    /// `map`.
    pub(super) fn write_path(&self, path: &mut String, map: impl Fn(f64, f64) -> (f64, f64)) {
        for &[left, top, right, bottom] in &self.rects {
            let (x0, y0) = map(f64::from(left), f64::from(top));
            let (x1, y1) = map(f64::from(right), f64::from(bottom));
            write_rect_path(path, x0, y0, x1, y1);
        }
    }
}

/// Decodes a region record. `max_rects` bounds the rectangle count.
pub(super) fn parse_region(data: &[u8], max_rects: usize) -> Result<Region, RegionError> {
    let size = usize::from(read_u16(data, 0).ok_or(RegionError::Truncated)?);
    let bounds = PictRect::read(data, 2).ok_or(RegionError::Truncated)?;
    if size <= 10 {
        let rects = if bounds.is_nonempty() {
            vec![[
                i32::from(bounds.left),
                i32::from(bounds.top),
                i32::from(bounds.right),
                i32::from(bounds.bottom),
            ]]
        } else {
            Vec::new()
        };
        return Ok(Region { bounds, rects });
    }
    let data = data.get(..size).ok_or(RegionError::Truncated)?;
    let mut offset = 10usize;
    let mut read = || -> Result<i16, RegionError> {
        let value = read_i16(data, offset).ok_or(RegionError::Truncated)?;
        offset += 2;
        Ok(value)
    };
    // Sorted horizontal inversion points currently in effect.
    let mut active: Vec<i32> = Vec::new();
    let mut previous_y: Option<i32> = None;
    let mut bands: Vec<(i32, i32, Vec<i32>)> = Vec::new();
    let mut rect_count = 0usize;
    loop {
        let y = read()?;
        if y == END {
            break;
        }
        let y = i32::from(y);
        if let Some(top) = previous_y
            && y > top
            && active.len() >= 2
        {
            match bands.last_mut() {
                Some((_, bottom, spans)) if *bottom == top && *spans == active => *bottom = y,
                _ => {
                    rect_count = rect_count.saturating_add(active.len() / 2);
                    if rect_count > max_rects {
                        return Err(RegionError::TooComplex);
                    }
                    bands.push((top, y, active.clone()));
                },
            }
        }
        loop {
            let x = read()?;
            if x == END {
                break;
            }
            let x = i32::from(x);
            match active.binary_search(&x) {
                Ok(index) => {
                    active.remove(index);
                },
                Err(index) => active.insert(index, x),
            }
        }
        previous_y = Some(y);
    }
    let rects = bands
        .into_iter()
        .flat_map(|(top, bottom, spans)| {
            spans
                .chunks_exact(2)
                .map(|span| [span[0], top, span[1], bottom])
                .collect::<Vec<_>>()
        })
        .collect();
    Ok(Region { bounds, rects })
}

/// Writes an axis-aligned rectangle as SVG path data.
pub(super) fn write_rect_path(path: &mut String, x0: f64, y0: f64, x1: f64, y1: f64) {
    path.push('M');
    crate::svg_utils::write_num(path, x0);
    path.push(' ');
    crate::svg_utils::write_num(path, y0);
    path.push('H');
    crate::svg_utils::write_num(path, x1);
    path.push('V');
    crate::svg_utils::write_num(path, y1);
    path.push('H');
    crate::svg_utils::write_num(path, x0);
    path.push('Z');
}

#[cfg(test)]
pub(super) mod tests {
    use super::*;

    pub(in crate::pict) fn region_record(bounds: [i16; 4], lines: &[(i16, &[i16])]) -> Vec<u8> {
        let mut body = Vec::new();
        for value in bounds {
            body.extend_from_slice(&value.to_be_bytes());
        }
        if !lines.is_empty() {
            for (y, xs) in lines {
                body.extend_from_slice(&y.to_be_bytes());
                for x in *xs {
                    body.extend_from_slice(&x.to_be_bytes());
                }
                body.extend_from_slice(&END.to_be_bytes());
            }
            body.extend_from_slice(&END.to_be_bytes());
        }
        let mut data = ((body.len() + 2) as u16).to_be_bytes().to_vec();
        data.extend_from_slice(&body);
        data
    }

    #[test]
    fn rectangular_region_is_its_bounds() {
        let region = parse_region(&region_record([1, 2, 3, 4], &[]), 10).unwrap();
        assert_eq!(region.rects, vec![[2, 1, 4, 3]]);
        assert!(region.is_rectangular());
    }

    #[test]
    fn inversion_points_toggle_spans_and_merge_bands() {
        // An L shape: rows 0..2 cover x 0..4, rows 2..4 cover x 0..2.
        let region = parse_region(
            &region_record([0, 0, 4, 4], &[(0, &[0, 4]), (2, &[2, 4]), (4, &[0, 2])]),
            10,
        )
        .unwrap();
        assert_eq!(region.rects, vec![[0, 0, 4, 2], [0, 2, 2, 4]]);
        assert!(!region.is_rectangular());

        let mut path = String::new();
        region.write_path(&mut path, |x, y| (x, y));
        assert_eq!(path, "M0 0H4V2H0ZM0 2H2V4H0Z");
    }

    #[test]
    fn identical_scan_lines_merge_vertically() {
        let region = parse_region(
            &region_record([0, 0, 4, 5], &[(0, &[1, 5]), (2, &[]), (4, &[1, 5])]),
            10,
        )
        .unwrap();
        assert_eq!(region.rects, vec![[1, 0, 5, 4]]);
        assert!(!region.is_rectangular());
    }

    #[test]
    fn truncated_and_oversized_regions_are_rejected() {
        let mut data = region_record([0, 0, 4, 4], &[(0, &[0, 4]), (4, &[0, 4])]);
        data.truncate(data.len() - 2);
        assert_eq!(parse_region(&data, 10), Err(RegionError::Truncated));
        let data = region_record(
            [0, 0, 4, 8],
            &[(0, &[0, 1, 2, 3]), (1, &[4, 5]), (2, &[0, 1, 2, 3, 4, 5])],
        );
        assert_eq!(parse_region(&data, 2), Err(RegionError::TooComplex));
    }
}
