//! Bounding box calculation for WMF files without placeable headers
//!
//! When a WMF file lacks a placeable header, we must scan all drawing records
//! to determine the effective drawing area. This matches libwmf's wmf_scan behavior.

use super::super::constants::record;
use super::super::parser::WmfRecord;
use crate::common::binary::{read_i16_le, read_u16_le};

/// Calculate bounding box by scanning WMF records
pub struct BoundsCalculator;

impl BoundsCalculator {
    /// Scan WMF records to determine bounding box (matches libwmf wmf_scan)
    pub fn scan_records(records: &[WmfRecord]) -> (i16, i16, i16, i16) {
        let mut left = i16::MAX;
        let mut top = i16::MAX;
        let mut right = i16::MIN;
        let mut bottom = i16::MIN;
        let mut font_height = 0i16;

        for rec in records {
            match rec.function {
                record::RECTANGLE | record::ELLIPSE | record::ROUND_RECT
                    if rec.params.len() >= 8 =>
                {
                    let b = read_i16_le(&rec.params, 0).unwrap_or(0);
                    let r = read_i16_le(&rec.params, 2).unwrap_or(0);
                    let t = read_i16_le(&rec.params, 4).unwrap_or(0);
                    let l = read_i16_le(&rec.params, 6).unwrap_or(0);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, l, t);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, r, b);
                },
                record::POLYGON | record::POLYLINE if rec.params.len() >= 2 => {
                    let count = read_i16_le(&rec.params, 0).unwrap_or(0) as usize;
                    for i in 0..count.min((rec.params.len() - 2) / 4) {
                        let x = read_i16_le(&rec.params, 2 + i * 4).unwrap_or(0);
                        let y = read_i16_le(&rec.params, 4 + i * 4).unwrap_or(0);
                        Self::update(&mut left, &mut top, &mut right, &mut bottom, x, y);
                    }
                },
                record::LINE_TO | record::MOVE_TO | record::SET_PIXEL_V
                    if rec.params.len() >= 4 =>
                {
                    let y = read_i16_le(&rec.params, 0).unwrap_or(0);
                    let x = read_i16_le(&rec.params, 2).unwrap_or(0);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, x, y);
                },
                record::ARC | record::PIE | record::CHORD if rec.params.len() >= 16 => {
                    let yend = read_i16_le(&rec.params, 0).unwrap_or(0);
                    let xend = read_i16_le(&rec.params, 2).unwrap_or(0);
                    let ystart = read_i16_le(&rec.params, 4).unwrap_or(0);
                    let xstart = read_i16_le(&rec.params, 6).unwrap_or(0);
                    let b = read_i16_le(&rec.params, 8).unwrap_or(0);
                    let r = read_i16_le(&rec.params, 10).unwrap_or(0);
                    let t = read_i16_le(&rec.params, 12).unwrap_or(0);
                    let l = read_i16_le(&rec.params, 14).unwrap_or(0);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, l, t);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, r, b);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, xstart, ystart);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, xend, yend);
                },
                record::CREATE_FONT_INDIRECT if rec.params.len() >= 2 => {
                    font_height = read_i16_le(&rec.params, 0).unwrap_or(0).abs();
                },
                record::TEXT_OUT if rec.params.len() >= 6 => {
                    let len = read_u16_le(&rec.params, 0).unwrap_or(0) as usize;
                    let off = 2 + len.div_ceil(2) * 2;
                    if rec.params.len() >= off + 4 {
                        let y = read_i16_le(&rec.params, off).unwrap_or(0);
                        let x = read_i16_le(&rec.params, off + 2).unwrap_or(0);
                        Self::update(&mut left, &mut top, &mut right, &mut bottom, x, y);
                        Self::update(
                            &mut left,
                            &mut top,
                            &mut right,
                            &mut bottom,
                            x,
                            y + font_height,
                        );
                    }
                },
                record::EXT_TEXT_OUT if rec.params.len() >= 8 => {
                    let y = read_i16_le(&rec.params, 0).unwrap_or(0);
                    let x = read_i16_le(&rec.params, 2).unwrap_or(0);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, x, y);
                    Self::update(
                        &mut left,
                        &mut top,
                        &mut right,
                        &mut bottom,
                        x,
                        y + font_height,
                    );
                },
                _ => {},
            }
        }

        // Return bounds or default if empty
        if left != i16::MAX && right != i16::MIN && top != i16::MAX && bottom != i16::MIN {
            (left, top, right, bottom)
        } else {
            (0, 0, 1000, 1000)
        }
    }

    #[inline]
    fn update(left: &mut i16, top: &mut i16, right: &mut i16, bottom: &mut i16, x: i16, y: i16) {
        *left = (*left).min(x);
        *top = (*top).min(y);
        *right = (*right).max(x);
        *bottom = (*bottom).max(y);
    }
}
