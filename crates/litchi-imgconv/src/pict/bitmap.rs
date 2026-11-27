//! `BitMap`, `PixMap` and `PixPat` structures embedded in PICT records.
//!
//! Layouts follow *Inside Macintosh: Imaging With QuickDraw*, chapter 4
//! ("Color QuickDraw": `PixMap`, `ColorTable`, `PixPat`) and Appendix A
//! ("Picture Opcodes": `BitsRect`, `BitsRgn`, `PackBitsRect`, `PackBitsRgn`,
//! `DirectBitsRect`, `DirectBitsRgn`, `BkPixPat`, `PnPixPat`, `FillPixPat`).

use super::data::{Unpacked, unpack_bits, unpack_words};
use super::types::{PictRect, read_u16};

/// Bytes of a `PixMap` record after `baseAddr`: `rowBytes` through
/// `pmReserved`.
const PIXMAP_BYTES: usize = 46;
/// Bytes of a `BitMap` record after `baseAddr`: `rowBytes` and `bounds`.
const BITMAP_BYTES: usize = 10;

/// Pixel-decoding ceilings applied before any pixel buffer is allocated.
#[derive(Debug, Clone, Copy)]
pub(super) struct PixelLimits {
    pub(super) max_width: u32,
    pub(super) max_height: u32,
    pub(super) max_pixels: u64,
}

/// How a pixel map stores its rows.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RowEncoding {
    /// Rows are stored unpacked, `stored` bytes each.
    Raw { stored: usize },
    /// Rows are PackBits-compressed with a leading byte count.
    Packed { item: PackItem, unpacked: usize },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PackItem {
    Byte,
    Word,
}

/// A `BitMap` or `PixMap` header together with its colour table.
#[derive(Debug, Clone)]
pub(super) struct PixMap {
    pub(super) row_bytes: usize,
    pub(super) bounds: PictRect,
    pub(super) pack_type: u16,
    pub(super) pixel_size: u16,
    pub(super) cmp_count: u16,
    /// Whether the record is a monochrome `BitMap` (no `PixMap` flag).
    pub(super) is_bitmap: bool,
    /// Colour table for indexed pixels, indexed by pixel value.
    pub(super) palette: Vec<[u8; 3]>,
}

impl PixMap {
    fn rows(&self) -> usize {
        usize::try_from(self.bounds.height()).unwrap_or(0)
    }

    fn columns(&self) -> usize {
        usize::try_from(self.bounds.width()).unwrap_or(0)
    }

    /// Row encoding for pixel data that follows this header.
    ///
    /// `always_raw` is set for `BitsRect`/`BitsRgn`, whose rows are never
    /// packed. Otherwise rows narrower than eight bytes are stored unpacked,
    /// `packType` 1 means unpacked, `packType` 2 drops the pad byte of 32-bit
    /// pixels, 16-bit maps pack two-byte items and 32-bit maps pack each
    /// component plane separately.
    fn encoding(&self, always_raw: bool) -> RowEncoding {
        if always_raw || self.row_bytes < 8 || self.pack_type == 1 {
            return RowEncoding::Raw {
                stored: self.row_bytes,
            };
        }
        match self.pixel_size {
            32 if self.pack_type == 2 => RowEncoding::Raw {
                stored: self.row_bytes / 4 * 3,
            },
            32 => RowEncoding::Packed {
                item: PackItem::Byte,
                unpacked: self
                    .columns()
                    .saturating_mul(usize::from(self.cmp_count.clamp(1, 4))),
            },
            16 => RowEncoding::Packed {
                item: PackItem::Word,
                unpacked: self.row_bytes,
            },
            _ => RowEncoding::Packed {
                item: PackItem::Byte,
                unpacked: self.row_bytes,
            },
        }
    }
}

/// Parses a `BitMap` or `PixMap` header (without `baseAddr`) and, for
/// indexed pixel maps that carry one, the colour table that follows.
///
/// `direct` is set for `DirectBitsRect`/`DirectBitsRgn`, whose pixel maps
/// are always `PixMap` records without a colour table.
fn parse_pixmap(data: &[u8], start: usize, direct: bool) -> Option<(PixMap, usize)> {
    let row_bytes_word = read_u16(data, start)?;
    let bounds = PictRect::read(data, start.checked_add(2)?)?;
    let row_bytes = usize::from(row_bytes_word & 0x3FFF);
    if !direct && row_bytes_word & 0x8000 == 0 {
        data.get(start..start.checked_add(BITMAP_BYTES)?)?;
        let pixmap = PixMap {
            row_bytes,
            bounds,
            pack_type: 0,
            pixel_size: 1,
            cmp_count: 1,
            is_bitmap: true,
            palette: vec![[255, 255, 255], [0, 0, 0]],
        };
        return Some((pixmap, start + BITMAP_BYTES));
    }
    data.get(start..start.checked_add(PIXMAP_BYTES)?)?;
    let pack_type = read_u16(data, start + 12)?;
    let pixel_size = read_u16(data, start + 28)?;
    let cmp_count = read_u16(data, start + 30)?;
    let mut pos = start + PIXMAP_BYTES;
    let mut palette = Vec::new();
    if !direct {
        let (table, end) = parse_color_table(data, pos, pixel_size)?;
        palette = table;
        pos = end;
    }
    let pixmap = PixMap {
        row_bytes,
        bounds,
        pack_type,
        pixel_size,
        cmp_count,
        is_bitmap: false,
        palette,
    };
    Some((pixmap, pos))
}

/// Parses a `ColorTable`: `ctSeed`, `ctFlags`, `ctSize` (entry count minus
/// one) and `ctSize + 1` `ColorSpec` entries of a value and an `RGBColor`.
///
/// When bit 15 of `ctFlags` is set the table is a device table whose entry
/// order gives the pixel value; otherwise each entry's value field does.
fn parse_color_table(data: &[u8], start: usize, pixel_size: u16) -> Option<(Vec<[u8; 3]>, usize)> {
    let flags = read_u16(data, start.checked_add(4)?)?;
    let size = read_u16(data, start.checked_add(6)?)?;
    // A negative ctSize (0xFFFF) denotes an empty table.
    let entries = if size >= 0x8000 {
        0
    } else {
        usize::from(size) + 1
    };
    let end = start.checked_add(8)?.checked_add(entries.checked_mul(8)?)?;
    data.get(start..end)?;
    let capacity = if pixel_size <= 8 {
        1usize << pixel_size.max(1)
    } else {
        256
    };
    let mut palette = vec![[0u8; 3]; capacity];
    for index in 0..entries {
        let entry = start + 8 + index * 8;
        let value = if flags & 0x8000 != 0 {
            index
        } else {
            usize::from(read_u16(data, entry)?)
        };
        let color = [data[entry + 2], data[entry + 4], data[entry + 6]];
        if let Some(slot) = palette.get_mut(value) {
            *slot = color;
        }
    }
    Some((palette, end))
}

/// Returns the byte length of packed or unpacked pixel rows starting at
/// `start`, or `None` when the input ends first.
fn pixel_data_len(data: &[u8], start: usize, pixmap: &PixMap, always_raw: bool) -> Option<usize> {
    let rows = pixmap.rows();
    if rows == 0 || pixmap.columns() == 0 || pixmap.row_bytes == 0 {
        return Some(0);
    }
    match pixmap.encoding(always_raw) {
        RowEncoding::Raw { stored } => {
            let length = stored.checked_mul(rows)?;
            data.get(start..start.checked_add(length)?)?;
            Some(length)
        },
        RowEncoding::Packed { .. } => {
            let wide = pixmap.row_bytes > 250;
            let mut pos = start;
            for _ in 0..rows {
                let (count, width) = packed_count(data, pos, wide)?;
                pos = pos.checked_add(width)?.checked_add(count)?;
            }
            data.get(start..pos)?;
            Some(pos - start)
        },
    }
}

fn packed_count(data: &[u8], pos: usize, wide: bool) -> Option<(usize, usize)> {
    if wide {
        Some((usize::from(read_u16(data, pos)?), 2))
    } else {
        Some((usize::from(*data.get(pos)?), 1))
    }
}

/// A parsed bitmap opcode (`BitsRect` through `DirectBitsRgn`).
#[derive(Debug, Clone)]
pub(super) struct BitsRecord<'a> {
    pub(super) pixmap: PixMap,
    pub(super) src_rect: PictRect,
    pub(super) dst_rect: PictRect,
    pub(super) mode: u16,
    /// The mask region of the `*Rgn` variants.
    pub(super) mask_region: Option<&'a [u8]>,
    always_raw: bool,
    pixels: &'a [u8],
}

/// Parses a bitmap opcode's data and returns it with its total length.
pub(super) fn parse_bits_record(data: &[u8], opcode: u16) -> Option<(BitsRecord<'_>, usize)> {
    let direct = matches!(opcode, 0x009A | 0x009B);
    let has_region = matches!(opcode, 0x0091 | 0x0099 | 0x009B);
    let always_raw = matches!(opcode, 0x0090 | 0x0091);
    let start = if direct { 4 } else { 0 };
    let (pixmap, mut pos) = parse_pixmap(data, start, direct)?;
    let src_rect = PictRect::read(data, pos)?;
    let dst_rect = PictRect::read(data, pos.checked_add(8)?)?;
    let mode = read_u16(data, pos.checked_add(16)?)?;
    pos = pos.checked_add(18)?;
    let mut mask_region = None;
    if has_region {
        let size = usize::from(read_u16(data, pos)?);
        let end = pos.checked_add(size.max(2))?;
        mask_region = Some(data.get(pos..end)?);
        pos = end;
    }
    let length = pixel_data_len(data, pos, &pixmap, always_raw)?;
    let end = pos.checked_add(length)?;
    let record = BitsRecord {
        pixmap,
        src_rect,
        dst_rect,
        mode,
        mask_region,
        always_raw,
        pixels: data.get(pos..end)?,
    };
    Some((record, end))
}

/// A `PixPat` from `BkPixPat`, `PnPixPat` or `FillPixPat`.
#[derive(Debug, Clone)]
pub(super) struct PixPat<'a> {
    /// The monochrome pattern used by devices without colour.
    pub(super) pat1: [u8; 8],
    pub(super) kind: PixPatKind<'a>,
}

#[derive(Debug, Clone)]
pub(super) enum PixPatKind<'a> {
    /// `patType` 0: an old-style pattern drawn in the port colours.
    Bits,
    /// `patType` 1: a full-colour pixel pattern.
    Color { pixmap: PixMap, pixels: &'a [u8] },
    /// `patType` 2: a dither pattern approximating an RGB colour.
    Dither([u8; 3]),
}

/// Parses a `PixPat` record and returns it with its length.
pub(super) fn parse_pixpat(data: &[u8]) -> Option<(PixPat<'_>, usize)> {
    let pat_type = read_u16(data, 0)?;
    let pat1: [u8; 8] = data.get(2..10)?.try_into().ok()?;
    match pat_type {
        0 => Some((
            PixPat {
                pat1,
                kind: PixPatKind::Bits,
            },
            10,
        )),
        1 => {
            let (pixmap, pos) = parse_pixmap(data, 10, false)?;
            if pixmap.is_bitmap {
                return None;
            }
            let length = pixel_data_len(data, pos, &pixmap, false)?;
            let end = pos.checked_add(length)?;
            let pixels = data.get(pos..end)?;
            Some((
                PixPat {
                    pat1,
                    kind: PixPatKind::Color { pixmap, pixels },
                },
                end,
            ))
        },
        2 => {
            let rgb = data.get(10..16)?;
            Some((
                PixPat {
                    pat1,
                    kind: PixPatKind::Dither([rgb[0], rgb[2], rgb[4]]),
                },
                16,
            ))
        },
        _ => None,
    }
}

/// Decoded pixels in RGBA order, one row after another.
#[derive(Debug, Clone)]
pub(super) struct DecodedPixels {
    pub(super) width: u32,
    pub(super) height: u32,
    pub(super) rgba: Vec<u8>,
    /// For one-bit sources, whether each pixel was a set (foreground) bit.
    pub(super) mono: Option<Vec<bool>>,
    /// Some packed rows decoded to fewer bytes than the row holds.
    pub(super) short_rows: bool,
    /// Some packed rows held more data than the row holds.
    pub(super) overlong_rows: bool,
}

/// Why pixel data could not be decoded.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum DecodeError {
    /// The pixel depth or component layout is not a QuickDraw format.
    Unsupported(String),
    /// The pixel map exceeds the caller's limits.
    TooLarge(String),
}

impl BitsRecord<'_> {
    /// Decodes the whole pixel map.
    pub(super) fn decode(&self, limits: PixelLimits) -> Result<DecodedPixels, DecodeError> {
        decode_pixels(&self.pixmap, self.pixels, self.always_raw, limits)
    }
}

/// Decodes the pixels of a colour `PixPat`.
pub(super) fn decode_pixpat(
    pixmap: &PixMap,
    pixels: &[u8],
    limits: PixelLimits,
) -> Result<DecodedPixels, DecodeError> {
    decode_pixels(pixmap, pixels, false, limits)
}

fn decode_pixels(
    pixmap: &PixMap,
    pixels: &[u8],
    always_raw: bool,
    limits: PixelLimits,
) -> Result<DecodedPixels, DecodeError> {
    let width = pixmap.columns();
    let height = pixmap.rows();
    let width32 = u32::try_from(width).map_err(|_| DecodeError::TooLarge("width".into()))?;
    let height32 = u32::try_from(height).map_err(|_| DecodeError::TooLarge("height".into()))?;
    let pixel_count = u64::from(width32) * u64::from(height32);
    if width32 > limits.max_width || height32 > limits.max_height || pixel_count > limits.max_pixels
    {
        return Err(DecodeError::TooLarge(format!(
            "pixel map is {width32}x{height32}"
        )));
    }
    match (pixmap.pixel_size, pixmap.cmp_count) {
        (1 | 2 | 4 | 8, _) | (16, _) => {},
        (32, 3 | 4) => {},
        (32, count) if pixmap.pack_type == 1 || pixmap.pack_type == 2 || pixmap.row_bytes < 8 => {
            if count == 0 {
                return Err(DecodeError::Unsupported(
                    "32-bit pixels without components".into(),
                ));
            }
        },
        (size, count) => {
            return Err(DecodeError::Unsupported(format!(
                "{size}-bit pixels with {count} components"
            )));
        },
    }
    let byte_len = usize::try_from(pixel_count)
        .ok()
        .and_then(|count| count.checked_mul(4))
        .ok_or_else(|| DecodeError::TooLarge("pixel buffer".into()))?;
    let mut rgba = Vec::new();
    rgba.try_reserve_exact(byte_len)
        .map_err(|_| DecodeError::TooLarge("pixel buffer allocation".into()))?;
    rgba.resize(byte_len, 0);
    let mut mono = (pixmap.pixel_size == 1).then(|| vec![false; width * height]);

    let encoding = pixmap.encoding(always_raw);
    let row_len = match encoding {
        RowEncoding::Raw { stored } => stored,
        RowEncoding::Packed { unpacked, .. } => unpacked,
    };
    let mut row = vec![0u8; row_len];
    let wide = pixmap.row_bytes > 250;
    let mut pos = 0usize;
    let mut short_rows = false;
    let mut overlong_rows = false;
    for y in 0..height {
        row.fill(0);
        match encoding {
            RowEncoding::Raw { stored } => {
                let source = pixels.get(pos..pos.saturating_add(stored)).unwrap_or(&[]);
                row[..source.len()].copy_from_slice(source);
                pos = pos.saturating_add(stored);
            },
            RowEncoding::Packed { item, .. } => {
                let Some((count, field)) = packed_count(pixels, pos, wide) else {
                    short_rows = true;
                    break;
                };
                let start = pos + field;
                let end = start.saturating_add(count).min(pixels.len());
                let packed = &pixels[start..end];
                let Unpacked { filled, overrun } = match item {
                    PackItem::Byte => unpack_bits(packed, &mut row),
                    PackItem::Word => unpack_words(packed, &mut row),
                };
                short_rows |= filled < row.len();
                overlong_rows |= overrun;
                pos = start.saturating_add(count);
            },
        }
        let out = &mut rgba[y * width * 4..(y + 1) * width * 4];
        decode_row(pixmap, &row, width, out);
        if let Some(mono) = mono.as_mut() {
            for (x, bit) in mono[y * width..(y + 1) * width].iter_mut().enumerate() {
                *bit = row
                    .get(x / 8)
                    .is_some_and(|byte| byte & (0x80 >> (x % 8)) != 0);
            }
        }
    }
    Ok(DecodedPixels {
        width: width32,
        height: height32,
        rgba,
        mono,
        short_rows,
        overlong_rows,
    })
}

fn decode_row(pixmap: &PixMap, row: &[u8], width: usize, out: &mut [u8]) {
    match pixmap.pixel_size {
        1 | 2 | 4 | 8 => {
            let bits = usize::from(pixmap.pixel_size);
            let mask = (1u16 << bits) - 1;
            for x in 0..width {
                let bit = x * bits;
                let byte = row.get(bit / 8).copied().unwrap_or(0);
                let shift = 8 - bits - bit % 8;
                let index = usize::from((u16::from(byte) >> shift) & mask);
                let color = pixmap.palette.get(index).copied().unwrap_or([0, 0, 0]);
                out[x * 4..x * 4 + 4].copy_from_slice(&[color[0], color[1], color[2], 255]);
            }
        },
        16 => {
            for x in 0..width {
                let value = read_u16(row, x * 2).unwrap_or(0);
                let expand = |component: u16| {
                    let component = (component & 0x1F) as u8;
                    (component << 3) | (component >> 2)
                };
                out[x * 4..x * 4 + 4].copy_from_slice(&[
                    expand(value >> 10),
                    expand(value >> 5),
                    expand(value),
                    255,
                ]);
            }
        },
        _ => {
            let planar = matches!(pixmap.encoding(false), RowEncoding::Packed { .. });
            let raw_rgb = pixmap.pack_type == 2 && pixmap.row_bytes >= 8;
            for x in 0..width {
                let color = if raw_rgb {
                    let at = x * 3;
                    [byte(row, at), byte(row, at + 1), byte(row, at + 2)]
                } else if planar {
                    // Planes are stored in alpha (when present), red, green,
                    // blue order; QuickDraw ignores the alpha plane.
                    let skip = if pixmap.cmp_count >= 4 { width } else { 0 };
                    [
                        byte(row, skip + x),
                        byte(row, skip + width + x),
                        byte(row, skip + 2 * width + x),
                    ]
                } else {
                    let at = x * 4;
                    [byte(row, at + 1), byte(row, at + 2), byte(row, at + 3)]
                };
                out[x * 4..x * 4 + 4].copy_from_slice(&[color[0], color[1], color[2], 255]);
            }
        },
    }
}

fn byte(row: &[u8], index: usize) -> u8 {
    row.get(index).copied().unwrap_or(0)
}

/// Returns the default colour table of an indexed depth, used by images whose
/// description refers to the system colour table (*Inside Macintosh:
/// Imaging With QuickDraw*, "The Color Manager": the standard 2-, 4- and
/// 8-bit system palettes).
pub(super) fn default_palette(pixel_size: u16) -> Vec<[u8; 3]> {
    match pixel_size {
        1 => vec![[255, 255, 255], [0, 0, 0]],
        2 => vec![
            [255, 255, 255],
            [0xAC, 0xAC, 0xAC],
            [0x55, 0x55, 0x55],
            [0, 0, 0],
        ],
        4 => vec![
            [0xFF, 0xFF, 0xFF],
            [0xFC, 0xF3, 0x05],
            [0xFF, 0x64, 0x02],
            [0xDD, 0x08, 0x06],
            [0xF2, 0x08, 0x84],
            [0x46, 0x00, 0xA5],
            [0x00, 0x00, 0xD4],
            [0x02, 0xAB, 0xEA],
            [0x1F, 0xB7, 0x14],
            [0x00, 0x64, 0x11],
            [0x56, 0x2C, 0x05],
            [0x90, 0x71, 0x3A],
            [0xC0, 0xC0, 0xC0],
            [0x80, 0x80, 0x80],
            [0x40, 0x40, 0x40],
            [0x00, 0x00, 0x00],
        ],
        _ => {
            const LEVELS: [u8; 6] = [0xFF, 0xCC, 0x99, 0x66, 0x33, 0x00];
            const RAMP: [u8; 10] = [0xEE, 0xDD, 0xBB, 0xAA, 0x88, 0x77, 0x55, 0x44, 0x22, 0x11];
            let mut palette = Vec::with_capacity(256);
            for red in LEVELS {
                for green in LEVELS {
                    for blue in LEVELS {
                        palette.push([red, green, blue]);
                    }
                }
            }
            // The cube's final entry is black, which the table stores last.
            palette.pop();
            palette.extend(RAMP.iter().map(|&level| [level, 0, 0]));
            palette.extend(RAMP.iter().map(|&level| [0, level, 0]));
            palette.extend(RAMP.iter().map(|&level| [0, 0, level]));
            palette.extend(RAMP.iter().map(|&level| [level, level, level]));
            palette.push([0, 0, 0]);
            palette
        },
    }
}

#[cfg(test)]
pub(super) mod tests {
    use super::*;

    const LIMITS: PixelLimits = PixelLimits {
        max_width: 64,
        max_height: 64,
        max_pixels: 4096,
    };

    pub(in crate::pict) fn push_rect(data: &mut Vec<u8>, rect: [i16; 4]) {
        for value in rect {
            data.extend_from_slice(&value.to_be_bytes());
        }
    }

    /// Builds a `PixMap` header (without `baseAddr`).
    pub(in crate::pict) fn pixmap_header(
        row_bytes: u16,
        bounds: [i16; 4],
        pack_type: u16,
        pixel_size: u16,
        cmp_count: u16,
    ) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&(row_bytes | 0x8000).to_be_bytes());
        push_rect(&mut data, bounds);
        data.extend_from_slice(&0u16.to_be_bytes()); // pmVersion
        data.extend_from_slice(&pack_type.to_be_bytes());
        data.extend_from_slice(&0u32.to_be_bytes()); // packSize
        data.extend_from_slice(&0x0048_0000u32.to_be_bytes()); // hRes
        data.extend_from_slice(&0x0048_0000u32.to_be_bytes()); // vRes
        let pixel_type: u16 = if pixel_size > 8 { 16 } else { 0 };
        data.extend_from_slice(&pixel_type.to_be_bytes());
        data.extend_from_slice(&pixel_size.to_be_bytes());
        data.extend_from_slice(&cmp_count.to_be_bytes());
        let cmp_size: u16 = if pixel_size == 32 {
            8
        } else {
            pixel_size.min(8)
        };
        data.extend_from_slice(&cmp_size.to_be_bytes());
        data.extend_from_slice(&[0; 12]); // planeBytes, pmTable, pmReserved
        assert_eq!(data.len(), PIXMAP_BYTES);
        data
    }

    pub(in crate::pict) fn color_table(colors: &[[u8; 3]]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0u32.to_be_bytes());
        data.extend_from_slice(&0u16.to_be_bytes());
        data.extend_from_slice(&(colors.len() as u16 - 1).to_be_bytes());
        for (index, color) in colors.iter().enumerate() {
            data.extend_from_slice(&(index as u16).to_be_bytes());
            for &component in color {
                data.extend_from_slice(&[component, component]);
            }
        }
        data
    }

    fn mono_bits_rect(rows: &[[u8; 2]]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&2u16.to_be_bytes()); // rowBytes, BitMap
        push_rect(&mut data, [0, 0, rows.len() as i16, 10]);
        push_rect(&mut data, [0, 0, rows.len() as i16, 10]);
        push_rect(&mut data, [5, 5, 5 + rows.len() as i16, 15]);
        data.extend_from_slice(&0u16.to_be_bytes()); // srcCopy
        for row in rows {
            data.extend_from_slice(row);
        }
        data
    }

    #[test]
    fn unpacked_monochrome_bitmap_decodes_rows() {
        let data = mono_bits_rect(&[[0b1010_0000, 0b0100_0000], [0, 0]]);
        let (record, length) = parse_bits_record(&data, 0x0090).unwrap();
        assert_eq!(length, data.len());
        assert_eq!(record.dst_rect.left, 5);
        let decoded = record.decode(LIMITS).unwrap();
        assert_eq!((decoded.width, decoded.height), (10, 2));
        assert_eq!(&decoded.rgba[0..8], &[0, 0, 0, 255, 255, 255, 255, 255]);
        let mono = decoded.mono.unwrap();
        assert!(mono[0] && !mono[1] && mono[2] && mono[9]);
    }

    #[test]
    fn short_bitmap_rows_are_never_packed() {
        // PackBitsRect stores rows narrower than eight bytes unpacked.
        let data = mono_bits_rect(&[[0xFF, 0xC0]]);
        let (record, length) = parse_bits_record(&data, 0x0098).unwrap();
        assert_eq!(length, data.len());
        assert!(
            record
                .decode(LIMITS)
                .unwrap()
                .mono
                .unwrap()
                .iter()
                .all(|&bit| bit)
        );
    }

    #[test]
    fn packed_indexed_rows_use_byte_counts_and_color_table() {
        let mut data = pixmap_header(8, [0, 0, 2, 8], 0, 8, 1);
        data.extend_from_slice(&color_table(&[[0, 0, 0], [0xFF, 0, 0], [0, 0xFF, 0]]));
        push_rect(&mut data, [0, 0, 2, 8]);
        push_rect(&mut data, [0, 0, 2, 8]);
        data.extend_from_slice(&0u16.to_be_bytes());
        data.extend_from_slice(&[2, 0xF9, 0x01]); // eight pixels of index 1
        data.extend_from_slice(&[4, 0x01, 0x02, 0x00, 0xFB]); // truncated run
        let (record, length) = parse_bits_record(&data, 0x0098).unwrap();
        assert_eq!(length, data.len());
        let decoded = record.decode(LIMITS).unwrap();
        assert_eq!(&decoded.rgba[0..4], &[0xFF, 0, 0, 255]);
        assert_eq!(&decoded.rgba[32..40], &[0, 0xFF, 0, 255, 0, 0, 0, 255]);
        assert!(decoded.short_rows);
    }

    #[test]
    fn wide_rows_use_two_byte_counts() {
        let mut data = pixmap_header(252, [0, 0, 1, 252], 0, 8, 1);
        data.extend_from_slice(&color_table(&[[0, 0, 0], [9, 9, 9]]));
        push_rect(&mut data, [0, 0, 1, 252]);
        push_rect(&mut data, [0, 0, 1, 252]);
        data.extend_from_slice(&0u16.to_be_bytes());
        data.extend_from_slice(&4u16.to_be_bytes());
        data.extend_from_slice(&[0x82, 0x01, 0x84, 0x00]); // 127 ones, 125 zeros
        let limits = PixelLimits {
            max_width: 300,
            ..LIMITS
        };
        let (record, length) = parse_bits_record(&data, 0x0098).unwrap();
        assert_eq!(length, data.len());
        let decoded = record.decode(limits).unwrap();
        assert_eq!(
            &decoded.rgba[126 * 4..128 * 4],
            &[9, 9, 9, 255, 0, 0, 0, 255]
        );
        assert!(!decoded.short_rows);
    }

    #[test]
    fn direct_32_bit_planar_rows_skip_alpha_plane() {
        let mut data = vec![0, 0, 0, 0xFF]; // baseAddr
        data.extend_from_slice(&pixmap_header(8, [0, 0, 1, 2], 4, 32, 4));
        push_rect(&mut data, [0, 0, 1, 2]);
        push_rect(&mut data, [0, 0, 1, 2]);
        data.extend_from_slice(&0u16.to_be_bytes());
        // alpha, red, green, blue planes of two pixels each
        data.extend_from_slice(&[9, 7, 0, 0, 10, 20, 30, 40, 50, 60]);
        let (record, length) = parse_bits_record(&data, 0x009A).unwrap();
        assert_eq!(length, data.len());
        let decoded = record.decode(LIMITS).unwrap();
        assert_eq!(&decoded.rgba, &[10, 30, 50, 255, 20, 40, 60, 255]);
    }

    #[test]
    fn direct_16_bit_rows_pack_words() {
        let mut data = vec![0, 0, 0, 0xFF];
        data.extend_from_slice(&pixmap_header(8, [0, 0, 1, 4], 3, 16, 3));
        push_rect(&mut data, [0, 0, 1, 4]);
        push_rect(&mut data, [0, 0, 1, 4]);
        data.extend_from_slice(&0u16.to_be_bytes());
        data.extend_from_slice(&[3, 0xFD, 0x7C, 0x00]); // four pure-red pixels
        let (record, _) = parse_bits_record(&data, 0x009A).unwrap();
        let decoded = record.decode(LIMITS).unwrap();
        assert_eq!(&decoded.rgba[12..16], &[255, 0, 0, 255]);
    }

    #[test]
    fn oversized_pixel_maps_are_rejected_before_allocation() {
        let mut data = pixmap_header(8, [0, 0, 100, 100], 1, 8, 1);
        data.extend_from_slice(&color_table(&[[0, 0, 0]]));
        let (pixmap, _) = parse_pixmap(&data, 0, false).unwrap();
        assert!(matches!(
            decode_pixels(&pixmap, &[], false, LIMITS),
            Err(DecodeError::TooLarge(_))
        ));
    }

    #[test]
    fn pixpat_lengths_follow_pattern_type() {
        let mut dither = vec![0, 2];
        dither.extend_from_slice(&[0xAA; 8]);
        dither.extend_from_slice(&[0x12, 0x12, 0x34, 0x34, 0x56, 0x56]);
        let (pattern, length) = parse_pixpat(&dither).unwrap();
        assert_eq!(length, 16);
        assert!(matches!(
            pattern.kind,
            PixPatKind::Dither([0x12, 0x34, 0x56])
        ));

        let mut color = vec![0, 1];
        color.extend_from_slice(&[0x55; 8]);
        color.extend_from_slice(&pixmap_header(2, [0, 0, 8, 8], 0, 2, 1));
        color.extend_from_slice(&color_table(&[[0, 0, 0], [255, 255, 255]]));
        color.extend_from_slice(&[0x11; 16]);
        let (pattern, length) = parse_pixpat(&color).unwrap();
        assert_eq!(length, color.len());
        let PixPatKind::Color { pixmap, pixels } = pattern.kind else {
            panic!("colour pattern expected");
        };
        let decoded = decode_pixpat(&pixmap, pixels, LIMITS).unwrap();
        assert_eq!(&decoded.rgba[0..8], &[0, 0, 0, 255, 255, 255, 255, 255]);
    }

    #[test]
    fn default_eight_bit_palette_has_cube_ramps_and_black() {
        let palette = default_palette(8);
        assert_eq!(palette.len(), 256);
        assert_eq!(palette[0], [255, 255, 255]);
        assert_eq!(palette[214], [0, 0, 0x33]);
        assert_eq!(palette[215], [0xEE, 0, 0]);
        assert_eq!(palette[255], [0, 0, 0]);
    }
}
