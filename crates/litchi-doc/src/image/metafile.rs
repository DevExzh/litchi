//! Inline pictures whose PICF carries a raw metafile instead of OfficeArt.
//!
//! Word writes inline pictures as [MS-DOC] PICFAndOfficeArtData, whose
//! `mfpf.mm` is MM_SHAPE or MM_SHAPEFILE. Older writers and pasted metafile
//! pictures use the Windows METAFILEPICT form instead: `mfpf.mm` is a Windows
//! mapping mode and a standard WMF ([MS-WMF] 2.3.2.2 META_HEADER) follows the
//! 68-byte PICF. Word for the Macintosh stores a placeholder WMF ("use Word 6.0c
//! or later to view Macintosh picture") followed by the QuickDraw PICT itself,
//! without the 512-byte PICT file header.
//!
//! Such a picture is returned as an owned OfficeArt metafile BLIP so that the
//! rest of the image pipeline treats it like any other stored metafile.

use std::borrow::Cow;

use litchi_odraw::image::builder::BlipBuilder;
use litchi_odraw::image::{Blip, File, Kind, Limits, Point, Rect};

use super::{ImageError, PictureFields};

/// Windows mapping modes `MM_TEXT` through `MM_ANISOTROPIC`.
const MAPPING_MODES: std::ops::RangeInclusive<i16> = 1..=8;
const WMF_HEADER_LEN: usize = 18;
const META_EOF: u16 = 0x0000;
const META_SETWINDOWORG: u16 = 0x020B;
const META_SETWINDOWEXT: u16 = 0x020C;
const EMU_PER_TWIP: i32 = 635;
const EMU_PER_HIMETRIC: i32 = 360;

/// The picture a METAFILEPICT-form PICF stores.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Payload<'data> {
    /// A standard WMF, ending with its EOF record when it has one.
    Wmf(&'data [u8]),
    /// A Macintosh PICT that follows the placeholder WMF.
    Pict(&'data [u8]),
}

/// The metafile in `picture`, the bytes that follow the 68-byte PICF, when
/// `fields` declares the METAFILEPICT form and a WMF header is present.
pub(super) fn payload<'data>(
    fields: &PictureFields,
    picture: &'data [u8],
) -> Option<Payload<'data>> {
    let mm = fields.mm;
    if !MAPPING_MODES.contains(&mm) || !is_wmf_header(picture) {
        return None;
    }
    let Some(end) = scan(picture).end else {
        return Some(Payload::Wmf(picture));
    };
    let (wmf, rest) = picture.split_at(end);
    Some(if is_pict_header(rest) {
        Payload::Pict(rest)
    } else {
        Payload::Wmf(wmf)
    })
}

/// An owned OfficeArt BLIP holding `payload`, sized from the PICF.
pub(super) fn file(
    fields: &PictureFields,
    payload: Payload<'_>,
) -> Result<File<'static>, ImageError> {
    let (kind, data, bounds) = match payload {
        Payload::Wmf(data) => (Kind::Wmf, data, scan(data).window().unwrap_or_default()),
        Payload::Pict(data) => (Kind::Pict, data, pict_frame(data)),
    };
    let limits = Limits::default();
    let mut record = Vec::new();
    BlipBuilder::meta(kind, Cow::Borrowed(data), bounds, extent(fields))
        .and_then(|blip| blip.write(&mut record, limits))
        .map_err(|_error| {
            ImageError::ExtractImageFailed(litchi_odraw::Error::MalformedImage {
                reason: "PICF metafile does not fit an OfficeArt BLIP",
            })
        })?;
    let blip = Blip::parse(&record, limits).map_err(ImageError::ExtractImageFailed)?;
    File::new(blip, None, 0, limits)
        .into_owned()
        .map_err(ImageError::ExtractImageFailed)
}

/// The physical picture size in EMUs: the PICF goal size in twips, or the
/// METAFILEPICT extent in HIMETRIC units when no goal size is recorded.
fn extent(fields: &PictureFields) -> Point {
    let axis = |goal: i16, ext: i16| {
        if goal > 0 {
            i32::from(goal) * EMU_PER_TWIP
        } else if ext > 0 {
            i32::from(ext) * EMU_PER_HIMETRIC
        } else {
            0
        }
    };
    Point {
        x: axis(fields.dxa_goal, fields.x_ext),
        y: axis(fields.dya_goal, fields.y_ext),
    }
}

fn u16_at(data: &[u8], offset: usize) -> Option<u16> {
    let bytes = data.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn i16_at(data: &[u8], offset: usize) -> Option<i16> {
    u16_at(data, offset).map(|value| i16::from_le_bytes(value.to_le_bytes()))
}

fn be_i16_at(data: &[u8], offset: usize) -> Option<i16> {
    let bytes = data.get(offset..offset.checked_add(2)?)?;
    Some(i16::from_be_bytes([bytes[0], bytes[1]]))
}

/// [MS-WMF] 2.3.2.2: memory or disk metafile, nine-word header, version 1 or 3.
fn is_wmf_header(data: &[u8]) -> bool {
    data.len() >= WMF_HEADER_LEN
        && matches!(u16_at(data, 0), Some(1 | 2))
        && u16_at(data, 2) == Some(9)
        && matches!(u16_at(data, 4), Some(0x0100 | 0x0300))
}

/// A PICT picture without its file header: picSize and picFrame, then the
/// version 1 (`0x1101`) or version 2 (`0x0011 0x02FF`) opcode.
fn is_pict_header(data: &[u8]) -> bool {
    matches!(data.get(10..12), Some([0x11, 0x01]))
        || matches!(data.get(10..14), Some([0x00, 0x11, 0x02, 0xFF]))
}

fn pict_frame(data: &[u8]) -> Rect {
    let coordinate = |offset| be_i16_at(data, offset).map_or(0, i32::from);
    Rect {
        left: coordinate(4),
        top: coordinate(2),
        right: coordinate(8),
        bottom: coordinate(6),
    }
}

/// Facts read from the WMF record sequence. The header's own size field is
/// not trusted: the Macintosh placeholder WMF declares a wrong one.
#[derive(Debug, Default)]
struct Scan {
    /// The byte length through the EOF record.
    end: Option<usize>,
    window_origin: Option<(i16, i16)>,
    window_extent: Option<(i16, i16)>,
}

impl Scan {
    /// The first window rectangle, which is the picture frame in logical units.
    fn window(&self) -> Option<Rect> {
        let (width, height) = self.window_extent?;
        let (x, y) = self.window_origin.unwrap_or_default();
        Some(Rect {
            left: i32::from(x),
            top: i32::from(y),
            right: i32::from(x) + i32::from(width),
            bottom: i32::from(y) + i32::from(height),
        })
    }
}

fn scan(wmf: &[u8]) -> Scan {
    let mut scan = Scan::default();
    let mut offset = WMF_HEADER_LEN;
    loop {
        let Some(size_words) = wmf
            .get(offset..offset.saturating_add(4))
            .and_then(|bytes| <[u8; 4]>::try_from(bytes).ok())
            .map(u32::from_le_bytes)
        else {
            return scan;
        };
        let Some(function) = u16_at(wmf, offset + 4) else {
            return scan;
        };
        let Some(end) = usize::try_from(size_words)
            .ok()
            .and_then(|words| words.checked_mul(2))
            .filter(|len| *len >= 6)
            .and_then(|len| offset.checked_add(len))
            .filter(|end| *end <= wmf.len())
        else {
            return scan;
        };
        // Both records store y before x.
        let point = || Some((i16_at(wmf, offset + 8)?, i16_at(wmf, offset + 6)?));
        match function {
            META_EOF => {
                scan.end = Some(end);
                return scan;
            },
            META_SETWINDOWORG if scan.window_origin.is_none() && end >= offset + 10 => {
                scan.window_origin = point();
            },
            META_SETWINDOWEXT if scan.window_extent.is_none() && end >= offset + 10 => {
                scan.window_extent = point();
            },
            _ => {},
        }
        offset = end;
    }
}

#[cfg(test)]
pub(super) mod tests {
    use super::*;

    fn record(function: u16, params: &[i16]) -> Vec<u8> {
        let words = u32::try_from(3 + params.len()).unwrap();
        let mut bytes = words.to_le_bytes().to_vec();
        bytes.extend_from_slice(&function.to_le_bytes());
        for param in params {
            bytes.extend_from_slice(&param.to_le_bytes());
        }
        bytes
    }

    /// A WMF that sets its window to (10, 20)-(110, 70) and draws a line.
    pub(in crate::image) fn wmf() -> Vec<u8> {
        let body = [
            record(META_SETWINDOWORG, &[20, 10]),
            record(META_SETWINDOWEXT, &[50, 100]),
            record(0x0213, &[60, 100]),
            record(META_EOF, &[]),
        ]
        .concat();
        let words = u32::try_from((WMF_HEADER_LEN + body.len()) / 2).unwrap();
        let mut bytes = Vec::new();
        for value in [1u16, 9, 0x0300] {
            bytes.extend_from_slice(&value.to_le_bytes());
        }
        bytes.extend_from_slice(&words.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&5u32.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&body);
        bytes
    }

    /// A version 2 PICT header framed (0, 0)-(283, 424), then its end opcode.
    pub(in crate::image) fn pict() -> Vec<u8> {
        let mut bytes = vec![0, 0, 0, 0, 0, 0, 0x01, 0x1B, 0x01, 0xA8];
        bytes.extend_from_slice(&[0x00, 0x11, 0x02, 0xFF, 0x00, 0xFF]);
        bytes
    }

    pub(in crate::image) fn fields(mm: i16) -> PictureFields {
        let mut bytes = vec![0; 68];
        bytes[6..8].copy_from_slice(&mm.to_le_bytes());
        bytes[8..10].copy_from_slice(&2540i16.to_le_bytes());
        bytes[10..12].copy_from_slice(&1270i16.to_le_bytes());
        bytes[28..30].copy_from_slice(&1440i16.to_le_bytes());
        bytes[30..32].copy_from_slice(&720i16.to_le_bytes());
        PictureFields::try_parse(&bytes, 0).unwrap()
    }

    #[test]
    fn only_mapping_modes_with_a_wmf_header_carry_a_raw_metafile() {
        let wmf = wmf();
        assert_eq!(payload(&fields(8), &wmf), Some(Payload::Wmf(&wmf)));
        assert_eq!(payload(&fields(1), &wmf), Some(Payload::Wmf(&wmf)));
        for mm in [0, 9, 0x64, 0x66, 94, 99] {
            assert_eq!(payload(&fields(mm), &wmf), None, "mm {mm}");
        }
        let mut not_wmf = wmf.clone();
        not_wmf[2] = 10;
        assert_eq!(payload(&fields(8), &not_wmf), None);
        assert_eq!(payload(&fields(8), &wmf[..17]), None);
    }

    #[test]
    fn wmf_ends_at_its_eof_record_and_a_following_pict_is_the_picture() {
        let wmf = wmf();
        let padded = [wmf.as_slice(), &[0; 6]].concat();
        assert_eq!(payload(&fields(8), &padded), Some(Payload::Wmf(&wmf)));

        let pict = pict();
        let mac = [wmf.as_slice(), &pict].concat();
        assert_eq!(payload(&fields(8), &mac), Some(Payload::Pict(&pict)));

        // Without an EOF record the whole payload stays the metafile.
        let truncated = &wmf[..wmf.len() - 6];
        assert_eq!(
            payload(&fields(8), truncated),
            Some(Payload::Wmf(truncated))
        );
    }

    #[test]
    fn wmf_blip_takes_its_bounds_from_the_window_and_its_size_from_the_goal() {
        let wmf = wmf();
        let file = file(&fields(8), Payload::Wmf(&wmf)).unwrap();
        assert_eq!(file.kind(), Kind::Wmf);
        assert_eq!(file.data().unwrap(), wmf.as_slice());
        let blip = file.blip().unwrap();
        let Blip::Wmf(meta) = blip else {
            panic!("expected a WMF BLIP");
        };
        let header = meta.header();
        assert_eq!(
            header.bounds,
            Rect {
                left: 10,
                top: 20,
                right: 110,
                bottom: 70
            }
        );
        assert_eq!(
            header.extent,
            Point {
                x: 914_400,
                y: 457_200
            }
        );
    }

    #[test]
    fn metafile_size_falls_back_to_the_himetric_extent() {
        let mut fields = fields(8);
        fields.dxa_goal = 0;
        fields.dya_goal = 0;
        assert_eq!(
            extent(&fields),
            Point {
                x: 914_400,
                y: 457_200
            }
        );
        fields.x_ext = 0;
        assert_eq!(extent(&fields).x, 0);
    }

    #[test]
    fn pict_blip_takes_its_bounds_from_the_picture_frame() {
        let pict = pict();
        let file = file(&fields(8), Payload::Pict(&pict)).unwrap();
        assert_eq!(file.kind(), Kind::Pict);
        assert_eq!(file.data().unwrap(), pict.as_slice());
        let Blip::Pict(meta) = file.blip().unwrap() else {
            panic!("expected a PICT BLIP");
        };
        assert_eq!(
            meta.header().bounds,
            Rect {
                left: 0,
                top: 0,
                right: 424,
                bottom: 283
            }
        );
    }

    #[test]
    fn malformed_records_end_the_scan_without_a_window() {
        let mut wmf = wmf();
        // A record shorter than its own six-byte header.
        wmf[WMF_HEADER_LEN..WMF_HEADER_LEN + 4].copy_from_slice(&2u32.to_le_bytes());
        let scan = scan(&wmf);
        assert!(scan.end.is_none());
        assert!(scan.window().is_none());
    }
}
