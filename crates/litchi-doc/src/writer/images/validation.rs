//! Bounded image sniffing and native OfficeArt dimension validation.

use super::super::core::WriteError;
use super::codec::{ASSUMED_DPI, TWIPS_PER_INCH};
use super::model::Picture;
use crate::parts::images::PictureType;
use litchi_odraw::image::{Kind as BlipKind, Rect};

/// Sniff the picture format, rejecting unsupported or unrecognised data.
pub(super) fn sniff_kind(data: &[u8]) -> Result<BlipKind, WriteError> {
    let kind = match PictureType::from_data(data) {
        PictureType::Emf => Some(BlipKind::Emf),
        PictureType::Wmf => Some(BlipKind::Wmf),
        PictureType::Jpeg => Some(BlipKind::Jpeg),
        PictureType::Png => Some(BlipKind::Png),
        PictureType::Bmp | PictureType::Dib => Some(BlipKind::Dib),
        PictureType::Tiff => Some(BlipKind::Tiff),
        PictureType::Unknown if dib_dimensions(data).is_some() => Some(BlipKind::Dib),
        PictureType::Unknown if pict_dimensions(data).is_some() => Some(BlipKind::Pict),
        _ => None,
    };
    kind.ok_or_else(|| {
        WriteError::InvalidData(
            "DOC picture data is not a supported native OfficeArt format".to_string(),
        )
    })
}

/// Validate that explicit format metadata agrees with the encoded bytes.
pub(super) fn validate_kind(kind: BlipKind, data: &[u8]) -> Result<(), WriteError> {
    let valid = match kind {
        BlipKind::Emf => PictureType::from_data(data) == PictureType::Emf,
        BlipKind::Wmf => valid_wmf_data(data),
        BlipKind::Pict => pict_dimensions(data).is_some(),
        BlipKind::Jpeg | BlipKind::CmykJpeg => PictureType::from_data(data) == PictureType::Jpeg,
        BlipKind::Png => PictureType::from_data(data) == PictureType::Png,
        BlipKind::Dib => dib_dimensions(data).is_some(),
        BlipKind::Tiff => PictureType::from_data(data) == PictureType::Tiff,
        BlipKind::Error | BlipKind::Unknown | BlipKind::Other(_) => false,
    };
    if !valid {
        return Err(WriteError::InvalidData(format!(
            "DOC picture bytes do not match explicit {kind:?} kind"
        )));
    }
    if data.len() > i32::MAX as usize - 1024 {
        return Err(WriteError::InvalidData(
            "DOC picture exceeds the signed 32-bit PICF block limit".to_string(),
        ));
    }
    Ok(())
}

fn valid_wmf_data(data: &[u8]) -> bool {
    (data.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A]) && data.len() >= 22)
        || (data.starts_with(&[0x01, 0x00, 0x09, 0x00]) && data.len() >= 18)
}

/// Convert a pixel count to twips at the assumed screen resolution.
fn pixels_to_twips(pixels: u32) -> u32 {
    pixels.saturating_mul(TWIPS_PER_INCH) / ASSUMED_DPI
}

/// Extract the pixel dimensions of an image without decoding it.
pub(super) fn intrinsic_dimensions_twips(kind: BlipKind, data: &[u8]) -> Option<(u32, u32)> {
    match kind {
        BlipKind::Png => pixels_as_twips(png_pixel_dimensions(data)?),
        BlipKind::Jpeg | BlipKind::CmykJpeg => pixels_as_twips(jpeg_pixel_dimensions(data)?),
        BlipKind::Dib => pixels_as_twips(dib_dimensions(data)?),
        BlipKind::Tiff => pixels_as_twips(tiff_dimensions(data)?),
        BlipKind::Emf => emf_dimensions_twips(data),
        BlipKind::Wmf => wmf_dimensions_twips(data),
        BlipKind::Pict => pixels_as_twips(pict_dimensions(data)?),
        BlipKind::Error | BlipKind::Unknown | BlipKind::Other(_) => None,
    }
}

fn pixels_as_twips((width, height): (u32, u32)) -> Option<(u32, u32)> {
    Some((pixels_to_twips(width), pixels_to_twips(height)))
}

/// Length of the PNG file signature in bytes.
pub(super) const PNG_SIGNATURE_LEN: usize = 8;
/// Length of the fixed part of a PNG chunk header (length + type).
pub(super) const PNG_CHUNK_HEADER_LEN: usize = 8;
/// Declared length of the IHDR chunk payload.
pub(super) const PNG_IHDR_LEN: u32 = 13;

/// Read width and height from the PNG IHDR chunk.
pub(super) fn png_pixel_dimensions(data: &[u8]) -> Option<(u32, u32)> {
    const IHDR_TYPE: &[u8; 4] = b"IHDR";

    let header = data.get(PNG_SIGNATURE_LEN..PNG_SIGNATURE_LEN + PNG_CHUNK_HEADER_LEN)?;
    let declared_len = u32::from_be_bytes(header[..4].try_into().ok()?);
    if declared_len != PNG_IHDR_LEN || &header[4..] != IHDR_TYPE {
        return None;
    }
    let dimensions = data.get(
        PNG_SIGNATURE_LEN + PNG_CHUNK_HEADER_LEN..PNG_SIGNATURE_LEN + PNG_CHUNK_HEADER_LEN + 8,
    )?;
    let width = u32::from_be_bytes(dimensions[..4].try_into().ok()?);
    let height = u32::from_be_bytes(dimensions[4..].try_into().ok()?);
    (width > 0 && height > 0).then_some((width, height))
}

/// JPEG marker segment prefix byte.
pub(super) const JPEG_MARKER_PREFIX: u8 = 0xFF;
/// JPEG start-of-image marker.
pub(super) const JPEG_MARKER_SOI: u8 = 0xD8;
/// JPEG start-of-scan marker (ends the header search).
pub(super) const JPEG_MARKER_SOS: u8 = 0xDA;

/// Read width and height from the first JPEG SOF (start of frame) segment.
pub(super) fn jpeg_pixel_dimensions(data: &[u8]) -> Option<(u32, u32)> {
    if data.len() < 4 || data[0] != JPEG_MARKER_PREFIX || data[1] != JPEG_MARKER_SOI {
        return None;
    }
    let mut offset = 2usize;
    while let Some(&marker) = data.get(offset + 1) {
        if data[offset] != JPEG_MARKER_PREFIX {
            return None;
        }
        match marker {
            // Start-of-frame markers, excluding DHT (0xC4), JPG (0xC8) and
            // DAC (0xCC) which share the marker range.
            0xC0..=0xCF if !matches!(marker, 0xC4 | 0xC8 | 0xCC) => {
                // Segment: length (2), precision (1), height (2), width (2).
                let segment = data.get(offset + 2..offset + 9)?;
                let height = u32::from(u16::from_be_bytes(segment[3..5].try_into().ok()?));
                let width = u32::from(u16::from_be_bytes(segment[5..7].try_into().ok()?));
                return (width > 0 && height > 0).then_some((width, height));
            },
            // Standalone markers without a length field (SOI, RSTn, TEM).
            JPEG_MARKER_SOI | 0x01 | 0xD0..=0xD7 => offset += 2,
            JPEG_MARKER_SOS => return None,
            _ => {
                let length =
                    u16::from_be_bytes(data.get(offset + 2..offset + 4)?.try_into().ok()?) as usize;
                offset = offset.checked_add(2 + length)?;
            },
        }
    }
    None
}

/// Read dimensions from a Windows DIB, accepting and skipping a BMP file
/// header when present.
fn dib_dimensions(data: &[u8]) -> Option<(u32, u32)> {
    let dib = if data.starts_with(b"BM") {
        data.get(14..)?
    } else {
        data
    };
    let header_size = u32::from_le_bytes(dib.get(..4)?.try_into().ok()?);
    let (width, height) = match header_size {
        // BITMAPCOREHEADER uses unsigned 16-bit dimensions.
        12 => (
            u32::from(u16::from_le_bytes(dib.get(4..6)?.try_into().ok()?)),
            u32::from(u16::from_le_bytes(dib.get(6..8)?.try_into().ok()?)),
        ),
        // BITMAPINFOHEADER and all later compatible headers use signed i32.
        40 | 52 | 56 | 108 | 124 => {
            let width = i32::from_le_bytes(dib.get(4..8)?.try_into().ok()?);
            let height = i32::from_le_bytes(dib.get(8..12)?.try_into().ok()?);
            (u32::try_from(width).ok()?, height.unsigned_abs())
        },
        _ => return None,
    };
    (width > 0 && height > 0).then_some((width, height))
}

/// Read the first TIFF IFD's ImageWidth and ImageLength values without
/// decoding image strips.
fn tiff_dimensions(data: &[u8]) -> Option<(u32, u32)> {
    let little_endian = match data.get(..4)? {
        b"II\x2A\0" => true,
        b"MM\0\x2A" => false,
        _ => return None,
    };
    let read_u16 = |bytes: &[u8]| -> Option<u16> {
        let bytes: [u8; 2] = bytes.get(..2)?.try_into().ok()?;
        Some(if little_endian {
            u16::from_le_bytes(bytes)
        } else {
            u16::from_be_bytes(bytes)
        })
    };
    let read_u32 = |bytes: &[u8]| -> Option<u32> {
        let bytes: [u8; 4] = bytes.get(..4)?.try_into().ok()?;
        Some(if little_endian {
            u32::from_le_bytes(bytes)
        } else {
            u32::from_be_bytes(bytes)
        })
    };
    let ifd_offset = usize::try_from(read_u32(data.get(4..)?)?).ok()?;
    let entry_count = usize::from(read_u16(data.get(ifd_offset..)?)?);
    // Every entry is fixed-width, so this also bounds the loop by file size.
    let entries = data.get(ifd_offset.checked_add(2)?..)?;
    let entry_bytes = entry_count.checked_mul(12)?;
    let entries = entries.get(..entry_bytes)?;
    let mut width = None;
    let mut height = None;
    for entry in entries.chunks_exact(12) {
        let tag = read_u16(entry)?;
        if !matches!(tag, 256 | 257) {
            continue;
        }
        let field_type = read_u16(&entry[2..])?;
        let count = read_u32(&entry[4..])?;
        if count != 1 {
            return None;
        }
        let value = match field_type {
            3 => u32::from(read_u16(&entry[8..])?),
            4 => read_u32(&entry[8..])?,
            _ => return None,
        };
        if tag == 256 {
            width = Some(value);
        } else {
            height = Some(value);
        }
    }
    let dimensions = (width?, height?);
    (dimensions.0 > 0 && dimensions.1 > 0).then_some(dimensions)
}

/// Read physical dimensions from the EMF header's frame rectangle, whose
/// coordinates are expressed in hundredths of a millimetre.
fn emf_dimensions_twips(data: &[u8]) -> Option<(u32, u32)> {
    if PictureType::from_data(data) != PictureType::Emf {
        return None;
    }
    let left = i32::from_le_bytes(data.get(24..28)?.try_into().ok()?);
    let top = i32::from_le_bytes(data.get(28..32)?.try_into().ok()?);
    let right = i32::from_le_bytes(data.get(32..36)?.try_into().ok()?);
    let bottom = i32::from_le_bytes(data.get(36..40)?.try_into().ok()?);
    let width = u32::try_from(right.checked_sub(left)?).ok()?;
    let height = u32::try_from(bottom.checked_sub(top)?).ok()?;
    hundredth_mm_as_twips(width, height)
}

fn hundredth_mm_as_twips(width: u32, height: u32) -> Option<(u32, u32)> {
    // One inch is 2540 hundredths of a millimetre.
    let width = u32::try_from(u64::from(width).checked_mul(TWIPS_PER_INCH.into())? / 2540).ok()?;
    let height =
        u32::try_from(u64::from(height).checked_mul(TWIPS_PER_INCH.into())? / 2540).ok()?;
    (width > 0 && height > 0).then_some((width, height))
}

/// Read a placeable WMF's bounds and units-per-inch. Non-placeable WMFs do
/// not carry physical extents and therefore require explicit dimensions.
fn wmf_dimensions_twips(data: &[u8]) -> Option<(u32, u32)> {
    if !data.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A]) {
        return None;
    }
    let left = i32::from(i16::from_le_bytes(data.get(6..8)?.try_into().ok()?));
    let top = i32::from(i16::from_le_bytes(data.get(8..10)?.try_into().ok()?));
    let right = i32::from(i16::from_le_bytes(data.get(10..12)?.try_into().ok()?));
    let bottom = i32::from(i16::from_le_bytes(data.get(12..14)?.try_into().ok()?));
    let units_per_inch = u32::from(u16::from_le_bytes(data.get(14..16)?.try_into().ok()?));
    if units_per_inch == 0 {
        return None;
    }
    let width = u32::try_from(right.checked_sub(left)?).ok()?;
    let height = u32::try_from(bottom.checked_sub(top)?).ok()?;
    let width_twips =
        u32::try_from(u64::from(width).checked_mul(TWIPS_PER_INCH.into())? / units_per_inch as u64)
            .ok()?;
    let height_twips = u32::try_from(
        u64::from(height).checked_mul(TWIPS_PER_INCH.into())? / units_per_inch as u64,
    )
    .ok()?;
    (width_twips > 0 && height_twips > 0).then_some((width_twips, height_twips))
}

/// Read a PICT frame rectangle. PICT resources can be bare or preceded by the
/// conventional 512-byte file header.
fn pict_dimensions(data: &[u8]) -> Option<(u32, u32)> {
    let (_, top, left, bottom, right) = pict_frame(data)?;
    let width = u32::try_from(right.checked_sub(left)?).ok()?;
    let height = u32::try_from(bottom.checked_sub(top)?).ok()?;
    (width > 0 && height > 0).then_some((width, height))
}

fn pict_frame(data: &[u8]) -> Option<(usize, i32, i32, i32, i32)> {
    for base in [0usize, 512] {
        let Some(header) = data.get(base..base.checked_add(14)?) else {
            continue;
        };
        let top = i32::from(i16::from_be_bytes(header[2..4].try_into().ok()?));
        let left = i32::from(i16::from_be_bytes(header[4..6].try_into().ok()?));
        let bottom = i32::from(i16::from_be_bytes(header[6..8].try_into().ok()?));
        let right = i32::from(i16::from_be_bytes(header[8..10].try_into().ok()?));
        let version = &header[10..14];
        if !version.starts_with(&[0x11, 0x01]) && version != [0x00, 0x11, 0x02, 0xFF] {
            continue;
        }
        return Some((base, top, left, bottom, right));
    }
    None
}

/// Clipping bounds stored in an OfficeArtMetafileHeader.
pub(super) fn metafile_bounds(picture: &Picture) -> Result<Rect, WriteError> {
    let data = &picture.data;
    let bounds = match picture.kind {
        BlipKind::Emf if data.len() >= 24 => Rect {
            left: read_i32(data, 8)?,
            top: read_i32(data, 12)?,
            right: read_i32(data, 16)?,
            bottom: read_i32(data, 20)?,
        },
        BlipKind::Wmf if data.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A]) => Rect {
            left: i32::from(read_i16(data, 6)?),
            top: i32::from(read_i16(data, 8)?),
            right: i32::from(read_i16(data, 10)?),
            bottom: i32::from(read_i16(data, 12)?),
        },
        BlipKind::Pict => {
            let (_, top, left, bottom, right) = pict_frame(data).ok_or_else(|| {
                WriteError::InvalidData("validated PICT picture has no frame".to_string())
            })?;
            Rect {
                left,
                top,
                right,
                bottom,
            }
        },
        _ => Rect {
            left: 0,
            top: 0,
            right: i32::try_from(picture.width_twips)
                .map_err(|_| WriteError::InvalidData("picture width exceeds i32".to_string()))?,
            bottom: i32::try_from(picture.height_twips)
                .map_err(|_| WriteError::InvalidData("picture height exceeds i32".to_string()))?,
        },
    };
    Ok(bounds)
}

fn read_i32(data: &[u8], offset: usize) -> Result<i32, WriteError> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| WriteError::InvalidData("metafile bounds offset overflows".to_string()))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| WriteError::InvalidData("metafile bounds are truncated".to_string()))?;
    Ok(i32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn read_i16(data: &[u8], offset: usize) -> Result<i16, WriteError> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| WriteError::InvalidData("metafile bounds offset overflows".to_string()))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| WriteError::InvalidData("metafile bounds are truncated".to_string()))?;
    Ok(i16::from_le_bytes([bytes[0], bytes[1]]))
}
