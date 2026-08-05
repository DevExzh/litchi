//! Images writer for DOC files
//!
//! Generates OfficeArtWordDrawing blocks ([MS-DOC] 2.9.172) for inline
//! pictures: a PICF picture descriptor ([MS-DOC] 2.9.161) followed by an
//! OfficeArtSpContainer and an OfficeArtFBSE record carrying the embedded
//! BLIP ([MS-ODRAW] 2.2). Blocks are appended to the Data stream; a
//! 0x0001 picture character in the text stream references each block through
//! sprmCPicLocation. Image bytes are stored without re-encoding; BMP file
//! headers are removed because OfficeArt stores their DIB payload.

use super::core::DocWriteError;
use crate::parts::images::PictureType;
use crate::parts::spa::{
    SPA_LEN, ShapeHorizontalOrigin, ShapeTextWrap, ShapeVerticalOrigin, ShapeWrapSide, Spa,
};
use litchi_odraw::image::write::{BlipBuilder, digest};
use litchi_odraw::image::{Kind as BlipKind, Point, Rect};

/// Size of the PICF picture descriptor header ([MS-DOC] 2.9.161).
const PICF_HEADER_LEN: usize = 0x44;
/// PICF `cbHeader` value, matching [`PICF_HEADER_LEN`].
const PICF_CB_HEADER: i16 = 0x44;
/// PICF `mfp.mm` value Word writes for OfficeArt shape pictures.
const PICF_MM_SHAPE: i16 = 0x64;
/// Unscaled picture factor in permille (1000 = 100%).
const SCALE_100_PERCENT: i16 = 1000;

// Escher record types ([MS-ODRAW] 2.3).
const RECORD_DGG_CONTAINER: u16 = 0xF000;
const RECORD_BSTORE_CONTAINER: u16 = 0xF001;
const RECORD_DG_CONTAINER: u16 = 0xF002;
const RECORD_SPGR_CONTAINER: u16 = 0xF003;
const RECORD_SP_CONTAINER: u16 = 0xF004;
const RECORD_DGG: u16 = 0xF006;
const RECORD_BSE: u16 = 0xF007;
const RECORD_DG: u16 = 0xF008;
const RECORD_SPGR: u16 = 0xF009;
const RECORD_SP: u16 = 0xF00A;
const RECORD_OPT: u16 = 0xF00B;
const RECORD_CLIENT_ANCHOR: u16 = 0xF010;
const RECORD_CLIENT_DATA: u16 = 0xF011;

// Escher record versions.
const VERSION_CONTAINER: u16 = 0xF;
const VERSION_DG: u16 = 0x0;
const VERSION_SPGR: u16 = 0x1;
const VERSION_SP: u16 = 0x2;
const VERSION_OPT: u16 = 0x3;
const VERSION_ATOM: u16 = 0x0;

/// MSOSHAPETYPE value for a picture frame shape.
const SHAPE_TYPE_PICTURE_FRAME: u16 = 0x4B;
/// OfficeArtFSP `fHaveAnchor` flag.
const SP_FLAG_HAVE_ANCHOR: u32 = 0x0200;
/// OfficeArtFSP `fHaveShapeType` flag.
const SP_FLAG_HAVE_SHAPE_TYPE: u32 = 0x0800;
/// OfficeArtFSP `fGroup` flag, set on the group shape of a drawing.
const SP_FLAG_GROUP: u32 = 0x0001;
/// OfficeArtFSP `fPatriarch` flag, set on the topmost group shape.
const SP_FLAG_PATRIARCH: u32 = 0x0004;
/// OfficeArt `pib` property (0x0104) with the fBid bit set, meaning the
/// value is a 1-based index of the BSE within the same drawing block.
const OPT_PIB_BLIP_INDEX: u16 = 0x4104;
/// The single BSE stored inside each OfficeArtWordDrawing block.
const OPT_PIB_FIRST_BSE: u32 = 1;
/// Payload length of the empty ClientAnchor record used for inline pictures.
const CLIENT_ANCHOR_PAYLOAD_LEN: u32 = 4;
/// Payload length of the OfficeArtFSP record (spid + flags).
const SP_PAYLOAD_LEN: u32 = 8;
/// Payload length of the OfficeArtOPT record holding one simple property.
const OPT_PAYLOAD_LEN: u32 = 6;
/// Total length of the OfficeArtSpContainer record including its header.
const SHAPE_CONTAINER_LEN: u32 = (RECORD_HEADER_LEN as u32 + SP_PAYLOAD_LEN)
    + (RECORD_HEADER_LEN as u32 + OPT_PAYLOAD_LEN)
    + (RECORD_HEADER_LEN as u32 + CLIENT_ANCHOR_PAYLOAD_LEN)
    + RECORD_HEADER_LEN as u32;

/// Length of an Escher record header in bytes.
const RECORD_HEADER_LEN: usize = 8;
/// Fixed OfficeArtFBSE payload length when no name is stored.
const BSE_HEADER_LEN: usize = 36;
/// OfficeArtFBSE `foDelay` value when the BLIP is embedded in the BSE record
/// (no delay-stream position).
const BSE_NO_DELAY_STREAM: u32 = u32::MAX;

/// Shape id assigned to the first inline picture. Word numbers inline shapes
/// starting at 1025 in documents without an OfficeArtDgContainer.
pub(crate) const FIRST_PICTURE_SHAPE_ID: u32 = 1025;
/// Shape id of the group shape that parents the shapes of a drawing.
const GROUP_SHAPE_ID: u32 = 1024;
/// Number of shape ids in one OfficeArt drawing cluster.
const SHAPE_IDS_PER_CLUSTER: u32 = 1024;

/// Twips per inch; the writer assumes 96 DPI when converting pixel sizes.
const TWIPS_PER_INCH: u32 = 1440;
/// Assumed screen resolution for pixel-sized images.
const ASSUMED_DPI: u32 = 96;
/// Largest dimension expressible in the signed 16-bit PICF goal fields.
const MAX_PICF_DIMENSION_TWIPS: u32 = i16::MAX as u32;

/// An inline picture to be embedded in a DOC document.
///
/// Encoded bytes are stored as-is except that a 14-byte BMP file header is
/// removed to obtain the DIB payload required by OfficeArt.
#[derive(Debug, Clone)]
pub struct DocPicture {
    /// Raw BLIP file data.
    data: Vec<u8>,
    /// Detected native OfficeArt kind.
    kind: BlipKind,
    /// Display width in twips.
    width_twips: u32,
    /// Display height in twips.
    height_twips: u32,
}

impl DocPicture {
    /// Create a picture from raw image bytes.
    ///
    /// The format is sniffed from the byte signature and the display
    /// dimensions are derived from bitmap pixels or metafile bounds. Returns
    /// an error for unsupported formats or when the dimensions cannot be
    /// determined; use [`Self::from_parts`] to supply dimensions explicitly.
    pub fn new(data: Vec<u8>) -> Result<Self, DocWriteError> {
        let kind = sniff_kind(&data)?;
        let dimensions = intrinsic_dimensions_twips(kind, &data).ok_or_else(|| {
            DocWriteError::InvalidData(
                "DOC picture dimensions are unreadable; use DocPicture::from_parts".to_string(),
            )
        })?;
        Self::from_parts_as(data, kind, dimensions.0, dimensions.1)
    }

    /// Create a picture from raw image bytes and explicit display dimensions
    /// in twips (1/1440 inch).
    pub fn from_parts(
        data: Vec<u8>,
        width_twips: u32,
        height_twips: u32,
    ) -> Result<Self, DocWriteError> {
        let kind = sniff_kind(&data)?;
        Self::from_parts_as(data, kind, width_twips, height_twips)
    }

    /// Create a picture with an explicit native OfficeArt format and display
    /// dimensions. This is useful for headerless DIB and PICT data whose
    /// format cannot always be inferred unambiguously.
    pub fn from_parts_as(
        mut data: Vec<u8>,
        kind: BlipKind,
        width_twips: u32,
        height_twips: u32,
    ) -> Result<Self, DocWriteError> {
        validate_kind(kind, &data)?;
        if kind == BlipKind::Dib && data.starts_with(b"BM") {
            data.drain(..14);
        }
        let picture = Self {
            data,
            kind,
            width_twips: 0,
            height_twips: 0,
        };
        picture.with_dimensions_twips(width_twips, height_twips)
    }

    /// Override the display dimensions in twips (1/1440 inch).
    ///
    /// Dimensions must fit the signed 16-bit PICF goal fields, i.e. they must
    /// be in `1..=32767` (about 22.7 inches at 100% scale).
    pub fn with_dimensions_twips(
        mut self,
        width_twips: u32,
        height_twips: u32,
    ) -> Result<Self, DocWriteError> {
        for dimension in [width_twips, height_twips] {
            if !(1..=MAX_PICF_DIMENSION_TWIPS).contains(&dimension) {
                return Err(DocWriteError::InvalidData(format!(
                    "DOC picture dimension {dimension} twips is outside 1..={MAX_PICF_DIMENSION_TWIPS}"
                )));
            }
        }
        self.width_twips = width_twips;
        self.height_twips = height_twips;
        Ok(self)
    }

    /// Raw image bytes.
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Detected native OfficeArt kind.
    pub const fn kind(&self) -> BlipKind {
        self.kind
    }

    /// Display width in twips.
    pub fn width_twips(&self) -> u32 {
        self.width_twips
    }

    /// Display height in twips.
    pub fn height_twips(&self) -> u32 {
        self.height_twips
    }
}

/// Sniff the picture format, rejecting unsupported or unrecognised data.
fn sniff_kind(data: &[u8]) -> Result<BlipKind, DocWriteError> {
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
        DocWriteError::InvalidData(
            "DOC picture data is not a supported native OfficeArt format".to_string(),
        )
    })
}

/// Validate that explicit format metadata agrees with the encoded bytes.
fn validate_kind(kind: BlipKind, data: &[u8]) -> Result<(), DocWriteError> {
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
        return Err(DocWriteError::InvalidData(format!(
            "DOC picture bytes do not match explicit {kind:?} kind"
        )));
    }
    if data.len() > i32::MAX as usize - 1024 {
        return Err(DocWriteError::InvalidData(
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
fn intrinsic_dimensions_twips(kind: BlipKind, data: &[u8]) -> Option<(u32, u32)> {
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
const PNG_SIGNATURE_LEN: usize = 8;
/// Length of the fixed part of a PNG chunk header (length + type).
const PNG_CHUNK_HEADER_LEN: usize = 8;
/// Declared length of the IHDR chunk payload.
const PNG_IHDR_LEN: u32 = 13;

/// Read width and height from the PNG IHDR chunk.
fn png_pixel_dimensions(data: &[u8]) -> Option<(u32, u32)> {
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
const JPEG_MARKER_PREFIX: u8 = 0xFF;
/// JPEG start-of-image marker.
const JPEG_MARKER_SOI: u8 = 0xD8;
/// JPEG start-of-scan marker (ends the header search).
const JPEG_MARKER_SOS: u8 = 0xDA;

/// Read width and height from the first JPEG SOF (start of frame) segment.
fn jpeg_pixel_dimensions(data: &[u8]) -> Option<(u32, u32)> {
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
fn metafile_bounds(picture: &DocPicture) -> Result<Rect, DocWriteError> {
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
                DocWriteError::InvalidData("validated PICT picture has no frame".to_string())
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
                .map_err(|_| DocWriteError::InvalidData("picture width exceeds i32".to_string()))?,
            bottom: i32::try_from(picture.height_twips).map_err(|_| {
                DocWriteError::InvalidData("picture height exceeds i32".to_string())
            })?,
        },
    };
    Ok(bounds)
}

fn read_i32(data: &[u8], offset: usize) -> Result<i32, DocWriteError> {
    let end = offset.checked_add(4).ok_or_else(|| {
        DocWriteError::InvalidData("metafile bounds offset overflows".to_string())
    })?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| DocWriteError::InvalidData("metafile bounds are truncated".to_string()))?;
    Ok(i32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn read_i16(data: &[u8], offset: usize) -> Result<i16, DocWriteError> {
    let end = offset.checked_add(2).ok_or_else(|| {
        DocWriteError::InvalidData("metafile bounds offset overflows".to_string())
    })?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| DocWriteError::InvalidData("metafile bounds are truncated".to_string()))?;
    Ok(i16::from_le_bytes([bytes[0], bytes[1]]))
}

fn picture_blip(picture: &DocPicture) -> Result<BlipBuilder<'_>, DocWriteError> {
    if picture.kind.is_meta() {
        let width = i32::try_from(picture.width_twips)
            .ok()
            .and_then(|value| value.checked_mul(635))
            .ok_or_else(|| {
                DocWriteError::InvalidData("metafile width extent exceeds i32".to_string())
            })?;
        let height = i32::try_from(picture.height_twips)
            .ok()
            .and_then(|value| value.checked_mul(635))
            .ok_or_else(|| {
                DocWriteError::InvalidData("metafile height extent exceeds i32".to_string())
            })?;
        Ok(BlipBuilder::meta(
            picture.kind,
            picture.data(),
            metafile_bounds(picture)?,
            Point {
                x: width,
                y: height,
            },
        )?)
    } else {
        Ok(BlipBuilder::bitmap(picture.kind, picture.data())?.tag(0xFF))
    }
}

/// Append an Escher record header ([MS-ODRAW] 2.2.1).
pub(crate) fn write_record_header(
    out: &mut Vec<u8>,
    version: u16,
    instance: u16,
    record_type: u16,
    payload_len: u32,
) {
    let ver_inst = (instance << 4) | version;
    out.extend_from_slice(&ver_inst.to_le_bytes());
    out.extend_from_slice(&record_type.to_le_bytes());
    out.extend_from_slice(&payload_len.to_le_bytes());
}

/// Append an OfficeArtOPT record holding simple (non-complex) properties as
/// (opid, value) pairs.
pub(crate) fn write_opt_record(out: &mut Vec<u8>, properties: &[(u16, u32)]) {
    write_record_header(
        out,
        VERSION_OPT,
        properties.len() as u16,
        RECORD_OPT,
        properties.len() as u32 * 6,
    );
    for &(opid, value) in properties {
        out.extend_from_slice(&opid.to_le_bytes());
        out.extend_from_slice(&value.to_le_bytes());
    }
}

/// Append the PICF picture descriptor ([MS-DOC] 2.9.161).
///
/// The 68-byte header is followed by the OfficeArt shape data; `lcb` covers
/// both, matching the layout Word writes for inline pictures.
fn write_picf(out: &mut Vec<u8>, picture: &DocPicture, lcb: u32) {
    out.extend_from_slice(&(lcb as i32).to_le_bytes());
    out.extend_from_slice(&PICF_CB_HEADER.to_le_bytes());
    out.extend_from_slice(&PICF_MM_SHAPE.to_le_bytes());
    out.extend_from_slice(&(picture.width_twips as i16).to_le_bytes()); // xExt
    out.extend_from_slice(&(picture.height_twips as i16).to_le_bytes()); // yExt
    out.extend_from_slice(&0i16.to_le_bytes()); // swHMF
    out.extend_from_slice(&0i32.to_le_bytes()); // grf: low byte 0x00 = Word 2000 picture block
    out.extend_from_slice(&0i32.to_le_bytes()); // padding
    out.extend_from_slice(&0i16.to_le_bytes()); // mmPM
    out.extend_from_slice(&0i32.to_le_bytes()); // padding2
    out.extend_from_slice(&(picture.width_twips as i16).to_le_bytes()); // dxaGoal
    out.extend_from_slice(&(picture.height_twips as i16).to_le_bytes()); // dyaGoal
    out.extend_from_slice(&SCALE_100_PERCENT.to_le_bytes()); // mx
    out.extend_from_slice(&SCALE_100_PERCENT.to_le_bytes()); // my
    out.extend_from_slice(&[0; 8]); // crop left/top/right/bottom
    out.push(0); // fReserved
    out.push(0); // bpp
    out.extend_from_slice(&[0; 16]); // brcTop/brcLeft/brcBottom/brcRight
    out.extend_from_slice(&0i16.to_le_bytes()); // dxaReserved3
    out.extend_from_slice(&0i16.to_le_bytes()); // dyaReserved3
    out.extend_from_slice(&0i16.to_le_bytes()); // cProps
}

/// Append the OfficeArtSpContainer describing the inline picture shape.
///
/// The shape is a picture frame whose single `pib` property references the
/// BSE that follows inside the same drawing block. Inline pictures carry an
/// empty ClientAnchor record, as Word writes for 0x0001-anchored pictures.
fn write_shape_container(out: &mut Vec<u8>, shape_id: u32) {
    let container_payload_len = SHAPE_CONTAINER_LEN - RECORD_HEADER_LEN as u32;

    write_record_header(
        out,
        VERSION_CONTAINER,
        0,
        RECORD_SP_CONTAINER,
        container_payload_len,
    );
    // OfficeArtFSP: picture frame shape with an anchor and explicit type.
    write_record_header(
        out,
        VERSION_SP,
        SHAPE_TYPE_PICTURE_FRAME,
        RECORD_SP,
        SP_PAYLOAD_LEN,
    );
    out.extend_from_slice(&shape_id.to_le_bytes());
    out.extend_from_slice(&(SP_FLAG_HAVE_ANCHOR | SP_FLAG_HAVE_SHAPE_TYPE).to_le_bytes());
    // OfficeArtOPT: pib referencing the adjacent BSE (1-based index).
    write_record_header(out, VERSION_OPT, 1, RECORD_OPT, OPT_PAYLOAD_LEN);
    out.extend_from_slice(&OPT_PIB_BLIP_INDEX.to_le_bytes());
    out.extend_from_slice(&OPT_PIB_FIRST_BSE.to_le_bytes());
    // OfficeArtClientAnchor: empty for inline pictures.
    write_record_header(
        out,
        VERSION_ATOM,
        0,
        RECORD_CLIENT_ANCHOR,
        CLIENT_ANCHOR_PAYLOAD_LEN,
    );
    out.extend_from_slice(&[0; CLIENT_ANCHOR_PAYLOAD_LEN as usize]);
}

/// Append an OfficeArtFBSE record with the embedded OfficeArtBlip record for
/// a picture. Used both for the Data-stream picture blocks and for the
/// BStoreContainer inside the drawing group of floating pictures.
fn write_bse_with_embedded_blip(
    out: &mut Vec<u8>,
    picture: &DocPicture,
) -> Result<(), DocWriteError> {
    let blip = picture_blip(picture)?;
    let blip_record_len = blip.wire_len()?;
    let bse_payload_len = u32::try_from(BSE_HEADER_LEN)
        .ok()
        .and_then(|header| header.checked_add(blip_record_len))
        .ok_or_else(|| DocWriteError::InvalidData("DOC FBSE length exceeds u32".to_string()))?;

    write_record_header(
        out,
        2,
        u16::from(picture.kind.raw()),
        RECORD_BSE,
        bse_payload_len,
    );
    out.push(picture.kind.raw()); // btWin32
    out.push(picture.kind.raw()); // btMacOS
    out.extend_from_slice(digest(picture.data()).as_bytes()); // rgbUid
    out.extend_from_slice(&0u16.to_le_bytes()); // tag
    out.extend_from_slice(&blip_record_len.to_le_bytes()); // size
    out.extend_from_slice(&1u32.to_le_bytes()); // cRef
    out.extend_from_slice(&BSE_NO_DELAY_STREAM.to_le_bytes()); // foDelay: BLIP is embedded
    out.push(0); // usage
    out.push(0); // cbName
    out.push(0); // unused2
    out.push(0); // unused3

    blip.write(out)?;
    Ok(())
}

/// Append an OfficeArtWordDrawing block (PICF + shape container + BSE with an
/// embedded BLIP) to the Data stream.
pub(crate) fn write_picture_block(
    picture: &DocPicture,
    shape_id: u32,
    out: &mut Vec<u8>,
) -> Result<(), DocWriteError> {
    let blip_record_len = picture_blip(picture)?.wire_len()?;
    let bse_payload_len = u32::try_from(BSE_HEADER_LEN)
        .ok()
        .and_then(|header| header.checked_add(blip_record_len))
        .ok_or_else(|| DocWriteError::InvalidData("DOC FBSE length exceeds u32".to_string()))?;

    let block_start = out.len();
    // lcb covers the PICF header plus everything that follows it.
    let lcb = u32::try_from(PICF_HEADER_LEN)
        .ok()
        .and_then(|value| value.checked_add(SHAPE_CONTAINER_LEN))
        .and_then(|value| value.checked_add(RECORD_HEADER_LEN as u32))
        .and_then(|value| value.checked_add(bse_payload_len))
        .ok_or_else(|| DocWriteError::InvalidData("DOC picture block exceeds u32".to_string()))?;
    write_picf(out, picture, lcb);
    write_shape_container(out, shape_id);
    write_bse_with_embedded_blip(out, picture)?;

    let actual = out.len().checked_sub(block_start).ok_or_else(|| {
        DocWriteError::InvalidData("DOC picture block length underflow".to_string())
    })?;
    if actual != usize::try_from(lcb).unwrap_or(usize::MAX) {
        return Err(DocWriteError::InvalidData(
            "DOC picture block length mismatch".to_string(),
        ));
    }
    Ok(())
}

// ============================================================================
// Floating pictures: PlcfSpa and OfficeArtContent (DggInfo)
// ============================================================================

/// Position and wrapping of a floating picture.
///
/// The position is the top-left corner of the picture in twips, relative to
/// the origins selected by [`ShapeHorizontalOrigin`] and
/// [`ShapeVerticalOrigin`]. The size comes from the [`DocPicture`] display
/// dimensions. Defaults match a typical Word floating picture: page-relative
/// offsets, square wrapping on both sides, in front of the text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FloatingPosition {
    /// Left offset in twips relative to the horizontal origin.
    left_twips: i32,
    /// Top offset in twips relative to the vertical origin.
    top_twips: i32,
    /// Horizontal position origin (Spa `bx`).
    horizontal_origin: ShapeHorizontalOrigin,
    /// Vertical position origin (Spa `by`).
    vertical_origin: ShapeVerticalOrigin,
    /// Text-wrapping style (Spa `wr`).
    wrap: ShapeTextWrap,
    /// Wrap side restriction (Spa `wrk`).
    wrap_side: ShapeWrapSide,
    /// Whether the picture appears behind the text (Spa `fBelowText`).
    behind_text: bool,
    /// Whether the anchor is locked to its paragraph (Spa `fAnchorLock`).
    anchor_locked: bool,
}

impl FloatingPosition {
    /// Create a position from offsets in twips, defaulting to page-relative
    /// origins and square wrapping in front of the text.
    pub fn new(left_twips: i32, top_twips: i32) -> Self {
        Self {
            left_twips,
            top_twips,
            horizontal_origin: ShapeHorizontalOrigin::Page,
            vertical_origin: ShapeVerticalOrigin::Page,
            wrap: ShapeTextWrap::Square,
            wrap_side: ShapeWrapSide::Both,
            behind_text: false,
            anchor_locked: false,
        }
    }

    /// Set the horizontal and vertical position origins.
    pub fn with_origins(
        mut self,
        horizontal: ShapeHorizontalOrigin,
        vertical: ShapeVerticalOrigin,
    ) -> Self {
        self.horizontal_origin = horizontal;
        self.vertical_origin = vertical;
        self
    }

    /// Set the text-wrapping style.
    pub fn with_text_wrap(mut self, wrap: ShapeTextWrap) -> Self {
        self.wrap = wrap;
        self
    }

    /// Set the wrap side restriction.
    pub fn with_wrap_side(mut self, wrap_side: ShapeWrapSide) -> Self {
        self.wrap_side = wrap_side;
        self
    }

    /// Place the picture behind (or in front of) the text.
    pub fn behind_text(mut self, behind_text: bool) -> Self {
        self.behind_text = behind_text;
        self
    }

    /// Lock the anchor to its paragraph.
    pub fn lock_anchor(mut self, anchor_locked: bool) -> Self {
        self.anchor_locked = anchor_locked;
        self
    }
}

/// The visual content of a floating shape in the drawing layer.
pub(crate) enum FloatingShapeContent<'a> {
    /// A picture frame whose BLIP is stored in the blip store.
    Picture(&'a DocPicture),
    /// A primitive preset-geometry shape (rectangle, ellipse, ...).
    Primitive(&'a super::shapes::DocDrawingShape),
}

/// Everything the table-stream builders need to know about one floating
/// picture or primitive shape anchored in the Main Document.
pub(crate) struct FloatingShapeInfo<'a> {
    /// Character position of the 0x0008 anchor character (Main Document CP).
    pub anchor_cp: u32,
    /// Shape id, shared with the picture's Data-stream block when present.
    pub shape_id: u32,
    /// What the shape renders.
    pub content: FloatingShapeContent<'a>,
    /// Display width in twips.
    pub width_twips: u32,
    /// Display height in twips.
    pub height_twips: u32,
    /// Position and wrapping.
    pub position: &'a FloatingPosition,
    /// Textbox story text when the shape is a text box.
    pub text: Option<&'a str>,
}

impl FloatingShapeInfo<'_> {
    /// The MSOSPT shape type for the OfficeArtFSP record instance.
    fn shape_type(&self) -> u16 {
        if self.text.is_some() {
            return super::shapes::MSOSPT_TEXT_BOX;
        }
        match &self.content {
            FloatingShapeContent::Picture(_) => SHAPE_TYPE_PICTURE_FRAME,
            FloatingShapeContent::Primitive(shape) => shape.kind().shape_type(),
        }
    }

    /// Build the Spa record for this shape.
    fn spa(&self) -> Spa {
        let left = self.position.left_twips;
        let top = self.position.top_twips;
        Spa {
            shape_id: self.shape_id,
            left,
            top,
            right: left + self.width_twips as i32,
            bottom: top + self.height_twips as i32,
            horizontal_origin: self.position.horizontal_origin,
            vertical_origin: self.position.vertical_origin,
            wrap: self.position.wrap,
            wrap_side: self.position.wrap_side,
            below_text: self.position.behind_text,
            anchor_locked: self.position.anchor_locked,
        }
    }
}

/// Build the PlcfSpa for the Main Document ([MS-DOC] 2.8.27).
///
/// `shapes` must be ordered by ascending anchor CP (guaranteed by the
/// writer's append-only story). `final_cp` is the document's ccpText; the
/// final CP entry is undefined per spec but must exceed all anchor CPs.
pub(crate) fn build_plcf_spa(shapes: &[FloatingShapeInfo<'_>], final_cp: u32) -> Vec<u8> {
    let mut out = Vec::with_capacity((shapes.len() + 1) * 4 + shapes.len() * SPA_LEN);
    for shape in shapes {
        out.extend_from_slice(&shape.anchor_cp.to_le_bytes());
    }
    out.extend_from_slice(&final_cp.to_le_bytes());
    for shape in shapes {
        out.extend_from_slice(&shape.spa().to_bytes());
    }
    out
}

/// OfficeArtWordDrawing `dgglbl` value for the Main Document drawing.
const DGGLBL_MAIN_DOCUMENT: u8 = 0x00;
/// OfficeArtWordDrawing `dgglbl` value for the Header Document drawing.
const DGGLBL_HEADER_DOCUMENT: u8 = 0x01;
/// Shape id assigned to the first header-story shape. Each drawing owns a
/// cluster of shape ids ([MS-ODRAW] OfficeArtIDCL): the Main Document
/// drawing uses the cluster starting at 1024, the Header Document drawing
/// the next one.
pub(crate) const HEADER_FIRST_SHAPE_ID: u32 = FIRST_PICTURE_SHAPE_ID + SHAPE_IDS_PER_CLUSTER;
/// OfficeArtFDG `csp` counting mode: shapes plus the group shape.
const DG_GROUP_SHAPE_COUNT: u32 = 1;
/// OfficeArtFSP payload length (spid + flags).
const FSP_PAYLOAD_LEN: u32 = 8;
/// OfficeArtSpgr payload length (empty group bounds rectangle).
const SPGR_PAYLOAD_LEN: u32 = 16;
/// Word's ClientAnchor payload: a 4-byte index into the PlcfSpa aCP array.
const WORD_CLIENT_ANCHOR_LEN: u32 = 4;
/// OfficeArtClientData payload length used by Word for shapes.
const CLIENT_DATA_LEN: u32 = 4;
/// OfficeArtIDCL `dgid` of the Main Document drawing.
const DGID_MAIN_DOCUMENT: u32 = 1;
/// OfficeArtIDCL `dgid` of the Header Document drawing.
const DGID_HEADER_DOCUMENT: u32 = 2;

/// Compute the OfficeArtFDGG `spidMax`: the start of the next shape-id
/// cluster beyond the highest allocated shape id.
fn spid_max(highest_shape_id: u32) -> u32 {
    highest_shape_id
        .saturating_add(SHAPE_IDS_PER_CLUSTER - highest_shape_id % SHAPE_IDS_PER_CLUSTER)
}

/// Append one OfficeArtWordDrawing element (dgglbl byte + OfficeArtDgContainer
/// holding the drawing's group shape and one shape per floating picture or
/// primitive). `dgid` is the drawing identifier (1 = Main, 2 = Header).
fn write_dg_container(
    out: &mut Vec<u8>,
    dgglbl: u8,
    dgid: u16,
    first_shape_id: u32,
    shapes: &[FloatingShapeInfo<'_>],
    bse_index_start: u32,
) {
    out.push(dgglbl);
    let dg_container_start = out.len();
    write_record_header(out, VERSION_CONTAINER, 0, RECORD_DG_CONTAINER, 0);

    // OfficeArtFDG: shape count (including the group shape) and next free spid.
    let shape_count = shapes.len() as u32 + DG_GROUP_SHAPE_COUNT;
    write_record_header(out, VERSION_DG, dgid, RECORD_DG, 8);
    out.extend_from_slice(&shape_count.to_le_bytes()); // csp
    out.extend_from_slice(&(first_shape_id + shapes.len() as u32).to_le_bytes()); // spidCur

    // OfficeArtSpgrContainer: the drawing's group shape plus all shapes.
    let spgr_container_start = out.len();
    write_record_header(out, VERSION_CONTAINER, 0, RECORD_SPGR_CONTAINER, 0);

    // Group shape: empty bounds rectangle and a group/patriarch FSP.
    let group_container_start = out.len();
    write_record_header(out, VERSION_CONTAINER, 0, RECORD_SP_CONTAINER, 0);
    write_record_header(out, VERSION_SPGR, 0, RECORD_SPGR, SPGR_PAYLOAD_LEN);
    out.extend_from_slice(&[0; SPGR_PAYLOAD_LEN as usize]);
    write_record_header(out, VERSION_SP, 0, RECORD_SP, FSP_PAYLOAD_LEN);
    out.extend_from_slice(&GROUP_SHAPE_ID.to_le_bytes());
    out.extend_from_slice(&(SP_FLAG_GROUP | SP_FLAG_PATRIARCH).to_le_bytes());
    patch_record_len(out, group_container_start);

    // One shape per floating picture or primitive. Pictures reference their
    // BSE through a 1-based pib index assigned in document order; text boxes
    // reference their FTXBXS entry the same way through the TXID (each
    // drawing's textbox PLC is indexed independently).
    let mut bse_index = bse_index_start;
    let mut ftxbxs_index = 0;
    for (index, shape) in shapes.iter().enumerate() {
        let shape_start = out.len();
        write_record_header(out, VERSION_CONTAINER, 0, RECORD_SP_CONTAINER, 0);
        write_record_header(
            out,
            VERSION_SP,
            shape.shape_type(),
            RECORD_SP,
            FSP_PAYLOAD_LEN,
        );
        out.extend_from_slice(&shape.shape_id.to_le_bytes());
        out.extend_from_slice(&(SP_FLAG_HAVE_ANCHOR | SP_FLAG_HAVE_SHAPE_TYPE).to_le_bytes());
        match &shape.content {
            // OfficeArtOPT: pib referencing this picture's BSE.
            FloatingShapeContent::Picture(_) => {
                write_opt_record(out, &[(OPT_PIB_BLIP_INDEX, bse_index)]);
                bse_index += 1;
            },
            // OfficeArtOPT: fill/line colors and boolean style properties.
            FloatingShapeContent::Primitive(primitive) => {
                super::shapes::write_shape_opt(out, primitive);
            },
        }
        // ClientAnchor: index of this shape's anchor CP in the PlcfSpa.
        write_record_header(
            out,
            VERSION_ATOM,
            0,
            RECORD_CLIENT_ANCHOR,
            WORD_CLIENT_ANCHOR_LEN,
        );
        out.extend_from_slice(&(index as u32).to_le_bytes());
        // OfficeArtClientData: present but unused.
        write_record_header(out, VERSION_ATOM, 0, RECORD_CLIENT_DATA, CLIENT_DATA_LEN);
        out.extend_from_slice(&0u32.to_le_bytes());
        // OfficeArtClientTextbox: links a text box to its FTXBXS entry.
        if shape.text.is_some() {
            super::shapes::write_client_textbox(out, ftxbxs_index);
            ftxbxs_index += 1;
        }
        patch_record_len(out, shape_start);
    }

    patch_record_len(out, spgr_container_start);
    patch_record_len(out, dg_container_start);
}

/// Build the OfficeArtContent referenced by `fcDggInfo` ([MS-DOC] 2.9.171):
/// an OfficeArtDggContainer (drawing defaults plus the blip store) followed
/// by one OfficeArtWordDrawing per non-empty drawing — the Main Document
/// drawing first, then the Header Document drawing.
///
/// `allocated_main_shapes` counts every shape id allocated in the Main
/// Document cluster (inline and floating pictures plus main-story shapes);
/// it only feeds the advisory spidMax/cluster bookkeeping.
pub(crate) fn build_dgg_info(
    main_shapes: &[FloatingShapeInfo<'_>],
    header_shapes: &[FloatingShapeInfo<'_>],
    allocated_main_shapes: u32,
) -> Result<Vec<u8>, DocWriteError> {
    let mut out = Vec::new();

    // ── OfficeArtDggContainer ──
    let dgg_container_start = out.len();
    write_record_header(&mut out, VERSION_CONTAINER, 0, RECORD_DGG_CONTAINER, 0);

    // OfficeArtFDGG: spidMax, cluster table, saved shape/drawing counts.
    let has_main_drawing = !main_shapes.is_empty();
    let has_header_drawing = !header_shapes.is_empty();
    let drawing_count = u32::from(has_main_drawing) + u32::from(has_header_drawing);
    // [MS-ODRAW] 2.2.47 defines cidcl as the number of OfficeArtIDCL
    // records plus one. There can be at most the main and header drawings.
    let cluster_count = drawing_count + 1;
    let highest_shape_id = if has_header_drawing {
        HEADER_FIRST_SHAPE_ID + header_shapes.len() as u32
    } else {
        FIRST_PICTURE_SHAPE_ID + allocated_main_shapes
    };
    let dgg_payload_len = 16 + 8 * drawing_count;
    write_record_header(&mut out, VERSION_ATOM, 0, RECORD_DGG, dgg_payload_len);
    out.extend_from_slice(&spid_max(highest_shape_id).to_le_bytes()); // spidMax
    out.extend_from_slice(&cluster_count.to_le_bytes()); // cidcl
    let saved_shapes = main_shapes.len() as u32
        + header_shapes.len() as u32
        + drawing_count * DG_GROUP_SHAPE_COUNT;
    out.extend_from_slice(&saved_shapes.to_le_bytes()); // cspSaved
    out.extend_from_slice(&drawing_count.to_le_bytes()); // cdgSaved
    // rgidcl: one OfficeArtIDCL per drawing, each owning a spid cluster.
    if has_main_drawing {
        out.extend_from_slice(&DGID_MAIN_DOCUMENT.to_le_bytes());
        out.extend_from_slice(&(DG_GROUP_SHAPE_COUNT + allocated_main_shapes).to_le_bytes());
    }
    if has_header_drawing {
        out.extend_from_slice(&DGID_HEADER_DOCUMENT.to_le_bytes());
        out.extend_from_slice(&(DG_GROUP_SHAPE_COUNT + header_shapes.len() as u32).to_le_bytes());
    }

    // OfficeArtBStoreContainer: one FBSE (with embedded BLIP) per picture in
    // either drawing. Main-drawing pictures are indexed first; the header
    // drawing's pib indices continue after them.
    let bstore_start = out.len();
    let main_picture_count = main_shapes
        .iter()
        .filter(|shape| matches!(shape.content, FloatingShapeContent::Picture(_)))
        .count();
    let picture_count = main_picture_count
        + header_shapes
            .iter()
            .filter(|shape| matches!(shape.content, FloatingShapeContent::Picture(_)))
            .count();
    let picture_count = u16::try_from(picture_count).map_err(|_| {
        DocWriteError::InvalidData("DOC BStore picture count exceeds u16".to_string())
    })?;
    if picture_count > 0x0fff {
        return Err(DocWriteError::InvalidData(
            "DOC BStore contains more than 4095 pictures".to_string(),
        ));
    }
    write_record_header(
        &mut out,
        VERSION_CONTAINER,
        picture_count,
        RECORD_BSTORE_CONTAINER,
        0,
    );
    for shape in main_shapes.iter().chain(header_shapes.iter()) {
        if let FloatingShapeContent::Picture(picture) = shape.content {
            write_bse_with_embedded_blip(&mut out, picture)?;
        }
    }
    patch_record_len(&mut out, bstore_start);

    patch_record_len(&mut out, dgg_container_start);

    // ── OfficeArtWordDrawing elements, main drawing first ──
    if has_main_drawing {
        write_dg_container(
            &mut out,
            DGGLBL_MAIN_DOCUMENT,
            DGID_MAIN_DOCUMENT as u16,
            FIRST_PICTURE_SHAPE_ID,
            main_shapes,
            OPT_PIB_FIRST_BSE,
        );
    }
    if has_header_drawing {
        write_dg_container(
            &mut out,
            DGGLBL_HEADER_DOCUMENT,
            DGID_HEADER_DOCUMENT as u16,
            HEADER_FIRST_SHAPE_ID,
            header_shapes,
            OPT_PIB_FIRST_BSE + main_picture_count as u32,
        );
    }

    Ok(out)
}

/// Back-patch the payload length of a container record header written with a
/// zero placeholder length.
fn patch_record_len(out: &mut [u8], record_start: usize) {
    let payload_len = (out.len() - record_start - RECORD_HEADER_LEN) as u32;
    out[record_start + 4..record_start + 8].copy_from_slice(&payload_len.to_le_bytes());
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Minimal byte sequences recognised by the format sniffer.
    fn png_bytes() -> Vec<u8> {
        let mut data = vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        data.extend_from_slice(&PNG_IHDR_LEN.to_be_bytes());
        data.extend_from_slice(b"IHDR");
        data.extend_from_slice(&32u32.to_be_bytes()); // width
        data.extend_from_slice(&16u32.to_be_bytes()); // height
        data.extend_from_slice(&[8, 2, 0, 0, 0]); // bit depth, RGB, compression, filter, interlace
        data.extend_from_slice(&[0; 4]); // CRC (not validated here)
        data
    }

    fn jpeg_bytes() -> Vec<u8> {
        let mut data = vec![0xFF, 0xD8, 0xFF, 0xE0];
        data.extend_from_slice(&16u16.to_be_bytes()); // APP0 length
        data.extend_from_slice(b"JFIF\0");
        data.extend_from_slice(&[0; 9]); // rest of APP0
        data.extend_from_slice(&[0xFF, 0xC0]);
        data.extend_from_slice(&17u16.to_be_bytes()); // SOF0 length
        data.push(8); // precision
        data.extend_from_slice(&24u16.to_be_bytes()); // height
        data.extend_from_slice(&48u16.to_be_bytes()); // width
        data
    }

    #[test]
    fn sniff_kind_detects_png_and_jpeg() {
        assert_eq!(sniff_kind(&png_bytes()).unwrap(), BlipKind::Png);
        assert_eq!(sniff_kind(&jpeg_bytes()).unwrap(), BlipKind::Jpeg);
        assert!(sniff_kind(b"not an image").is_err());
    }

    #[test]
    fn png_pixel_dimensions_reads_ihdr() {
        assert_eq!(png_pixel_dimensions(&png_bytes()), Some((32, 16)));
        assert_eq!(png_pixel_dimensions(&png_bytes()[..20]), None);
    }

    #[test]
    fn jpeg_pixel_dimensions_reads_sof0() {
        assert_eq!(jpeg_pixel_dimensions(&jpeg_bytes()), Some((48, 24)));
        assert_eq!(jpeg_pixel_dimensions(&jpeg_bytes()[..10]), None);
    }

    #[test]
    fn doc_picture_new_derives_dimensions_at_96_dpi() {
        let picture = DocPicture::new(png_bytes()).unwrap();
        assert_eq!(picture.kind(), BlipKind::Png);
        assert_eq!(picture.width_twips(), 32 * TWIPS_PER_INCH / ASSUMED_DPI);
        assert_eq!(picture.height_twips(), 16 * TWIPS_PER_INCH / ASSUMED_DPI);
    }

    #[test]
    fn doc_picture_rejects_unknown_format_and_bad_dimensions() {
        assert!(DocPicture::new(b"garbage".to_vec()).is_err());
        assert!(DocPicture::from_parts(png_bytes(), 0, 100).is_err());
        assert!(DocPicture::from_parts(png_bytes(), 100, MAX_PICF_DIMENSION_TWIPS + 1).is_err());
    }

    /// Parse an Escher record header, returning (version, instance, type, payload).
    fn parse_record_header(block: &[u8], offset: usize) -> (u16, u16, u16, u32) {
        let ver_inst = u16::from_le_bytes(block[offset..offset + 2].try_into().unwrap());
        let record_type = u16::from_le_bytes(block[offset + 2..offset + 4].try_into().unwrap());
        let length = u32::from_le_bytes(block[offset + 4..offset + 8].try_into().unwrap());
        (ver_inst & 0xF, ver_inst >> 4, record_type, length)
    }

    #[test]
    fn picture_block_layout_matches_ms_doc() {
        let picture = DocPicture::new(png_bytes()).unwrap();
        let mut block = Vec::new();
        write_picture_block(&picture, FIRST_PICTURE_SHAPE_ID, &mut block).unwrap();

        // PICF header.
        let lcb = i32::from_le_bytes(block[0..4].try_into().unwrap()) as usize;
        assert_eq!(lcb, block.len());
        assert_eq!(
            i16::from_le_bytes(block[4..6].try_into().unwrap()),
            PICF_CB_HEADER
        );
        assert_eq!(
            i16::from_le_bytes(block[6..8].try_into().unwrap()),
            PICF_MM_SHAPE
        );
        // Block type byte (grf low byte) 0x00 + mm 0x64 is what the reader
        // recognises as a Word 2000 picture.
        assert_eq!(block[0x0E], 0);
        let dxa_goal = i16::from_le_bytes(block[0x1C..0x1E].try_into().unwrap());
        let dya_goal = i16::from_le_bytes(block[0x1E..0x20].try_into().unwrap());
        assert_eq!(dxa_goal as u32, picture.width_twips());
        assert_eq!(dya_goal as u32, picture.height_twips());

        // SpContainer follows the PICF.
        let (ver, _inst, record_type, len) = parse_record_header(&block, PICF_HEADER_LEN);
        assert_eq!((ver, record_type), (VERSION_CONTAINER, RECORD_SP_CONTAINER));
        let bse_offset = PICF_HEADER_LEN + RECORD_HEADER_LEN + len as usize;

        // BSE record with the embedded BLIP.
        let (ver, inst, record_type, len) = parse_record_header(&block, bse_offset);
        assert_eq!((ver, record_type), (2, RECORD_BSE));
        assert_eq!(inst, u16::from(BlipKind::Png.raw()));
        let bse = &block[bse_offset + RECORD_HEADER_LEN..];
        assert_eq!(bse[0], BlipKind::Png.raw()); // btWin32
        assert_eq!(bse[1], BlipKind::Png.raw()); // btMacOS
        let blip_size = u32::from_le_bytes(bse[20..24].try_into().unwrap()) as usize;
        let c_ref = u32::from_le_bytes(bse[24..28].try_into().unwrap());
        let fo_delay = u32::from_le_bytes(bse[28..32].try_into().unwrap());
        assert_eq!(c_ref, 1);
        assert_eq!(fo_delay, BSE_NO_DELAY_STREAM); // embedded
        assert_eq!(blip_size + BSE_HEADER_LEN, len as usize);

        // Embedded BLIP record.
        let blip = &bse[BSE_HEADER_LEN..BSE_HEADER_LEN + blip_size];
        let parsed = litchi_odraw::image::Blip::parse(blip).unwrap();
        assert_eq!(parsed.kind(), BlipKind::Png);
        assert_eq!(parsed.record().instance(), 0x6e0);
        // The BLIP UID matches the BSE rgbUid, and the payload is byte-identical.
        assert_eq!(parsed.uids().unwrap().effective(), digest(picture.data()));
        assert_eq!(parsed.data(), picture.data());
    }

    #[test]
    fn picture_uid_is_the_md4_digest_of_blip_data() {
        let picture = DocPicture::new(png_bytes()).unwrap();
        assert_eq!(
            digest(picture.data()),
            picture_blip(&picture).unwrap().uid()
        );
        // RFC 1320 test vector.
        assert_eq!(
            digest(b"abc").bytes(),
            [
                0xa4, 0x48, 0x01, 0x7a, 0xaf, 0x21, 0xd8, 0x52, 0x5f, 0xc1, 0x0a, 0xe8, 0x7a, 0xa6,
                0x72, 0x9d,
            ]
        );
        assert_eq!(
            digest(
                b"12345678901234567890123456789012345678901234567890123456789012345678901234567890"
            )
            .bytes(),
            [
                0xe3, 0x3b, 0x4d, 0xdc, 0x9c, 0x38, 0xf2, 0x19, 0x9c, 0x3e, 0x7b, 0x16, 0x4f, 0xcc,
                0x05, 0x36,
            ]
        );
    }

    #[test]
    fn picture_block_bse_parses_with_crate_reader() {
        use litchi_odraw::Record;
        use litchi_odraw::image::{Blip, Context, Entry, Storage};

        let picture = DocPicture::new(jpeg_bytes()).unwrap();
        let mut block = Vec::new();
        write_picture_block(&picture, FIRST_PICTURE_SHAPE_ID, &mut block).unwrap();

        let (_ver, _inst, _record_type, sp_len) = parse_record_header(&block, PICF_HEADER_LEN);
        let bse_offset = PICF_HEADER_LEN + RECORD_HEADER_LEN + sp_len as usize;
        let (record, _) = Record::parse(&block, bse_offset).unwrap();
        let bse = Entry::parse(record).unwrap();
        assert_eq!(bse.kind().unwrap(), BlipKind::Jpeg);
        assert!(matches!(
            bse.storage().unwrap(),
            Storage::Embedded(Blip::Jpeg(_))
        ));
        let blip = bse.resolve(Context::new()).unwrap().unwrap();
        assert_eq!(blip.data(), picture.data());
    }

    // ── Floating pictures ──

    use crate::parts::spa::{
        ShapeHorizontalOrigin, ShapeTextWrap, ShapeVerticalOrigin, ShapeWrapSide,
    };

    fn floating_shapes<'a>(
        png: &'a DocPicture,
        jpeg: &'a DocPicture,
        positions: &'a [FloatingPosition],
    ) -> Vec<FloatingShapeInfo<'a>> {
        vec![
            FloatingShapeInfo {
                anchor_cp: 12,
                shape_id: FIRST_PICTURE_SHAPE_ID,
                content: FloatingShapeContent::Picture(png),
                width_twips: png.width_twips(),
                height_twips: png.height_twips(),
                position: &positions[0],
                text: None,
            },
            FloatingShapeInfo {
                anchor_cp: 30,
                shape_id: FIRST_PICTURE_SHAPE_ID + 1,
                content: FloatingShapeContent::Picture(jpeg),
                width_twips: jpeg.width_twips(),
                height_twips: jpeg.height_twips(),
                position: &positions[1],
                text: None,
            },
        ]
    }

    fn sample_positions() -> [FloatingPosition; 2] {
        [
            FloatingPosition::new(1440, 720)
                .with_origins(ShapeHorizontalOrigin::Page, ShapeVerticalOrigin::Paragraph)
                .with_text_wrap(ShapeTextWrap::Square)
                .lock_anchor(true),
            FloatingPosition::new(2880, 1440)
                .with_text_wrap(ShapeTextWrap::None)
                .behind_text(true),
        ]
    }

    #[test]
    fn floating_position_builder_defaults() {
        let position = FloatingPosition::new(100, 200);
        assert_eq!(position.horizontal_origin, ShapeHorizontalOrigin::Page);
        assert_eq!(position.vertical_origin, ShapeVerticalOrigin::Page);
        assert_eq!(position.wrap, ShapeTextWrap::Square);
        assert_eq!(position.wrap_side, ShapeWrapSide::Both);
        assert!(!position.behind_text);
        assert!(!position.anchor_locked);
    }

    #[test]
    fn plcf_spa_layout_matches_ms_doc() {
        let png = DocPicture::new(png_bytes()).unwrap();
        let jpeg = DocPicture::new(jpeg_bytes()).unwrap();
        let positions = sample_positions();
        let shapes = floating_shapes(&png, &jpeg, &positions);

        let plcf = build_plcf_spa(&shapes, 500);
        let anchors = crate::parts::spa::parse_plcf_spa(&plcf).unwrap();
        assert_eq!(anchors.len(), 2);
        assert_eq!(anchors[0].cp, 12);
        assert_eq!(anchors[1].cp, 30);

        let first = &anchors[0].spa;
        assert_eq!(first.shape_id, FIRST_PICTURE_SHAPE_ID);
        assert_eq!((first.left, first.top), (1440, 720));
        assert_eq!(first.width() as u32, png.width_twips());
        assert_eq!(first.height() as u32, png.height_twips());
        assert_eq!(first.horizontal_origin, ShapeHorizontalOrigin::Page);
        assert_eq!(first.vertical_origin, ShapeVerticalOrigin::Paragraph);
        assert_eq!(first.wrap, ShapeTextWrap::Square);
        assert!(first.anchor_locked);

        let second = &anchors[1].spa;
        assert_eq!(second.shape_id, FIRST_PICTURE_SHAPE_ID + 1);
        assert_eq!(second.wrap, ShapeTextWrap::None);
        assert!(second.below_text);

        // The final CP is ccpText and exceeds every anchor CP.
        let final_cp = u32::from_le_bytes(plcf[8..12].try_into().unwrap());
        assert_eq!(final_cp, 500);
    }

    /// Collect all Escher records in a byte slice as (offset, ver, instance, type, len).
    /// Handles the OfficeArtWordDrawing dgglbl byte between top-level records.
    fn collect_records(data: &[u8], start: usize, end: usize) -> Vec<(usize, u16, u16, u16, u32)> {
        fn walk(
            data: &[u8],
            start: usize,
            end: usize,
            records: &mut Vec<(usize, u16, u16, u16, u32)>,
        ) {
            let mut offset = start;
            while offset + RECORD_HEADER_LEN <= end {
                let (ver, inst, record_type, len) = parse_record_header(data, offset);
                records.push((offset, ver, inst, record_type, len));
                if ver == VERSION_CONTAINER {
                    walk(
                        data,
                        offset + RECORD_HEADER_LEN,
                        offset + RECORD_HEADER_LEN + len as usize,
                        records,
                    );
                }
                offset += RECORD_HEADER_LEN + len as usize;
            }
        }

        let mut records = Vec::new();
        let (_ver, _inst, _record_type, len) = parse_record_header(data, start);
        let mut drawing_offset = start + RECORD_HEADER_LEN + len as usize;
        walk(data, start, drawing_offset, &mut records);
        // Each OfficeArtWordDrawing: dgglbl byte followed by a DgContainer.
        while drawing_offset + 1 + RECORD_HEADER_LEN <= end {
            let dgglbl = data[drawing_offset];
            assert!(dgglbl <= 1, "dgglbl must be 0 (main) or 1 (header)");
            let (_ver, _inst, _record_type, len) = parse_record_header(data, drawing_offset + 1);
            walk(
                data,
                drawing_offset + 1,
                drawing_offset + 1 + RECORD_HEADER_LEN + len as usize,
                &mut records,
            );
            drawing_offset += 1 + RECORD_HEADER_LEN + len as usize;
        }
        records
    }

    #[test]
    fn dgg_info_layout_matches_ms_odraw() {
        let png = DocPicture::new(png_bytes()).unwrap();
        let jpeg = DocPicture::new(jpeg_bytes()).unwrap();
        let positions = sample_positions();
        let shapes = floating_shapes(&png, &jpeg, &positions);

        let dgg = build_dgg_info(&shapes, &[], 3).unwrap();
        let records = collect_records(&dgg, 0, dgg.len());

        // Top level: DggContainer, then dgglbl + DgContainer.
        let (off, ver, _inst, record_type, len) = records[0];
        assert_eq!(
            (off, ver, record_type),
            (0, VERSION_CONTAINER, RECORD_DGG_CONTAINER)
        );
        let dg_container_offset = off + RECORD_HEADER_LEN + len as usize;
        assert_eq!(dgg[dg_container_offset], DGGLBL_MAIN_DOCUMENT);

        // Dgg: cluster table and counts.
        let dgg_record = records
            .iter()
            .find(|record| record.3 == RECORD_DGG)
            .unwrap();
        let dgg_payload = &dgg[dgg_record.0 + RECORD_HEADER_LEN..];
        let spid_max = u32::from_le_bytes(dgg_payload[0..4].try_into().unwrap());
        assert_eq!(spid_max, 2 * SHAPE_IDS_PER_CLUSTER);
        assert_eq!(
            u32::from_le_bytes(dgg_payload[4..8].try_into().unwrap()),
            2 // one OfficeArtIDCL plus one, per [MS-ODRAW] 2.2.47
        );
        // cspSaved = 2 shapes + group; cdgSaved = 1 drawing.
        assert_eq!(
            u32::from_le_bytes(dgg_payload[8..12].try_into().unwrap()),
            3
        );
        assert_eq!(
            u32::from_le_bytes(dgg_payload[12..16].try_into().unwrap()),
            1
        );

        // BStoreContainer holds one BSE per picture.
        let bstore = records
            .iter()
            .find(|record| record.3 == RECORD_BSTORE_CONTAINER)
            .unwrap();
        assert_eq!(bstore.2, 2);
        let bses: Vec<_> = records
            .iter()
            .filter(|record| record.3 == RECORD_BSE)
            .collect();
        assert_eq!(bses.len(), 2);

        // Dg: shape count includes the group shape; spidCur is next free.
        let dg = records.iter().find(|record| record.3 == RECORD_DG).unwrap();
        let dg_payload = &dgg[dg.0 + RECORD_HEADER_LEN..];
        assert_eq!(u32::from_le_bytes(dg_payload[0..4].try_into().unwrap()), 3);
        assert_eq!(
            u32::from_le_bytes(dg_payload[4..8].try_into().unwrap()),
            FIRST_PICTURE_SHAPE_ID + 2
        );

        // Group shape with fGroup|fPatriarch, then one picture shape per picture.
        let sps: Vec<_> = records
            .iter()
            .filter(|record| record.3 == RECORD_SP)
            .collect();
        assert_eq!(sps.len(), 3);
        let (group_spid, group_flags) = (
            u32::from_le_bytes(dgg[sps[0].0 + 8..sps[0].0 + 12].try_into().unwrap()),
            u32::from_le_bytes(dgg[sps[0].0 + 12..sps[0].0 + 16].try_into().unwrap()),
        );
        assert_eq!(group_spid, GROUP_SHAPE_ID);
        assert_eq!(group_flags, SP_FLAG_GROUP | SP_FLAG_PATRIARCH);
        for (index, sp) in sps[1..].iter().enumerate() {
            let spid = u32::from_le_bytes(dgg[sp.0 + 8..sp.0 + 12].try_into().unwrap());
            assert_eq!(spid, FIRST_PICTURE_SHAPE_ID + index as u32);
            assert_eq!(sp.2, SHAPE_TYPE_PICTURE_FRAME);
        }

        // OPT pib references the BSEs 1-based, in order.
        let opts: Vec<_> = records
            .iter()
            .filter(|record| record.3 == RECORD_OPT)
            .collect();
        assert_eq!(opts.len(), 2);
        for (index, opt) in opts.iter().enumerate() {
            let prop = u16::from_le_bytes(dgg[opt.0 + 8..opt.0 + 10].try_into().unwrap());
            let value = u32::from_le_bytes(dgg[opt.0 + 10..opt.0 + 14].try_into().unwrap());
            assert_eq!(prop, OPT_PIB_BLIP_INDEX);
            assert_eq!(value, index as u32 + OPT_PIB_FIRST_BSE);
        }

        // ClientAnchor records index into the PlcfSpa aCP array, in order.
        let anchors: Vec<_> = records
            .iter()
            .filter(|record| record.3 == RECORD_CLIENT_ANCHOR)
            .collect();
        assert_eq!(anchors.len(), 2);
        for (index, anchor) in anchors.iter().enumerate() {
            let value = u32::from_le_bytes(dgg[anchor.0 + 8..anchor.0 + 12].try_into().unwrap());
            assert_eq!(value, index as u32);
        }

        // Spids across the drawing are unique.
        let mut spids: Vec<u32> = sps
            .iter()
            .map(|sp| u32::from_le_bytes(dgg[sp.0 + 8..sp.0 + 12].try_into().unwrap()))
            .collect();
        spids.sort_unstable();
        spids.dedup();
        assert_eq!(spids.len(), 3);
    }

    #[test]
    fn dgg_info_bse_blips_parse_with_crate_reader() {
        use litchi_odraw::Record;
        use litchi_odraw::image::{Context, Entry};

        let png = DocPicture::new(png_bytes()).unwrap();
        let jpeg = DocPicture::new(jpeg_bytes()).unwrap();
        let positions = sample_positions();
        let shapes = floating_shapes(&png, &jpeg, &positions);

        let dgg = build_dgg_info(&shapes, &[], 2).unwrap();
        let records = collect_records(&dgg, 0, dgg.len());
        let bses: Vec<_> = records
            .iter()
            .filter(|record| record.3 == RECORD_BSE)
            .collect();
        assert_eq!(bses.len(), 2);

        let expected = [(BlipKind::Png, png.data()), (BlipKind::Jpeg, jpeg.data())];
        for (bse_record, (kind, payload)) in bses.iter().zip(expected.iter()) {
            let (record, _) = Record::parse(&dgg, bse_record.0).unwrap();
            let bse = Entry::parse(record).unwrap();
            assert_eq!(bse.kind().unwrap(), *kind);
            let blip = bse.resolve(Context::new()).unwrap().unwrap();
            assert_eq!(blip.data(), *payload);
        }
    }

    #[test]
    fn spid_max_rounds_up_to_next_cluster() {
        assert_eq!(spid_max(FIRST_PICTURE_SHAPE_ID), 2 * SHAPE_IDS_PER_CLUSTER);
        assert_eq!(
            spid_max(FIRST_PICTURE_SHAPE_ID + 1),
            2 * SHAPE_IDS_PER_CLUSTER
        );
        assert_eq!(
            spid_max(FIRST_PICTURE_SHAPE_ID + SHAPE_IDS_PER_CLUSTER),
            3 * SHAPE_IDS_PER_CLUSTER
        );
    }

    #[test]
    fn dgg_info_with_header_drawing_uses_own_cluster() {
        use crate::writer::shapes::DocDrawingShape;

        let png = DocPicture::new(png_bytes()).unwrap();
        let jpeg = DocPicture::new(jpeg_bytes()).unwrap();
        let positions = sample_positions();
        let main_shapes = floating_shapes(&png, &jpeg, &positions);

        let rect = DocDrawingShape::new(crate::writer::shapes::DocShapeKind::Rectangle, 1440, 720)
            .unwrap();
        let position = FloatingPosition::new(720, 360);
        let header_shapes = vec![FloatingShapeInfo {
            anchor_cp: 0,
            shape_id: HEADER_FIRST_SHAPE_ID,
            content: FloatingShapeContent::Primitive(&rect),
            width_twips: rect.width_twips(),
            height_twips: rect.height_twips(),
            position: &position,
            text: Some("hdr"),
        }];

        let dgg = build_dgg_info(&main_shapes, &header_shapes, 2).unwrap();
        let records = collect_records(&dgg, 0, dgg.len());

        // Two drawings: main (dgglbl 0, Dg instance 1) and header (dgglbl 1,
        // Dg instance 2). The Dgg cluster table has two entries, while cidcl
        // is the entry count plus one.
        let dgg_record = records
            .iter()
            .find(|record| record.3 == RECORD_DGG)
            .unwrap();
        let dgg_payload = &dgg[dgg_record.0 + RECORD_HEADER_LEN..];
        assert_eq!(u32::from_le_bytes(dgg_payload[4..8].try_into().unwrap()), 3);
        assert_eq!(
            u32::from_le_bytes(dgg_payload[0..4].try_into().unwrap()),
            3 * SHAPE_IDS_PER_CLUSTER
        );
        let dg_records: Vec<_> = records
            .iter()
            .filter(|record| record.3 == RECORD_DG)
            .collect();
        assert_eq!(dg_records.len(), 2);
        assert_eq!(dg_records[0].2, 1);
        assert_eq!(dg_records[1].2, 2);

        // dgglbl bytes: main then header.
        let dg_container_offsets: Vec<usize> = records
            .iter()
            .filter(|record| record.3 == RECORD_DG_CONTAINER)
            .map(|record| record.0)
            .collect();
        assert_eq!(dg_container_offsets.len(), 2);
        assert_eq!(dgg[dg_container_offsets[0] - 1], 0);
        assert_eq!(dgg[dg_container_offsets[1] - 1], 1);

        // The header shape uses the header cluster spid and its own TXID
        // numbering (0x10000 for the first header text box).
        let sps: Vec<_> = records
            .iter()
            .filter(|record| record.3 == RECORD_SP)
            .collect();
        let spids: Vec<u32> = sps
            .iter()
            .map(|sp| u32::from_le_bytes(dgg[sp.0 + 8..sp.0 + 12].try_into().unwrap()))
            .collect();
        assert!(spids.contains(&HEADER_FIRST_SHAPE_ID));

        let txid_records: Vec<_> = records.iter().filter(|record| record.3 == 0xF00D).collect();
        assert_eq!(txid_records.len(), 1);
        let txid = u32::from_le_bytes(
            dgg[txid_records[0].0 + 8..txid_records[0].0 + 12]
                .try_into()
                .unwrap(),
        );
        assert_eq!(txid, 0x0001_0000);
    }
}
