//! Images writer for DOC files
//!
//! Generates OfficeArtWordDrawing blocks ([MS-DOC] 2.9.172) for inline
//! pictures: a PICF picture descriptor ([MS-DOC] 2.9.161) followed by an
//! OfficeArtSpContainer and an OfficeArtFBSE record carrying the embedded
//! bitmap BLIP ([MS-ODRAW] 2.2). Blocks are appended to the Data stream; a
//! 0x0001 picture character in the text stream references each block through
//! sprmCPicLocation. Image bytes are stored verbatim after format sniffing —
//! no re-encoding is performed.

use super::core::DocWriteError;
use crate::doc::parts::images::PictureType;
use crate::doc::parts::spa::{
    SPA_LEN, ShapeHorizontalOrigin, ShapeTextWrap, ShapeVerticalOrigin, ShapeWrapSide, Spa,
};

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
const RECORD_BLIP_JPEG: u16 = 0xF01D;
const RECORD_BLIP_PNG: u16 = 0xF01E;

// Escher record versions.
const VERSION_CONTAINER: u16 = 0xF;
const VERSION_DG: u16 = 0x0;
const VERSION_SPGR: u16 = 0x1;
const VERSION_BSE: u16 = 0x2;
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
/// Length of a BLIP record UID in bytes.
const BLIP_UID_LEN: usize = 16;
/// BLIP payload marker byte for embedded (non-external) picture data.
const BLIP_EMBEDDED_MARKER: u8 = 0xFF;
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

/// MSOBLIPTYPE value for JPEG blips.
const MSO_BLIP_TYPE_JPEG: u8 = 0x05;
/// MSOBLIPTYPE value for PNG blips.
const MSO_BLIP_TYPE_PNG: u8 = 0x06;
/// OfficeArtBlip record instance for single-UID JPEG blips.
const BLIP_INSTANCE_JPEG: u16 = 0x46A;
/// OfficeArtBlip record instance for single-UID PNG blips.
const BLIP_INSTANCE_PNG: u16 = 0x6E0;

/// Bitmap formats supported by the DOC picture writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PictureFormat {
    /// JPEG image (OfficeArtBlipJPEG).
    Jpeg,
    /// PNG image (OfficeArtBlipPNG).
    Png,
}

impl PictureFormat {
    /// Detect the picture format from raw bytes via the crate's format detection.
    fn sniff(data: &[u8]) -> Option<Self> {
        match PictureType::from_data(data) {
            PictureType::Jpeg => Some(Self::Jpeg),
            PictureType::Png => Some(Self::Png),
            _ => None,
        }
    }

    /// OfficeArtBlip record type for this format.
    fn blip_record_type(self) -> u16 {
        match self {
            Self::Jpeg => RECORD_BLIP_JPEG,
            Self::Png => RECORD_BLIP_PNG,
        }
    }

    /// OfficeArtBlip record instance for the single-UID record variant.
    fn blip_instance(self) -> u16 {
        match self {
            Self::Jpeg => BLIP_INSTANCE_JPEG,
            Self::Png => BLIP_INSTANCE_PNG,
        }
    }

    /// MSOBLIPTYPE value used in the BSE header.
    fn mso_blip_type(self) -> u8 {
        match self {
            Self::Jpeg => MSO_BLIP_TYPE_JPEG,
            Self::Png => MSO_BLIP_TYPE_PNG,
        }
    }
}

/// An inline picture to be embedded in a DOC document.
///
/// The image bytes are stored as-is; the writer only sniffs the format and
/// records the display dimensions (in twips, 1/1440 inch).
#[derive(Debug, Clone)]
pub struct DocPicture {
    /// Raw image bytes (PNG or JPEG), stored verbatim.
    data: Vec<u8>,
    /// Detected image format.
    format: PictureFormat,
    /// Display width in twips.
    width_twips: u32,
    /// Display height in twips.
    height_twips: u32,
}

impl DocPicture {
    /// Create a picture from raw image bytes.
    ///
    /// The format is sniffed from the byte signature and the display
    /// dimensions are derived from the pixel dimensions at 96 DPI. Returns an
    /// error for unsupported formats or when the pixel dimensions cannot be
    /// determined; use [`Self::from_parts`] to supply dimensions explicitly.
    pub fn new(data: Vec<u8>) -> Result<Self, DocWriteError> {
        let format = sniff_format(&data)?;
        let (width_px, height_px) = pixel_dimensions(format, &data).ok_or_else(|| {
            DocWriteError::InvalidData(
                "DOC picture pixel dimensions are unreadable; use DocPicture::from_parts"
                    .to_string(),
            )
        })?;
        Self::from_parts(data, pixels_to_twips(width_px), pixels_to_twips(height_px))
    }

    /// Create a picture from raw image bytes and explicit display dimensions
    /// in twips (1/1440 inch).
    pub fn from_parts(
        data: Vec<u8>,
        width_twips: u32,
        height_twips: u32,
    ) -> Result<Self, DocWriteError> {
        let format = sniff_format(&data)?;
        let picture = Self {
            data,
            format,
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

    /// Detected image format.
    pub fn format(&self) -> PictureFormat {
        self.format
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
fn sniff_format(data: &[u8]) -> Result<PictureFormat, DocWriteError> {
    PictureFormat::sniff(data).ok_or_else(|| {
        DocWriteError::InvalidData(
            "DOC picture data is not a supported bitmap format (PNG or JPEG)".to_string(),
        )
    })
}

/// Convert a pixel count to twips at the assumed screen resolution.
fn pixels_to_twips(pixels: u32) -> u32 {
    pixels.saturating_mul(TWIPS_PER_INCH) / ASSUMED_DPI
}

/// Extract the pixel dimensions of an image without decoding it.
fn pixel_dimensions(format: PictureFormat, data: &[u8]) -> Option<(u32, u32)> {
    match format {
        PictureFormat::Png => png_pixel_dimensions(data),
        PictureFormat::Jpeg => jpeg_pixel_dimensions(data),
    }
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
    let dimensions = data
        .get(PNG_SIGNATURE_LEN + PNG_CHUNK_HEADER_LEN..PNG_SIGNATURE_LEN + PNG_CHUNK_HEADER_LEN + 8)?;
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
                let length = u16::from_be_bytes(
                    data.get(offset + 2..offset + 4)?.try_into().ok()?,
                ) as usize;
                offset = offset.checked_add(2 + length)?;
            },
        }
    }
    None
}

/// FNV-1a offset basis and prime, used to derive BLIP record UIDs.
const FNV_OFFSET_BASIS: u64 = 0xcbf2_9ce4_8422_2325;
const FNV_PRIME: u64 = 0x0000_0100_0000_01b3;

/// Compute the FNV-1a hash of `bytes` seeded with `seed`.
fn fnv1a(seed: u64, bytes: &[u8]) -> u64 {
    let mut hash = seed;
    for &byte in bytes {
        hash ^= u64::from(byte);
        hash = hash.wrapping_mul(FNV_PRIME);
    }
    hash
}

/// Derive the 16-byte UID shared by a BSE and its embedded BLIP record.
///
/// Word uses the UID to deduplicate identical blips; mixing in the shape id
/// keeps the UIDs of repeated images distinct.
fn picture_uid(picture: &DocPicture, shape_id: u32) -> [u8; BLIP_UID_LEN] {
    let first = fnv1a(FNV_OFFSET_BASIS ^ u64::from(shape_id), &picture.data);
    let second = fnv1a(first ^ FNV_PRIME, &picture.data);
    let mut uid = [0u8; BLIP_UID_LEN];
    uid[..8].copy_from_slice(&first.to_le_bytes());
    uid[8..].copy_from_slice(&second.to_le_bytes());
    uid
}

/// Append an Escher record header ([MS-ODRAW] 2.2.1).
fn write_record_header(
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
    write_record_header(out, VERSION_SP, SHAPE_TYPE_PICTURE_FRAME, RECORD_SP, SP_PAYLOAD_LEN);
    out.extend_from_slice(&shape_id.to_le_bytes());
    out.extend_from_slice(&(SP_FLAG_HAVE_ANCHOR | SP_FLAG_HAVE_SHAPE_TYPE).to_le_bytes());
    // OfficeArtOPT: pib referencing the adjacent BSE (1-based index).
    write_record_header(out, VERSION_OPT, 1, RECORD_OPT, OPT_PAYLOAD_LEN);
    out.extend_from_slice(&OPT_PIB_BLIP_INDEX.to_le_bytes());
    out.extend_from_slice(&OPT_PIB_FIRST_BSE.to_le_bytes());
    // OfficeArtClientAnchor: empty for inline pictures.
    write_record_header(out, VERSION_ATOM, 0, RECORD_CLIENT_ANCHOR, CLIENT_ANCHOR_PAYLOAD_LEN);
    out.extend_from_slice(&[0; CLIENT_ANCHOR_PAYLOAD_LEN as usize]);
}

/// Append an OfficeArtFBSE record with the embedded OfficeArtBlip record for
/// a picture. Used both for the Data-stream picture blocks and for the
/// BStoreContainer inside the drawing group of floating pictures.
fn write_bse_with_embedded_blip(out: &mut Vec<u8>, picture: &DocPicture, shape_id: u32) {
    let blip_payload_len = (BLIP_UID_LEN + 1 + picture.data.len()) as u32;
    let blip_record_len = RECORD_HEADER_LEN as u32 + blip_payload_len;
    let bse_payload_len = BSE_HEADER_LEN as u32 + blip_record_len;

    write_record_header(
        out,
        VERSION_BSE,
        u16::from(picture.format.mso_blip_type()),
        RECORD_BSE,
        bse_payload_len,
    );
    out.push(picture.format.mso_blip_type()); // btWin32
    out.push(picture.format.mso_blip_type()); // btMacOS
    out.extend_from_slice(&picture_uid(picture, shape_id)); // rgbUid
    out.extend_from_slice(&0u16.to_le_bytes()); // tag
    out.extend_from_slice(&blip_record_len.to_le_bytes()); // size
    out.extend_from_slice(&1u32.to_le_bytes()); // cRef
    out.extend_from_slice(&BSE_NO_DELAY_STREAM.to_le_bytes()); // foDelay: BLIP is embedded
    out.push(0); // usage
    out.push(0); // cbName
    out.push(0); // unused2
    out.push(0); // unused3

    // OfficeArtBlip: single-UID bitmap record with the raw image bytes.
    write_record_header(
        out,
        VERSION_ATOM,
        picture.format.blip_instance(),
        picture.format.blip_record_type(),
        blip_payload_len,
    );
    out.extend_from_slice(&picture_uid(picture, shape_id));
    out.push(BLIP_EMBEDDED_MARKER);
    out.extend_from_slice(&picture.data);
}

/// Append an OfficeArtWordDrawing block (PICF + shape container + BSE with an
/// embedded BLIP) to the Data stream.
pub(crate) fn write_picture_block(picture: &DocPicture, shape_id: u32, out: &mut Vec<u8>) {
    let blip_payload_len = (BLIP_UID_LEN + 1 + picture.data.len()) as u32;
    let blip_record_len = RECORD_HEADER_LEN as u32 + blip_payload_len;
    let bse_payload_len = BSE_HEADER_LEN as u32 + blip_record_len;

    let block_start = out.len();
    // lcb covers the PICF header plus everything that follows it.
    let lcb = PICF_HEADER_LEN as u32
        + SHAPE_CONTAINER_LEN
        + RECORD_HEADER_LEN as u32
        + bse_payload_len;
    write_picf(out, picture, lcb);
    write_shape_container(out, shape_id);
    write_bse_with_embedded_blip(out, picture, shape_id);

    debug_assert_eq!(out.len() - block_start, lcb as usize);
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

/// Everything the table-stream builders need to know about one floating
/// picture anchored in the Main Document.
pub(crate) struct FloatingShapeInfo<'a> {
    /// Character position of the 0x0008 anchor character (Main Document CP).
    pub anchor_cp: u32,
    /// Shape id, shared with the picture's Data-stream block.
    pub shape_id: u32,
    /// The picture itself.
    pub picture: &'a DocPicture,
    /// Position and wrapping.
    pub position: &'a FloatingPosition,
}

impl FloatingShapeInfo<'_> {
    /// Build the Spa record for this shape.
    fn spa(&self) -> Spa {
        let left = self.position.left_twips;
        let top = self.position.top_twips;
        Spa {
            shape_id: self.shape_id,
            left,
            top,
            right: left + self.picture.width_twips() as i32,
            bottom: top + self.picture.height_twips() as i32,
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
/// OfficeArtFDGG cluster-table entry count for a single drawing.
const DGG_SINGLE_CLUSTER: u32 = 1;
/// OfficeArtFDG `recInstance`: identifier of the drawing (1-based).
const DG_INSTANCE_FIRST_DRAWING: u16 = 1;
/// OfficeArtIDCL `dgid` of the single drawing written by this writer.
const IDCL_FIRST_DRAWING: u32 = 1;

/// Compute the OfficeArtFDGG `spidMax`: the start of the next shape-id
/// cluster beyond every allocated picture shape id.
fn spid_max(total_pictures: u32) -> u32 {
    let highest = FIRST_PICTURE_SHAPE_ID + total_pictures;
    highest
        .checked_add(SHAPE_IDS_PER_CLUSTER - highest % SHAPE_IDS_PER_CLUSTER)
        .unwrap_or(u32::MAX)
}

/// Build the OfficeArtContent referenced by `fcDggInfo` ([MS-DOC] 2.9.171):
/// an OfficeArtDggContainer (drawing defaults plus the blip store) followed
/// by the Main Document's OfficeArtWordDrawing (dgglbl + OfficeArtDgContainer
/// holding one picture-frame shape per floating picture).
pub(crate) fn build_dgg_info(shapes: &[FloatingShapeInfo<'_>], total_pictures: u32) -> Vec<u8> {
    let mut out = Vec::new();

    // ── OfficeArtDggContainer ──
    let dgg_container_start = out.len();
    write_record_header(&mut out, VERSION_CONTAINER, 0, RECORD_DGG_CONTAINER, 0);

    // OfficeArtFDGG: spidMax, cluster table, saved shape/drawing counts.
    const DGG_PAYLOAD_LEN: u32 = 24; // 4 header fields + one OfficeArtIDCL
    write_record_header(&mut out, VERSION_ATOM, 0, RECORD_DGG, DGG_PAYLOAD_LEN);
    out.extend_from_slice(&spid_max(total_pictures).to_le_bytes()); // spidMax
    out.extend_from_slice(&DGG_SINGLE_CLUSTER.to_le_bytes()); // cidcl
    let shape_count = shapes.len() as u32 + DG_GROUP_SHAPE_COUNT;
    out.extend_from_slice(&shape_count.to_le_bytes()); // cspSaved
    out.extend_from_slice(&DGG_SINGLE_CLUSTER.to_le_bytes()); // cdgSaved
    out.extend_from_slice(&IDCL_FIRST_DRAWING.to_le_bytes()); // rgidcl[0].dgid
    out.extend_from_slice(&shape_count.to_le_bytes()); // rgidcl[0].cspidCur

    // OfficeArtBStoreContainer: one FBSE (with embedded BLIP) per picture.
    let bstore_start = out.len();
    write_record_header(&mut out, VERSION_CONTAINER, shapes.len() as u16, RECORD_BSTORE_CONTAINER, 0);
    for shape in shapes {
        write_bse_with_embedded_blip(&mut out, shape.picture, shape.shape_id);
    }
    patch_record_len(&mut out, bstore_start);

    patch_record_len(&mut out, dgg_container_start);

    // ── OfficeArtWordDrawing for the Main Document ──
    out.push(DGGLBL_MAIN_DOCUMENT);
    let dg_container_start = out.len();
    write_record_header(&mut out, VERSION_CONTAINER, 0, RECORD_DG_CONTAINER, 0);

    // OfficeArtFDG: shape count (including the group shape) and next free spid.
    write_record_header(
        &mut out,
        VERSION_DG,
        DG_INSTANCE_FIRST_DRAWING,
        RECORD_DG,
        8,
    );
    out.extend_from_slice(&shape_count.to_le_bytes()); // csp
    out.extend_from_slice(&(FIRST_PICTURE_SHAPE_ID + shapes.len() as u32).to_le_bytes()); // spidCur

    // OfficeArtSpgrContainer: the drawing's group shape plus all shapes.
    let spgr_container_start = out.len();
    write_record_header(&mut out, VERSION_CONTAINER, 0, RECORD_SPGR_CONTAINER, 0);

    // Group shape: empty bounds rectangle and a group/patriarch FSP.
    let group_container_start = out.len();
    write_record_header(&mut out, VERSION_CONTAINER, 0, RECORD_SP_CONTAINER, 0);
    write_record_header(&mut out, VERSION_SPGR, 0, RECORD_SPGR, SPGR_PAYLOAD_LEN);
    out.extend_from_slice(&[0; SPGR_PAYLOAD_LEN as usize]);
    write_record_header(&mut out, VERSION_SP, 0, RECORD_SP, FSP_PAYLOAD_LEN);
    out.extend_from_slice(&GROUP_SHAPE_ID.to_le_bytes());
    out.extend_from_slice(&(SP_FLAG_GROUP | SP_FLAG_PATRIARCH).to_le_bytes());
    patch_record_len(&mut out, group_container_start);

    // One picture-frame shape per floating picture.
    for (index, shape) in shapes.iter().enumerate() {
        let shape_start = out.len();
        write_record_header(&mut out, VERSION_CONTAINER, 0, RECORD_SP_CONTAINER, 0);
        write_record_header(
            &mut out,
            VERSION_SP,
            SHAPE_TYPE_PICTURE_FRAME,
            RECORD_SP,
            FSP_PAYLOAD_LEN,
        );
        out.extend_from_slice(&shape.shape_id.to_le_bytes());
        out.extend_from_slice(&(SP_FLAG_HAVE_ANCHOR | SP_FLAG_HAVE_SHAPE_TYPE).to_le_bytes());
        // OfficeArtOPT: pib referencing this shape's BSE (1-based index).
        write_record_header(&mut out, VERSION_OPT, 1, RECORD_OPT, OPT_PAYLOAD_LEN);
        out.extend_from_slice(&OPT_PIB_BLIP_INDEX.to_le_bytes());
        out.extend_from_slice(&(index as u32 + OPT_PIB_FIRST_BSE).to_le_bytes());
        // ClientAnchor: index of this shape's anchor CP in the PlcfSpa.
        write_record_header(&mut out, VERSION_ATOM, 0, RECORD_CLIENT_ANCHOR, WORD_CLIENT_ANCHOR_LEN);
        out.extend_from_slice(&(index as u32).to_le_bytes());
        // OfficeArtClientData: present but unused.
        write_record_header(&mut out, VERSION_ATOM, 0, RECORD_CLIENT_DATA, CLIENT_DATA_LEN);
        out.extend_from_slice(&0u32.to_le_bytes());
        patch_record_len(&mut out, shape_start);
    }

    patch_record_len(&mut out, spgr_container_start);
    patch_record_len(&mut out, dg_container_start);

    out
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
    fn sniff_format_detects_png_and_jpeg() {
        assert_eq!(PictureFormat::sniff(&png_bytes()), Some(PictureFormat::Png));
        assert_eq!(PictureFormat::sniff(&jpeg_bytes()), Some(PictureFormat::Jpeg));
        assert_eq!(PictureFormat::sniff(b"not an image"), None);
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
        assert_eq!(picture.format(), PictureFormat::Png);
        assert_eq!(picture.width_twips(), 32 * TWIPS_PER_INCH / ASSUMED_DPI);
        assert_eq!(picture.height_twips(), 16 * TWIPS_PER_INCH / ASSUMED_DPI);
    }

    #[test]
    fn doc_picture_rejects_unknown_format_and_bad_dimensions() {
        assert!(DocPicture::new(b"garbage".to_vec()).is_err());
        assert!(DocPicture::from_parts(png_bytes(), 0, 100).is_err());
        assert!(
            DocPicture::from_parts(png_bytes(), 100, MAX_PICF_DIMENSION_TWIPS + 1).is_err()
        );
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
        write_picture_block(&picture, FIRST_PICTURE_SHAPE_ID, &mut block);

        // PICF header.
        let lcb = i32::from_le_bytes(block[0..4].try_into().unwrap()) as usize;
        assert_eq!(lcb, block.len());
        assert_eq!(i16::from_le_bytes(block[4..6].try_into().unwrap()), PICF_CB_HEADER);
        assert_eq!(i16::from_le_bytes(block[6..8].try_into().unwrap()), PICF_MM_SHAPE);
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
        assert_eq!((ver, record_type), (VERSION_BSE, RECORD_BSE));
        assert_eq!(inst, u16::from(MSO_BLIP_TYPE_PNG));
        let bse = &block[bse_offset + RECORD_HEADER_LEN..];
        assert_eq!(bse[0], MSO_BLIP_TYPE_PNG); // btWin32
        assert_eq!(bse[1], MSO_BLIP_TYPE_PNG); // btMacOS
        let blip_size = u32::from_le_bytes(bse[20..24].try_into().unwrap()) as usize;
        let c_ref = u32::from_le_bytes(bse[24..28].try_into().unwrap());
        let fo_delay = u32::from_le_bytes(bse[28..32].try_into().unwrap());
        assert_eq!(c_ref, 1);
        assert_eq!(fo_delay, BSE_NO_DELAY_STREAM); // embedded
        assert_eq!(blip_size + BSE_HEADER_LEN, len as usize);

        // Embedded BLIP record.
        let blip = &bse[BSE_HEADER_LEN..BSE_HEADER_LEN + blip_size];
        let (ver, inst, record_type, len) = parse_record_header(blip, 0);
        assert_eq!((ver, record_type), (VERSION_ATOM, RECORD_BLIP_PNG));
        assert_eq!(inst, BLIP_INSTANCE_PNG);
        assert_eq!(len as usize + RECORD_HEADER_LEN, blip_size);
        // The BLIP UID matches the BSE rgbUid, and the payload is byte-identical.
        assert_eq!(&blip[8..8 + BLIP_UID_LEN], &bse[2..2 + BLIP_UID_LEN]);
        assert_eq!(blip[8 + BLIP_UID_LEN], BLIP_EMBEDDED_MARKER);
        assert_eq!(&blip[9 + BLIP_UID_LEN..], picture.data());
    }

    #[test]
    fn picture_uids_are_distinct_per_shape() {
        let picture = DocPicture::new(png_bytes()).unwrap();
        assert_ne!(
            picture_uid(&picture, FIRST_PICTURE_SHAPE_ID),
            picture_uid(&picture, FIRST_PICTURE_SHAPE_ID + 1)
        );
    }

    #[cfg(feature = "imgconv")]
    #[test]
    fn picture_block_bse_parses_with_crate_reader() {
        use litchi_imgconv::{Blip, BlipStoreEntry, BlipType};

        let picture = DocPicture::new(jpeg_bytes()).unwrap();
        let mut block = Vec::new();
        write_picture_block(&picture, FIRST_PICTURE_SHAPE_ID, &mut block);

        let (_ver, _inst, _record_type, sp_len) = parse_record_header(&block, PICF_HEADER_LEN);
        let bse_offset = PICF_HEADER_LEN + RECORD_HEADER_LEN + sp_len as usize;
        let bse_payload = &block[bse_offset + RECORD_HEADER_LEN..];

        let bse = BlipStoreEntry::parse(bse_payload).unwrap();
        assert_eq!(bse.blip_type, BlipType::Jpeg);
        assert!(!bse.is_delay_loaded());

        let blip = Blip::parse(&bse_payload[BSE_HEADER_LEN..BSE_HEADER_LEN + bse.size as usize])
            .unwrap();
        assert_eq!(blip.blip_type(), Some(BlipType::Jpeg));
        assert_eq!(blip.picture_data(), picture.data());
    }

    // ── Floating pictures ──

    use crate::doc::parts::spa::{
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
                picture: png,
                position: &positions[0],
            },
            FloatingShapeInfo {
                anchor_cp: 30,
                shape_id: FIRST_PICTURE_SHAPE_ID + 1,
                picture: jpeg,
                position: &positions[1],
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
        let anchors = crate::doc::parts::spa::parse_plcf_spa(&plcf).unwrap();
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
        let dgg_container_end = start + RECORD_HEADER_LEN + len as usize;
        walk(data, start, dgg_container_end, &mut records);
        // Skip the OfficeArtWordDrawing dgglbl byte before the DgContainer.
        walk(data, dgg_container_end + 1, end, &mut records);
        records
    }

    #[test]
    fn dgg_info_layout_matches_ms_odraw() {
        let png = DocPicture::new(png_bytes()).unwrap();
        let jpeg = DocPicture::new(jpeg_bytes()).unwrap();
        let positions = sample_positions();
        let shapes = floating_shapes(&png, &jpeg, &positions);

        let dgg = build_dgg_info(&shapes, 3);
        let records = collect_records(&dgg, 0, dgg.len());

        // Top level: DggContainer, then dgglbl + DgContainer.
        let (off, ver, _inst, record_type, len) = records[0];
        assert_eq!((off, ver, record_type), (0, VERSION_CONTAINER, RECORD_DGG_CONTAINER));
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
            DGG_SINGLE_CLUSTER
        );
        // cspSaved = 2 shapes + group; cdgSaved = 1 drawing.
        assert_eq!(u32::from_le_bytes(dgg_payload[8..12].try_into().unwrap()), 3);
        assert_eq!(u32::from_le_bytes(dgg_payload[12..16].try_into().unwrap()), 1);

        // BStoreContainer holds one BSE per picture.
        let bstore = records
            .iter()
            .find(|record| record.3 == RECORD_BSTORE_CONTAINER)
            .unwrap();
        assert_eq!(bstore.2, 2);
        let bses: Vec<_> = records.iter().filter(|record| record.3 == RECORD_BSE).collect();
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
        let sps: Vec<_> = records.iter().filter(|record| record.3 == RECORD_SP).collect();
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
        let opts: Vec<_> = records.iter().filter(|record| record.3 == RECORD_OPT).collect();
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
            let value = u32::from_le_bytes(
                dgg[anchor.0 + 8..anchor.0 + 12].try_into().unwrap(),
            );
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

    #[cfg(feature = "imgconv")]
    #[test]
    fn dgg_info_bse_blips_parse_with_crate_reader() {
        use litchi_imgconv::{Blip, BlipStoreEntry, BlipType};

        let png = DocPicture::new(png_bytes()).unwrap();
        let jpeg = DocPicture::new(jpeg_bytes()).unwrap();
        let positions = sample_positions();
        let shapes = floating_shapes(&png, &jpeg, &positions);

        let dgg = build_dgg_info(&shapes, 2);
        let records = collect_records(&dgg, 0, dgg.len());
        let bses: Vec<_> = records.iter().filter(|record| record.3 == RECORD_BSE).collect();
        assert_eq!(bses.len(), 2);

        let expected = [
            (BlipType::Png, png.data()),
            (BlipType::Jpeg, jpeg.data()),
        ];
        for (bse_record, (blip_type, payload)) in bses.iter().zip(expected.iter()) {
            let bse_payload =
                &dgg[bse_record.0 + RECORD_HEADER_LEN..bse_record.0 + RECORD_HEADER_LEN + bse_record.4 as usize];
            let bse = BlipStoreEntry::parse(bse_payload).unwrap();
            assert_eq!(bse.blip_type, *blip_type);
            assert!(!bse.is_delay_loaded());
            let blip =
                Blip::parse(&bse_payload[BSE_HEADER_LEN..BSE_HEADER_LEN + bse.size as usize])
                    .unwrap();
            assert_eq!(blip.picture_data(), *payload);
        }
    }

    #[test]
    fn spid_max_rounds_up_to_next_cluster() {
        assert_eq!(spid_max(0), 2 * SHAPE_IDS_PER_CLUSTER);
        assert_eq!(spid_max(1), 2 * SHAPE_IDS_PER_CLUSTER);
        assert_eq!(
            spid_max(SHAPE_IDS_PER_CLUSTER),
            3 * SHAPE_IDS_PER_CLUSTER
        );
    }
}
