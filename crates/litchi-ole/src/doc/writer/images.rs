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

/// Size of the PICF picture descriptor header ([MS-DOC] 2.9.161).
const PICF_HEADER_LEN: usize = 0x44;
/// PICF `cbHeader` value, matching [`PICF_HEADER_LEN`].
const PICF_CB_HEADER: i16 = 0x44;
/// PICF `mfp.mm` value Word writes for OfficeArt shape pictures.
const PICF_MM_SHAPE: i16 = 0x64;
/// Unscaled picture factor in permille (1000 = 100%).
const SCALE_100_PERCENT: i16 = 1000;

// Escher record types ([MS-ODRAW] 2.3).
const RECORD_SP_CONTAINER: u16 = 0xF004;
const RECORD_BSE: u16 = 0xF007;
const RECORD_SP: u16 = 0xF00A;
const RECORD_OPT: u16 = 0xF00B;
const RECORD_CLIENT_ANCHOR: u16 = 0xF010;
const RECORD_BLIP_JPEG: u16 = 0xF01D;
const RECORD_BLIP_PNG: u16 = 0xF01E;

// Escher record versions.
const VERSION_CONTAINER: u16 = 0xF;
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

    // OfficeArtFBSE with the embedded OfficeArtBlip record.
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

    debug_assert_eq!(out.len() - block_start, lcb as usize);
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
}
