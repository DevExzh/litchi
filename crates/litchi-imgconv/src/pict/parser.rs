//! PICT picture framing.
//!
//! A picture is a `picSize` word, the `picFrame` rectangle and a sequence of
//! big-endian opcode records. Version 1 pictures use one-byte opcodes;
//! version 2 pictures use two-byte opcodes whose records start on even
//! offsets. Record lengths follow *Inside Macintosh: Imaging With
//! QuickDraw*, Appendix A ("Picture Opcodes"), including the lengths the
//! book defines for reserved opcodes so readers can skip them.

use litchi_core::error::{Error, Result};

use super::bitmap;
use super::types::{PictRect, read_fixed, read_i16, read_u16, read_u32};

/// PICT file version
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PictVersion {
    /// Version 1 (original format)
    V1,
    /// Version 2 (extended format)
    V2,
}

/// PICT file header
///
/// PICT files may have an optional 512-byte header (used by some applications)
/// followed by the actual PICT data.
#[derive(Debug, Clone)]
pub struct PictHeader {
    /// Version of the PICT file
    pub version: PictVersion,
    /// Picture frame (top, left, bottom, right)
    pub frame: (i16, i16, i16, i16),
    /// Whether this file has the 512-byte header
    pub has_512_header: bool,
}

impl PictHeader {
    /// Parse PICT header from data.
    ///
    /// Standalone PICT files start with a 512-byte application header;
    /// pictures stored in documents and on the clipboard do not. Both
    /// positions are scored, preferring the start of the data on a tie.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let candidate = detect_header(data).ok_or_else(|| {
            Error::ParseError("data does not start with a PICT picture header".into())
        })?;
        Ok(Self::from_candidate(&candidate))
    }

    fn from_candidate(candidate: &HeaderCandidate) -> Self {
        Self {
            version: candidate.version,
            frame: (
                candidate.frame.top,
                candidate.frame.left,
                candidate.frame.bottom,
                candidate.frame.right,
            ),
            has_512_header: candidate.offset == 512,
        }
    }

    /// Get width of the picture
    pub fn width(&self) -> i32 {
        i32::from(self.frame.3) - i32::from(self.frame.1)
    }

    /// Get height of the picture
    pub fn height(&self) -> i32 {
        i32::from(self.frame.2) - i32::from(self.frame.0)
    }
}

/// PICT opcode
///
/// PICT files are composed of opcodes that describe drawing operations
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PictOpcode {
    /// No operation
    Nop = 0x0000,
    /// Clipping region
    Clip = 0x0001,
    /// Background pattern
    BkPat = 0x0002,
    /// Text font
    TxFont = 0x0003,
    /// Text face
    TxFace = 0x0004,
    /// Text mode
    TxMode = 0x0005,
    /// Version opcode
    Version = 0x0011,
    /// Extended version 2 header
    HeaderOp = 0x0C00,
    /// End of picture
    EndPic = 0x00FF,
    /// Direct bits rect (includes bitmap data)
    DirectBitsRect = 0x009A,
    /// Packed direct bits rect
    PackedDirectBitsRect = 0x009B,
    /// Compressed QuickTime image
    CompressedQuickTime = 0x8200,
}

impl PictOpcode {
    /// Create from u16 value
    pub fn from_u16(value: u16) -> Option<Self> {
        match value {
            0x0000 => Some(Self::Nop),
            0x0001 => Some(Self::Clip),
            0x0002 => Some(Self::BkPat),
            0x0003 => Some(Self::TxFont),
            0x0004 => Some(Self::TxFace),
            0x0005 => Some(Self::TxMode),
            0x0011 => Some(Self::Version),
            0x0C00 => Some(Self::HeaderOp),
            0x00FF => Some(Self::EndPic),
            0x009A => Some(Self::DirectBitsRect),
            0x009B => Some(Self::PackedDirectBitsRect),
            0x8200 => Some(Self::CompressedQuickTime),
            _ => None,
        }
    }
}

/// PICT record
#[derive(Debug, Clone)]
pub struct PictRecord {
    /// Opcode
    pub opcode: u16,
    /// Record data
    pub data: Vec<u8>,
}

/// A non-fatal fact discovered while framing a picture.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PictParserWarning {
    /// Stable diagnostic code; `noncanonical-*` codes do not change the
    /// rendered picture.
    pub code: &'static str,
    /// Human-readable explanation.
    pub message: String,
}

/// Contents of the version 2 `HeaderOp` record (opcode `0x0C00`).
#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct HeaderOp {
    /// `-2` for an extended version 2 picture, `-1` for a basic one.
    pub version: i16,
    /// Native horizontal resolution in dots per inch (extended version 2).
    pub h_res: f64,
    /// Native vertical resolution in dots per inch (extended version 2).
    pub v_res: f64,
    /// Picture bounds at the native resolution (extended version 2).
    pub source_rect: PictRect,
}

/// PICT file parser
pub struct PictParser {
    /// PICT header
    pub header: PictHeader,
    /// The picture's records in order, except picture comments
    /// (`ShortComment` and `LongComment`), which carry application data that
    /// QuickDraw playback ignores.
    pub records: Vec<PictRecord>,
    /// Raw PICT data
    data: Vec<u8>,
    /// Offset of each record's opcode from the start of the picture data.
    offsets: Vec<usize>,
    header_op: Option<HeaderOp>,
    warnings: Vec<PictParserWarning>,
}

impl PictParser {
    /// Create a new PICT parser from raw data under the caller's resource
    /// policy. The input and record ceilings are checked before any owned
    /// copies or record-vector capacity are created.
    ///
    /// A picture that ends without an `EndPic` opcode, or inside a record,
    /// keeps every complete record and reports the condition through
    /// [`PictParser::warnings`].
    pub fn new(data: &[u8], limits: &crate::Limits) -> Result<Self> {
        if limits.max_records == 0 || limits.max_uncompressed_bytes == 0 {
            return Err(Error::ParseError(
                "PICT parser limits must be greater than zero".into(),
            ));
        }
        if data.len() > limits.max_uncompressed_bytes {
            return Err(Error::ParseError(format!(
                "PICT input exceeds limit {} bytes",
                limits.max_uncompressed_bytes
            )));
        }
        let candidate = detect_header(data).ok_or_else(|| {
            Error::ParseError("data does not start with a PICT picture header".into())
        })?;
        let header = PictHeader::from_candidate(&candidate);
        let mut warnings = Vec::new();
        if candidate.offset != 0 && candidate.offset != 512 {
            warnings.push(PictParserWarning {
                code: "noncanonical-pict-leading-bytes",
                message: format!(
                    "picture data starts after {} unrelated bytes",
                    candidate.offset
                ),
            });
        }

        let picture = &data[candidate.offset..];
        let two_byte_opcodes = header.version == PictVersion::V2;
        let mut offset = 10usize;
        let mut records = Vec::new();
        let mut offsets = Vec::new();
        let mut header_op = None;
        let mut ended = false;
        let mut stopped = false;
        while offset < picture.len() {
            let (opcode, data_start) = if two_byte_opcodes {
                match read_u16(picture, offset) {
                    Some(opcode) => (opcode, offset + 2),
                    // A lone pad byte at the end of the input.
                    None => break,
                }
            } else {
                (u16::from(picture[offset]), offset + 1)
            };
            if opcode == 0x00FF {
                offset = data_start;
                ended = true;
                break;
            }
            let comment = matches!(opcode, 0x00A0 | 0x00A1);
            if !comment && records.len() >= limits.max_records {
                return Err(Error::ParseError(format!(
                    "PICT record count exceeds limit {}",
                    limits.max_records
                )));
            }
            let length = match record_data_len(opcode, picture, data_start) {
                Ok(length) => length,
                Err(failure) => {
                    warnings.push(framing_warning(opcode, failure, records.len()));
                    stopped = true;
                    break;
                },
            };
            let data_end = data_start + length;
            if !comment {
                let record_data = picture[data_start..data_end].to_vec();
                if opcode == 0x0C00 {
                    header_op = header_op.or_else(|| parse_header_op(&record_data));
                }
                records.push(PictRecord {
                    opcode,
                    data: record_data,
                });
                offsets.push(offset);
            }
            offset = data_end;
            if two_byte_opcodes && offset % 2 == 1 {
                offset += 1;
            }
        }
        if !ended && !stopped {
            // picSize holds the length of the picture data, of which version 2
            // keeps only the low 16 bits. A picture that ends there, with or
            // without room for its end-of-picture opcode, lost nothing but
            // that opcode; one that ends earlier was cut short.
            let consumed = offset.min(picture.len());
            let declared = read_u16(picture, 0).map(usize::from);
            let end_opcode = if two_byte_opcodes { 2 } else { 1 };
            let complete = [consumed, consumed + end_opcode].iter().any(|&length| {
                let length = if two_byte_opcodes {
                    length & 0xFFFF
                } else {
                    length
                };
                declared == Some(length)
            });
            warnings.push(if complete {
                PictParserWarning {
                    code: "noncanonical-pict-missing-end",
                    message: format!(
                        "picture ends after {} records without an end-of-picture opcode",
                        records.len()
                    ),
                }
            } else {
                PictParserWarning {
                    code: "pict-truncated",
                    message: format!(
                        "picture ends after {} records and {consumed} bytes without an end-of-picture opcode; its header declares {} bytes",
                        records.len(),
                        declared.unwrap_or(0)
                    ),
                }
            });
        } else if ended && offset < picture.len() {
            warnings.push(PictParserWarning {
                code: "noncanonical-pict-trailing-data",
                message: format!(
                    "{} bytes follow the end-of-picture opcode",
                    picture.len() - offset
                ),
            });
        }

        Ok(Self {
            header,
            records,
            data: data.to_vec(),
            offsets,
            header_op,
            warnings,
        })
    }

    /// Non-fatal facts discovered while framing the picture.
    pub fn warnings(&self) -> &[PictParserWarning] {
        &self.warnings
    }

    /// Offset of record `index`'s opcode from the start of the picture data.
    pub(super) fn record_offset(&self, index: usize) -> Option<usize> {
        self.offsets.get(index).copied()
    }

    /// The version 2 header record, when the picture has one.
    pub(super) fn header_op(&self) -> Option<HeaderOp> {
        self.header_op
    }

    /// The picture frame.
    pub(super) fn frame(&self) -> PictRect {
        PictRect {
            top: self.header.frame.0,
            left: self.header.frame.1,
            bottom: self.header.frame.2,
            right: self.header.frame.3,
        }
    }

    /// Get the raw PICT data
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Get width
    pub fn width(&self) -> i32 {
        self.header.width()
    }

    /// Get height
    pub fn height(&self) -> i32 {
        self.header.height()
    }

    /// Get aspect ratio
    pub fn aspect_ratio(&self) -> f64 {
        let w = self.width() as f64;
        let h = self.height() as f64;
        if h == 0.0 { 1.0 } else { w / h }
    }
}

/// Why a record could not be framed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FramingFailure {
    /// The input ends inside the record.
    Truncated,
    /// The record's own length fields are inconsistent.
    Malformed,
}

fn framing_warning(opcode: u16, failure: FramingFailure, records: usize) -> PictParserWarning {
    match failure {
        FramingFailure::Truncated => PictParserWarning {
            code: "pict-truncated",
            message: format!(
                "picture ends inside opcode 0x{opcode:04X}; drew the {records} complete records"
            ),
        },
        FramingFailure::Malformed => PictParserWarning {
            code: "pict-malformed-record",
            message: format!(
                "opcode 0x{opcode:04X} has inconsistent length fields; drew the {records} records before it"
            ),
        },
    }
}

fn parse_header_op(data: &[u8]) -> Option<HeaderOp> {
    let version = read_i16(data, 0)?;
    if version == -2 {
        Some(HeaderOp {
            version,
            h_res: read_fixed(data, 4)?,
            v_res: read_fixed(data, 8)?,
            source_rect: PictRect::read(data, 12)?,
        })
    } else {
        Some(HeaderOp {
            version,
            h_res: 72.0,
            v_res: 72.0,
            source_rect: PictRect::default(),
        })
    }
}

/// Returns the length of the data that follows `opcode` at `offset`.
fn record_data_len(
    opcode: u16,
    picture: &[u8],
    offset: usize,
) -> std::result::Result<usize, FramingFailure> {
    use FramingFailure::{Malformed, Truncated};
    let word = |at: usize| read_u16(picture, at).map(usize::from).ok_or(Truncated);
    let long = |at: usize| {
        read_u32(picture, at)
            .and_then(|value| usize::try_from(value).ok())
            .ok_or(Truncated)
    };
    let byte = |at: usize| picture.get(at).copied().map(usize::from).ok_or(Truncated);
    let length = match opcode {
        0x0000
        | 0x0017..=0x0019
        | 0x001C
        | 0x001E
        | 0x0038..=0x003F
        | 0x0048..=0x004F
        | 0x0058..=0x005F
        | 0x0078..=0x007F
        | 0x0088..=0x008F
        | 0x00B0..=0x00CF
        | 0x8000..=0x80FF => 0,
        // Clip, regions and polygons: the size word counts the whole record.
        0x0001 | 0x0070..=0x0077 | 0x0080..=0x0087 => {
            let size = word(offset)?;
            if size < 10 {
                return Err(Malformed);
            }
            size
        },
        0x0004 | 0x0011 => 1,
        0x0003 | 0x0005 | 0x0008 | 0x000D | 0x0015 | 0x0016 | 0x0023 | 0x00A0 => 2,
        0x0006 | 0x0007 | 0x000B | 0x000C | 0x000E | 0x000F | 0x0021 | 0x0068..=0x006F => 4,
        0x001A | 0x001B | 0x001D | 0x001F | 0x0022 => 6,
        0x0002
        | 0x0009
        | 0x000A
        | 0x0010
        | 0x0020
        | 0x0030..=0x0037
        | 0x0040..=0x0047
        | 0x0050..=0x0057 => 8,
        0x0060..=0x0067 => 12,
        0x0012..=0x0014 => {
            let (_, length) =
                bitmap::parse_pixpat(picture.get(offset..).unwrap_or(&[])).ok_or(Truncated)?;
            length
        },
        0x0028 => 5 + byte(offset + 4)?,
        0x0029 | 0x002A => 2 + byte(offset + 1)?,
        0x002B => 3 + byte(offset + 2)?,
        0x0024..=0x002F | 0x0092..=0x0097 | 0x009C..=0x009F | 0x00A2..=0x00AF => 2 + word(offset)?,
        0x0090 | 0x0091 | 0x0098..=0x009B => {
            let (_, length) =
                bitmap::parse_bits_record(picture.get(offset..).unwrap_or(&[]), opcode)
                    .ok_or(Truncated)?;
            length
        },
        0x00A1 => 4 + word(offset + 2)?,
        0x00D0..=0x00FE | 0x8100..=0xFFFF => 4 + long(offset)?,
        0x0100..=0x7FFF => usize::from(opcode >> 8) * 2,
        // `EndPic` is handled by the caller.
        0x00FF => 0,
    };
    let end = offset.checked_add(length).ok_or(Truncated)?;
    if end > picture.len() {
        return Err(Truncated);
    }
    Ok(length)
}

#[derive(Debug, Clone, Copy)]
struct HeaderCandidate {
    offset: usize,
    version: PictVersion,
    frame: PictRect,
    confidence: i32,
}

/// Locates the picture header.
///
/// Offsets 0 and 512 are scored by the plausibility of the frame and by the
/// version opcode that must follow it, allowing the leading `Nop` opcodes
/// some writers insert. When neither qualifies, a bounded prefix is scanned
/// for an exact header, which recovers pictures preceded by unrelated bytes.
fn detect_header(data: &[u8]) -> Option<HeaderCandidate> {
    let at_start = header_candidate(data, 0, 9);
    let after_file_header = header_candidate(data, 512, 9);
    let best = match (at_start, after_file_header) {
        (Some(start), Some(file)) if file.confidence > start.confidence => Some(file),
        (Some(start), _) => Some(start),
        (None, file) => file,
    };
    if best.is_some() {
        return best;
    }
    (1..=1024usize)
        .filter(|&offset| offset != 512)
        .filter_map(|offset| header_candidate(data, offset, 0))
        .find(|candidate| candidate.confidence == FULL_CONFIDENCE)
}

const FULL_CONFIDENCE: i32 = 20;

fn header_candidate(data: &[u8], offset: usize, max_nops: usize) -> Option<HeaderCandidate> {
    let picture = data.get(offset..)?;
    let frame = PictRect::read(picture, 2)?;
    if frame.left > frame.right || frame.top > frame.bottom {
        return None;
    }
    let mut confidence = FULL_CONFIDENCE;
    let extreme = |value: i16| !(-2048..=2048).contains(&value);
    if [frame.top, frame.left, frame.bottom, frame.right]
        .into_iter()
        .any(extreme)
        || !frame.is_nonempty()
    {
        confidence -= 3;
    } else if frame.width() < 8 || frame.height() < 8 {
        confidence -= 1;
    }
    let after_frame = picture.get(10..)?;
    // Version 1 `Nop` opcodes are single zero bytes; version 2 `Nop` opcodes
    // are zero words preceding the zero high byte of the `0x0011` opcode.
    let zeros = after_frame
        .iter()
        .take(max_nops * 2 + 2)
        .take_while(|&&byte| byte == 0)
        .count();
    let (version, nops) = match after_frame.get(zeros..zeros + 2)? {
        [0x11, 0x01] => (PictVersion::V1, zeros),
        [0x11, 0x02] if zeros % 2 == 1 && after_frame.get(zeros + 2) == Some(&0xFF) => {
            (PictVersion::V2, zeros / 2)
        },
        _ => return None,
    };
    if nops > max_nops {
        return None;
    }
    confidence -= i32::try_from(nops).unwrap_or(i32::MAX);
    Some(HeaderCandidate {
        offset,
        version,
        frame,
        confidence,
    })
}

#[cfg(test)]
pub(super) mod tests {
    use super::*;

    pub(in crate::pict) fn version2_picture(frame: [i16; 4], records: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_be_bytes());
        for value in frame {
            data.extend_from_slice(&value.to_be_bytes());
        }
        data.extend_from_slice(&[0x00, 0x11, 0x02, 0xFF]);
        data.extend_from_slice(&[0x0C, 0x00, 0xFF, 0xFE, 0, 0]);
        data.extend_from_slice(&0x0048_0000u32.to_be_bytes());
        data.extend_from_slice(&0x0048_0000u32.to_be_bytes());
        for value in frame {
            data.extend_from_slice(&value.to_be_bytes());
        }
        data.extend_from_slice(&[0; 4]);
        data.extend_from_slice(records);
        data
    }

    fn version1_picture(records: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_be_bytes());
        for value in [0i16, 0, 10, 10] {
            data.extend_from_slice(&value.to_be_bytes());
        }
        data.extend_from_slice(&[0x11, 0x01]);
        data.extend_from_slice(records);
        data
    }

    #[test]
    fn test_pict_opcodes() {
        assert_eq!(PictOpcode::from_u16(0x00FF), Some(PictOpcode::EndPic));
        assert_eq!(PictOpcode::from_u16(0x0011), Some(PictOpcode::Version));
    }

    #[test]
    fn version2_records_are_word_aligned() {
        let mut records = Vec::new();
        records.extend_from_slice(&0x0004u16.to_be_bytes()); // TxFace
        records.extend_from_slice(&[1, 0]); // data and pad byte
        records.extend_from_slice(&0x0029u16.to_be_bytes()); // DHText
        records.extend_from_slice(&[3, 2, b'h', b'i']);
        records.extend_from_slice(&0x00FFu16.to_be_bytes());
        let data = version2_picture([0, 0, 100, 100], &records);
        let parser = PictParser::new(&data, &crate::Limits::default()).unwrap();
        let opcodes: Vec<u16> = parser.records.iter().map(|record| record.opcode).collect();
        assert_eq!(opcodes, [0x0011, 0x0C00, 0x0004, 0x0029]);
        assert_eq!(parser.records[2].data, [1]);
        assert_eq!(parser.records[3].data, [3, 2, b'h', b'i']);
        assert!(parser.warnings().is_empty());
        let header_op = parser.header_op().unwrap();
        assert_eq!(header_op.version, -2);
        assert_eq!(header_op.h_res, 72.0);
    }

    #[test]
    fn reserved_opcodes_use_their_defined_lengths() {
        let mut records = Vec::new();
        records.extend_from_slice(&0x0024u16.to_be_bytes()); // length word
        records.extend_from_slice(&[0, 2, 9, 9]);
        records.extend_from_slice(&0x00B5u16.to_be_bytes()); // no data
        records.extend_from_slice(&0x00D3u16.to_be_bytes()); // 32-bit length
        records.extend_from_slice(&[0, 0, 0, 2, 7, 7]);
        records.extend_from_slice(&0x0300u16.to_be_bytes()); // 2 * high byte
        records.extend_from_slice(&[0; 6]);
        records.extend_from_slice(&0x8012u16.to_be_bytes()); // no data
        records.extend_from_slice(&0x8300u16.to_be_bytes()); // 32-bit length
        records.extend_from_slice(&[0, 0, 0, 1, 5, 0]);
        records.extend_from_slice(&0x00A1u16.to_be_bytes()); // LongComment
        records.extend_from_slice(&[0, 100, 0, 2, 1, 2]);
        records.extend_from_slice(&0x00FFu16.to_be_bytes());
        let data = version2_picture([0, 0, 10, 10], &records);
        let parser = PictParser::new(&data, &crate::Limits::default()).unwrap();
        // The LongComment is framed but not kept.
        let lengths: Vec<usize> = parser.records[2..].iter().map(|r| r.data.len()).collect();
        assert_eq!(lengths, [4, 0, 6, 6, 0, 5]);
        assert_eq!(parser.record_offset(2), Some(40));
        assert!(parser.warnings().is_empty());
    }

    #[test]
    fn version1_uses_one_byte_opcodes_without_padding() {
        let mut records = vec![0x04, 0x01]; // TxFace
        records.extend_from_slice(&[0x0D, 0, 12]); // TxSize
        records.extend_from_slice(&[0x31, 0, 1, 0, 1, 0, 5, 0, 5]); // paintRect
        records.push(0xFF);
        let data = version1_picture(&records);
        let parser = PictParser::new(&data, &crate::Limits::default()).unwrap();
        assert_eq!(parser.header.version, PictVersion::V1);
        let opcodes: Vec<u16> = parser.records.iter().map(|record| record.opcode).collect();
        assert_eq!(opcodes, [0x11, 0x04, 0x0D, 0x31]);
        assert!(parser.warnings().is_empty());
    }

    #[test]
    fn version1_tolerates_a_leading_nop() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_be_bytes());
        for value in [0i16, 0, 10, 10] {
            data.extend_from_slice(&value.to_be_bytes());
        }
        data.extend_from_slice(&[0x00, 0x11, 0x01, 0xFF]);
        let parser = PictParser::new(&data, &crate::Limits::default()).unwrap();
        assert_eq!(parser.header.version, PictVersion::V1);
        assert_eq!(parser.records.len(), 2);
    }

    #[test]
    fn file_header_is_detected_and_short_inputs_do_not_panic() {
        let picture = version2_picture([0, 0, 20, 30], &0x00FFu16.to_be_bytes());
        let mut file = vec![0u8; 512];
        file.extend_from_slice(&picture);
        let header = PictHeader::parse(&file).unwrap();
        assert!(header.has_512_header);
        assert_eq!(header.width(), 30);
        assert!(!PictHeader::parse(&picture).unwrap().has_512_header);
        for length in 0..file.len().min(530) {
            let _ = PictHeader::parse(&file[..length]);
            let _ = PictParser::new(&file[..length], &crate::Limits::default());
        }
    }

    #[test]
    fn leading_garbage_is_skipped_and_reported() {
        let mut data = vec![0xAB; 7];
        data.extend_from_slice(&version2_picture([0, 0, 20, 30], &0x00FFu16.to_be_bytes()));
        let parser = PictParser::new(&data, &crate::Limits::default()).unwrap();
        assert_eq!(parser.records.len(), 2);
        assert_eq!(parser.warnings()[0].code, "noncanonical-pict-leading-bytes");
        assert_eq!(parser.warnings().len(), 1);
    }

    #[test]
    fn missing_end_is_noncanonical_only_at_the_declared_size() {
        let mut records = Vec::new();
        records.extend_from_slice(&0x0003u16.to_be_bytes());
        records.extend_from_slice(&1u16.to_be_bytes());
        let mut data = version2_picture([0, 0, 10, 10], &records);
        let parser = PictParser::new(&data, &crate::Limits::default()).unwrap();
        assert_eq!(parser.records.len(), 3);
        // The picture ends before the size its header declares.
        assert_eq!(parser.warnings()[0].code, "pict-truncated");
        // It ends where its header says, or just before the end opcode.
        for declared in [data.len(), data.len() + 2] {
            data[..2].copy_from_slice(&u16::try_from(declared).unwrap().to_be_bytes());
            let parser = PictParser::new(&data, &crate::Limits::default()).unwrap();
            assert_eq!(parser.warnings()[0].code, "noncanonical-pict-missing-end");
        }
        // Version 2 keeps only the low 16 bits of the size.
        let mut long = records.clone();
        while long.len() < 0x1_0000 {
            long.extend_from_slice(&0u16.to_be_bytes()); // NOP
        }
        let mut data = version2_picture([0, 0, 10, 10], &long);
        let declared = u16::try_from(data.len() & 0xFFFF).unwrap();
        data[..2].copy_from_slice(&declared.to_be_bytes());
        let parser = PictParser::new(&data, &crate::Limits::default()).unwrap();
        assert_eq!(parser.warnings()[0].code, "noncanonical-pict-missing-end");

        records.extend_from_slice(&0x0020u16.to_be_bytes()); // Line needs 8 bytes
        records.extend_from_slice(&[0, 1]);
        let data = version2_picture([0, 0, 10, 10], &records);
        let parser = PictParser::new(&data, &crate::Limits::default()).unwrap();
        assert_eq!(parser.records.len(), 3);
        assert_eq!(parser.warnings().len(), 1);
        assert_eq!(parser.warnings()[0].code, "pict-truncated");
    }

    #[test]
    fn trailing_data_after_end_is_noncanonical() {
        let mut records = 0x00FFu16.to_be_bytes().to_vec();
        records.extend_from_slice(&[0; 6]);
        let data = version2_picture([0, 0, 10, 10], &records);
        let parser = PictParser::new(&data, &crate::Limits::default()).unwrap();
        assert_eq!(parser.warnings()[0].code, "noncanonical-pict-trailing-data");
    }

    #[test]
    fn parser_requires_and_enforces_caller_limits() {
        let data = version2_picture([0, 0, 10, 10], &0x00FFu16.to_be_bytes());

        let mut limits = crate::Limits {
            max_uncompressed_bytes: data.len(),
            max_records: 2,
            ..crate::Limits::default()
        };
        assert_eq!(PictParser::new(&data, &limits).unwrap().records.len(), 2);

        limits.max_uncompressed_bytes = data.len() - 1;
        assert!(PictParser::new(&data, &limits).is_err());

        limits.max_uncompressed_bytes = data.len();
        limits.max_records = 1;
        assert!(PictParser::new(&data, &limits).is_err());

        limits.max_records = 0;
        assert!(PictParser::new(&data, &limits).is_err());
    }
}
