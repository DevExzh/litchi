//! Typed Word shape models projected from OfficeArt.

use litchi_odraw::{Container, Record, RecordKind};

use litchi_odraw::shape::Shape as OfficeArtShape;
pub use litchi_odraw::shape::{Bounds, Kind};

use std::io;

/// One unknown OfficeArt record retained from a Word shape container.
///
/// OfficeArt is extensible and [MS-ODRAW] allows producers to add records that
/// a particular reader does not understand. The record is stored as its exact
/// wire representation, including the eight-byte record header, so a later
/// layer can inspect or replay it without a lossy decode/re-encode cycle.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnknownRecord {
    bytes: Box<[u8]>,
}

impl UnknownRecord {
    pub(crate) fn from_record(record: &Record<'_>) -> Self {
        let version_instance = u16::from(record.version()) | (record.instance() << 4);
        let mut bytes = Vec::with_capacity(8 + record.data().len());
        bytes.extend_from_slice(&version_instance.to_le_bytes());
        bytes.extend_from_slice(&record.raw_kind().to_le_bytes());
        bytes.extend_from_slice(&record.len().to_le_bytes());
        bytes.extend_from_slice(record.data());
        Self {
            bytes: bytes.into_boxed_slice(),
        }
    }

    /// Returns the exact OfficeArt record bytes, including its header.
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Returns the record kind value exactly as it appeared on the wire.
    pub fn raw_kind(&self) -> u16 {
        u16::from_le_bytes([self.bytes[2], self.bytes[3]])
    }

    /// Returns the four-bit record version.
    pub fn version(&self) -> u8 {
        (u16::from_le_bytes([self.bytes[0], self.bytes[1]]) & 0x000F) as u8
    }

    /// Returns the twelve-bit record instance.
    pub fn instance(&self) -> u16 {
        u16::from_le_bytes([self.bytes[0], self.bytes[1]]) >> 4
    }

    /// Returns the record payload without its header.
    pub fn data(&self) -> &[u8] {
        &self.bytes[8..]
    }
}

/// Shape information extracted from a Word document's OfficeArt drawing.
///
/// The model is owned because textbox text is resolved from separate Word
/// stories after OfficeArt parsing. Shape children and unknown-record payloads
/// are consequently independent of the OLE stream borrowed during decoding.
#[derive(Debug, Clone)]
pub struct Shape {
    /// Format-neutral OfficeArt shape family.
    pub shape_type: Kind,
    /// OfficeArt shape identifier (`spid`).
    pub shape_id: u32,
    /// Text content extracted from the shape's Word textbox story, if any.
    pub text: Option<String>,
    /// Whether this shape is a group shape.
    pub is_group: bool,
    /// Child shapes, in OfficeArt source order.
    pub children: Vec<Shape>,
    /// Fill color as `(R, G, B)`, when an explicit `fillColor` is present.
    pub fill_color: Option<(u8, u8, u8)>,
    /// Line color as `(R, G, B)`, when an explicit `lineColor` is present.
    pub line_color: Option<(u8, u8, u8)>,
    /// The raw MSOSPT preset-geometry value ([MS-ODRAW] 2.4.24), when present.
    pub native_shape_type: Option<u16>,
    /// Typed `FSPGR` coordinate-space bounds for group shapes.
    pub group_bounds: Option<Bounds>,
    /// Unknown records directly owned by this shape's OfficeArt container.
    pub unknown_records: Vec<UnknownRecord>,
    pub(crate) text_link: bool,
}

impl Shape {
    /// Project a host-neutral OfficeArt shape into the Word drawing facade.
    pub(crate) fn from_office_art(shape: &OfficeArtShape<'_>) -> io::Result<Self> {
        // In Word, OfficeArtClientTextbox contains only a TXID into the Word
        // textbox story. Text itself is resolved by `Document::text_boxes` and
        // cannot be decoded from OfficeArt bytes in isolation.
        let text_link = if let Some(textbox) = shape.textbox() {
            let _: &[u8; 4] = textbox
                .data()
                .try_into()
                .map_err(|_| invalid_data("Word OfficeArtClientTextbox payload is not one TXID"))?;
            true
        } else {
            false
        };

        let children = shape
            .children()
            .iter()
            .map(Self::from_office_art)
            .collect::<io::Result<_>>()?;

        Ok(Self {
            shape_type: shape.kind(),
            shape_id: shape.id(),
            text: None,
            is_group: matches!(shape.kind(), Kind::Group | Kind::Table),
            children,
            fill_color: shape.props().get_fill_color(),
            line_color: shape.props().get_line_color(),
            native_shape_type: Some(shape.native_kind().raw()),
            group_bounds: shape.group_bounds().copied(),
            unknown_records: collect_unknown_records(shape)?,
            text_link,
        })
    }
}

fn collect_unknown_records(shape: &OfficeArtShape<'_>) -> io::Result<Vec<UnknownRecord>> {
    let mut records = Vec::new();
    collect_unknown_children(shape.meta(), &mut records)?;

    // For a group shape, `container` wraps the group-header shape and its
    // child shapes. Scan its direct unknown siblings as well, but avoid
    // visiting the same SpContainer twice for ordinary shapes.
    let container = shape.container();
    let meta = shape.meta();
    let same_record = container.record().data().as_ptr() == meta.record().data().as_ptr()
        && container.record().data().len() == meta.record().data().len();
    if !same_record {
        collect_unknown_children(container, &mut records)?;
    }
    Ok(records)
}

fn collect_unknown_children(
    container: &Container<'_>,
    records: &mut Vec<UnknownRecord>,
) -> io::Result<()> {
    for child in container.children() {
        let child = child.map_err(invalid_data)?;
        if matches!(child.kind(), RecordKind::Unknown(_)) {
            records.push(UnknownRecord::from_record(&child));
        }
    }
    Ok(())
}

fn invalid_data(error: impl ToString) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidData, error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_odraw::Record;

    #[test]
    fn bounds_keep_signed_wire_coordinates() {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0x0001_u16.to_le_bytes());
        bytes.extend_from_slice(&0xF009_u16.to_le_bytes());
        bytes.extend_from_slice(&16_u32.to_le_bytes());
        for coordinate in [-120_i32, 40, 960, 880] {
            bytes.extend_from_slice(&coordinate.to_le_bytes());
        }
        let (record, consumed) = Record::parse(&bytes, 0).expect("valid FSPGR record");
        assert_eq!(consumed, bytes.len());
        let bounds = Bounds::from_record(&record).expect("typed FSPGR bounds");
        assert_eq!(bounds, Bounds::new(-120, 40, 960, 880));
        assert_eq!(bounds.width(), Some(1080));
        assert_eq!(bounds.height(), Some(840));
    }

    #[test]
    fn unknown_record_round_trips_exact_wire_bytes() {
        let bytes = [
            0x37, 0x12, // version 7, instance 0x123
            0x34, 0xF1, // producer-defined record kind
            0x03, 0x00, 0x00, 0x00, // payload length
            0xA5, 0x00, 0xFE,
        ];
        let (record, consumed) = Record::parse(&bytes, 0).expect("valid unknown record");
        assert_eq!(consumed, bytes.len());
        let unknown = UnknownRecord::from_record(&record);
        assert_eq!(unknown.bytes(), bytes);
        assert_eq!(unknown.raw_kind(), 0xF134);
        assert_eq!(unknown.version(), 7);
        assert_eq!(unknown.instance(), 0x123);
        assert_eq!(unknown.data(), &bytes[8..]);
    }
}
