//! Test-only Prost oracle: decode an IWA message payload with its generated
//! schema and compare two payloads on every known field.
//!
//! A payload is canonicalised by decoding it with its Prost type and
//! re-encoding it: Prost drops unknown fields and writes known fields in
//! field-number order, so two payloads are equal on every known field exactly
//! when their canonical encodings are equal. Message type numbers overlap
//! between applications, so the table is chosen per format. A case whose
//! changed objects carry a type not listed here fails with a message naming
//! the type; register it with [`KnownFields::with`].

use std::collections::{BTreeMap, BTreeSet};
use std::fmt::Debug;

use litchi_iwa_protos::{kn, tn, tp, tsa, tsce, tsd, tsp, tss, tst, tswp};
use prost::Message;

use crate::harness::Format;

/// One payload decoded through its schema.
pub struct Decoded {
    /// Known fields only, re-encoded in field-number order.
    pub canonical: Vec<u8>,
    /// Pretty `Debug` rendering of the decoded message.
    pub debug: String,
    /// Top-level field numbers present in the payload but not in the schema.
    pub unknown_fields: BTreeSet<u32>,
}

/// Decoder of one message type.
pub type Decoder = fn(&[u8]) -> Result<Decoded, prost::DecodeError>;

/// Message type number to schema name and decoder, for one format.
pub struct KnownFields {
    decoders: BTreeMap<u32, (&'static str, Decoder)>,
}

impl KnownFields {
    /// The schemas registered for `format` (shared TSP/TST/TSWP/TSS types
    /// plus the format's own document types).
    pub fn for_format(format: Format) -> Self {
        let mut known = Self {
            decoders: BTreeMap::new(),
        };
        known = known
            .with(
                401,
                "TSS.StylesheetArchive",
                decode_as::<tss::StylesheetArchive>,
            )
            .with(
                2_001,
                "TSWP.StorageArchive",
                decode_as::<tswp::StorageArchive>,
            )
            .with(
                6_001,
                "TST.TableModelArchive",
                decode_as::<tst::TableModelArchive>,
            )
            .with(6_002, "TST.Tile", decode_as::<tst::Tile>)
            .with(6_005, "TST.TableDataList", decode_as::<tst::TableDataList>)
            .with(
                6_006,
                "TST.HeaderStorageBucket",
                decode_as::<tst::HeaderStorageBucket>,
            )
            .with(
                6_011,
                "TST.TableDataListSegment",
                decode_as::<tst::TableDataListSegment>,
            )
            .with(
                11_006,
                "TSP.PackageMetadata",
                decode_as::<tsp::PackageMetadata>,
            )
            // The calculation engine is shared by the three applications
            // (record 0376a: Keynote and Pages formula writes).
            .with(
                4_000,
                "TSCE.CalculationEngineArchive",
                decode_as::<tsce::CalculationEngineArchive>,
            )
            .with(
                4_008,
                "TSCE.FormulaOwnerDependenciesArchive",
                decode_as::<tsce::FormulaOwnerDependenciesArchive>,
            )
            .with(
                4_009,
                "TSCE.CellRecordTileArchive",
                decode_as::<tsce::CellRecordTileArchive>,
            )
            .with(
                4_010,
                "TSCE.RangePrecedentsTileArchive",
                decode_as::<tsce::RangePrecedentsTileArchive>,
            );
        match format {
            Format::Numbers => known
                .with(1, "TN.DocumentArchive", decode_as::<tn::DocumentArchive>)
                .with(2, "TN.SheetArchive", decode_as::<tn::SheetArchive>)
                .with(
                    4_000,
                    "TSCE.CalculationEngineArchive",
                    decode_as::<tsce::CalculationEngineArchive>,
                )
                .with(
                    4_008,
                    "TSCE.FormulaOwnerDependenciesArchive",
                    decode_as::<tsce::FormulaOwnerDependenciesArchive>,
                )
                .with(
                    6_000,
                    "TST.TableInfoArchive",
                    decode_as::<tst::TableInfoArchive>,
                ),
            Format::Pages => known
                .with(
                    10_000,
                    "TP.DocumentArchive",
                    decode_as::<tp::DocumentArchive>,
                )
                .with(10_011, "TP.SectionArchive", decode_as::<tp::SectionArchive>),
            Format::Keynote => known
                .with(1, "KN.DocumentArchive", decode_as::<kn::DocumentArchive>)
                .with(2, "KN.ShowArchive", decode_as::<kn::ShowArchive>)
                .with(4, "KN.SlideNodeArchive", decode_as::<kn::SlideNodeArchive>)
                .with(5, "KN.SlideArchive", decode_as::<kn::SlideArchive>)
                // Slide drawables (wave W1-C, record 0372); the type numbers
                // are those of `tools/gen_keynote_reference_schema.py`.
                .with(
                    7,
                    "KN.PlaceholderArchive",
                    decode_as::<kn::PlaceholderArchive>,
                )
                .with(8, "KN.BuildArchive", decode_as::<kn::BuildArchive>)
                .with(15, "KN.NoteArchive", decode_as::<kn::NoteArchive>)
                .with(
                    633,
                    "TSA.CaptionInfoArchive",
                    decode_as::<tsa::CaptionInfoArchive>,
                )
                .with(
                    2_011,
                    "TSWP.ShapeInfoArchive",
                    decode_as::<tswp::ShapeInfoArchive>,
                )
                .with(3_005, "TSD.ImageArchive", decode_as::<tsd::ImageArchive>)
                .with(3_006, "TSD.MaskArchive", decode_as::<tsd::MaskArchive>)
                .with(3_007, "TSD.MovieArchive", decode_as::<tsd::MovieArchive>)
                .with(
                    3_056,
                    "TSD.CommentStorageArchive",
                    decode_as::<tsd::CommentStorageArchive>,
                )
                .with(
                    3_097,
                    "TSD.StandinCaptionArchive",
                    decode_as::<tsd::StandinCaptionArchive>,
                ),
        }
    }

    /// Register (or replace) the schema of one message type.
    #[must_use]
    pub fn with(mut self, message_type: u32, name: &'static str, decoder: Decoder) -> Self {
        self.decoders.insert(message_type, (name, decoder));
        self
    }

    /// Remove the schema of one message type.
    #[must_use]
    pub fn without(mut self, message_type: u32) -> Self {
        self.decoders.remove(&message_type);
        self
    }

    /// The schema name and decoder of `message_type`, if registered.
    pub fn get(&self, message_type: u32) -> Option<(&'static str, Decoder)> {
        self.decoders.get(&message_type).copied()
    }
}

/// Decode `data` as `M` and canonicalise it.
pub fn decode_as<M: Message + Default + Debug>(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    Ok(decoded(M::decode(data)?, data))
}

/// Canonicalise an already decoded (and possibly normalised) `message`
/// whose source payload was `data`. Cases that accept a documented
/// divergence register a decoder that normalises the message first.
pub fn decoded<M: Message + Debug>(message: M, data: &[u8]) -> Decoded {
    let canonical = message.encode_to_vec();
    let known = top_level_fields(&canonical);
    let unknown_fields = top_level_fields(data).difference(&known).copied().collect();
    Decoded {
        canonical,
        debug: format!("{message:#?}"),
        unknown_fields,
    }
}

/// Top-level field numbers of a protobuf payload (best effort: scanning
/// stops at the first malformed tag).
pub fn top_level_fields(data: &[u8]) -> BTreeSet<u32> {
    let mut fields = BTreeSet::new();
    let mut offset = 0;
    while offset < data.len() {
        let Some(tag) = varint(data, &mut offset) else {
            break;
        };
        let Ok(number) = u32::try_from(tag >> 3) else {
            break;
        };
        let skipped = match tag & 7 {
            0 => varint(data, &mut offset).map(|_| 0),
            1 => Some(8),
            2 => varint(data, &mut offset).and_then(|length| usize::try_from(length).ok()),
            5 => Some(4),
            _ => None,
        };
        let Some(skip) = skipped else {
            break;
        };
        let Some(end) = offset.checked_add(skip).filter(|end| *end <= data.len()) else {
            break;
        };
        fields.insert(number);
        offset = end;
    }
    fields
}

fn varint(data: &[u8], offset: &mut usize) -> Option<u64> {
    let mut value = 0_u64;
    for shift in (0..64).step_by(7) {
        let byte = *data.get(*offset)?;
        *offset += 1;
        value |= u64::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            return Some(value);
        }
    }
    None
}
