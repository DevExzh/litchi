//! Selection of the rendering track for a metafile that carries EMF+ records.
//!
//! [MS-EMFPLUS] 1.3.1 defines two kinds of EMF+ metafile. A dual metafile
//! carries EMF records that alone define the whole picture, and EMF+-only
//! metafiles rely on EMF+ records plus the EMF records that follow an
//! `EmfPlusGetDC` record. GDI+ recognizes an EMF+ metafile only by the
//! `EmfPlusHeader` record in the first record after `EMR_HEADER`, so a stream
//! without one never selects the EMF+ track.

use super::renderer::MetafileKind;
use super::{
    EMR_COMMENT, EmfPlusRecordIter, ParserLimits, RecordType, try_extract_emfplus_comment_body,
};

/// Facts about the EMF+ records of one metafile, gathered before playback.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct StreamFacts {
    /// Whether any `EMR_COMMENT_EMFPLUS` record is present.
    pub has_emfplus_comment: bool,
    /// The kind declared by the `EmfPlusHeader` record that starts the EMF+
    /// stream.
    ///
    /// GDI+ reads that header only from the record right after `EMR_HEADER`,
    /// so `None` means GDI+ treats the metafile as a classic EMF: there are no
    /// EMF+ records, they start later (for example a dual metafile copied into
    /// another one), or the first comment cannot be framed.
    pub kind: Option<MetafileKind>,
    /// Number of `EmfPlusGetDC` records in the stream.
    pub get_dc_count: usize,
}

impl StreamFacts {
    /// Whether the metafile declares that its classic records define the whole
    /// picture on their own.
    #[must_use]
    pub const fn is_dual(&self) -> bool {
        matches!(self.kind, Some(MetafileKind::Dual))
    }
}

/// The records a metafile is rendered from.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TrackChoice {
    /// Play the classic EMF records only.
    Classic {
        /// Why the EMF+ records were not played, when the metafile has any.
        ///
        /// The code starts with `noncanonical-`: the classic records define the
        /// picture a GDI player shows.
        note: Option<(&'static str, &'static str)>,
    },
    /// Play the EMF+ records.
    EmfPlus {
        /// Whether the classic records remain a complete fallback rendering
        /// when the EMF+ track fails or would need an approximation.
        classic_fallback: bool,
    },
}

/// Choose the rendering track of one metafile from its EMF+ stream facts.
///
/// A dual metafile's classic records alone define the whole picture
/// ([MS-EMFPLUS] 1.3.1), so they are the fallback rendering whenever the EMF+
/// records fail. When such a metafile uses `EmfPlusGetDC`, the complete classic
/// track is chosen: GDI+ itself records the drawing between `EmfPlusGetDC` and
/// the next EMF+ record as classic records only, so the classic track contains
/// every part of the picture, while the EMF+ drawing records next to them carry
/// images and brushes that the EMF+ renderer may not reproduce exactly.
///
/// An EMF+-only metafile has no complete classic rendering, so its EMF+
/// failures stay failures and an `EmfPlusGetDC` record needs record-level
/// interleaving of the two tracks, which is not implemented.
pub fn choose_track(facts: StreamFacts) -> Result<TrackChoice, &'static str> {
    if !facts.has_emfplus_comment {
        return Ok(TrackChoice::Classic { note: None });
    }
    match facts.kind {
        None => Ok(TrackChoice::Classic {
            note: Some((
                "noncanonical-emfplus-not-leading",
                "EMF+ records that do not start with an EmfPlusHeader right after the EMF header are ignored, as GDI+ plays such a metafile as classic EMF",
            )),
        }),
        Some(MetafileKind::Dual) if facts.get_dc_count > 0 => Ok(TrackChoice::Classic {
            note: Some((
                "noncanonical-emfplus-dual-getdc",
                "dual metafile with GetDC records rendered from its complete classic EMF track",
            )),
        }),
        Some(MetafileKind::Dual) => Ok(TrackChoice::EmfPlus {
            classic_fallback: true,
        }),
        Some(MetafileKind::EmfPlusOnly) if facts.get_dc_count > 0 => {
            Err("EMF+ Only GetDC requires ordered classic/EMF+ SVG interleaving")
        },
        Some(MetafileKind::EmfPlusOnly) => Ok(TrackChoice::EmfPlus {
            classic_fallback: false,
        }),
    }
}

/// Collect the stream facts of one metafile from its records.
///
/// `records` are `(record type, EmfRecord::data)` pairs of every record after
/// `EMR_HEADER`, in stream order. EMF+ records are only framed, never decoded,
/// and scanning of a comment stops at its first framing error because the EMF+
/// renderer reports that error.
#[must_use]
pub fn scan_stream<'a, I>(records: I, limits: ParserLimits) -> StreamFacts
where
    I: IntoIterator<Item = (u32, &'a [u8])>,
{
    let mut facts = StreamFacts::default();
    for (index, (record_type, body)) in records.into_iter().enumerate() {
        if record_type != EMR_COMMENT {
            continue;
        }
        let Ok(Some(payload)) = try_extract_emfplus_comment_body(body, limits) else {
            continue;
        };
        facts.has_emfplus_comment = true;
        let Ok(framed) = EmfPlusRecordIter::new(payload, limits) else {
            continue;
        };
        for (position, record) in framed.enumerate() {
            let Ok(record) = record else {
                break;
            };
            if index == 0 && position == 0 && record.header.record_type == RecordType::Header {
                facts.kind = Some(if record.header.flags.raw() & 1 != 0 {
                    MetafileKind::Dual
                } else {
                    MetafileKind::EmfPlusOnly
                });
            }
            if record.header.record_type == RecordType::GetDc {
                facts.get_dc_count = facts.get_dc_count.saturating_add(1);
            }
        }
    }
    facts
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(kind: RecordType, flags: u16, data: &[u8]) -> Vec<u8> {
        let size = 12_u32 + u32::try_from(data.len()).unwrap();
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&kind.raw().to_le_bytes());
        bytes.extend_from_slice(&flags.to_le_bytes());
        bytes.extend_from_slice(&size.to_le_bytes());
        bytes.extend_from_slice(&u32::try_from(data.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(data);
        bytes
    }

    fn comment(payload: &[u8]) -> Vec<u8> {
        let data_size = 4_u32 + u32::try_from(payload.len()).unwrap();
        let mut body = Vec::new();
        body.extend_from_slice(&data_size.to_le_bytes());
        body.extend_from_slice(&super::super::EMFPLUS_COMMENT_IDENTIFIER.to_le_bytes());
        body.extend_from_slice(payload);
        while (body.len() + 8) % 4 != 0 {
            body.push(0);
        }
        body
    }

    /// Scan records that follow `EMR_HEADER`: EMF+ comment bodies, with `None`
    /// standing for a classic drawing record.
    fn scan(records: &[Option<Vec<u8>>]) -> StreamFacts {
        const EMR_RECTANGLE: u32 = 43;
        scan_stream(
            records.iter().map(|record| match record {
                Some(body) => (EMR_COMMENT, body.as_slice()),
                None => (EMR_RECTANGLE, &[0_u8; 16][..]),
            }),
            ParserLimits::default(),
        )
    }

    fn stream(records: &[Vec<u8>]) -> Option<Vec<u8>> {
        Some(comment(&records.concat()))
    }

    const DUAL_GET_DC: TrackChoice = TrackChoice::Classic {
        note: Some((
            "noncanonical-emfplus-dual-getdc",
            "dual metafile with GetDC records rendered from its complete classic EMF track",
        )),
    };

    #[test]
    fn dual_header_with_get_dc_selects_the_classic_track() {
        let facts = scan(&[stream(&[
            record(RecordType::Header, 1, &[0; 16]),
            record(RecordType::GetDc, 0, &[]),
            record(RecordType::EndOfFile, 0, &[]),
        ])]);
        assert_eq!(facts.kind, Some(MetafileKind::Dual));
        assert_eq!(facts.get_dc_count, 1);
        assert_eq!(choose_track(facts), Ok(DUAL_GET_DC));
    }

    #[test]
    fn dual_header_without_get_dc_keeps_the_classic_fallback() {
        let facts = scan(&[stream(&[
            record(RecordType::Header, 1, &[0; 16]),
            record(RecordType::EndOfFile, 0, &[]),
        ])]);
        assert_eq!(
            choose_track(facts),
            Ok(TrackChoice::EmfPlus {
                classic_fallback: true
            })
        );
    }

    #[test]
    fn emfplus_only_stream_without_get_dc_has_no_fallback() {
        let facts = scan(&[stream(&[
            record(RecordType::Header, 0, &[0; 16]),
            record(RecordType::EndOfFile, 0, &[]),
        ])]);
        assert_eq!(
            choose_track(facts),
            Ok(TrackChoice::EmfPlus {
                classic_fallback: false
            })
        );
    }

    #[test]
    fn emfplus_only_get_dc_needs_interleaving() {
        let facts = scan(&[stream(&[
            record(RecordType::Header, 0, &[0; 16]),
            record(RecordType::GetDc, 0, &[]),
            record(RecordType::EndOfFile, 0, &[]),
        ])]);
        assert!(choose_track(facts).is_err());
    }

    #[test]
    fn get_dc_is_counted_across_comment_boundaries() {
        let facts = scan(&[
            stream(&[
                record(RecordType::Header, 1, &[0; 16]),
                record(RecordType::GetDc, 0, &[]),
            ]),
            None,
            stream(&[
                record(RecordType::GetDc, 0, &[]),
                record(RecordType::EndOfFile, 0, &[]),
            ]),
        ]);
        assert_eq!(facts.get_dc_count, 2);
        assert!(facts.is_dual());
    }

    #[test]
    fn stream_without_a_leading_header_plays_the_classic_track() {
        let facts = scan(&[stream(&[
            record(RecordType::GetDc, 0, &[]),
            record(RecordType::Header, 1, &[0; 16]),
            record(RecordType::EndOfFile, 0, &[]),
        ])]);
        assert!(facts.has_emfplus_comment);
        assert_eq!(facts.kind, None);
        assert!(matches!(
            choose_track(facts),
            Ok(TrackChoice::Classic {
                note: Some(("noncanonical-emfplus-not-leading", _))
            })
        ));
    }

    #[test]
    fn emfplus_stream_after_classic_records_plays_the_classic_track() {
        // A dual metafile played into another metafile leaves its EMF+
        // records after the outer metafile's own classic records.
        let facts = scan(&[
            None,
            stream(&[
                record(RecordType::Header, 1, &[0; 16]),
                record(RecordType::EndOfFile, 0, &[]),
            ]),
        ]);
        assert_eq!(facts.kind, None);
        assert!(matches!(
            choose_track(facts),
            Ok(TrackChoice::Classic {
                note: Some(("noncanonical-emfplus-not-leading", _))
            })
        ));
    }

    #[test]
    fn an_unframed_emfplus_comment_has_no_kind() {
        let facts = scan(&[Some(vec![0_u8; 8])]);
        assert!(!facts.has_emfplus_comment);
        assert_eq!(choose_track(facts), Ok(TrackChoice::Classic { note: None }));
    }

    #[test]
    fn metafile_without_emfplus_comments_plays_classic_records() {
        let facts = scan_stream(std::iter::empty(), ParserLimits::default());
        assert_eq!(choose_track(facts), Ok(TrackChoice::Classic { note: None }));
    }
}
