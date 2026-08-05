//! Lossless, bounded views over BIFF chart substreams.
//!
//! A standalone Microsoft Graph workbook uses a chart BOF with document type
//! `0x8000`. Excel workbooks, including `Excel.Chart` OLE payloads embedded by
//! PowerPoint, use document type `0x0020`. [`Refs`] recognizes both without
//! depending on either host format.

use std::iter::FusedIterator;

use litchi_biff::{Kind as RecordKind, RecordRef, Records};

use crate::limits::as_u64;
use crate::{Error, Limits, Result};

pub mod axis;
pub mod cache;
mod codec;
pub mod format;
pub mod group;
pub mod layout;
mod model;

pub use model::{
    Ai, Binding, Cache, CellRef, Chart, Context, Count, DataKind, Edit, Family, Group, GroupId,
    Label, Legend, Link, Order, Owner, Props, Raw, Rect, Role, RowCol, Series, Source, Value,
    ValueRef, XlValue,
};

const BOF: RecordKind = RecordKind::from_wire(0x0809);
const EOF: RecordKind = RecordKind::from_wire(0x000A);
const BOF_BYTES: usize = 16;
const GRAPH_VERSION: u16 = 0x0680;
const GRAPH_DOC_TYPE: u16 = 0x8000;
const EXCEL_VERSION: u16 = 0x0600;
const EXCEL_DOC_TYPE: u16 = 0x0020;

/// Producer grammar used by a chart-substream BOF.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// Standalone Microsoft Graph chart sheet (`0x0680`, `0x8000`).
    Graph,
    /// Excel-hosted chart substream (`0x0600`, `0x0020`).
    Excel,
}

/// Borrowed, lossless view of one complete chart BOF-through-EOF substream.
#[derive(Debug, Clone, Copy)]
pub struct Ref<'a> {
    bytes: &'a [u8],
    kind: Kind,
    offset: usize,
    limits: Limits,
}

impl<'a> Ref<'a> {
    /// Validates one exact chart substream with conservative limits.
    pub fn open(bytes: &'a [u8]) -> Result<Self> {
        Self::with_limits(bytes, Limits::default())
    }

    /// Validates one exact chart substream with explicit limits.
    pub fn with_limits(bytes: &'a [u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        check_bytes(bytes, limits)?;

        let mut refs = Refs::with_limits(bytes, limits)?;
        let chart = refs.next().ok_or(Error::InvalidChart {
            offset: 0,
            reason: "missing chart BOF",
        })??;
        if chart.bytes.len() != bytes.len() {
            return chart_error(0, "records appear outside the chart substream");
        }
        if let Some(next) = refs.next() {
            next?;
            return chart_error(
                chart.bytes.len(),
                "more than one chart substream is present",
            );
        }
        Ok(chart)
    }

    pub(crate) const fn from_validated(
        bytes: &'a [u8],
        kind: Kind,
        offset: usize,
        limits: Limits,
    ) -> Self {
        Self {
            bytes,
            kind,
            offset,
            limits,
        }
    }

    /// Exact BOF-through-EOF bytes, including every unknown record.
    pub const fn as_bytes(self) -> &'a [u8] {
        self.bytes
    }

    /// Chart BOF grammar detected from the input.
    pub const fn kind(self) -> Kind {
        self.kind
    }

    /// Byte offset relative to the Workbook supplied to discovery.
    pub const fn offset(self) -> usize {
        self.offset
    }

    /// Resource bounds under which this chart was validated.
    pub const fn limits(self) -> Limits {
        self.limits
    }

    /// Traverse all records in original order without allocation.
    pub fn records(self) -> Records<'a> {
        Records::new(self.bytes)
    }

    /// Copies this bounded substream into an owned chart.
    pub fn own(self) -> Result<Stream> {
        self.own_with(self.limits)
    }

    /// Copies this substream under explicit bounds.
    pub fn own_with(self, limits: Limits) -> Result<Stream> {
        let limits = limits.validate()?;
        Ref::with_limits(self.bytes, limits)?;
        if self.bytes.len() > limits.biff.max_output_bytes {
            return Err(Error::LimitExceeded {
                resource: "output bytes",
                observed: as_u64(self.bytes.len()),
                maximum: as_u64(limits.biff.max_output_bytes),
            });
        }
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.bytes.len())
            .map_err(|_| Error::Allocation {
                resource: "chart bytes",
            })?;
        bytes.extend_from_slice(self.bytes);
        Ok(Stream {
            bytes,
            kind: self.kind,
            offset: self.offset,
            limits,
        })
    }
}

/// Move-owned, lossless chart substream.
#[derive(Debug)]
pub struct Stream {
    bytes: Vec<u8>,
    kind: Kind,
    offset: usize,
    limits: Limits,
}

impl Stream {
    /// Takes ownership and validates without copying the input allocation.
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Self::with_limits(bytes, Limits::default())
    }

    /// Takes ownership and validates under explicit resource bounds.
    pub fn with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let chart = Ref::with_limits(&bytes, limits)?;
        let kind = chart.kind;
        let limits = chart.limits;
        Ok(Self {
            bytes,
            kind,
            offset: 0,
            limits,
        })
    }

    /// Borrow this chart without copying or revalidation.
    pub fn as_ref(&self) -> Ref<'_> {
        Ref::from_validated(&self.bytes, self.kind, self.offset, self.limits)
    }

    /// Exact BOF-through-EOF bytes.
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Chart BOF grammar detected from the input.
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Original Workbook-relative offset, or zero for an isolated input.
    pub const fn offset(&self) -> usize {
        self.offset
    }

    /// Traverse all records in original order without allocation.
    pub fn records(&self) -> Records<'_> {
        self.as_ref().records()
    }

    /// Recover the original allocation without copying.
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    pub(crate) fn relimit(mut self, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        Ref::with_limits(&self.bytes, limits)?;
        if self.bytes.len() > limits.biff.max_output_bytes {
            return Err(Error::LimitExceeded {
                resource: "output bytes",
                observed: as_u64(self.bytes.len()),
                maximum: as_u64(limits.biff.max_output_bytes),
            });
        }
        self.limits = limits;
        Ok(self)
    }
}

/// Move-owned Workbook bytes containing one or more validated chart substreams.
///
/// Unlike the root [`crate::Workbook`], this host-neutral capability does not
/// require the standalone Microsoft Graph globals-plus-chart topology. It is
/// suitable for Excel workbooks whose charts are nested in worksheet streams.
#[derive(Debug)]
pub struct Book {
    bytes: Vec<u8>,
    limits: Limits,
    charts: usize,
}

impl Book {
    /// Takes ownership, validates every discovered chart, and requires one.
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Self::with_limits(bytes, Limits::default())
    }

    /// Takes ownership and validates every chart under explicit bounds.
    pub fn with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let mut charts = Refs::with_limits(&bytes, limits)?;
        let mut count = 0usize;
        for chart in &mut charts {
            chart?;
            count = count.checked_add(1).ok_or(Error::SizeOverflow {
                resource: "chart count",
            })?;
        }
        if count == 0 {
            return chart_error(0, "Workbook contains no chart substream");
        }
        Ok(Self {
            bytes,
            limits,
            charts: count,
        })
    }

    /// Exact caller-supplied Workbook stream bytes.
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Traverse every validated chart in source order without allocation.
    pub fn charts(&self) -> Refs<'_> {
        Refs::from_validated(&self.bytes, self.limits)
    }

    /// Resource limits under which all charts were validated.
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Number of validated chart substreams.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.charts
    }

    /// Whether this Workbook contains no charts.
    ///
    /// A successfully constructed `Book` is never empty.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.charts == 0
    }

    /// Recover the original Workbook allocation without copying.
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

/// Allocation-free discovery of chart substreams in a caller-owned Workbook.
///
/// Errors are yielded at the first malformed chart or exhausted resource
/// bound. After an error the iterator is fused.
#[derive(Debug)]
pub struct Refs<'a> {
    bytes: &'a [u8],
    records: Records<'a>,
    limits: Limits,
    active: Option<Active>,
    charts: usize,
    done: bool,
}

#[derive(Debug, Clone, Copy)]
struct Active {
    start: usize,
    kind: Kind,
    records: usize,
}

impl<'a> Refs<'a> {
    /// Scans any BIFF Workbook stream for standalone or Excel chart BOFs.
    pub fn open_workbook(bytes: &'a [u8]) -> Result<Self> {
        Self::with_limits(bytes, Limits::default())
    }

    /// Scans a BIFF Workbook stream with explicit resource bounds.
    pub fn with_limits(bytes: &'a [u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        check_bytes(bytes, limits)?;
        let records = Records::with_limits(bytes, limits.biff)?;
        Ok(Self {
            bytes,
            records,
            limits,
            active: None,
            charts: 0,
            done: false,
        })
    }

    pub(crate) fn from_validated(bytes: &'a [u8], limits: Limits) -> Self {
        Self {
            bytes,
            records: Records::new(bytes),
            limits,
            active: None,
            charts: 0,
            done: false,
        }
    }

    fn fail(&mut self, error: Error) -> Option<Result<Ref<'a>>> {
        self.done = true;
        Some(Err(error))
    }
}

impl<'a> Iterator for Refs<'a> {
    type Item = Result<Ref<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }

        loop {
            let record = match self.records.next() {
                Some(Ok(record)) => record,
                Some(Err(error)) => return self.fail(error.into()),
                None => {
                    self.done = true;
                    return self
                        .active
                        .map(|active| chart_error(active.start, "chart BOF has no matching EOF"));
                },
            };

            if let Some(mut active) = self.active {
                active.records = match active.records.checked_add(1) {
                    Some(records) => records,
                    None => {
                        return self.fail(Error::SizeOverflow {
                            resource: "chart record count",
                        });
                    },
                };
                if active.records > self.limits.max_chart_records {
                    return self.fail(Error::LimitExceeded {
                        resource: "chart record count",
                        observed: as_u64(active.records),
                        maximum: as_u64(self.limits.max_chart_records),
                    });
                }
                self.active = Some(active);

                if record.kind() == BOF {
                    return self.fail(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "nested BOF in chart substream",
                    });
                }
                if record.kind() != EOF {
                    continue;
                }
                if !record.payload().is_empty() {
                    return self.fail(Error::InvalidChart {
                        offset: record.offset(),
                        reason: "chart EOF has a non-empty payload",
                    });
                }

                let end = match record.offset().checked_add(record.encoded().len()) {
                    Some(end) => end,
                    None => {
                        return self.fail(Error::SizeOverflow {
                            resource: "chart substream",
                        });
                    },
                };
                let Some(bytes) = self.bytes.get(active.start..end) else {
                    return self.fail(Error::InvalidChart {
                        offset: active.start,
                        reason: "chart range exceeds Workbook bytes",
                    });
                };
                self.active = None;
                return Some(Ok(Ref::from_validated(
                    bytes,
                    active.kind,
                    active.start,
                    self.limits,
                )));
            }

            if record.kind() != BOF {
                continue;
            }
            let kind = match parse_bof(record) {
                Ok(Some(kind)) => kind,
                Ok(None) => continue,
                Err(error) => return self.fail(error),
            };
            let next = match self.charts.checked_add(1) {
                Some(next) => next,
                None => {
                    return self.fail(Error::SizeOverflow {
                        resource: "chart count",
                    });
                },
            };
            if next > self.limits.max_charts {
                return self.fail(Error::LimitExceeded {
                    resource: "chart count",
                    observed: as_u64(next),
                    maximum: as_u64(self.limits.max_charts),
                });
            }
            self.charts = next;
            self.active = Some(Active {
                start: record.offset(),
                kind,
                records: 1,
            });
        }
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        if self.done { (0, Some(0)) } else { (0, None) }
    }
}

impl FusedIterator for Refs<'_> {}

fn parse_bof(record: RecordRef<'_>) -> Result<Option<Kind>> {
    let payload = record.payload();
    if payload.len() < 4 {
        return chart_error(record.offset(), "BOF payload is shorter than four bytes");
    }
    let version = le_u16(payload, 0).ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "BOF version is truncated",
    })?;
    let doc_type = le_u16(payload, 2).ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "BOF document type is truncated",
    })?;
    let kind = match (version, doc_type) {
        (GRAPH_VERSION, GRAPH_DOC_TYPE) => Kind::Graph,
        (EXCEL_VERSION, EXCEL_DOC_TYPE) => Kind::Excel,
        (_, GRAPH_DOC_TYPE) | (_, EXCEL_DOC_TYPE) => {
            return chart_error(record.offset(), "chart BOF has an unsupported BIFF version");
        },
        _ => return Ok(None),
    };
    if payload.len() != BOF_BYTES {
        return chart_error(record.offset(), "chart BOF payload is not 16 bytes");
    }
    Ok(Some(kind))
}

fn check_bytes(bytes: &[u8], limits: Limits) -> Result<()> {
    if bytes.len() > limits.max_workbook_bytes {
        return Err(Error::LimitExceeded {
            resource: "Workbook bytes",
            observed: as_u64(bytes.len()),
            maximum: as_u64(limits.max_workbook_bytes),
        });
    }
    Ok(())
}

fn le_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    let pair = bytes.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_le_bytes([*pair.first()?, *pair.get(1)?]))
}

fn chart_error<T>(offset: usize, reason: &'static str) -> Result<T> {
    Err(Error::InvalidChart { offset, reason })
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_biff::{Encoder, Kind as RecordKind, Record as OwnedRecord, Records};

    const UNKNOWN: RecordKind = RecordKind::from_wire(0x7777);

    fn bof(version: u16, doc_type: u16, tail: u8) -> [u8; BOF_BYTES] {
        let mut bytes = [tail; BOF_BYTES];
        bytes[0..2].copy_from_slice(&version.to_le_bytes());
        bytes[2..4].copy_from_slice(&doc_type.to_le_bytes());
        bytes
    }

    fn push(out: &mut Encoder, kind: RecordKind, payload: &[u8]) {
        out.push(kind, payload).expect("bounded test record");
    }

    fn chart(kind: Kind, unknown: &[u8]) -> Vec<u8> {
        let mut out = Encoder::new();
        match kind {
            Kind::Graph => push(&mut out, BOF, &bof(GRAPH_VERSION, GRAPH_DOC_TYPE, 0xA5)),
            Kind::Excel => push(&mut out, BOF, &bof(EXCEL_VERSION, EXCEL_DOC_TYPE, 0x5A)),
        }
        push(&mut out, UNKNOWN, unknown);
        push(&mut out, EOF, &[]);
        out.finish()
    }

    fn hosted_workbook() -> (Vec<u8>, usize, Vec<u8>, usize, Vec<u8>) {
        let first = chart(Kind::Excel, &[0x10, 0x20, 0x30]);
        let second = chart(Kind::Excel, &[0xFE, 0xDC]);
        let mut out = Encoder::new();
        push(&mut out, BOF, &bof(EXCEL_VERSION, 0x0005, 0));
        push(&mut out, UNKNOWN, &[1]);
        push(&mut out, EOF, &[]);
        push(&mut out, BOF, &bof(EXCEL_VERSION, 0x0010, 0));
        push(&mut out, UNKNOWN, &[2, 3]);
        let first_offset = out.as_bytes().len();
        for record in Records::new(&first) {
            out.push_ref(record.expect("first chart"))
                .expect("copy first chart");
        }
        push(&mut out, UNKNOWN, &[4]);
        push(&mut out, EOF, &[]);
        let second_offset = out.as_bytes().len();
        for record in Records::new(&second) {
            out.push_ref(record.expect("second chart"))
                .expect("copy second chart");
        }
        (out.finish(), first_offset, first, second_offset, second)
    }

    #[test]
    fn discovers_hosted_charts_in_order_without_copying() {
        let (bytes, first_offset, first_bytes, second_offset, second_bytes) = hosted_workbook();
        let pointer = bytes.as_ptr();
        let capacity = bytes.capacity();
        let book = Book::open(bytes).expect("host Workbook");
        assert_eq!(book.as_bytes().as_ptr(), pointer);
        assert_eq!(book.len(), 2);
        assert!(!book.is_empty());
        let mut charts = book.charts();

        let first = charts.next().expect("first").expect("valid first");
        assert_eq!(first.kind(), Kind::Excel);
        assert_eq!(first.offset(), first_offset);
        assert_eq!(first.as_bytes(), first_bytes);
        assert_eq!(
            first.as_bytes().as_ptr(),
            book.as_bytes()[first_offset..].as_ptr()
        );

        let owned = first.own().expect("bounded owned chart");
        assert_eq!(owned.offset(), first_offset);
        assert_eq!(owned.as_bytes(), first_bytes);

        let second = charts.next().expect("second").expect("valid second");
        assert_eq!(second.offset(), second_offset);
        assert_eq!(second.as_bytes(), second_bytes);
        assert!(charts.next().is_none());

        let bytes = book.into_bytes();
        assert_eq!(bytes.as_ptr(), pointer);
        assert_eq!(bytes.capacity(), capacity);
    }

    #[test]
    fn owned_chart_and_record_reuse_input_allocations_exactly() {
        let bytes = chart(Kind::Graph, &[0xCA, 0xFE, 0xBA, 0xBE]);
        let pointer = bytes.as_ptr();
        let capacity = bytes.capacity();
        let owned = Stream::open(bytes).expect("chart");
        assert_eq!(owned.as_bytes().as_ptr(), pointer);
        assert_eq!(
            owned.as_bytes().len(),
            owned
                .records()
                .map(|r| r.expect("record").encoded().len())
                .sum()
        );

        let mut replay = Encoder::new();
        for record in owned.records() {
            replay
                .push_ref(record.expect("valid record"))
                .expect("exact replay");
        }
        assert_eq!(replay.finish(), owned.as_bytes());
        let bytes = owned.into_bytes();
        assert_eq!(bytes.as_ptr(), pointer);
        assert_eq!(bytes.capacity(), capacity);

        let frame = Records::new(&bytes)
            .nth(1)
            .expect("unknown frame")
            .expect("valid frame")
            .encoded()
            .to_vec();
        let frame_pointer = frame.as_ptr();
        let record = OwnedRecord::open(frame).expect("owned record");
        assert_eq!(record.kind(), UNKNOWN);
        assert_eq!(record.as_ref().encoded(), record.as_bytes());
        let frame = record.into_bytes();
        assert_eq!(frame.as_ptr(), frame_pointer);
    }

    #[test]
    fn rejects_nested_truncated_and_malformed_chart_boundaries() {
        let chart_bof = bof(EXCEL_VERSION, EXCEL_DOC_TYPE, 0);

        let mut nested = Encoder::new();
        push(&mut nested, BOF, &chart_bof);
        push(&mut nested, BOF, &chart_bof);
        push(&mut nested, EOF, &[]);
        assert!(matches!(
            Refs::open_workbook(&nested.finish())
                .expect("scanner")
                .next(),
            Some(Err(Error::InvalidChart {
                reason: "nested BOF in chart substream",
                ..
            }))
        ));

        let mut missing = Encoder::new();
        push(&mut missing, BOF, &chart_bof);
        push(&mut missing, UNKNOWN, &[1]);
        assert!(matches!(
            Refs::open_workbook(&missing.finish())
                .expect("scanner")
                .next(),
            Some(Err(Error::InvalidChart {
                reason: "chart BOF has no matching EOF",
                ..
            }))
        ));

        let mut nonempty = Encoder::new();
        push(&mut nonempty, BOF, &chart_bof);
        push(&mut nonempty, EOF, &[1]);
        assert!(matches!(
            Book::open(nonempty.finish()),
            Err(Error::InvalidChart {
                reason: "chart EOF has a non-empty payload",
                ..
            })
        ));

        let mut short = Encoder::new();
        push(&mut short, BOF, &[0, 6, 0x20]);
        assert!(matches!(
            Book::open(short.finish()),
            Err(Error::InvalidChart {
                reason: "BOF payload is shorter than four bytes",
                ..
            })
        ));

        let mut version = Encoder::new();
        push(&mut version, BOF, &bof(0x0500, EXCEL_DOC_TYPE, 0));
        assert!(matches!(
            Book::open(version.finish()),
            Err(Error::InvalidChart {
                reason: "chart BOF has an unsupported BIFF version",
                ..
            })
        ));
    }

    #[test]
    fn enforces_workbook_chart_and_per_chart_bounds() {
        let (bytes, _, _, _, _) = hosted_workbook();
        let mut count = Refs::with_limits(
            &bytes,
            Limits {
                max_charts: 1,
                ..Limits::default()
            },
        )
        .expect("scanner");
        assert!(count.next().expect("first chart").is_ok());
        assert!(matches!(
            count.next(),
            Some(Err(Error::LimitExceeded {
                resource: "chart count",
                ..
            }))
        ));
        assert!(count.next().is_none());

        let one = chart(Kind::Excel, &[1]);
        assert!(matches!(
            Book::with_limits(
                one,
                Limits {
                    max_chart_records: 2,
                    ..Limits::default()
                }
            ),
            Err(Error::LimitExceeded {
                resource: "chart record count",
                ..
            })
        ));

        assert!(matches!(
            Refs::with_limits(
                &bytes,
                Limits {
                    max_workbook_bytes: bytes.len() - 1,
                    ..Limits::default()
                }
            ),
            Err(Error::LimitExceeded {
                resource: "Workbook bytes",
                ..
            })
        ));

        assert!(matches!(
            Book::open(vec![0x0A, 0, 0, 0]),
            Err(Error::InvalidChart {
                reason: "Workbook contains no chart substream",
                ..
            })
        ));
    }
}
