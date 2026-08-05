//! BIFF8 `RealTimeData` record (MS-XLS 2.4.214): real-time data (RTD)
//! topics in the workbook globals substream.
//!
//! There is one `RealTimeData` record per RTD topic; the `RTD` production
//! (MS-XLS 2.1) is `RealTimeData *ContinueFrt`, so the logical payload is the
//! record body followed by any `ContinueFrt` bodies concatenated. Each record
//! carries the topic as an `XLUnicodeStringSegmentedRTD` (MS-XLS 2.5.298)
//! whose first sub-string is the RTD server ProgID and whose second is the
//! server name, the last value returned by the server as an `RTDOper`
//! variant (MS-XLS 2.5.224), and the cells subscribed to the topic as
//! `RTDEItem` entries (MS-XLS 2.5.223).
//!
//! Adjacent records share topic prefixes: `ichSamePrefix` counts the leading
//! characters this topic has in common with the previous record's topic, and
//! the stored string holds only the remainder. The fully reconstructed topic
//! is exposed as [`Record::topic`]; pass the previous topic to
//! [`Record::parse`] so the prefix can be re-applied.
//!
//! Everything in this module is INERT: ProgIDs, server names, and topics are
//! stored verbatim and no RTD server is ever located, launched, or queried.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub(crate) use codec::{CONTINUE_FRT_RECORD_TYPE, REAL_TIME_DATA_RECORD_TYPE};

// The crate-root facade is migrated separately; keep its existing exports
// working while the contextual owner uses prefix-free canonical names.
pub use model::{Cell as RtdCell, Record as RealTimeData, Value as RtdValue};
