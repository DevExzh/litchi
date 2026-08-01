//! BIFF8 chart collection marker records of the Chart Sheet substream
//! (MS-XLS 2.1):
//!
//! - **Begin** (0x1033): the beginning of a chart record collection
//!   (MS-XLS 2.4.17).
//! - **End** (0x1034): the end of a chart record collection (MS-XLS 2.4.99).
//! - **PlotArea** (0x1035): an empty record marking that the following
//!   `Frame` record specifies the plot area (MS-XLS 2.4.197).
//!
//! All three records are fieldless; the readers validate the empty payload
//! so the records round-trip byte-exactly. Everything is INERT: no record
//! collections are constructed.
//!
//! # References
//!
//! - MS-XLS 2.4.17 (Begin), 2.4.99 (End), 2.4.197 (PlotArea)

use super::{XlsError, XlsResult};

macro_rules! empty_record {
    (
        $(#[$meta:meta])*
        $name:ident
    ) => {
        $(#[$meta])*
        #[derive(Debug, Clone, Copy, PartialEq, Eq)]
        pub struct $name;

        impl $name {
            /// Parse the record payload, which MUST be empty.
            pub fn parse(data: &[u8]) -> XlsResult<Self> {
                if !data.is_empty() {
                    return Err(XlsError::InvalidLength {
                        expected: 0,
                        found: data.len(),
                    });
                }
                Ok(Self)
            }

            /// Serialize back to a complete record payload.
            pub fn to_payload(&self) -> Vec<u8> {
                Vec::new()
            }
        }
    };
}

empty_record! {
    /// Typed `Begin` record content (MS-XLS 2.4.17): the beginning of a
    /// chart record collection. The record has no fields.
    XlsBegin
}

empty_record! {
    /// Typed `End` record content (MS-XLS 2.4.99): the end of a chart record
    /// collection. The record has no fields.
    XlsEnd
}

empty_record! {
    /// Typed `PlotArea` record content (MS-XLS 2.4.197): marks that the
    /// following `Frame` record specifies the plot area. The record has no
    /// fields.
    XlsPlotArea
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip_empty_payloads() {
        assert_eq!(XlsBegin::parse(&[]).unwrap().to_payload(), Vec::<u8>::new());
        assert_eq!(XlsEnd::parse(&[]).unwrap().to_payload(), Vec::<u8>::new());
        assert_eq!(
            XlsPlotArea::parse(&[]).unwrap().to_payload(),
            Vec::<u8>::new()
        );
    }

    #[test]
    fn rejects_non_empty_payloads() {
        assert!(XlsBegin::parse(&[0]).is_err());
        assert!(XlsEnd::parse(&[0, 0]).is_err());
        assert!(XlsPlotArea::parse(&[0xFF]).is_err());
    }
}
