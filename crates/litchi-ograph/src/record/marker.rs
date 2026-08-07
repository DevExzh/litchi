//! Empty collection and plot-area markers (`[MS-OGRAPH]` 2.4.11, 2.4.42,
//! and 2.4.78).

use crate::{Result, record};
use litchi_biff::{Encoder, Kind, RecordRef};

macro_rules! marker {
    ($(#[$meta:meta])* $name:ident, $kind:literal) => {
        $(#[$meta])*
        #[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
        pub struct $name;

        impl $name {
            /// BIFF record identifier.
            pub const KIND: Kind = Kind::from_wire($kind);

            /// Decodes a framed record and rejects a non-empty payload.
            ///
            /// # Errors
            ///
            /// Returns [`crate::Error::WrongRecord`] if the record identifier
            /// differs, or [`crate::Error::InvalidRecordLength`] if the
            /// payload is not empty.
            pub fn parse(input: RecordRef<'_>) -> Result<Self> {
                record::payload(input, Self::KIND, 0)?;
                Ok(Self)
            }

            /// Decodes a payload supplied by an embedding host.
            ///
            /// # Errors
            ///
            /// Returns [`crate::Error::InvalidRecordLength`] if the payload is
            /// not empty.
            pub fn from_payload(payload: &[u8]) -> Result<Self> {
                record::payload_bytes(Self::KIND, payload, 0)?;
                Ok(Self)
            }

            /// Appends the complete record to a bounded encoder.
            ///
            /// # Errors
            ///
            /// Returns [`crate::Error::Biff`] if the encoder rejects the
            /// record.
            pub fn write(self, out: &mut Encoder) -> Result<()> {
                out.push(Self::KIND, &[])?;
                Ok(())
            }
        }
    };
}

marker! {
    /// Beginning of a chart record collection.
    Begin, 0x1033
}

marker! {
    /// End of a chart record collection.
    End, 0x1034
}

marker! {
    /// Indicates that the following frame belongs to the plot area.
    PlotArea, 0x1035
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic by design"
    )]
    use super::*;
    use litchi_biff::{Error as BiffError, Limits, Records, Resource};

    #[test]
    fn marker_is_empty_and_typed() {
        let mut out = Encoder::new();
        Begin.write(&mut out).expect("write");
        End.write(&mut out).expect("write");
        let input = out.finish();
        assert_eq!(input, [0x33, 0x10, 0, 0, 0x34, 0x10, 0, 0]);

        let mut records = Records::new(&input);
        let begin = records.next().expect("record").expect("valid");
        assert_eq!(begin.offset(), 0);
        assert_eq!(Begin::parse(begin).expect("Begin"), Begin);
        let end = records.next().expect("record").expect("valid");
        assert_eq!(end.offset(), 4);
        assert!(End::parse(begin).is_err());
        assert_eq!(End::parse(end).expect("End"), End);
        assert!(PlotArea::from_payload(&[0]).is_err());
    }

    #[test]
    fn framing_bounds_and_malformed_payloads_are_typed() {
        let input = [0x33, 0x10, 0, 0, 0x34, 0x10, 0, 0];
        let mut limits = Limits::default();
        limits.max_records = 1;
        let mut records = Records::with_limits(&input, limits).expect("limits");
        assert!(records.next().expect("first record").is_ok());
        assert!(matches!(
            records.next(),
            Some(Err(BiffError::LimitExceeded {
                resource: Resource::RecordCount,
                observed: 2,
                maximum: 1,
            }))
        ));

        let malformed = [0x33, 0x10, 1, 0];
        assert!(matches!(
            Records::new(&malformed).next(),
            Some(Err(BiffError::TruncatedPayload {
                offset: 0,
                declared: 1,
                available: 0,
                ..
            }))
        ));
    }
}
