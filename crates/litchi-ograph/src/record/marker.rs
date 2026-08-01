//! Empty collection and plot-area markers (`[MS-OGRAPH]` 2.4.11, 2.4.42,
//! and 2.4.78).

use crate::raw::{Encoder, Kind, RecordRef};
use crate::{Result, record};

macro_rules! marker {
    ($(#[$meta:meta])* $name:ident, $kind:literal) => {
        $(#[$meta])*
        #[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
        pub struct $name;

        impl $name {
            /// BIFF record identifier.
            pub const KIND: Kind = Kind::new($kind);

            /// Decodes a framed record and rejects a non-empty payload.
            pub fn parse(input: RecordRef<'_>) -> Result<Self> {
                record::payload(input, Self::KIND, 0)?;
                Ok(Self)
            }

            /// Decodes a payload supplied by an embedding host.
            pub fn from_payload(payload: &[u8]) -> Result<Self> {
                record::payload_bytes(Self::KIND, payload, 0)?;
                Ok(Self)
            }

            /// Appends the complete record to a bounded encoder.
            pub fn write(self, out: &mut Encoder) -> Result<()> {
                out.push(Self::KIND, &[])
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
    use super::*;
    use crate::raw::Records;

    #[test]
    fn marker_is_empty_and_typed() {
        let mut out = Encoder::new();
        Begin.write(&mut out).expect("write");
        let input = out.finish();
        let record = Records::new(&input).next().expect("record").expect("valid");
        assert_eq!(Begin::parse(record).expect("Begin"), Begin);
        assert!(End::parse(record).is_err());
        assert!(PlotArea::from_payload(&[0]).is_err());
    }
}
