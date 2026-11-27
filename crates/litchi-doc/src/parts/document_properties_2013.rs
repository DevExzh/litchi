//! Typed Word 2013 Document Properties extension.

use super::document_properties_97::DopExtensionError;
use super::document_properties_2010::Dop2010;

const DOP2010_SIZE: usize = 690;
const DOP2013_SIZE: usize = 694;

/// Typed Word 2013 DOP extension.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Dop2013 {
    /// Whether chart data labels and point properties track reference identities.
    pub chart_tracking_reference_based: bool,
}

impl Dop2013 {
    /// Parses the complete Word 2013 DOP generation.
    ///
    /// # Errors
    ///
    /// Returns [`DopExtensionError`] when this or an earlier DOP prefix is
    /// invalid or reserved chart-tracking bits are set.
    pub fn parse(dop: &[u8]) -> Result<Self, DopExtensionError> {
        if dop.len() < DOP2013_SIZE {
            return Err(DopExtensionError::new("Dop2013 is shorter than 694 bytes"));
        }
        Dop2010::parse(dop)?;
        let flags = le_u32(dop, DOP2010_SIZE);
        if flags & !1 != 0 {
            return Err(DopExtensionError::new(
                "Dop2013 reserved chart-tracking bits are nonzero",
            ));
        }
        Ok(Self {
            chart_tracking_reference_based: flags & 1 != 0,
        })
    }

    /// Writes the Word 2013 extension without normalizing older DOP bytes.
    ///
    /// # Errors
    ///
    /// Returns [`DopExtensionError`] when the target is shorter than the
    /// complete Word 2013 DOP generation.
    pub fn write_into(self, dop: &mut [u8]) -> Result<(), DopExtensionError> {
        if dop.len() < DOP2013_SIZE {
            return Err(DopExtensionError::new(
                "Dop2013 target is shorter than 694 bytes",
            ));
        }
        put_u32(
            dop,
            DOP2010_SIZE,
            u32::from(self.chart_tracking_reference_based),
        );
        Ok(())
    }
}

fn le_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

fn put_u32(data: &mut [u8], offset: usize, value: u32) {
    data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

#[cfg(test)]
mod tests {
    use super::*;

    fn valid_dop2013() -> Vec<u8> {
        let mut dop = vec![0u8; DOP2013_SIZE];
        dop[0x190..0x19a].copy_from_slice(&[0xa5, 0x06, 0xc0, 0x07, 0xb4, 0, 0xb4, 0, 1, 0x81]);
        let math = 616 + 24;
        put_u32(&mut dop, math, 1 << 4 | 1 << 11 | 1 << 12);
        dop[math + 14..math + 18].copy_from_slice(&120i32.to_le_bytes());
        dop[math + 18..math + 22].copy_from_slice(&120i32.to_le_bytes());
        put_u32(&mut dop, 674, 1);
        dop
    }

    #[test]
    fn parses_and_writes_chart_tracking_flag() {
        let dop = valid_dop2013();
        let mut value = Dop2013::parse(&dop).unwrap();
        assert!(!value.chart_tracking_reference_based);
        value.chart_tracking_reference_based = true;
        let mut output = dop.clone();
        value.write_into(&mut output).unwrap();
        assert!(
            Dop2013::parse(&output)
                .unwrap()
                .chart_tracking_reference_based
        );
    }

    #[test]
    fn rejects_reserved_bits_and_short_input() {
        let mut dop = valid_dop2013();
        put_u32(&mut dop, DOP2010_SIZE, 2);
        assert!(Dop2013::parse(&dop).is_err());
        assert!(Dop2013::parse(&dop[..DOP2013_SIZE - 1]).is_err());
    }
}
