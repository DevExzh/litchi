//! Typed, lossless Word 2010 Document Properties extension.

use super::document_properties_97::DopExtensionError;
use super::document_properties_2007::Dop2007;

const DOP2007_SIZE: usize = 674;
const DOP2010_SIZE: usize = 690;
const EXTENSION_SIZE: usize = DOP2010_SIZE - DOP2007_SIZE;
const DOCUMENT_ID: usize = 0;
const DISCARD_IMAGE_FLAGS: usize = 8;
const IMAGE_DPI: usize = 12;

/// A nonzero 31-bit identifier establishing the paragraph-ID context.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct DocumentId(u32);

impl DocumentId {
    /// Creates an identifier in the specification-defined nonzero 31-bit domain.
    ///
    /// # Errors
    ///
    /// Returns [`DopExtensionError`] for zero or a value with bit 31 set.
    pub fn try_new(value: u32) -> Result<Self, DopExtensionError> {
        if (1..0x8000_0000).contains(&value) {
            Ok(Self(value))
        } else {
            Err(DopExtensionError::new(format!(
                "Dop2010 document id {value:#010x} is outside 1..0x80000000"
            )))
        }
    }

    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }
}

/// Typed, lossless Word 2010 DOP extension.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Dop2010 {
    raw: [u8; EXTENSION_SIZE],
    pub document_id: DocumentId,
    pub discard_cropped_image_data: bool,
    /// Resolution Word uses when saving document images.
    pub image_dpi: u32,
}

impl Dop2010 {
    /// Parses the complete Word 2010 DOP generation.
    ///
    /// # Errors
    ///
    /// Returns [`DopExtensionError`] when this or an earlier DOP prefix is
    /// invalid, the document ID is out of range, or reserved bits are set.
    pub fn parse(dop: &[u8]) -> Result<Self, DopExtensionError> {
        if dop.len() < DOP2010_SIZE {
            return Err(DopExtensionError::new("Dop2010 is shorter than 690 bytes"));
        }
        Dop2007::parse(dop)?;
        let extension = &dop[DOP2007_SIZE..DOP2010_SIZE];
        let flags = le_u32(extension, DISCARD_IMAGE_FLAGS);
        if flags & !1 != 0 {
            return Err(DopExtensionError::new(
                "Dop2010 reserved discard-image bits are nonzero",
            ));
        }
        let mut raw = [0u8; EXTENSION_SIZE];
        raw.copy_from_slice(extension);
        Ok(Self {
            raw,
            document_id: DocumentId::try_new(le_u32(extension, DOCUMENT_ID))?,
            discard_cropped_image_data: flags & 1 != 0,
            image_dpi: le_u32(extension, IMAGE_DPI),
        })
    }

    /// Writes the Word 2010 extension without normalizing older DOP bytes.
    ///
    /// # Errors
    ///
    /// Returns [`DopExtensionError`] when the target is too short or the
    /// document ID is outside its specification-defined domain.
    pub fn write_into(mut self, dop: &mut [u8]) -> Result<(), DopExtensionError> {
        if dop.len() < DOP2010_SIZE {
            return Err(DopExtensionError::new(
                "Dop2010 target is shorter than 690 bytes",
            ));
        }
        DocumentId::try_new(self.document_id.get())?;
        put_u32(&mut self.raw, DOCUMENT_ID, self.document_id.get());
        put_u32(
            &mut self.raw,
            DISCARD_IMAGE_FLAGS,
            u32::from(self.discard_cropped_image_data),
        );
        put_u32(&mut self.raw, IMAGE_DPI, self.image_dpi);
        dop[DOP2007_SIZE..DOP2010_SIZE].copy_from_slice(&self.raw);
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

    fn valid_dop2010() -> Vec<u8> {
        let mut dop = vec![0u8; DOP2010_SIZE];
        dop[0x190..0x19a].copy_from_slice(&[0xa5, 0x06, 0xc0, 0x07, 0xb4, 0, 0xb4, 0, 1, 0x81]);
        let math = 616 + 24;
        put_u32(&mut dop, math, 1 << 4 | 1 << 11 | 1 << 12);
        dop[math + 14..math + 18].copy_from_slice(&120i32.to_le_bytes());
        dop[math + 18..math + 22].copy_from_slice(&120i32.to_le_bytes());
        put_u32(&mut dop, DOP2007_SIZE + DOCUMENT_ID, 0x1234_5678);
        dop
    }

    #[test]
    fn parses_round_trips_and_mutates_word_2010_settings() {
        let mut dop = valid_dop2010();
        dop[DOP2007_SIZE + 4..DOP2007_SIZE + 8].copy_from_slice(&[0xde, 0xad, 0xbe, 0xef]);
        put_u32(&mut dop, DOP2007_SIZE + DISCARD_IMAGE_FLAGS, 1);
        put_u32(&mut dop, DOP2007_SIZE + IMAGE_DPI, 220);
        let mut parsed = Dop2010::parse(&dop).unwrap();
        assert_eq!(parsed.document_id.get(), 0x1234_5678);
        assert!(parsed.discard_cropped_image_data);
        parsed.image_dpi = 300;
        let mut output = dop.clone();
        parsed.write_into(&mut output).unwrap();
        assert_eq!(
            &output[DOP2007_SIZE + 4..DOP2007_SIZE + 8],
            &[0xde, 0xad, 0xbe, 0xef]
        );
        assert_eq!(Dop2010::parse(&output).unwrap().image_dpi, 300);
    }

    #[test]
    fn rejects_invalid_document_id_reserved_bits_and_short_input() {
        let mut zero = valid_dop2010();
        put_u32(&mut zero, DOP2007_SIZE + DOCUMENT_ID, 0);
        assert!(Dop2010::parse(&zero).is_err());

        let mut high = valid_dop2010();
        put_u32(&mut high, DOP2007_SIZE + DOCUMENT_ID, 0x8000_0000);
        assert!(Dop2010::parse(&high).is_err());

        let mut flags = valid_dop2010();
        put_u32(&mut flags, DOP2007_SIZE + DISCARD_IMAGE_FLAGS, 2);
        assert!(Dop2010::parse(&flags).is_err());
        assert!(Dop2010::parse(&flags[..DOP2010_SIZE - 1]).is_err());
    }
}
