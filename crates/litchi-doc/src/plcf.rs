//! Borrowed property lists with character positions (PLCF).

use litchi_core::binary;

/// A validated, zero-copy view over a DOC PLCF record.
pub(crate) struct Plcf<'data> {
    data: &'data [u8],
    element_size: usize,
    count: usize,
    properties_start: usize,
}

impl<'data> Plcf<'data> {
    pub(crate) fn parse(data: &'data [u8], element_size: usize) -> Option<Self> {
        if element_size == 0 {
            return None;
        }

        let payload_len = data.len().checked_sub(4)?;
        let stride = 4usize.checked_add(element_size)?;
        let count = payload_len / stride;
        let positions_len = count.checked_add(1)?.checked_mul(4)?;
        let properties_len = count.checked_mul(element_size)?;
        let end = positions_len.checked_add(properties_len)?;
        data.get(..end)?;

        Some(Self {
            data,
            element_size,
            count,
            properties_start: positions_len,
        })
    }

    #[inline]
    pub(crate) const fn count(&self) -> usize {
        self.count
    }

    #[inline]
    pub(crate) fn position(&self, index: usize) -> Option<u32> {
        if index > self.count {
            return None;
        }
        binary::read_u32_le(self.data, index.checked_mul(4)?).ok()
    }

    #[inline]
    pub(crate) fn property(&self, index: usize) -> Option<&'data [u8]> {
        if index >= self.count {
            return None;
        }
        let start = self
            .properties_start
            .checked_add(index.checked_mul(self.element_size)?)?;
        let end = start.checked_add(self.element_size)?;
        self.data.get(start..end)
    }

    pub(crate) fn range(&self, index: usize) -> Option<(u32, u32)> {
        Some((self.position(index)?, self.position(index.checked_add(1)?)?))
    }
}

#[cfg(test)]
mod tests {
    use super::Plcf;

    #[test]
    fn parses_without_copying_properties() {
        let data = [
            0x00, 0x00, 0x00, 0x00, // CP 0
            0x0A, 0x00, 0x00, 0x00, // CP 10
            0x14, 0x00, 0x00, 0x00, // CP 20
            0x01, 0x02, // property 1
            0x03, 0x04, // property 2
        ];

        let plcf = Plcf::parse(&data, 2).unwrap();
        assert_eq!(plcf.count(), 2);
        assert_eq!(plcf.position(0), Some(0));
        assert_eq!(plcf.position(2), Some(20));
        assert_eq!(plcf.property(1), Some(&data[14..16]));
        assert_eq!(plcf.range(1), Some((10, 20)));
    }

    #[test]
    fn rejects_invalid_layouts_and_indices_without_panicking() {
        assert!(Plcf::parse(&[], 2).is_none());
        assert!(Plcf::parse(&[0; 4], 0).is_none());
        assert!(Plcf::parse(&[0; 4], usize::MAX).is_none());

        let plcf = Plcf::parse(&[0; 4], 2).unwrap();
        assert_eq!(plcf.count(), 0);
        assert_eq!(plcf.position(1), None);
        assert_eq!(plcf.property(usize::MAX), None);
        assert_eq!(plcf.range(usize::MAX), None);
    }
}
