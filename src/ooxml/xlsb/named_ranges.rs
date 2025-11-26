//! Named range support for XLSB

use crate::common::binary;
use crate::ooxml::xlsb::error::{XlsbError, XlsbResult};
use crate::ooxml::xlsb::records::wide_str_with_len;

/// Named range definition
///
/// Represents a defined name (named range) in the workbook.
#[derive(Debug, Clone)]
pub struct NamedRange {
    /// Name of the range
    pub name: String,
    /// Formula defining the range (as raw bytes)
    pub formula: Option<Vec<u8>>,
    /// Sheet ID (None for global scope)
    pub sheet_id: Option<u32>,
    /// Whether the name is hidden
    pub hidden: bool,
    /// Whether the name is a function
    pub function: bool,
}

impl NamedRange {
    /// Create a new named range
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::xlsb::named_ranges::NamedRange;
    ///
    /// let range = NamedRange::new("MyRange".to_string(), None);
    /// ```
    pub fn new(name: String, sheet_id: Option<u32>) -> Self {
        NamedRange {
            name,
            formula: None,
            sheet_id,
            hidden: false,
            function: false,
        }
    }

    /// Set formula bytes
    pub fn with_formula(mut self, formula: Vec<u8>) -> Self {
        self.formula = Some(formula);
        self
    }

    /// Set hidden flag
    pub fn with_hidden(mut self, hidden: bool) -> Self {
        self.hidden = hidden;
        self
    }

    /// Parse from XLSB BrtName record
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 8 {
            return Err(XlsbError::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }

        let flags = binary::read_u32_le_at(data, 0)?;
        let hidden = (flags & 0x0001) != 0;
        let function = (flags & 0x0002) != 0;

        // Sheet ID (-1 for global scope, otherwise sheet-specific)
        let sheet_id_raw = binary::read_u32_le_at(data, 4)? as i32;
        let sheet_id = if sheet_id_raw == -1 {
            None
        } else {
            Some(sheet_id_raw as u32)
        };

        let mut offset = 8;

        // Read name
        let (name, consumed) = wide_str_with_len(&data[offset..])?;
        offset += consumed;

        // Read formula if present
        let formula = if offset < data.len() {
            let formula_len = binary::read_u32_le_at(data, offset)? as usize;
            offset += 4;
            if data.len() >= offset + formula_len {
                Some(data[offset..offset + formula_len].to_vec())
            } else {
                None
            }
        } else {
            None
        };

        Ok(NamedRange {
            name,
            formula,
            sheet_id,
            hidden,
            function,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_named_range_builder() {
        let range = NamedRange::new("MyRange".to_string(), None)
            .with_hidden(true)
            .with_formula(vec![1, 2, 3]);

        assert_eq!(range.name, "MyRange");
        assert!(range.hidden);
        assert_eq!(range.formula, Some(vec![1, 2, 3]));
    }
}
