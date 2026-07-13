//! XLSB workbook calculation properties (`BrtCalcProp`).

use crate::xlsb::error::{XlsbError, XlsbResult};
use litchi_core::binary;

/// Workbook calculation mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[repr(u32)]
pub enum CalculationMode {
    Manual = 0,
    #[default]
    Automatic = 1,
    AutomaticExceptTables = 2,
}

/// Calculation policy stored in the workbook's `BrtCalcProp` record.
#[derive(Debug, Clone, PartialEq)]
pub struct CalculationProperties {
    pub recalculation_id: u32,
    pub mode: CalculationMode,
    pub iteration_count: u32,
    pub iteration_delta: f64,
    pub user_thread_count: i32,
    pub full_calculation_on_load: bool,
    pub a1_references: bool,
    pub iterative_calculation: bool,
    pub full_precision: bool,
    pub some_formulas_uncalculated: bool,
    pub recalculate_before_save: bool,
    pub multithreaded_calculation: bool,
    pub user_set_thread_count: bool,
    pub ignore_dependencies: bool,
}

impl Default for CalculationProperties {
    fn default() -> Self {
        Self {
            recalculation_id: 0x0001_EB1D,
            mode: CalculationMode::Automatic,
            iteration_count: 100,
            iteration_delta: 0.001,
            user_thread_count: 1,
            full_calculation_on_load: false,
            a1_references: true,
            iterative_calculation: false,
            full_precision: true,
            some_formulas_uncalculated: false,
            recalculate_before_save: true,
            multithreaded_calculation: true,
            user_set_thread_count: false,
            ignore_dependencies: false,
        }
    }
}

impl CalculationProperties {
    pub(crate) fn parse(data: &[u8]) -> XlsbResult<Self> {
        if !matches!(data.len(), 25 | 26) {
            return Err(XlsbError::InvalidLength {
                expected: 26,
                found: data.len(),
            });
        }
        let mode = match binary::read_u32_le_at(data, 4)? {
            0 => CalculationMode::Manual,
            1 => CalculationMode::Automatic,
            2 => CalculationMode::AutomaticExceptTables,
            value => {
                return Err(XlsbError::Unrecognized {
                    typ: "BrtCalcProp fAutoRecalc".to_string(),
                    val: value.to_string(),
                });
            },
        };
        let flags = if data.len() == 25 {
            u16::from(data[24])
        } else {
            binary::read_u16_le_at(data, 24)?
        };
        if data.len() == 26 && flags & !0x01FF != 0 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtCalcProp flags".to_string(),
                val: format!("0x{flags:04X}"),
            });
        }
        let properties = Self {
            recalculation_id: binary::read_u32_le_at(data, 0)?,
            mode,
            iteration_count: binary::read_u32_le_at(data, 8)?,
            iteration_delta: binary::read_f64_le_at(data, 12)?,
            user_thread_count: binary::read_u32_le_at(data, 20)? as i32,
            full_calculation_on_load: flags & 0x0001 != 0,
            a1_references: flags & 0x0002 != 0,
            iterative_calculation: flags & 0x0004 != 0,
            full_precision: flags & 0x0008 != 0,
            some_formulas_uncalculated: flags & 0x0010 != 0,
            recalculate_before_save: flags & 0x0020 != 0,
            multithreaded_calculation: flags & 0x0040 != 0,
            user_set_thread_count: flags & 0x0080 != 0,
            ignore_dependencies: flags & 0x0100 != 0,
        };
        properties.validate()?;
        Ok(properties)
    }

    pub(crate) fn validate(&self) -> XlsbResult<()> {
        if !self.iteration_delta.is_finite() || self.iteration_delta < 0.0 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtCalcProp xnumDelta".to_string(),
                val: self.iteration_delta.to_string(),
            });
        }
        if self.iterative_calculation && self.iteration_count == 0 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtCalcProp cCalcCount".to_string(),
                val: "0 while iterative calculation is enabled".to_string(),
            });
        }
        if self.user_set_thread_count
            && self.multithreaded_calculation
            && !(1..=1024).contains(&self.user_thread_count)
        {
            return Err(XlsbError::Unrecognized {
                typ: "BrtCalcProp cUserThreadCount".to_string(),
                val: self.user_thread_count.to_string(),
            });
        }
        Ok(())
    }

    pub(crate) fn flags(&self) -> u16 {
        u16::from(self.full_calculation_on_load)
            | (u16::from(self.a1_references) << 1)
            | (u16::from(self.iterative_calculation) << 2)
            | (u16::from(self.full_precision) << 3)
            | (u16::from(self.some_formulas_uncalculated) << 4)
            | (u16::from(self.recalculate_before_save) << 5)
            | (u16::from(self.multithreaded_calculation) << 6)
            | (u16::from(self.user_set_thread_count) << 7)
            | (u16::from(self.ignore_dependencies) << 8)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_default_calculation_properties() {
        let expected = CalculationProperties::default();
        let mut data = Vec::new();
        data.extend_from_slice(&expected.recalculation_id.to_le_bytes());
        data.extend_from_slice(&(expected.mode as u32).to_le_bytes());
        data.extend_from_slice(&expected.iteration_count.to_le_bytes());
        data.extend_from_slice(&expected.iteration_delta.to_le_bytes());
        data.extend_from_slice(&expected.user_thread_count.to_le_bytes());
        data.extend_from_slice(&expected.flags().to_le_bytes());
        assert_eq!(CalculationProperties::parse(&data).unwrap(), expected);
    }

    #[test]
    fn rejects_invalid_calculation_properties() {
        let mut data = vec![0; 26];
        data[4..8].copy_from_slice(&3u32.to_le_bytes());
        assert!(matches!(
            CalculationProperties::parse(&data),
            Err(XlsbError::Unrecognized { .. })
        ));
    }

    #[test]
    fn parses_excel_beta_single_byte_flags() {
        let data = [
            0x63, 0xDD, 0x01, 0x00, 1, 0, 0, 0, 100, 0, 0, 0, 0xFC, 0xA9, 0xF1, 0xD2, 0x4D, 0x62,
            0x50, 0x3F, 1, 0, 0, 0, 0x6A,
        ];
        let properties = CalculationProperties::parse(&data).unwrap();
        assert_eq!(properties.mode, CalculationMode::Automatic);
        assert!(properties.a1_references);
        assert!(properties.full_precision);
        assert!(properties.recalculate_before_save);
        assert!(properties.multithreaded_calculation);
        assert!(!properties.ignore_dependencies);
    }
}
