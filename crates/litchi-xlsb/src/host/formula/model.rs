//! Lossless host metadata used by XLSB formulas.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExternalSheet {
    pub external_link: u32,
    pub first_sheet: i32,
    pub last_sheet: i32,
}
