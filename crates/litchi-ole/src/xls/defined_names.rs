//! BIFF8 internal defined-name (`Lbl`) parsing and public models.

use super::error::{XlsError, XlsResult};
use super::formula::{FormulaContext, render_formula};

pub(crate) const LBL_RECORD_TYPE: u16 = 0x0018;
const LBL_HEADER_LEN: usize = 14;
const FLAG_HIDDEN: u16 = 0x0001;
const FLAG_FUNCTION: u16 = 0x0002;
const FLAG_VBA: u16 = 0x0004;
const FLAG_PROCEDURE: u16 = 0x0008;
const FLAG_BUILT_IN: u16 = 0x0020;
const RESERVED_FLAG_MASK: u16 = 0x9000;

/// A non-macro internal defined name from the workbook globals substream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsDefinedName {
    /// One-based position in the complete `Lbl` collection, including macro slots.
    pub record_index: u32,
    pub name: String,
    pub scope: XlsNameScope,
    pub hidden: bool,
    pub kind: XlsDefinedNameKind,
    /// Rendered formula using the same leading-`=` convention as cell formulas.
    pub formula: Option<String>,
    /// Original `NameParsedFormula.rgce` bytes.
    pub formula_tokens: Vec<u8>,
}

impl XlsDefinedName {
    /// Whether the rendered definition contains a deleted reference.
    pub fn is_deleted(&self) -> bool {
        self.formula
            .as_deref()
            .is_some_and(|formula| formula.contains("#REF!"))
    }
}

/// Scope of a defined name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsNameScope {
    Workbook,
    Worksheet(usize),
}

/// User-defined or reserved built-in name kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsDefinedNameKind {
    User,
    BuiltIn(XlsBuiltInName),
}

/// Built-in BIFF8 defined-name identifiers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsBuiltInName {
    ConsolidateArea,
    AutoOpen,
    AutoClose,
    Extract,
    Database,
    Criteria,
    PrintArea,
    PrintTitles,
    Recorder,
    DataForm,
    AutoActivate,
    AutoDeactivate,
    SheetTitle,
    FilterDatabase,
}

impl XlsBuiltInName {
    fn from_code(code: u8) -> Option<Self> {
        Some(match code {
            0x00 => Self::ConsolidateArea,
            0x01 => Self::AutoOpen,
            0x02 => Self::AutoClose,
            0x03 => Self::Extract,
            0x04 => Self::Database,
            0x05 => Self::Criteria,
            0x06 => Self::PrintArea,
            0x07 => Self::PrintTitles,
            0x08 => Self::Recorder,
            0x09 => Self::DataForm,
            0x0A => Self::AutoActivate,
            0x0B => Self::AutoDeactivate,
            0x0C => Self::SheetTitle,
            0x0D => Self::FilterDatabase,
            _ => return None,
        })
    }

    fn canonical_name(self) -> &'static str {
        match self {
            Self::ConsolidateArea => "Consolidate_Area",
            Self::AutoOpen => "Auto_Open",
            Self::AutoClose => "Auto_Close",
            Self::Extract => "Extract",
            Self::Database => "Database",
            Self::Criteria => "Criteria",
            Self::PrintArea => "Print_Area",
            Self::PrintTitles => "Print_Titles",
            Self::Recorder => "Recorder",
            Self::DataForm => "Data_Form",
            Self::AutoActivate => "Auto_Activate",
            Self::AutoDeactivate => "Auto_Deactivate",
            Self::SheetTitle => "Sheet_Title",
            Self::FilterDatabase => "_FilterDatabase",
        }
    }
}

/// One parsed `Lbl` slot. Macro slots remain here so `PtgName` indices do not shift.
#[derive(Debug, Clone)]
pub(crate) struct DefinedNameSlot {
    record_index: u32,
    name: String,
    itab: u16,
    hidden: bool,
    is_macro: bool,
    kind: XlsDefinedNameKind,
    formula_tokens: Vec<u8>,
}

impl DefinedNameSlot {
    pub(crate) fn parse(data: &[u8], record_index: u32) -> XlsResult<Self> {
        if data.len() < LBL_HEADER_LEN + 1 {
            return invalid("Lbl is missing its fixed header or name flags");
        }
        let flags = u16::from_le_bytes([data[0], data[1]]);
        if flags & RESERVED_FLAG_MASK != 0 {
            return invalid("Lbl contains nonzero reserved option flags");
        }
        let function = flags & FLAG_FUNCTION != 0;
        let vba = flags & FLAG_VBA != 0;
        let procedure = flags & FLAG_PROCEDURE != 0;
        if (function || vba) && !procedure {
            return invalid("Lbl function/VBA flags require the procedure flag");
        }
        let function_group = (flags >> 6) & 0x3f;
        if function_group > 31 {
            return invalid("Lbl function group must be at most 31");
        }
        let shortcut = data[2];
        if (function || !procedure) && shortcut != 0 {
            return invalid("Lbl shortcut key is not valid for these macro flags");
        }
        if procedure
            && !function
            && shortcut != 0
            && !shortcut.is_ascii_alphabetic()
        {
            return invalid("Lbl macro shortcut key must be an ASCII letter");
        }

        let character_count = usize::from(data[3]);
        let formula_len = usize::from(u16::from_le_bytes([data[4], data[5]]));
        let itab = u16::from_le_bytes([data[8], data[9]]);
        let string_flags = data[LBL_HEADER_LEN];
        if string_flags & !0x01 != 0 {
            return invalid("Lbl name contains unsupported Unicode flags");
        }
        let char_width = if string_flags & 0x01 == 0 { 1 } else { 2 };
        let name_byte_len = character_count
            .checked_mul(char_width)
            .ok_or_else(|| invalid_error("Lbl name length overflows"))?;
        let name_start = LBL_HEADER_LEN + 1;
        let name_end = name_start
            .checked_add(name_byte_len)
            .ok_or_else(|| invalid_error("Lbl name range overflows"))?;
        let name_bytes = data
            .get(name_start..name_end)
            .ok_or_else(|| invalid_error("Lbl name is truncated"))?;
        let decoded_name = if char_width == 1 {
            name_bytes.iter().map(|byte| char::from(*byte)).collect()
        } else {
            let units: Vec<u16> = name_bytes
                .chunks_exact(2)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                .collect();
            String::from_utf16(&units)
                .map_err(|_| invalid_error("Lbl name contains invalid UTF-16"))?
        };

        let built_in = flags & FLAG_BUILT_IN != 0;
        let (name, kind) = if built_in {
            if character_count != 1 {
                return invalid("built-in Lbl name must contain exactly one character");
            }
            let code = decoded_name
                .encode_utf16()
                .next()
                .and_then(|value| u8::try_from(value).ok())
                .ok_or_else(|| invalid_error("built-in Lbl identifier is invalid"))?;
            let built_in = XlsBuiltInName::from_code(code)
                .ok_or_else(|| invalid_error("built-in Lbl identifier is out of range"))?;
            (
                built_in.canonical_name().to_string(),
                XlsDefinedNameKind::BuiltIn(built_in),
            )
        } else {
            if decoded_name.is_empty() || decoded_name.contains('\0') {
                return invalid("user-defined Lbl name must be nonempty and contain no NUL");
            }
            (decoded_name, XlsDefinedNameKind::User)
        };

        let formula_end = name_end
            .checked_add(formula_len)
            .ok_or_else(|| invalid_error("Lbl formula range overflows"))?;
        let formula_tokens = data
            .get(name_end..formula_end)
            .ok_or_else(|| invalid_error("Lbl formula is truncated"))?
            .to_vec();

        Ok(Self {
            record_index,
            name,
            itab,
            hidden: flags & FLAG_HIDDEN != 0,
            is_macro: function || vba || procedure,
            kind,
            formula_tokens,
        })
    }

    /// Symbol table entry. Macro slots intentionally remain `None`.
    pub(crate) fn symbol(&self) -> Option<String> {
        (!self.is_macro).then(|| self.name.clone())
    }

    pub(crate) fn into_public(
        self,
        sheet_count: usize,
        context: &FormulaContext,
    ) -> XlsResult<Option<XlsDefinedName>> {
        let scope = if self.itab == 0 {
            XlsNameScope::Workbook
        } else {
            let sheet_index = usize::from(self.itab - 1);
            if sheet_index >= sheet_count {
                return invalid("Lbl.itab is outside the BoundSheet8 collection");
            }
            XlsNameScope::Worksheet(sheet_index)
        };
        if self.is_macro {
            return Ok(None);
        }
        let formula = (!self.formula_tokens.is_empty())
            .then(|| render_formula(&self.formula_tokens, Some(context)))
            .flatten();
        Ok(Some(XlsDefinedName {
            record_index: self.record_index,
            name: self.name,
            scope,
            hidden: self.hidden,
            kind: self.kind,
            formula,
            formula_tokens: self.formula_tokens,
        }))
    }
}

fn invalid<T>(message: &str) -> XlsResult<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: &str) -> XlsError {
    XlsError::InvalidRecord {
        record_type: LBL_RECORD_TYPE,
        message: message.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn lbl(flags: u16, itab: u16, name: &str, unicode: bool, formula: &[u8]) -> Vec<u8> {
        let units: Vec<u16> = name.encode_utf16().collect();
        let cch = if unicode { units.len() } else { name.len() };
        let mut data = Vec::new();
        data.extend_from_slice(&flags.to_le_bytes());
        data.push(0);
        data.push(cch as u8);
        data.extend_from_slice(&(formula.len() as u16).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&itab.to_le_bytes());
        data.extend_from_slice(&[0; 4]);
        data.push(u8::from(unicode));
        if unicode {
            for unit in units {
                data.extend_from_slice(&unit.to_le_bytes());
            }
        } else {
            data.extend_from_slice(name.as_bytes());
        }
        data.extend_from_slice(formula);
        data
    }

    #[test]
    fn parses_compressed_unicode_hidden_and_scoped_names() {
        let ordinary = DefinedNameSlot::parse(&lbl(FLAG_HIDDEN, 2, "Rate", false, &[0x1e, 2, 0]), 1)
            .unwrap();
        assert_eq!(ordinary.name, "Rate");
        assert!(ordinary.hidden);
        assert_eq!(ordinary.itab, 2);
        assert_eq!(ordinary.formula_tokens, [0x1e, 2, 0]);

        let unicode = DefinedNameSlot::parse(&lbl(0, 0, "税率", true, &[0x1e, 3, 0]), 2)
            .unwrap();
        assert_eq!(unicode.name, "税率");
    }

    #[test]
    fn parses_every_built_in_name() {
        for code in 0u8..=0x0d {
            let slot = DefinedNameSlot::parse(
                &lbl(FLAG_BUILT_IN, 1, &char::from(code).to_string(), false, &[0x1e, 1, 0]),
                u32::from(code) + 1,
            )
            .unwrap();
            assert!(matches!(slot.kind, XlsDefinedNameKind::BuiltIn(_)));
        }
    }

    #[test]
    fn macro_slots_keep_indices_but_have_no_symbol_or_public_value() {
        let slot = DefinedNameSlot::parse(&lbl(FLAG_PROCEDURE, 0, "Macro", false, &[]), 7)
            .unwrap();
        assert!(slot.symbol().is_none());
        assert!(slot
            .into_public(1, &FormulaContext::default())
            .unwrap()
            .is_none());
    }

    #[test]
    fn rejects_malformed_names_and_scope() {
        assert!(DefinedNameSlot::parse(&[0; 14], 1).is_err());
        let mut truncated = lbl(0, 0, "Name", true, &[0x1e, 1, 0]);
        truncated.truncate(17);
        assert!(DefinedNameSlot::parse(&truncated, 1).is_err());
        assert!(DefinedNameSlot::parse(&lbl(FLAG_BUILT_IN, 0, "x", false, &[]), 1).is_err());
        let invalid_scope = DefinedNameSlot::parse(&lbl(0, 3, "Name", false, &[]), 1).unwrap();
        assert!(invalid_scope
            .into_public(2, &FormulaContext::default())
            .is_err());
    }
}
