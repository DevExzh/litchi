//! Worksheet web-extension binding records (MS-XLSB 2.4.868).

use litchi_ooxml_common::web;
use std::collections::HashSet;

use crate::package::error::{Error, Result};
use crate::package::formula::{CellParsedFormula, FormulaParser, Token};
use crate::package::frt::{parse_formula_header, serialize_formula_header};
use crate::raw::{Records, Writer, kind};

const MAX_BINDINGS: usize = 65_536;
const MAX_APP_REF_CODE_UNITS: usize = 32_767;

/// The reference range encoded by a `BrtWebExtension` FRT formula.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Range {
    /// Index into the workbook's `ExternSheet` (`Xti`) collection.
    pub external_sheet_index: u16,
    pub first_row: u32,
    pub last_row: u32,
    pub first_column: u32,
    pub last_column: u32,
}

/// One binary worksheet-side Office Add-in binding.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Binding {
    pub application_reference: String,
    pub range: Range,
    /// Exact `FRTFormula`, retained for lossless authoring.
    pub formula: CellParsedFormula,
}

impl Binding {
    /// Construct and validate a binding from its native formula.
    ///
    /// `valid_external_sheet` must verify that the referenced XTI resolves to
    /// one internal worksheet (`firstSheet >= 0` and `firstSheet == lastSheet`).
    pub fn new(
        application_reference: impl Into<String>,
        formula: CellParsedFormula,
        valid_external_sheet: impl FnOnce(u16) -> bool,
    ) -> Result<Self> {
        let application_reference = application_reference.into();
        validate_app_ref(&application_reference)?;
        let range = range_from_formula(&formula)?;
        if !valid_external_sheet(range.external_sheet_index) {
            return Err(invalid(
                "BrtWebExtension",
                "formula does not reference one internal worksheet",
            ));
        }
        Ok(Self {
            application_reference,
            range,
            formula,
        })
    }

    /// Parse one `BrtWebExtension` payload.
    pub fn parse_payload(
        data: &[u8],
        valid_external_sheet: impl FnOnce(u16) -> bool,
    ) -> Result<Self> {
        let (mut formulas, consumed) = parse_formula_header(data, "BrtWebExtension", 1)?;
        if formulas.len() != 1 {
            return Err(invalid(
                "BrtWebExtension",
                "FRTHeader must contain exactly one formula",
            ));
        }
        let formula = formulas.pop().ok_or_else(|| {
            invalid(
                "BrtWebExtension",
                "FRTHeader must contain exactly one formula",
            )
        })?;
        let string_data = data.get(consumed..).ok_or(Error::InvalidLength {
            expected: consumed,
            found: data.len(),
        })?;
        let application_reference = parse_wide_string_exact(string_data)?;
        Self::new(application_reference, formula, valid_external_sheet)
    }

    /// Serialize one `BrtWebExtension` payload.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        validate_app_ref(&self.application_reference)?;
        if range_from_formula(&self.formula)? != self.range {
            return Err(invalid(
                "BrtWebExtension",
                "cached range disagrees with its binary formula",
            ));
        }
        let mut output = serialize_formula_header(std::slice::from_ref(&self.formula), 1)?;
        write_wide_string(&mut output, &self.application_reference)?;
        Ok(output)
    }
}

/// Parse a complete `WEBEXTENSIONS` record collection.
pub fn parse_xlsb_web_extension_bindings(
    records: &[u8],
    mut valid_external_sheet: impl FnMut(u16) -> bool,
) -> Result<Vec<Binding>> {
    let mut iterator = Records::new(records);
    let begin = iterator
        .next()
        .ok_or_else(|| Error::UnexpectedEndOfStream("WEBEXTENSIONS".to_string()))??;
    if begin.kind() != kind::BEGIN_WEB_EXTENSIONS || !begin.payload().is_empty() {
        return Err(invalid(
            "WEBEXTENSIONS",
            "collection must start with empty BrtBeginWebExtensions",
        ));
    }
    let mut bindings = Vec::new();
    let mut app_refs = HashSet::new();
    loop {
        let record = iterator
            .next()
            .ok_or_else(|| Error::UnexpectedEndOfStream("WEBEXTENSIONS".to_string()))??;
        match record.kind() {
            kind::WEB_EXTENSION => {
                if bindings.len() == MAX_BINDINGS {
                    return Err(invalid("WEBEXTENSIONS", "binding count exceeds 65,536"));
                }
                let binding =
                    Binding::parse_payload(record.payload(), |index| valid_external_sheet(index))?;
                if !app_refs.insert(binding.application_reference.clone()) {
                    return Err(invalid("WEBEXTENSIONS", "duplicate binding appRef"));
                }
                bindings.push(binding);
            },
            kind::END_WEB_EXTENSIONS => {
                if !record.payload().is_empty() {
                    return Err(invalid("BrtEndWebExtensions", "end record must be empty"));
                }
                if bindings.is_empty() {
                    return Err(invalid(
                        "WEBEXTENSIONS",
                        "collection requires at least one binding",
                    ));
                }
                if iterator.next().is_some() {
                    return Err(invalid(
                        "WEBEXTENSIONS",
                        "records follow BrtEndWebExtensions",
                    ));
                }
                return Ok(bindings);
            },
            other => {
                return Err(invalid(
                    "WEBEXTENSIONS",
                    format!("unexpected record 0x{other:04X}"),
                ));
            },
        }
    }
}

/// Serialize a complete `WEBEXTENSIONS` record collection.
pub fn write_xlsb_web_extension_bindings(bindings: &[Binding]) -> Result<Vec<u8>> {
    if bindings.is_empty() || bindings.len() > MAX_BINDINGS {
        return Err(invalid(
            "WEBEXTENSIONS",
            "binding count must be in 1..=65,536",
        ));
    }
    let mut app_refs = HashSet::with_capacity(bindings.len());
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    writer.write_record(kind::BEGIN_WEB_EXTENSIONS, &[])?;
    for binding in bindings {
        if !app_refs.insert(&binding.application_reference) {
            return Err(invalid("WEBEXTENSIONS", "duplicate binding appRef"));
        }
        writer.write_record(kind::WEB_EXTENSION, &binding.to_payload()?)?;
    }
    writer.write_record(kind::END_WEB_EXTENSIONS, &[])?;
    Ok(output)
}

/// Require every binary worksheet `appRef` to resolve to one package binding.
pub fn validate_xlsb_web_extension_apprefs<'a>(
    worksheet_bindings: &[Binding],
    package_bindings: impl IntoIterator<Item = &'a web::Binding>,
) -> Result<()> {
    PackageAppRefs::new(package_bindings)?.validate(worksheet_bindings)
}

pub(crate) struct PackageAppRefs<'a> {
    values: HashSet<&'a str>,
}

impl<'a> PackageAppRefs<'a> {
    pub(crate) fn new(
        package_bindings: impl IntoIterator<Item = &'a web::Binding>,
    ) -> Result<Self> {
        let mut values = HashSet::new();
        let mut count = 0usize;
        for binding in package_bindings {
            count = count
                .checked_add(1)
                .ok_or_else(|| invalid("MS-OWEXML bindings", "package binding count overflow"))?;
            if count > MAX_BINDINGS {
                return Err(invalid(
                    "MS-OWEXML bindings",
                    "package binding count exceeds 65,536",
                ));
            }
            if values.len() == values.capacity() {
                values.try_reserve(1).map_err(|_| {
                    invalid(
                        "MS-OWEXML bindings",
                        "unable to reserve binding validation memory",
                    )
                })?;
            }
            if !values.insert(binding.app_ref()) {
                return Err(invalid(
                    "MS-OWEXML bindings",
                    "duplicate package binding appref",
                ));
            }
        }
        Ok(Self { values })
    }

    pub(crate) fn validate(&self, worksheet_bindings: &[Binding]) -> Result<()> {
        if worksheet_bindings.len() > MAX_BINDINGS {
            return Err(invalid(
                "WEBEXTENSIONS",
                "worksheet binding count exceeds 65,536",
            ));
        }
        for binding in worksheet_bindings {
            if !self.values.contains(binding.application_reference.as_str()) {
                return Err(invalid(
                    "BrtWebExtension.appRef",
                    format!(
                        "'{}' has no matching MS-OWEXML binding",
                        binding.application_reference
                    ),
                ));
            }
        }
        Ok(())
    }
}

fn range_from_formula(formula: &CellParsedFormula) -> Result<Range> {
    if formula
        .rgce
        .first()
        .is_none_or(|token| token & 0x60 != 0x20)
    {
        return Err(invalid(
            "BrtWebExtension",
            "binding formula root must use the REFERENCE operand class",
        ));
    }
    let tokens = FormulaParser::with_extra(&formula.rgce, &formula.rgcb).parse()?;
    if tokens.len() != 1 {
        return Err(invalid(
            "BrtWebExtension",
            "binding formula must be one reference expression",
        ));
    }
    let token = tokens.into_iter().next().ok_or_else(|| {
        invalid(
            "BrtWebExtension",
            "binding formula must be one reference expression",
        )
    })?;
    match token {
        Token::CellRef3d {
            sheet_index,
            row,
            col,
            ..
        } => Ok(Range {
            external_sheet_index: sheet_index,
            first_row: row,
            last_row: row,
            first_column: col,
            last_column: col,
        }),
        Token::AreaRef3d {
            sheet_index,
            row_first,
            row_last,
            col_first,
            col_last,
            ..
        } => Ok(Range {
            external_sheet_index: sheet_index,
            first_row: row_first,
            last_row: row_last,
            first_column: col_first,
            last_column: col_last,
        }),
        Token::CellRef { .. } | Token::AreaRef { .. } | Token::ReferenceError { .. } => {
            Err(invalid(
                "BrtWebExtension",
                "local and invalid reference tokens are forbidden",
            ))
        },
        _ => Err(invalid(
            "BrtWebExtension",
            "binding formula root is not a 3D reference",
        )),
    }
}

fn validate_app_ref(value: &str) -> Result<()> {
    let units = value.encode_utf16().count();
    if units == 0 || units > MAX_APP_REF_CODE_UNITS {
        return Err(invalid(
            "BrtWebExtension.appRef",
            "length must be in 1..=32,767 UTF-16 code units",
        ));
    }
    if value.chars().any(char::is_control) {
        return Err(invalid(
            "BrtWebExtension.appRef",
            "control characters are forbidden",
        ));
    }
    Ok(())
}

fn parse_wide_string_exact(data: &[u8]) -> Result<String> {
    let length_bytes = data.get(..4).ok_or(Error::InvalidLength {
        expected: 4,
        found: data.len(),
    })?;
    let length_bytes = <[u8; 4]>::try_from(length_bytes).map_err(|_| Error::InvalidLength {
        expected: 4,
        found: data.len(),
    })?;
    let count = u32::from_le_bytes(length_bytes) as usize;
    if count == 0 || count > MAX_APP_REF_CODE_UNITS {
        return Err(invalid("BrtWebExtension.appRef", "invalid string length"));
    }
    let expected = 4usize
        .checked_add(
            count
                .checked_mul(2)
                .ok_or_else(|| invalid("BrtWebExtension.appRef", "length overflow"))?,
        )
        .ok_or_else(|| invalid("BrtWebExtension.appRef", "length overflow"))?;
    if data.len() != expected {
        return Err(Error::InvalidLength {
            expected,
            found: data.len(),
        });
    }
    let encoded = data.get(4..).ok_or(Error::InvalidLength {
        expected,
        found: data.len(),
    })?;
    let units = encoded
        .chunks_exact(2)
        .map(|bytes| {
            <[u8; 2]>::try_from(bytes)
                .map(u16::from_le_bytes)
                .map_err(|_| invalid("BrtWebExtension.appRef", "invalid UTF-16 unit"))
        })
        .collect::<Result<Vec<_>>>()?;
    String::from_utf16(&units)
        .map_err(|_| invalid("BrtWebExtension.appRef", "invalid UTF-16 string"))
}

fn write_wide_string(output: &mut Vec<u8>, value: &str) -> Result<()> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if units.is_empty() || units.len() > MAX_APP_REF_CODE_UNITS {
        return Err(invalid("BrtWebExtension.appRef", "invalid string length"));
    }
    output.extend_from_slice(&(units.len() as u32).to_le_bytes());
    output.reserve(units.len() * 2);
    for unit in units {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}

fn invalid(typ: impl Into<String>, val: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: typ.into(),
        val: val.into(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::formula::FormulaCompiler;

    struct LyingHint<'a> {
        value: Option<&'a web::Binding>,
    }

    impl<'a> Iterator for LyingHint<'a> {
        type Item = &'a web::Binding;

        fn next(&mut self) -> Option<Self::Item> {
            self.value.take()
        }

        fn size_hint(&self) -> (usize, Option<usize>) {
            (usize::MAX, None)
        }
    }

    fn binding() -> Binding {
        // Public context-free compilation intentionally rejects 3D formulas;
        // construct the canonical PtgArea3d token directly.
        let binary = CellParsedFormula {
            rgce: vec![0x3B, 0, 0, 0, 0, 0, 0, 3, 0, 0, 0, 0, 0, 1, 0],
            rgcb: Vec::new(),
        };
        Binding::new("sales-table", binary, |index| index == 0).unwrap()
    }

    #[test]
    fn payload_and_collection_roundtrip() {
        let binding = binding();
        let payload = binding.to_payload().unwrap();
        assert_eq!(
            Binding::parse_payload(&payload, |index| index == 0).unwrap(),
            binding
        );
        let collection = write_xlsb_web_extension_bindings(std::slice::from_ref(&binding)).unwrap();
        assert_eq!(
            parse_xlsb_web_extension_bindings(&collection, |index| index == 0).unwrap(),
            [binding]
        );
    }

    #[test]
    fn rejects_invalid_xti_local_refs_and_trailing_payload() {
        let binding = binding();
        let payload = binding.to_payload().unwrap();
        assert!(Binding::parse_payload(&payload, |_| false).is_err());
        let local = FormulaCompiler::compile("$A$1:$B$4").unwrap();
        assert!(Binding::new("local", local, |_| true).is_err());
        let mut trailing = payload;
        trailing.push(0);
        assert!(Binding::parse_payload(&trailing, |_| true).is_err());
    }

    #[test]
    fn validates_package_appref_links() {
        let worksheet = [binding()];
        let package = [web::Binding::new("id", "table", "sales-table").unwrap()];
        validate_xlsb_web_extension_apprefs(&worksheet, &package).unwrap();
        validate_xlsb_web_extension_apprefs(
            &worksheet,
            LyingHint {
                value: Some(&package[0]),
            },
        )
        .unwrap();
        assert!(validate_xlsb_web_extension_apprefs(&worksheet, &[]).is_err());
    }
}
