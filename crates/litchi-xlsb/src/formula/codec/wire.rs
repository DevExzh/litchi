//! BIFF12 formula values and Parse Tree Generator (Ptg) codecs.
use super::super::model::read_u32_le_at;
///
/// The wire shapes follow [MS-XLSB] section 2.5.98. Workbook relationship and
/// name resolution intentionally remains in the OOXML host adapter.
///
/// [MS-XLSB] section 2.2.2 defines formulas as an RPN sequence of Ptg
/// structures; this module preserves unknown bytes at the containing formula
/// boundary while validating every modeled token and ancillary payload.
use super::super::model::*;
use super::super::{Error, MAX_CELL_FORMULA_BYTES, Result};
use super::validation::table_row_type_raw;

fn read_bytes(data: &[u8], offset: usize, length: usize) -> Result<&[u8]> {
    let end = offset
        .checked_add(length)
        .ok_or_else(|| Error::InvalidFormula("formula binary offset overflow".to_string()))?;
    data.get(offset..end).ok_or(Error::InvalidLength {
        expected: end,
        found: data.len(),
    })
}

pub(super) fn read_u16_le_at(data: &[u8], offset: usize) -> Result<u16> {
    let bytes = read_bytes(data, offset, 2)?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

pub(super) fn read_f64_le_at(data: &[u8], offset: usize) -> Result<f64> {
    let bytes = read_bytes(data, offset, 8)?;
    Ok(f64::from_le_bytes([
        bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
    ]))
}

impl ParsedFormula {
    /// Parse a `ParsedFormula`, returning the structure and bytes consumed.
    pub fn parse(data: &[u8]) -> Result<(Self, usize)> {
        if data.len() < 4 {
            return Err(Error::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }

        let cce = read_u32_le_at(data, 0)? as usize;
        if cce > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "cell formula token length {cce} exceeds {MAX_CELL_FORMULA_BYTES}"
            )));
        }
        let cb_offset = 4usize.checked_add(cce).ok_or_else(|| {
            Error::InvalidFormula("cell formula token length overflow".to_string())
        })?;
        if data.len() < cb_offset + 4 {
            return Err(Error::InvalidLength {
                expected: cb_offset + 4,
                found: data.len(),
            });
        }

        let cb = read_u32_le_at(data, cb_offset)? as usize;
        let end = cb_offset
            .checked_add(4)
            .and_then(|offset| offset.checked_add(cb))
            .ok_or_else(|| {
                Error::InvalidFormula("cell formula ancillary length overflow".to_string())
            })?;
        if data.len() < end {
            return Err(Error::InvalidLength {
                expected: end,
                found: data.len(),
            });
        }

        Ok((
            Self {
                rgce: data[4..cb_offset].to_vec(),
                rgcb: data[cb_offset + 4..end].to_vec(),
            },
            end,
        ))
    }

    /// Serialize this formula with its two length prefixes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        if self.rgce.len() > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "cell formula token length {} exceeds {MAX_CELL_FORMULA_BYTES}",
                self.rgce.len()
            )));
        }
        let cce = u32::try_from(self.rgce.len())
            .map_err(|_| Error::InvalidFormula("formula is too large".to_string()))?;
        let cb = u32::try_from(self.rgcb.len()).map_err(|_| {
            Error::InvalidFormula("formula ancillary data is too large".to_string())
        })?;
        let mut bytes = Vec::with_capacity(8 + self.rgce.len() + self.rgcb.len());
        bytes.extend_from_slice(&cce.to_le_bytes());
        bytes.extend_from_slice(&self.rgce);
        bytes.extend_from_slice(&cb.to_le_bytes());
        bytes.extend_from_slice(&self.rgcb);
        Ok(bytes)
    }

    /// Create the `PtgExp` placeholder stored in every array/shared formula
    /// cell record.
    pub fn exp(row: u32, col: u32) -> Result<Self> {
        if row >= 1_048_576 || col >= 16_384 {
            return Err(Error::InvalidCellReference(format!(
                "grouped formula cell ({row}, {col})"
            )));
        }
        let mut rgce = Vec::with_capacity(5);
        rgce.push(ptg_types::PTG_EXP);
        rgce.extend_from_slice(&row.to_le_bytes());
        Ok(Self {
            rgce,
            rgcb: col.to_le_bytes().to_vec(),
        })
    }

    /// Return the target cell encoded by a `PtgExp`/`PtgExtraCol` formula.
    pub fn exp_cell(&self) -> Result<Option<(u32, u32)>> {
        if self.rgce.first() != Some(&ptg_types::PTG_EXP) {
            return Ok(None);
        }
        if self.rgce.len() != 5 || self.rgcb.len() != 4 {
            return Err(Error::InvalidFormula(format!(
                "PtgExp requires 5 rgce bytes and 4 rgcb bytes, found {} and {}",
                self.rgce.len(),
                self.rgcb.len()
            )));
        }
        let row = read_u32_le_at(&self.rgce, 1)?;
        let col = read_u32_le_at(&self.rgcb, 0)?;
        if row >= 1_048_576 || col >= 16_384 {
            return Err(Error::InvalidCellReference(format!(
                "PtgExp target ({row}, {col})"
            )));
        }
        Ok(Some((row, col)))
    }
}

impl Group {
    pub fn parse_array(data: &[u8]) -> Result<Self> {
        if data.len() < 17 {
            return Err(Error::InvalidLength {
                expected: 17,
                found: data.len(),
            });
        }
        if data[16] & !1 != 0 {
            return Err(Error::InvalidFormula(format!(
                "BrtArrFmla has reserved flag bits 0x{:02X}",
                data[16] & !1
            )));
        }
        let range = Range::parse_binary(data)?;
        let (formula, consumed) = ParsedFormula::parse(&data[17..])?;
        if 17 + consumed != data.len() {
            return Err(Error::InvalidFormula(format!(
                "BrtArrFmla has {} trailing bytes",
                data.len() - 17 - consumed
            )));
        }
        Ok(Self {
            kind: GroupKind::Array,
            range,
            formula,
            always_calculate: data[16] & 1 != 0,
        })
    }

    pub fn parse_shared(data: &[u8]) -> Result<Self> {
        if data.len() < 16 {
            return Err(Error::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }
        let range = Range::parse_binary(data)?;
        let (formula, consumed) = ParsedFormula::parse(&data[16..])?;
        if 16 + consumed != data.len() {
            return Err(Error::InvalidFormula(format!(
                "BrtShrFmla has {} trailing bytes",
                data.len() - 16 - consumed
            )));
        }
        Ok(Self {
            kind: GroupKind::Shared,
            range,
            formula,
            always_calculate: false,
        })
    }

    pub fn to_record_data(&self) -> Result<Vec<u8>> {
        self.range.validate()?;
        let formula = self.formula.to_bytes()?;
        let flag_len = usize::from(self.kind == GroupKind::Array);
        let mut data = Vec::with_capacity(16 + flag_len + formula.len());
        data.extend_from_slice(&self.range.to_binary());
        if self.kind == GroupKind::Array {
            data.push(u8::from(self.always_calculate));
        }
        data.extend_from_slice(&formula);
        Ok(data)
    }
}

impl Token {
    /// Encode one of the XLSB extended tokens implemented by this model.
    ///
    /// The first vector is the `Rgce` token and the second is its corresponding
    /// `RgbExtra` payload. Other token families are intentionally rejected.
    pub fn to_extended_binary(&self) -> Result<(Vec<u8>, Vec<u8>)> {
        match self {
            Self::PivotName(index) => {
                let mut token = Vec::with_capacity(6);
                token.extend([ptg_types::PTG_EXTENDED, ptg_types::EPTG_SX_NAME]);
                token.extend_from_slice(&index.to_le_bytes());
                Ok((token, Vec::new()))
            },
            Self::TableReference(reference) => reference.to_extended_binary(),
            _ => Err(Error::InvalidFormula(
                "token is not an extended PtgList/PtgSxName token".to_string(),
            )),
        }
    }
}

impl TableReference {
    pub fn to_extended_binary(&self) -> Result<(Vec<u8>, Vec<u8>)> {
        let external = self.external.as_ref();
        if self.invalid {
            if self.row_type.is_some()
                || self.columns.is_some()
                || self.list_index.is_some()
                || external.is_some()
            {
                return Err(Error::InvalidFormula(
                    "invalid PtgList cannot carry resident or external table metadata".to_string(),
                ));
            }
        } else if external.is_some() {
            if self.row_type.is_some() || self.columns.is_some() || self.list_index.is_some() {
                return Err(Error::InvalidFormula(
                    "nonresident PtgList cannot carry resident table metadata".to_string(),
                ));
            }
        } else if self.row_type.is_none() || self.columns.is_none() || self.list_index.is_none() {
            return Err(Error::InvalidFormula(
                "resident PtgList is missing table metadata".to_string(),
            ));
        }

        let mut flags = match self.data_type {
            TableDataType::Reference => 0,
            TableDataType::Value => 1 << 10,
            TableDataType::Array => 2 << 10,
        };
        if self.square_bracket_space {
            flags |= 0x0080;
        }
        if self.comma_space {
            flags |= 0x0100;
        }
        if self.invalid {
            flags |= 0x1000;
        }
        if external.is_some() {
            flags |= 0x2000;
        }

        let (list_index, first, last) = if let (Some(row_type), Some(columns), Some(list_index)) =
            (self.row_type, self.columns, self.list_index)
        {
            if list_index == 0 || list_index == u32::MAX {
                return Err(Error::InvalidFormula(format!(
                    "PtgList has invalid table identifier {list_index}"
                )));
            }
            flags |= u16::from(table_row_type_raw(row_type)) << 2;
            match columns {
                TableColumns::All => (list_index, 0, 0),
                TableColumns::One(column) => {
                    if column >= 16_384 {
                        return Err(Error::InvalidFormula(
                            "PtgList column is outside worksheet bounds".to_string(),
                        ));
                    }
                    flags |= 1;
                    (list_index, column, column)
                },
                TableColumns::Range { first, last } => {
                    if first > last || last >= 16_384 {
                        return Err(Error::InvalidFormula(
                            "PtgList column range is invalid".to_string(),
                        ));
                    }
                    flags |= 2;
                    (list_index, first, last)
                },
            }
        } else {
            (0, 0, 0)
        };

        let mut token = Vec::with_capacity(14);
        token.extend([ptg_types::PTG_EXTENDED, ptg_types::EPTG_LIST]);
        token.extend_from_slice(&self.sheet_index.to_le_bytes());
        token.extend_from_slice(&flags.to_le_bytes());
        token.extend_from_slice(&list_index.to_le_bytes());
        token.extend_from_slice(&first.to_le_bytes());
        token.extend_from_slice(&last.to_le_bytes());
        let extra = external
            .map(write_extra_list)
            .transpose()?
            .unwrap_or_default();
        Ok((token, extra))
    }
}

fn write_extra_list(reference: &ExternalTableReference) -> Result<Vec<u8>> {
    let table_units = reference.table.encode_utf16().count();
    if table_units == 0 || table_units >= 256 {
        return Err(Error::InvalidFormula(format!(
            "PtgExtraList table length {table_units} is outside 1..=255"
        )));
    }
    let has_columns = !matches!(reference.columns, TableNamedColumns::All);
    let mut extra = Vec::new();
    extra.push(u8::from(has_columns));
    extra.extend_from_slice(&u16::from(table_row_type_raw(reference.row_type)).to_le_bytes());
    extra.extend_from_slice(&(table_units as u16).to_le_bytes());
    push_formula_utf16(&mut extra, &reference.table);
    match &reference.columns {
        TableNamedColumns::All => {},
        TableNamedColumns::One(name) => {
            extra.extend([0, 0, 1]);
            write_sxos(&mut extra, false, name)?;
        },
        TableNamedColumns::Range { first, last } => {
            extra.extend([0, 0, 2]);
            write_sxos(&mut extra, true, first)?;
            write_sxos(&mut extra, false, last)?;
        },
    }
    Ok(extra)
}

fn write_sxos(output: &mut Vec<u8>, not_last: bool, name: &str) -> Result<()> {
    let units = name.encode_utf16().count();
    if units == 0 || units > 1_048_576 {
        return Err(Error::InvalidFormula(format!(
            "structured-reference column length {units} is outside 1..=1048576"
        )));
    }
    output.push(u8::from(not_last));
    output.extend_from_slice(&2u16.to_le_bytes());
    output.extend_from_slice(&(units as u32).to_le_bytes());
    push_formula_utf16(output, name);
    Ok(())
}

fn push_formula_utf16(output: &mut Vec<u8>, value: &str) {
    output.extend(value.encode_utf16().flat_map(u16::to_le_bytes));
}
