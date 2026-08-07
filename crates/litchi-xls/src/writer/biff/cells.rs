//! Cell record BIFF8 writers.

use crate::{Error, Result};
use std::io::Write;

use super::{write_record, write_record_header};

fn encode_rk(value: f64) -> Option<u32> {
    let int_value = value as i32;
    if f64::from(int_value) == value && (-(1 << 29)..(1 << 29)).contains(&int_value) {
        return Some(((int_value as u32) << 2) | 0x02);
    }

    let scaled = value * 100.0;
    let scaled_int = scaled as i32;
    if f64::from(scaled_int) == scaled && (-(1 << 29)..(1 << 29)).contains(&scaled_int) {
        return Some(((scaled_int as u32) << 2) | 0x03);
    }

    None
}

/// Write NUMBER record (floating point cell)
///
/// Record type: 0x0203
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `row` - Row index (0-based)
/// * `col` - Column index (0-based)
/// * `value` - Cell value (f64)
pub(crate) fn write_number<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    value: f64,
) -> Result<()> {
    // BIFF8 stores row as a 16-bit index (0..65535)
    let row_u16 = u16::try_from(row).map_err(|_| {
        Error::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for NUMBER record",
            row
        ))
    })?;

    write_record_header(writer, 0x0203, 14)?;

    writer.write_all(&row_u16.to_le_bytes())?;
    writer.write_all(&col.to_le_bytes())?;

    // XF record index
    writer.write_all(&xf_index.to_le_bytes())?;

    // IEEE 754 floating point value
    writer.write_all(&value.to_le_bytes())?;

    Ok(())
}

pub(crate) fn write_mulrk<W: Write>(
    writer: &mut W,
    row: u32,
    first_col: u16,
    values: &[(u16, f64)],
) -> Result<()> {
    let row_u16 = u16::try_from(row).map_err(|_| {
        Error::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for MULRK record",
            row
        ))
    })?;
    if values.len() < 2 {
        return Err(Error::InvalidData(
            "MULRK record requires at least two contiguous numeric cells".to_string(),
        ));
    }
    let last_col = first_col
        .checked_add(
            u16::try_from(values.len() - 1).map_err(|_| {
                Error::InvalidData("MULRK column span exceeds BIFF8 limit".to_string())
            })?,
        )
        .ok_or_else(|| Error::InvalidData("MULRK last column exceeds BIFF8 limit".to_string()))?;
    let data_len = 6u16
        .checked_add(u16::try_from(values.len() * 6).map_err(|_| {
            Error::InvalidData("MULRK payload exceeds BIFF8 size limit".to_string())
        })?)
        .ok_or_else(|| Error::InvalidData("MULRK payload exceeds BIFF8 size limit".to_string()))?;

    write_record_header(writer, 0x00BD, data_len)?;
    writer.write_all(&row_u16.to_le_bytes())?;
    writer.write_all(&first_col.to_le_bytes())?;
    for (xf_index, value) in values {
        let rk = encode_rk(*value).ok_or_else(|| {
            Error::InvalidData(format!("Value {value} cannot be encoded as RK for MULRK"))
        })?;
        writer.write_all(&xf_index.to_le_bytes())?;
        writer.write_all(&rk.to_le_bytes())?;
    }
    writer.write_all(&last_col.to_le_bytes())?;
    Ok(())
}

/// Write LABELSST record (string cell with reference to SST)
///
/// Record type: 0x00FD
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `row` - Row index (0-based)
/// * `col` - Column index (0-based)
/// * `sst_index` - Index into shared string table
pub(crate) fn write_labelsst<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    sst_index: u32,
) -> Result<()> {
    // BIFF8 stores row as a 16-bit index (0..65535)
    let row_u16 = u16::try_from(row).map_err(|_| {
        Error::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for LABELSST record",
            row
        ))
    })?;

    // 2 (row) + 2 (col) + 2 (xf) + 4 (sst index) = 10 bytes
    write_record_header(writer, 0x00FD, 10)?;

    writer.write_all(&row_u16.to_le_bytes())?;
    writer.write_all(&col.to_le_bytes())?;

    // XF record index
    writer.write_all(&xf_index.to_le_bytes())?;

    // SST index
    writer.write_all(&sst_index.to_le_bytes())?;

    Ok(())
}

/// Write BOOLERR record (boolean or error cell)
///
/// Record type: 0x0205
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `row` - Row index (0-based)
/// * `col` - Column index (0-based)
/// * `value` - Boolean value
pub(crate) fn write_boolerr<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    value: bool,
) -> Result<()> {
    // BIFF8 stores row as a 16-bit index (0..65535)
    let row_u16 = u16::try_from(row).map_err(|_| {
        Error::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for BOOLERR record",
            row
        ))
    })?;

    // 2 (row) + 2 (col) + 2 (xf) + 1 (value) + 1 (is-error flag) = 8 bytes
    write_record_header(writer, 0x0205, 8)?;

    writer.write_all(&row_u16.to_le_bytes())?;
    writer.write_all(&col.to_le_bytes())?;

    // XF record index
    writer.write_all(&xf_index.to_le_bytes())?;

    // Boolean value (0 = false, 1 = true) + error flag (0 = boolean)
    writer.write_all(&[if value { 1 } else { 0 }, 0])?;

    Ok(())
}

/// Write a BIFF8 FORMULA record with an empty cached result.
///
/// The `fAlwaysCalc` flag requests recalculation when Excel opens the file.
pub(crate) fn write_formula<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    tokens: &[u8],
) -> Result<()> {
    write_formula_with_metadata(
        writer,
        row,
        col,
        xf_index,
        tokens,
        crate::FormulaMetadata::new().with_always_calculate(true),
    )
}

pub(crate) fn write_formula_with_metadata<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    tokens: &[u8],
    metadata: crate::FormulaMetadata,
) -> Result<()> {
    let row_u16 = u16::try_from(row).map_err(|_| {
        Error::InvalidData(format!(
            "Row index {row} exceeds BIFF8 limit 65535 for FORMULA record"
        ))
    })?;
    crate::formula_metadata::validate_for_write(&metadata)?;
    let shared_owner = metadata.shared_owner();
    let array_owner = metadata.array_owner();
    let shared_tokens = match shared_owner {
        Some(owner) => Some(
            owner
                .validate_cell(row_u16, col)
                .map(|_| owner.anchor_tokens())?,
        ),
        None => None,
    };
    let array_tokens = match array_owner {
        Some(owner) => {
            let col_u8 = u8::try_from(col).map_err(|_| {
                Error::InvalidFormula(format!(
                    "array-formula column {col} exceeds the BIFF8 limit"
                ))
            })?;
            let range = owner.range();
            if row_u16 < range.first().row()
                || row_u16 > range.last().row()
                || col_u8 < range.first().col()
                || col_u8 > range.last().col()
            {
                return Err(Error::InvalidFormula(format!(
                    "cell ({row_u16}, {col}) is outside its array-formula range"
                )));
            }
            Some(owner.anchor_tokens())
        },
        None => None,
    };
    let formula_tokens = match (shared_tokens.as_ref(), array_tokens.as_ref()) {
        (Some(tokens), None) => tokens.as_slice(),
        (None, Some(tokens)) => tokens.as_slice(),
        (None, None) => tokens,
        (Some(_), Some(_)) => {
            return Err(Error::InvalidFormula(
                "a Formula record cannot own both ShrFmla and Array metadata".to_string(),
            ));
        },
    };
    // Materialize and validate the complete Array payload before emitting its
    // anchor Formula, so semantic failures cannot leave an orphan Formula in
    // the caller's stream.
    let array_payload = match array_owner {
        Some(owner)
            if owner.anchor().row() == row_u16 && u16::from(owner.anchor().col()) == col =>
        {
            Some(owner.to_payload()?)
        },
        _ => None,
    };
    let flags = crate::formula_metadata::encode_flags(&metadata, formula_tokens)?;
    // A BIFF record payload is limited to 8,224 bytes. FORMULA contributes
    // 22 fixed bytes before the token stream.
    if formula_tokens.len() > 8_202 {
        return Err(Error::InvalidFormula(
            "Formula token stream exceeds BIFF8 record limit".to_string(),
        ));
    }
    let token_len = u16::try_from(formula_tokens.len())
        .map_err(|_| Error::InvalidFormula("Formula token length exceeds u16".to_string()))?;
    let data_len = 22u16
        .checked_add(token_len)
        .ok_or_else(|| Error::InvalidFormula("Formula record length overflow".to_string()))?;

    write_record_header(writer, 0x0006, data_len)?;
    writer.write_all(&row_u16.to_le_bytes())?;
    writer.write_all(&col.to_le_bytes())?;
    writer.write_all(&xf_index.to_le_bytes())?;
    // FormulaValue special cached EMPTY: type, reserved, data, reserved[3], marker.
    writer.write_all(&[3, 0, 0, 0, 0, 0, 0xff, 0xff])?;
    writer.write_all(&flags.to_le_bytes())?;
    writer.write_all(&metadata.calculation_cache().to_le_bytes())?;
    writer.write_all(&token_len.to_le_bytes())?;
    writer.write_all(formula_tokens)?;

    if let Some(owner) = shared_owner
        && owner.anchor().row() == row_u16
        && u16::from(owner.anchor().col()) == col
    {
        write_shared_formula(writer, owner)?;
    }
    if let Some(payload) = array_payload {
        write_record(writer, 0x0221, &payload)?;
    }
    Ok(())
}

/// Write the ShrFmla record owned by an anchor Formula record.
fn write_shared_formula<W: Write>(
    writer: &mut W,
    owner: &crate::formula_metadata::Owner,
) -> Result<()> {
    const FIXED_PAYLOAD_SIZE: usize = crate::formula_metadata::shared::FIXED_PAYLOAD_SIZE;
    const MAX_FORMULA_BYTES: usize = crate::formula_metadata::shared::MAX_FORMULA_BYTES;
    let tokens = owner.tokens();
    if tokens.is_empty() {
        return Err(Error::InvalidFormula(
            "ShrFmla shared parsed formula cannot be empty".to_string(),
        ));
    }
    if tokens.len() > MAX_FORMULA_BYTES {
        return Err(Error::InvalidFormula(format!(
            "ShrFmla shared parsed formula exceeds the BIFF8 limit of {MAX_FORMULA_BYTES} bytes"
        )));
    }
    if tokens.first().is_some_and(|opcode| opcode & 0x7F == 0x01) {
        return Err(Error::UnsupportedFeature(
            "array/PtgExp formulas cannot be used as a shared ShrFmla owner".to_string(),
        ));
    }
    let payload_len = FIXED_PAYLOAD_SIZE
        .checked_add(tokens.len())
        .ok_or_else(|| Error::InvalidFormula("ShrFmla payload length overflow".to_string()))?;
    let payload_len = u16::try_from(payload_len)
        .map_err(|_| Error::InvalidFormula("ShrFmla payload exceeds u16".to_string()))?;
    let c_use = owner.c_use()?;

    write_record_header(writer, 0x04BC, payload_len)?;
    let range = owner.range();
    writer.write_all(&range.first().row().to_le_bytes())?;
    writer.write_all(&range.last().row().to_le_bytes())?;
    writer.write_all(&[range.first().col(), range.last().col()])?;
    writer.write_all(&[0, c_use])?;
    let token_len = u16::try_from(tokens.len())
        .map_err(|_| Error::InvalidFormula("ShrFmla token length exceeds u16".to_string()))?;
    writer.write_all(&token_len.to_le_bytes())?;
    writer.write_all(tokens)?;
    Ok(())
}

/// Write a TABLE record (MS-XLS 2.4.319) for a what-if data table.
///
/// Record type: 0x0236
pub(crate) fn write_table<W: Write>(writer: &mut W, table: &crate::DataTable) -> Result<()> {
    let payload = table.to_payload();
    write_record(writer, 0x0236, &payload)
}

#[cfg(test)]
mod tests {
    use super::write_formula;
    use crate::formula_metadata::shared::{Cell, Owner, Range, parse};
    use crate::records::{CellRecord, Encoding, FormulaValue};

    #[test]
    fn writes_formula_record_with_recalculation_and_empty_cache() {
        let tokens = [0x1e, 2, 0, 0x1e, 3, 0, 0x03];
        let mut bytes = Vec::new();
        write_formula(&mut bytes, 4, 5, 15, &tokens).unwrap();
        assert_eq!(&bytes[..4], &[0x06, 0x00, 29, 0]);
        assert_eq!(u16::from_le_bytes([bytes[18], bytes[19]]), 1);

        let record = CellRecord::parse(0x0006, &bytes[4..], &Encoding::Utf16Le).unwrap();
        assert!(
            matches!(
                record,
                CellRecord::Formula {
                    row: 4,
                    col: 5,
                    xf_index: 15,
                    value: FormulaValue::Empty,
                    ref metadata,
                    ref formula,
                } if formula == &tokens
                    && *metadata
                        == crate::FormulaMetadata::new().with_always_calculate(true)
            ),
            "{record:?}"
        );
    }

    #[test]
    fn writes_anchor_formula_then_one_shrfmla_with_ptg_exp_and_refu() {
        let anchor = Cell::new(0, 0);
        let owner = Owner::new(Range::try_new(0, 0, 1, 1).unwrap(), anchor, &[0x1E, 99, 0])
            .unwrap()
            .with_participants(&[anchor, Cell::new(1, 0)])
            .unwrap();
        let metadata = crate::FormulaMetadata::new().with_shared(owner);
        let mut bytes = Vec::new();

        super::write_formula_with_metadata(&mut bytes, 0, 0, 15, &[0x1E, 99, 0], metadata).unwrap();

        assert_eq!(&bytes[..4], &[0x06, 0, 27, 0]);
        assert_eq!(&bytes[4 + 22..4 + 27], &[0x01, 0, 0, 0, 0]);
        assert_eq!(u16::from_le_bytes([bytes[4 + 14], bytes[4 + 15]]), 0x0008);

        let shared_offset = 4 + 27;
        assert_eq!(
            &bytes[shared_offset..shared_offset + 4],
            &[0xBC, 0x04, 13, 0]
        );
        let parsed = parse(&bytes[shared_offset + 4..]).unwrap();
        assert_eq!(parsed.range, Range::try_new(0, 0, 1, 1).unwrap());
        assert_eq!(parsed.count, 2);
        assert_eq!(parsed.tokens, [0x1E, 99, 0]);
    }

    #[test]
    fn participating_formula_emits_ptg_exp_without_a_duplicate_shrfmla() {
        let anchor = Cell::new(0, 0);
        let owner = Owner::new(Range::try_new(0, 0, 1, 0).unwrap(), anchor, &[0x1E, 99, 0])
            .unwrap()
            .with_participants(&[anchor, Cell::new(1, 0)])
            .unwrap();
        let metadata = crate::FormulaMetadata::new().with_shared(owner);
        let mut bytes = Vec::new();

        super::write_formula_with_metadata(&mut bytes, 1, 0, 15, &[], metadata).unwrap();
        assert_eq!(bytes.len(), 4 + 22 + 5);
        assert_eq!(&bytes[4 + 22..], &[0x01, 0, 0, 0, 0]);
    }

    #[test]
    fn array_members_emit_empty_ptg_exp_formulas_and_only_anchor_emits_array() {
        let range = Range::try_new(0, 0, 1, 0).unwrap();
        let owner =
            crate::formula_metadata::array::Owner::from_compiled(range, vec![0x1e, 7, 0]).unwrap();
        let metadata = crate::FormulaMetadata::new()
            .with_always_calculate(true)
            .with_array(owner);

        let mut anchor = Vec::new();
        super::write_formula_with_metadata(&mut anchor, 0, 0, 15, &[], metadata.clone()).unwrap();
        assert_eq!(&anchor[..4], &[0x06, 0x00, 27, 0]);
        assert_eq!(&anchor[10..18], &[3, 0, 0, 0, 0, 0, 0xff, 0xff]);
        assert_eq!(u16::from_le_bytes([anchor[18], anchor[19]]) & 0x0008, 0);
        assert_eq!(&anchor[26..31], &[0x01, 0, 0, 0, 0]);
        assert_eq!(
            &anchor[31..],
            &[
                0x21, 0x02, 17, 0, // Array header
                0, 0, 1, 0, 0, 0, // Ref
                1, 0, // fAlwaysCalc, reserved=0
                0, 0, 0, 0, // unused
                3, 0, // cce
                0x1e, 7, 0, // rgce
            ]
        );

        let mut participant = Vec::new();
        super::write_formula_with_metadata(&mut participant, 1, 0, 15, &[], metadata).unwrap();
        assert_eq!(participant.len(), 31);
        assert_eq!(&participant[26..], &[0x01, 0, 0, 0, 0]);
    }
}
