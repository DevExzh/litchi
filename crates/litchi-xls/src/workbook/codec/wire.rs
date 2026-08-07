//! BIFF wire-level helpers shared by workbook semantic collectors.

use crate::error::{Error, Result};
use crate::formula::{FormulaContext, render_formula, render_shared_formula};

#[derive(Debug)]
pub(super) struct SharedFormulaTemplate {
    pub(super) first_row: u16,
    pub(super) last_row: u16,
    pub(super) first_col: u16,
    pub(super) last_col: u16,
    tokens: Vec<u8>,
    relative: bool,
}

impl SharedFormulaTemplate {
    pub(super) fn contains(&self, row: u16, col: u16) -> bool {
        (self.first_row..=self.last_row).contains(&row)
            && (self.first_col..=self.last_col).contains(&col)
    }

    pub(super) fn render(
        &self,
        context: Option<&FormulaContext>,
        row: u16,
        col: u16,
    ) -> Option<String> {
        if !self.contains(row, col) {
            return None;
        }
        if self.relative {
            render_shared_formula(&self.tokens, context, row, col)
        } else {
            render_formula(&self.tokens, context)
        }
    }
}

pub(super) fn parse_shared_formula_template(
    record_type: u16,
    data: &[u8],
) -> Result<SharedFormulaTemplate> {
    if record_type != 0x04bc {
        return Err(Error::UnexpectedRecordType {
            expected: 0x04bc,
            found: record_type,
        });
    }
    let fixed_size = 10usize;
    let length_offset = 8usize;
    if data.len() < fixed_size {
        return Err(Error::InvalidLength {
            expected: fixed_size,
            found: data.len(),
        });
    }
    let first_row = u16::from_le_bytes([data[0], data[1]]);
    let last_row = u16::from_le_bytes([data[2], data[3]]);
    let first_col = u16::from(data[4]);
    let last_col = u16::from(data[5]);
    if first_row > last_row || first_col > last_col {
        return Err(Error::InvalidRecord {
            record_type,
            message: "shared formula range is reversed".to_string(),
        });
    }
    let token_len = usize::from(u16::from_le_bytes([
        data[length_offset],
        data[length_offset + 1],
    ]));
    let end = fixed_size
        .checked_add(token_len)
        .ok_or_else(|| Error::InvalidRecord {
            record_type,
            message: "shared formula token length overflows".to_string(),
        })?;
    let tokens = data
        .get(fixed_size..end)
        .ok_or(Error::InvalidLength {
            expected: end,
            found: data.len(),
        })?
        .to_vec();
    if tokens.is_empty() {
        return Err(Error::InvalidRecord {
            record_type,
            message: "shared formula token stream is empty".to_string(),
        });
    }
    Ok(SharedFormulaTemplate {
        first_row,
        last_row,
        first_col,
        last_col,
        tokens,
        relative: true,
    })
}

pub(crate) fn pivot_cache_stream_paths(
    paths: impl IntoIterator<Item = Vec<String>>,
) -> Vec<(u16, Vec<String>)> {
    let mut cache_paths = paths
        .into_iter()
        .filter_map(|path| {
            if path.len() != 2 || !path[0].eq_ignore_ascii_case("_SX_DB_CUR") {
                return None;
            }
            let stream_id = u16::from_str_radix(&path[1], 16).ok()?;
            Some((stream_id, path))
        })
        .collect::<Vec<_>>();
    cache_paths.sort_unstable_by_key(|(stream_id, _)| *stream_id);
    cache_paths
}
