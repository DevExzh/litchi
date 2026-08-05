//! Structured table metadata used by XLSB formula resolution.
use super::*;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Definition {
    pub(super) table_id: u32,
    pub(super) sheet_index: usize,
    pub(super) display_name: String,
    pub(super) columns: std::sync::Arc<[String]>,
}

impl Definition {
    /// Build a validated table definition with columns in table-relative order.
    pub fn try_new(
        table_id: u32,
        sheet_index: usize,
        display_name: impl Into<String>,
        columns: Vec<String>,
    ) -> Result<Self> {
        if table_id == 0 || table_id == u32::MAX {
            return Err(Error::InvalidFormula(format!(
                "table identifier {table_id} is outside 1..=4294967294"
            )));
        }
        let display_name = display_name.into();
        validate_table_name(&display_name)?;
        if columns.is_empty() || columns.len() > 16_384 {
            return Err(Error::InvalidFormula(format!(
                "table {display_name:?} has {} columns, outside 1..=16384",
                columns.len()
            )));
        }
        for (index, column) in columns.iter().enumerate() {
            validate_table_column_name(column, index)?;
            if columns[..index]
                .iter()
                .any(|existing| excel_name_eq(existing, column))
            {
                return Err(Error::InvalidFormula(format!(
                    "table {display_name:?} contains duplicate column {column:?}"
                )));
            }
        }
        Ok(Self {
            table_id,
            sheet_index,
            display_name,
            columns: columns.into(),
        })
    }

    pub fn table_id(&self) -> u32 {
        self.table_id
    }

    pub fn sheet_index(&self) -> usize {
        self.sheet_index
    }

    pub fn display_name(&self) -> &str {
        &self.display_name
    }

    pub fn columns(&self) -> &[String] {
        &self.columns
    }
}
