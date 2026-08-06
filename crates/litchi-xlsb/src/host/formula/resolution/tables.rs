//! Resident and external structured-reference resolution.

use crate::package::error::{Error, Result};
use crate::package::formula::{TableColumns, TableNamedColumns, TableReference};

use super::super::{
    format_structured_reference, validate_named_table_columns, validate_table_name,
};
use super::Context;

impl Context {
    pub(super) fn resolve_table_reference(&self, reference: &TableReference) -> Result<String> {
        if let Some(external) = &reference.external {
            if reference.row_type.is_some()
                || reference.columns.is_some()
                || reference.list_index.is_some()
            {
                return Err(Error::InvalidFormula(
                    "nonresident structured reference also contains resident metadata".to_string(),
                ));
            }
            validate_table_name(&external.table)?;
            validate_named_table_columns(&external.columns)?;
            let prefix = self.resolve_external_table_prefix(reference.sheet_index)?;
            return Ok(format!(
                "{prefix}!{}",
                format_structured_reference(
                    &external.table,
                    external.row_type,
                    &external.columns,
                    reference.square_bracket_space,
                    reference.comma_space,
                )
            ));
        }

        let table_id = reference.list_index.ok_or_else(|| {
            Error::InvalidFormula("resident structured reference omits table ID".to_string())
        })?;
        let row_type = reference.row_type.ok_or_else(|| {
            Error::InvalidFormula("resident structured reference omits row type".to_string())
        })?;
        let columns = reference.columns.ok_or_else(|| {
            Error::InvalidFormula("resident structured reference omits columns".to_string())
        })?;
        let sheet = self.resolve_table_sheet(reference.sheet_index)?;
        let mut matches = self
            .tables
            .iter()
            .filter(|table| table.table_id == table_id);
        let table = matches.next().ok_or_else(|| {
            Error::InvalidFormula(format!(
                "structured reference names missing table ID {table_id}"
            ))
        })?;
        if matches.next().is_some() {
            return Err(Error::InvalidFormula(format!(
                "structured reference table ID {table_id} is ambiguous"
            )));
        }
        if table.sheet_index != sheet {
            return Err(Error::InvalidFormula(format!(
                "structured reference locates table ID {table_id} on worksheet {sheet}, but metadata places it on {}",
                table.sheet_index
            )));
        }
        let named_columns = match columns {
            TableColumns::All => TableNamedColumns::All,
            TableColumns::One(index) => {
                let name = table.columns.get(usize::from(index)).ok_or_else(|| {
                    Error::InvalidFormula(format!(
                        "structured-reference column {index} exceeds {} columns in table {:?}",
                        table.columns.len(),
                        table.display_name
                    ))
                })?;
                TableNamedColumns::One(name.clone())
            },
            TableColumns::Range { first, last } => {
                let first_name = table.columns.get(usize::from(first)).ok_or_else(|| {
                    Error::InvalidFormula(format!(
                        "structured-reference first column {first} exceeds {} columns",
                        table.columns.len()
                    ))
                })?;
                let last_name = table.columns.get(usize::from(last)).ok_or_else(|| {
                    Error::InvalidFormula(format!(
                        "structured-reference last column {last} exceeds {} columns",
                        table.columns.len()
                    ))
                })?;
                TableNamedColumns::Range {
                    first: first_name.clone(),
                    last: last_name.clone(),
                }
            },
        };
        Ok(format_structured_reference(
            &table.display_name,
            row_type,
            &named_columns,
            reference.square_bracket_space,
            reference.comma_space,
        ))
    }
}
