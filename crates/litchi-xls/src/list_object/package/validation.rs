//! Collection-level invariants for worksheet list-object packages.

use super::super::model::ListObject;
use super::super::{FEATURE11_RECORD_TYPE, invalid};
use crate::Result;
use std::collections::HashSet;

pub(super) fn validate_tables(tables: &[ListObject]) -> Result<()> {
    let mut ids = HashSet::new();
    let mut names = HashSet::new();
    for table in tables {
        if !ids.insert(table.id) || !names.insert(table.name.to_lowercase()) {
            return Err(invalid(FEATURE11_RECORD_TYPE, "duplicate table id or name"));
        }
    }
    Ok(())
}
