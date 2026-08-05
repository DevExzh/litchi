//! Package-part aggregation for auto-mark-file references.

use super::codec::parse_part;
use super::model::AlphabeticalIndexAutoMarkFile;
use super::{MAX_XML_BYTES, make_error};
use crate::variable_declaration::{Part, Scope};
use litchi_core::Result;
use std::collections::HashSet;

pub(crate) fn parse_auto_mark_file_parts(
    parts: &[(&str, Part)],
) -> Result<Vec<AlphabeticalIndexAutoMarkFile>> {
    let total = parts.iter().try_fold(0usize, |total, (xml, _)| {
        total
            .checked_add(xml.len())
            .ok_or_else(|| make_error("auto-mark-file XML size overflow"))
    })?;
    if total > MAX_XML_BYTES {
        return Err(make_error("auto-mark-file XML exceeds 64 MiB"));
    }

    let mut references = Vec::new();
    let mut scopes = HashSet::<(Part, Scope)>::new();
    let mut aggregate = 0usize;
    for (xml, part) in parts {
        parse_part(xml, *part, &mut references, &mut scopes, &mut aggregate)?;
    }
    Ok(references)
}
