//! Package-level aggregation for variable declarations.
//!
//! Parsing each XML part remains a codec concern; this boundary owns the
//! document-wide declaration/reference inventory and its cross-part limits.

use super::{Kind, MAX_XML_BYTES, Part, Scope, codec, model::Declarations};
use litchi_core::Result;
use std::collections::HashSet;

/// Parse content, styles, or flat XML parts as one validated document view.
pub(crate) fn parse_parts(parts: &[(&str, Part)]) -> Result<Declarations> {
    let total = parts.iter().try_fold(0usize, |size, (xml, _)| {
        size.checked_add(xml.len())
            .ok_or_else(|| codec::invalid("variable declaration XML size overflow"))
    })?;
    if total > MAX_XML_BYTES {
        return Err(codec::invalid("variable declaration XML exceeds 64 MiB"));
    }

    let mut result = Declarations::default();
    let mut names = HashSet::<(Kind, String)>::new();
    let mut containers = HashSet::<(Part, Scope, Kind)>::new();
    let mut uses = HashSet::new();
    let mut all_uses = Vec::<(Kind, String)>::new();
    let mut aggregate = 0usize;
    let mut declaration_count = 0usize;
    for (xml, part) in parts {
        codec::parse_part(
            xml,
            *part,
            &mut result,
            &mut names,
            &mut containers,
            &mut uses,
            &mut all_uses,
            &mut aggregate,
            &mut declaration_count,
        )?;
    }
    for (kind, name) in all_uses {
        if !names.contains(&(kind, name.clone())) {
            return Err(codec::invalid(format!(
                "ODF {kind:?} variable '{name}' is used without a declaration"
            )));
        }
    }
    let dde = crate::dde_connection::parse_dde_connection_parts(parts)?;
    result.dde_connections = dde.declarations;
    result.dde_connection_uses = dde.uses;
    result.bibliography_configuration =
        crate::bibliography_configuration::parse_bibliography_configuration_parts(parts)?;
    result.auto_mark_files = crate::auto_mark_file::parse_auto_mark_file_parts(parts)?;
    Ok(result)
}
