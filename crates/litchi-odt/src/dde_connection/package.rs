//! Package-wide aggregation for ODF DDE metadata.

use super::{MAX_XML_BYTES, codec, model::Connections};
use crate::generic::FlatMutationBudget;
use crate::variable_declaration::{Part, Scope};
use litchi_core::Result;
use litchi_odf_common::XmlStreamLimits;
use std::collections::HashSet;

/// Parse all supplied ODF XML parts and validate cross-part DDE references.
#[cfg(test)]
pub(crate) fn parse_dde_connection_parts_with_budget(
    parts: &[(&str, Part)],
    budget: Option<&FlatMutationBudget>,
) -> Result<Connections> {
    parse_dde_connection_parts_with_budget_and_xml_limits(parts, budget, XmlStreamLimits::default())
}

/// Parse all supplied ODF XML parts and validate cross-part DDE references.
pub(crate) fn parse_dde_connection_parts_with_budget_and_xml_limits(
    parts: &[(&str, Part)],
    budget: Option<&FlatMutationBudget>,
    xml_limits: XmlStreamLimits,
) -> Result<Connections> {
    if let Some(budget) = budget {
        budget.check()?;
    }
    let total = parts.iter().try_fold(0usize, |total, (xml, _)| {
        total
            .checked_add(xml.len())
            .ok_or_else(|| codec::make_error("DDE connection XML size overflow"))
    })?;
    if total > MAX_XML_BYTES {
        return codec::invalid("DDE connection XML exceeds 64 MiB");
    }

    let mut parsed = Connections::default();
    let mut names = HashSet::<String>::new();
    let mut containers = HashSet::<(Part, Scope)>::new();
    let mut aggregate = 0usize;
    for (xml, part) in parts {
        codec::parse_part(
            xml,
            *part,
            &mut parsed,
            &mut names,
            &mut containers,
            &mut aggregate,
            budget,
            xml_limits,
        )?;
    }
    for usage in &parsed.uses {
        if let Some(budget) = budget {
            budget.check()?;
            budget.consume_objects(1)?;
        }
        if !names.contains(&usage.connection_name) {
            return codec::invalid(format!(
                "DDE connection '{}' is used without a declaration",
                usage.connection_name
            ));
        }
    }
    Ok(parsed)
}
