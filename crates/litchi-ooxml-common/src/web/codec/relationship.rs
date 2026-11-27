//! Relationship-namespace decoding for web-extension XML.

use super::super::{STRICT_RELATIONSHIPS_NAMESPACE, TRANSITIONAL_RELATIONSHIPS_NAMESPACE};
use super::semantic::invalid;
use super::xml::{Node, attr};
use crate::Result;

pub(in crate::web) fn relationship_attr<'a>(
    node: &'a Node,
    local_name: &str,
) -> Result<Option<&'a str>> {
    let transitional = attr(node, TRANSITIONAL_RELATIONSHIPS_NAMESPACE, local_name);
    let strict = attr(node, STRICT_RELATIONSHIPS_NAMESPACE, local_name);
    if transitional.is_some() && strict.is_some() {
        invalid(format!(
            "{} has both Strict and Transitional r:{local_name}",
            node.local_name
        ))
    } else {
        Ok(transitional.or(strict))
    }
}
