#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Bounded glossary codec facade.

mod package;
mod semantic;
mod validation;
mod xml;

pub(in crate::glossary) use package::*;
pub(in crate::glossary) use semantic::*;
pub(in crate::glossary) use validation::*;
pub(in crate::glossary) use xml::*;

use super::model::{Catalog, Conformance};
use super::{Error, Result};

///
/// # Errors
///
/// Returns an error if the operation cannot be completed.
#[cfg(test)]
pub(super) fn read(xml: &[u8]) -> Result<(Catalog, Conformance)> {
    read_with_limits(xml, &super::Limits::DEFAULT)
}

pub(in crate::glossary) fn read_with_limits(
    xml: &[u8],
    limits: &super::Limits,
) -> Result<(Catalog, Conformance)> {
    // Keep admission in the one source read so every caller, including the
    // package owner, proves the selected policy before either DOM projection
    // retains strings or the MCE projection is parsed.
    super::graph::raw::admit_xml(xml, limits)?;
    read_admitted_with_limits(xml, limits)
}

/// Read XML whose exact bytes have already passed the raw XML admission under
/// the same policy. This preserves all DOM, MCE, semantic, and catalog checks
/// without scanning the identical bytes a second time.
pub(in crate::glossary) fn read_admitted_with_limits(
    xml: &[u8],
    limits: &super::Limits,
) -> Result<(Catalog, Conformance)> {
    let original = parse_dom_with_limits(xml, limits)?;
    let original_conformance = Conformance::from_word(original.ns.as_ref())?;
    let producer_entries = extract_producer_entries(original, original_conformance, limits)?;
    let mce_limits = litchi_ooxml_common::mce::Limits {
        max_input_bytes: limits.xml_bytes(),
        // The projected DOM retains the processed XML, so the expansion
        // budget must honor both the part and DOM policies before parse_dom.
        max_output_bytes: limits.xml_bytes().min(limits.dom_bytes()),
        max_depth: limits.depth(),
        max_namespace_bindings: limits.attributes(),
        max_directive_tokens: limits.tokens(),
        max_choices_per_alternate: limits.values(),
        max_attributes_per_element: limits
            .attributes()
            .min(litchi_ooxml_common::mce::ATTRIBUTES_PER_ELEMENT_CEILING),
    };
    // Glossary entries can carry the same tracked-change attributes as the
    // main document.  Keep Word 2023's dateUtc namespace understood in this
    // DOCX-owned MCE profile so an ignorable declaration does not erase it
    // during the semantic projection.
    let mut capabilities = litchi_ooxml_common::mce::Capabilities::default();
    capabilities.understand_namespace(crate::revision::WORD_2023_DATE_UTC_NAMESPACE);
    // Retained run views understand symEx even when producers mark it ignorable.
    capabilities.understand_namespace(crate::run_symbols::NAMESPACE);
    let xml =
        litchi_ooxml_common::mce::process_markup_compatibility(xml, &capabilities, &mce_limits)?
            .xml;
    super::graph::raw::admit_xml(xml.as_ref(), limits)?;
    let root = parse_dom_with_limits(xml.as_ref(), limits)?;
    let conformance = Conformance::from_word(root.ns.as_ref())?;
    if conformance != original_conformance {
        return Err(invalid(
            "MCE preprocessing changed the glossary conformance family",
        ));
    }
    validate_word_dialect(&root, conformance)?;
    let catalog = project(&root, limits, producer_entries.as_deref())?;
    drop(root);
    validate_catalog_limits(&catalog, limits)?;
    Ok((catalog, conformance))
}

/// Serialize a catalog canonically in the selected dialect.
///
/// # Errors
///
/// Returns an error if the operation cannot be completed.
#[cfg(test)]
pub(super) fn write(value: &Catalog, conformance: Conformance) -> Result<Vec<u8>> {
    write_with_limits(value, conformance, &super::Limits::DEFAULT)
}

pub(in crate::glossary) fn write_with_limits(
    value: &Catalog,
    conformance: Conformance,
    limits: &super::Limits,
) -> Result<Vec<u8>> {
    validate_catalog_limits(value, limits)?;
    let plan = plan_write_with_limits(value, conformance, limits)?;
    let xml = write_plan_with_limits(value, conformance, &plan, limits)?;
    super::graph::raw::admit_xml(&xml, limits)?;
    Ok(xml)
}

pub(in crate::glossary) fn write_plan_with_limits(
    value: &Catalog,
    conformance: Conformance,
    plan: &WritePlan,
    limits: &super::Limits,
) -> Result<Vec<u8>> {
    if plan.bytes > limits.xml_bytes() {
        return Err(Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::XmlBytes,
            actual: plan.bytes,
            maximum: limits.xml_bytes(),
        });
    }
    let mut xml = String::new();
    xml.try_reserve_exact(plan.bytes)
        .map_err(|source| Error::Allocation {
            resource: "glossary XML",
            source,
        })?;
    emit_catalog(&mut xml, value, conformance, plan)?;
    if xml.len() != plan.bytes {
        return Err(invalid("glossary XML write plan did not match output"));
    }
    Ok(xml.into_bytes())
}
