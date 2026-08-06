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
use super::{Error, MAX, MAX_DEPTH, MAX_VALUES, Result};

pub fn read(xml: &[u8]) -> Result<(Catalog, Conformance)> {
    if xml.len() > MAX {
        return Err(invalid("glossary document exceeds 32 MiB"));
    }
    let original = parse_dom(xml)?;
    let original_conformance = Conformance::from_word(original.ns.as_ref())?;
    let producer_entries = extract_producer_entries(original, original_conformance)?;
    let limits = litchi_ooxml_common::mce::Limits {
        max_input_bytes: MAX,
        max_output_bytes: MAX,
        max_depth: MAX_DEPTH,
        max_namespace_bindings: MAX_VALUES,
        max_directive_tokens: MAX_VALUES,
        max_choices_per_alternate: MAX_VALUES,
    };
    let xml = litchi_ooxml_common::mce::process_markup_compatibility(
        xml,
        &litchi_ooxml_common::mce::Capabilities::default(),
        &limits,
    )?
    .xml;
    if xml.len() > MAX {
        return Err(invalid("processed glossary document exceeds 32 MiB"));
    }
    let root = parse_dom(xml.as_ref())?;
    let conformance = Conformance::from_word(root.ns.as_ref())?;
    if conformance != original_conformance {
        return Err(invalid(
            "MCE preprocessing changed the glossary conformance family",
        ));
    }
    validate_word_dialect(&root, conformance)?;
    let mut catalog = project(&root)?;
    drop(root);
    attach_producer_entries(&mut catalog, producer_entries);
    catalog.rebuild_state()?;
    Ok((catalog, conformance))
}

/// Serialize a catalog canonically in the selected dialect.
pub fn write(value: &Catalog, conformance: Conformance) -> Result<Vec<u8>> {
    let plan = plan_write(value, conformance)?;
    let mut xml = String::new();
    xml.try_reserve_exact(plan.bytes)
        .map_err(|source| Error::Allocation {
            resource: "glossary XML",
            source,
        })?;
    emit_catalog(&mut xml, value, conformance, &plan)?;
    if xml.len() != plan.bytes {
        return Err(invalid("glossary XML write plan did not match output"));
    }
    Ok(xml.into_bytes())
}
