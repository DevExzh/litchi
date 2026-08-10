//! OPC package ownership for the workbook calculation-chain part.

use std::borrow::Cow;

use crate::error::{Error, Result, allocation, invalid};
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::codec::{read, read_with_projection, write};
use super::model::{
    CONTENT_TYPE, Chain, Conformance, MAX_XML_BYTES, RELATIONSHIP, STRICT_NS, STRICT_RELATIONSHIP,
    TRANSITIONAL_NS,
};

pub(super) const MAX_WORKBOOK_SHEETS: usize = 65_534;

/// Load the optional inert calculation chain and its relationship conformance.
/// Formula cells are parsed as metadata only; no formula is evaluated.
pub fn load(package: &OpcPackage) -> Result<Option<(Chain, Conformance)>> {
    let workbook_uri = main_workbook_uri(package)?;
    load_for_workbook(package, &workbook_uri)
}

/// Validate the package topology for the optional calculation chain without
/// decoding its inert XML payload.
pub(crate) fn validate_package(package: &OpcPackage) -> Result<()> {
    let workbook_uri = main_workbook_uri(package)?;
    let Some(relationship) = relationship(package, &workbook_uri)? else {
        validate_part_set(package, None)?;
        return Ok(());
    };
    validate_part_set(package, Some(&relationship.part_name))?;
    validate_part(package, &relationship.part_name)
}

/// Store a caller-authored inert calculation chain in a `SpreadsheetML` package.
///
/// The supplied order is serialized without recalculating formulas or inferring
/// dependencies. Existing calculation-chain graph violations are rejected
/// before any package part is changed. The requested conformance is applied to
/// both the part XML and its workbook relationship.
pub fn put(package: &mut OpcPackage, chain: &Chain, conformance: Conformance) -> Result<bool> {
    let mut staged = package.clone();
    let changed = put_staged(&mut staged, chain, conformance)?;
    if changed {
        *package = staged;
    }
    Ok(changed)
}

fn put_staged(package: &mut OpcPackage, chain: &Chain, conformance: Conformance) -> Result<bool> {
    let xml = write(chain, conformance)?;
    let workbook_uri = main_workbook_uri(package)?;
    validate_sheet_ids(package, &workbook_uri, chain)?;
    let existing = relationship(package, &workbook_uri)?;

    if let Some(existing) = existing {
        validate_part_set(package, Some(&existing.part_name))?;
        validate_part(package, &existing.part_name)?;
        let bytes_changed = package.get_part(&existing.part_name)?.blob() != xml;
        let relationship_changed = existing.conformance != conformance;
        if !bytes_changed && !relationship_changed {
            return Ok(false);
        }
        if package.is_signed() {
            return Err(Error::Signed);
        }
        if bytes_changed && read_with_projection(package.get_part(&existing.part_name)?.blob())?.1 {
            return Err(invalid(
                "cannot replace calculation-chain source projected through MCE",
            ));
        }
        if bytes_changed {
            package.get_part_mut(&existing.part_name)?.set_blob(xml);
        }
        if relationship_changed {
            let workbook = package.get_part_mut(&workbook_uri)?;
            workbook.rels_mut().remove(&existing.relationship_id);
            workbook.rels_mut().add_relationship(
                conformance.relationship_type().into(),
                existing.target_reference,
                existing.relationship_id,
                false,
            );
        }
    } else {
        validate_part_set(package, None)?;
        if package.is_signed() {
            return Err(Error::Signed);
        }
        let part_name = next_part_name(package)?;
        let relationship_id = next_relationship_id(package, &workbook_uri)?;
        let target = part_name.relative_ref(workbook_uri.base_uri());
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part_name.clone(),
            CONTENT_TYPE.into(),
            xml,
        )))?;
        let workbook = match package.get_part_mut(&workbook_uri) {
            Ok(workbook) => workbook,
            Err(error) => {
                package.remove_part(&part_name);
                return Err(error.into());
            },
        };
        workbook.rels_mut().add_relationship(
            conformance.relationship_type().into(),
            target,
            relationship_id,
            false,
        );
    }

    Ok(true)
}

/// Remove the workbook's calculation-chain relationship and its unreferenced part.
///
/// No formulas are changed. A target that is also referenced elsewhere in the
/// package is retained.
pub fn remove(package: &mut OpcPackage) -> Result<bool> {
    let mut staged = package.clone();
    let changed = remove_staged(&mut staged)?;
    if changed {
        *package = staged;
    }
    Ok(changed)
}

fn remove_staged(package: &mut OpcPackage) -> Result<bool> {
    let workbook_uri = main_workbook_uri(package)?;
    let Some(existing) = relationship(package, &workbook_uri)? else {
        validate_part_set(package, None)?;
        return Ok(false);
    };
    validate_part_set(package, Some(&existing.part_name))?;
    validate_part(package, &existing.part_name)?;
    let retain_part = part_is_referenced_elsewhere(
        package,
        &existing.part_name,
        &workbook_uri,
        &existing.relationship_id,
    )?;
    if package.is_signed() {
        return Err(Error::Signed);
    }

    package
        .get_part_mut(&workbook_uri)?
        .rels_mut()
        .remove(&existing.relationship_id);
    if !retain_part {
        package.remove_part(&existing.part_name);
    }
    Ok(true)
}

fn validate_sheet_ids(package: &OpcPackage, workbook_uri: &PackURI, chain: &Chain) -> Result<()> {
    let workbook = package.get_part(workbook_uri)?;
    if workbook.blob().len() > MAX_XML_BYTES {
        return Err(invalid(format!(
            "workbook XML exceeds {MAX_XML_BYTES} bytes while validating calculation-chain sheets"
        )));
    }
    let processed = process_workbook_mce_with_limit(workbook.blob(), MAX_XML_BYTES)?;
    if processed.len() > MAX_XML_BYTES {
        return Err(invalid(format!(
            "processed workbook XML exceeds {MAX_XML_BYTES} bytes while validating calculation-chain sheets"
        )));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    // The chain sheet domain provides a hard catalog cardinality bound. Reserve
    // its complete u32 representation fallibly once, so parsing and insertion
    // cannot trigger hidden allocator growth after MCE projection.
    let mut catalog = Vec::<u32>::new();
    catalog
        .try_reserve_exact(MAX_WORKBOOK_SHEETS)
        .map_err(|source| allocation("calculation-chain workbook sheet catalog", source))?;
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid workbook XML: {error}")))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"sheet"
                    && matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == TRANSITIONAL_NS.as_bytes() || value == STRICT_NS.as_bytes()) =>
            {
                let mut sheet_id = None;
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(|error| {
                        invalid(format!("invalid workbook sheet attribute: {error}"))
                    })?;
                    if attribute.key.as_ref() == b"sheetId" {
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                            .map_err(|error| {
                                invalid(format!("invalid workbook sheetId: {error}"))
                            })?;
                        let value = value.parse::<u32>().map_err(|_source| {
                            invalid("workbook sheetId is not an unsigned integer")
                        })?;
                        if sheet_id.replace(value).is_some() {
                            return Err(invalid("duplicate workbook sheetId attribute"));
                        }
                    }
                }
                let sheet_id =
                    sheet_id.ok_or_else(|| invalid("workbook sheet requires sheetId"))?;
                if catalog.len() >= MAX_WORKBOOK_SHEETS {
                    return Err(invalid(format!(
                        "workbook sheet catalog exceeds {MAX_WORKBOOK_SHEETS} entries"
                    )));
                }
                catalog.push(sheet_id);
            },
            Event::Eof => break,
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    catalog.sort_unstable();
    if let Some(duplicate) = catalog
        .windows(2)
        .find_map(|pair| (pair[0] == pair[1]).then_some(pair[0]))
    {
        return Err(invalid(format!("duplicate workbook sheet ID {duplicate}")));
    }
    for cell in chain.cells() {
        let sheet_id = u32::from(cell.sheet().get());
        if catalog.binary_search(&sheet_id).is_err() {
            return Err(invalid(format!(
                "calculation-chain sheet ID {sheet_id} does not resolve to a workbook sheet"
            )));
        }
    }
    Ok(())
}

pub(super) fn process_workbook_mce_with_limit(
    xml: &[u8],
    max_xml_bytes: usize,
) -> Result<Cow<'_, [u8]>> {
    let limits = litchi_ooxml_common::mce::Limits {
        max_input_bytes: max_xml_bytes,
        max_output_bytes: max_xml_bytes,
        ..litchi_ooxml_common::mce::Limits::default()
    };
    litchi_ooxml_common::mce::process_markup_compatibility(
        xml,
        &litchi_ooxml_common::mce::Capabilities::default(),
        &limits,
    )
    .map(|output| output.xml)
    .map_err(|error| invalid(format!("workbook MCE error: {error}")))
}

fn load_for_workbook(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<(Chain, Conformance)>> {
    let Some(relationship) = relationship(package, workbook_uri)? else {
        validate_part_set(package, None)?;
        return Ok(None);
    };
    validate_part_set(package, Some(&relationship.part_name))?;
    validate_part(package, &relationship.part_name)?;
    let part = package.get_part(&relationship.part_name)?;
    Ok(Some((read(part.blob())?, relationship.conformance)))
}

#[derive(Debug, Clone)]
struct Relationship {
    relationship_id: String,
    part_name: PackURI,
    target_reference: String,
    conformance: Conformance,
}

fn relationship(package: &OpcPackage, workbook_uri: &PackURI) -> Result<Option<Relationship>> {
    let workbook = package.get_part(workbook_uri)?;
    let mut relationships = workbook.rels().iter().filter(|relationship| {
        matches!(relationship.reltype(), RELATIONSHIP | STRICT_RELATIONSHIP)
    });
    let Some(relationship) = relationships.next() else {
        return Ok(None);
    };
    if relationships.next().is_some() {
        return Err(invalid(
            "workbook has multiple calculation-chain relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid("calculation-chain relationship cannot be external"));
    }
    let conformance = if relationship.reltype() == RELATIONSHIP {
        Conformance::Transitional
    } else {
        Conformance::Strict
    };
    Ok(Some(Relationship {
        relationship_id: relationship.r_id().to_string(),
        part_name: relationship.target_partname()?,
        target_reference: relationship.target_ref().to_string(),
        conformance,
    }))
}

fn validate_part(package: &OpcPackage, part_name: &PackURI) -> Result<()> {
    let part = package.get_part(part_name)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(invalid(format!(
            "calculation-chain part '{part_name}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    if part.rels().iter().next().is_some() {
        return Err(invalid("calculation-chain part cannot have relationships"));
    }
    Ok(())
}

fn validate_part_set(package: &OpcPackage, relationship_target: Option<&PackURI>) -> Result<()> {
    let mut parts = package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE);
    let part_name = parts.next().map(litchi_opc::Part::partname);
    if parts.next().is_some() {
        return Err(invalid(
            "package contains more than one calculation-chain part",
        ));
    }
    match (relationship_target, part_name) {
        (None, None) => Ok(()),
        (None, Some(_)) => Err(invalid(
            "package contains a calculation-chain part without a workbook relationship",
        )),
        (Some(_), None) => Ok(()),
        (Some(target), Some(part_name)) if part_name == target => Ok(()),
        (Some(_), Some(_)) => Err(invalid(
            "workbook calculation-chain relationship does not target the calculation-chain part",
        )),
    }
}

fn main_workbook_uri(package: &OpcPackage) -> Result<PackURI> {
    use litchi_opc::constants::content_type as ct;

    let workbook = package.main_document_part()?;
    if !matches!(
        workbook.content_type(),
        ct::SML_SHEET_MAIN
            | ct::SML_TEMPLATE_MAIN
            | ct::SML_SHEET_MACRO_MAIN
            | ct::SML_TEMPLATE_MACRO_MAIN
    ) {
        return Err(invalid(format!(
            "main document part '{}' is not an XML workbook",
            workbook.partname()
        )));
    }
    Ok(workbook.partname().clone())
}

fn next_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 0..=65_536u32 {
        let name = if suffix == 0 {
            "/xl/calcChain.xml".to_string()
        } else {
            format!("/xl/calcChain{suffix}.xml")
        };
        let candidate = PackURI::new(&name).map_err(invalid)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free calculation-chain part name"))
}

fn next_relationship_id(package: &OpcPackage, workbook_uri: &PackURI) -> Result<String> {
    let relationships = package.get_part(workbook_uri)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdCalcChain{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free calculation-chain relationship ID"))
}

fn part_is_referenced_elsewhere(
    package: &OpcPackage,
    target: &PackURI,
    owner: &PackURI,
    owner_relationship: &str,
) -> Result<bool> {
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if part.partname() == owner && relationship.r_id() == owner_relationship {
                continue;
            }
            if !relationship.is_external() && relationship.target_partname()? == *target {
                return Ok(true);
            }
        }
    }
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()? == *target {
            return Ok(true);
        }
    }
    Ok(false)
}
