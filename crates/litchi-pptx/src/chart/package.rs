//! OPC graph ownership for ordinary `DrawingML` chart parts.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part as OpcPart};
use litchi_opc::{OpcPackage, PackURI};

use super::codec;
use super::model::Chart;
use crate::package::signature::{refuse_signed_change, settle_unsigned_change};
use crate::parts::{is_relationship_type, validate_content_type};
use crate::{Error, Result};

/// Ordinary `DrawingML` chart content type.
pub const CONTENT_TYPE: &str = ct::DML_CHART;

/// Borrowed ordinary chart part.
pub struct Part<'a> {
    pub(crate) part: &'a dyn OpcPart,
}

impl<'a> Part<'a> {
    /// Validate and borrow one chart part.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_part(part: &'a dyn OpcPart) -> Result<Self> {
        validate_content_type(part, CONTENT_TYPE)?;
        Ok(Self { part })
    }

    /// The underlying OPC part.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &'a dyn OpcPart {
        self.part
    }
}

/// Resolve one ordinary chart by package part name.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load<'a>(package: &'a OpcPackage, part_name: &str) -> Result<Part<'a>> {
    let uri = uri(part_name, "chart part")?;
    Part::from_part(package.get_part(&uri)?)
}

/// Resolve all ordinary charts directly related to a slide or other source part.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn related<'a>(package: &'a OpcPackage, source: &dyn OpcPart) -> Result<Vec<Part<'a>>> {
    let mut charts = Vec::new();
    for relationship in source.rels().iter() {
        if !is_relationship_type(relationship.reltype(), rt::CHART, "chart") {
            continue;
        }
        if relationship.is_external() {
            return Err(Error::Relationship(
                "chart relationships cannot be external".to_owned(),
            ));
        }
        let target = relationship.target_partname()?;
        charts.push(Part::from_part(package.get_part(&target)?)?);
    }
    Ok(charts)
}

/// Add a chart part and relationship to an existing source part.
///
/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn add(package: &mut OpcPackage, source_name: &str, chart: &Chart) -> Result<String> {
    let source_uri = uri(source_name, "chart source")?;
    let source = package.get_part(&source_uri)?;
    let _ = source;
    let target = next_uri(package)?;
    let target_ref = target.relative_ref(source_uri.base_uri());
    let xml = codec::encode(chart)?.into_bytes();
    refuse_signed_change(package)?;
    package.add_part(Box::new(BlobPart::new(
        target.clone(),
        CONTENT_TYPE.to_owned(),
        xml,
    )));
    let relationship_id = package
        .get_part_mut(&source_uri)?
        .relate_to(&target_ref, rt::CHART);
    settle_unsigned_change(package)?;
    Ok(relationship_id)
}

fn next_uri(package: &OpcPackage) -> Result<PackURI> {
    let mut index = 1u32;
    loop {
        let candidate =
            PackURI::new(format!("/ppt/charts/chart{index}.xml")).map_err(Error::Uri)?;
        if !package.contains_part(&candidate) {
            return Ok(candidate);
        }
        index = index
            .checked_add(1)
            .ok_or_else(|| Error::Invalid("chart part index overflow".to_owned()))?;
    }
}

fn uri(value: &str, label: &str) -> Result<PackURI> {
    PackURI::new(value).map_err(|error| Error::Uri(format!("{label}: {error}")))
}
