//! DOCX's typed creator for the unified operations surface.

use litchi_core::Metadata;

use super::Docx;
use crate::create::{self, Conformance, CreateOptions, Flavor};
use crate::error::{Error, Result};
use crate::options::OpenOptions;

/// Build and adopt a DOCX package without a ZIP serialization round trip.
pub(crate) fn create(options: &CreateOptions) -> Result<Docx> {
    if options.conformance_value() == Conformance::Strict {
        return Err(create::unsupported_docx_option(
            "the DOCX creator does not author Strict conformance",
            "use Conformance::Transitional; Strict authoring is not integrated",
        ));
    }

    let mut package = match options.flavor_value() {
        Flavor::Document => litchi_docx::Package::new().map_err(Error::from)?,
        Flavor::Template => litchi_docx::Package::new_template().map_err(Error::from)?,
    };

    if let Some(metadata) = options.metadata_value() {
        apply_metadata(&mut package, metadata)?;
    }

    // A created package is already an owning leaf package.  Adopt it directly
    // so the unified handle does not decode, serialize, and reopen the ZIP.
    Docx::from_owned_package(package, &OpenOptions::new())
}

fn apply_metadata(package: &mut litchi_docx::Package, metadata: &Metadata) -> Result<()> {
    // Creation stores core properties only; any other field is refused
    // instead of silently disappearing.
    if litchi_core::metadata::Field::ALL.into_iter().any(|field| {
        metadata.is_set(field) && !litchi_ooxml_common::properties::CORE_FIELDS.contains(&field)
    }) {
        return Err(create::invalid_created_metadata());
    }
    let mut edit = package.edit_metadata();
    edit.set(metadata.clone());
    *package = edit.commit()?.into_package();
    Ok(())
}
