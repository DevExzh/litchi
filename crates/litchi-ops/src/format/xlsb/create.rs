//! XLSB's typed creator for the unified operations surface.

use litchi_core::FileFormat;

use super::Xlsb;
use crate::create::{Conformance, CreateOptions, Flavor};
use crate::error::{Code, Error, Result};
use crate::options::OpenOptions;

/// Build and adopt the native minimal XLSB workbook directly.
pub(crate) fn create(options: &CreateOptions) -> Result<Xlsb> {
    if options.flavor_value() != Flavor::Document {
        return Err(unsupported_option(
            "the XLSB creator does not author template workbooks",
            "use Flavor::Document; template authoring is not integrated",
        ));
    }
    if options.conformance_value() == Conformance::Strict {
        return Err(unsupported_option(
            "the XLSB creator does not author Strict conformance",
            "use Conformance::Transitional; Strict authoring is not integrated",
        ));
    }
    let package = litchi_xlsb::Package::new().map_err(Error::from)?;
    match options.metadata_value() {
        Some(metadata) => super::metadata::create(package, metadata),
        None => Xlsb::from_package(package, &OpenOptions::new()),
    }
}

fn unsupported_option(message: &'static str, suggestion: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message)
        .with_format(FileFormat::Xlsb)
        .with_suggestion(suggestion)
}
