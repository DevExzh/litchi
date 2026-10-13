//! PPTX's typed creator for the unified operations surface.

use litchi_core::FileFormat;

use super::Pptx;
use crate::create::{Conformance, CreateOptions, Flavor};
use crate::error::{Code, Error, Result};
use crate::options::OpenOptions;

/// Build and adopt the native minimal PPTX presentation directly.
pub(crate) fn create(options: &CreateOptions) -> Result<Pptx> {
    if options.flavor_value() != Flavor::Document {
        return Err(unsupported_option(
            "the PPTX creator does not author template presentations",
            "use Flavor::Document; template authoring is not integrated",
        ));
    }
    if options.conformance_value() == Conformance::Strict {
        return Err(unsupported_option(
            "the PPTX creator does not author Strict conformance",
            "use Conformance::Transitional; Strict authoring is not integrated",
        ));
    }
    let presentation = litchi_pptx::Package::new()?;
    match options.metadata_value() {
        Some(metadata) => super::metadata::create(presentation, metadata),
        None => Pptx::from_package(presentation, &OpenOptions::new()),
    }
}

fn unsupported_option(message: &'static str, suggestion: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message)
        .with_format(FileFormat::Pptx)
        .with_suggestion(suggestion)
}
