//! XLSX's typed creator for the unified operations surface.

use litchi_core::FileFormat;

use super::Xlsx;
use crate::create::{Conformance, CreateOptions, Flavor};
use crate::error::{Code, Error, Result};
use crate::options::OpenOptions;

/// Build and adopt the native minimal XLSX workbook directly.
pub(crate) fn create(options: &CreateOptions) -> Result<Xlsx> {
    if options.flavor_value() != Flavor::Document {
        return Err(unsupported_option(
            "the XLSX creator does not author template workbooks",
            "use Flavor::Document; template authoring is not integrated",
        ));
    }
    if options.conformance_value() == Conformance::Strict {
        return Err(unsupported_option(
            "the XLSX creator does not author Strict conformance",
            "use Conformance::Transitional; Strict authoring is not integrated",
        ));
    }
    if options.metadata_value().is_some() {
        return Err(unsupported_option(
            "the XLSX creator cannot author initial unified metadata",
            "omit initial metadata; unified XLSX metadata authoring is not integrated",
        ));
    }

    let workbook = litchi_xlsx::Workbook::create().map_err(Error::from)?;
    Xlsx::from_workbook(workbook, &OpenOptions::new())
}

fn unsupported_option(message: &'static str, suggestion: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message)
        .with_format(FileFormat::Xlsx)
        .with_suggestion(suggestion)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_creator_adopts_the_native_workbook() {
        let workbook = create(&CreateOptions::new()).expect("create XLSX");
        assert_eq!(workbook.format(), FileFormat::Xlsx);
        assert_eq!(workbook.worksheet_names().expect("sheet names"), ["Sheet1"]);
    }

    #[test]
    fn unsupported_creator_options_are_typed_refusals() {
        for options in [
            CreateOptions::new().flavor(Flavor::Template),
            CreateOptions::new().conformance(Conformance::Strict),
            CreateOptions::new().metadata(litchi_core::Metadata::default()),
        ] {
            let error = create(&options).expect_err("unsupported option");
            assert_eq!(error.code(), Code::UnsupportedCapability);
            assert_eq!(error.format(), Some(FileFormat::Xlsx));
        }
    }
}
