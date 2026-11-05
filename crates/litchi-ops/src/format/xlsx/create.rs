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
    let workbook = litchi_xlsx::Package::new().map_err(Error::from)?;
    match options.metadata_value() {
        Some(metadata) => super::metadata::create(workbook, metadata),
        None => Xlsx::from_workbook(workbook, &OpenOptions::new()),
    }
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
    fn initial_metadata_is_authored_and_unstorable_fields_are_refused() {
        let metadata = litchi_core::Metadata {
            title: Some("Budget".to_owned()),
            company: Some("Litchi".to_owned()),
            ..litchi_core::Metadata::default()
        };
        let workbook = create(&CreateOptions::new().metadata(metadata)).expect("create XLSX");
        assert_eq!(
            workbook.metadata().unwrap().unwrap().title.as_deref(),
            Some("Budget")
        );
        assert_eq!(
            workbook.metadata().unwrap().unwrap().company.as_deref(),
            Some("Litchi")
        );

        let error = create(&CreateOptions::new().metadata(litchi_core::Metadata {
            codepage: Some(1252),
            ..litchi_core::Metadata::default()
        }))
        .expect_err("codepage is not an OOXML property");
        assert_eq!(error.code(), Code::UnsupportedProperty);
        assert_eq!(error.property(), Some("codepage"));
    }

    #[test]
    fn unsupported_creator_options_are_typed_refusals() {
        for options in [
            CreateOptions::new().flavor(Flavor::Template),
            CreateOptions::new().conformance(Conformance::Strict),
        ] {
            let error = create(&options).expect_err("unsupported option");
            assert_eq!(error.code(), Code::UnsupportedCapability);
            assert_eq!(error.format(), Some(FileFormat::Xlsx));
        }
    }
}
