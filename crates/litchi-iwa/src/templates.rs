//! Native starting packages used by the zero-input application editors.

use crate::{IWorkPackage, Result};

const BLANK_PAGES: &[u8] = include_bytes!("../templates/blank.pages");
const BLANK_NUMBERS: &[u8] = include_bytes!("../templates/blank.numbers");
const BASIC_WHITE_KEYNOTE: &[u8] = include_bytes!("../templates/basic-white.key");

pub(crate) fn blank_pages_package() -> Result<IWorkPackage> {
    independent_package(BLANK_PAGES)
}

pub(crate) fn blank_numbers_package() -> Result<IWorkPackage> {
    independent_package(BLANK_NUMBERS)
}

pub(crate) fn basic_white_keynote_package() -> Result<IWorkPackage> {
    independent_package(BASIC_WHITE_KEYNOTE)
}

fn independent_package(template: &[u8]) -> Result<IWorkPackage> {
    let mut package = IWorkPackage::from_bytes(template)?;
    package.regenerate_document_identity()?;
    Ok(package)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Document;
    use crate::registry::Application;

    #[test]
    fn embedded_starting_packages_are_valid_native_documents() {
        for (package, expected) in [
            (blank_pages_package().unwrap(), Application::Pages),
            (blank_numbers_package().unwrap(), Application::Numbers),
            (basic_white_keynote_package().unwrap(), Application::Keynote),
        ] {
            let document = Document::from_bytes(&package.to_bytes().unwrap()).unwrap();
            assert_eq!(document.application(), expected);
            assert!(!document.objects().is_empty());
        }
    }
}
