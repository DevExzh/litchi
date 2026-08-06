//! Package-context loading and transactional publication for zoom owners.

use litchi_opc::{OpcPackage, PackURI};

use crate::parts::SlidePart;
use crate::{Error, Result};

use super::model::Owner;

pub(crate) fn load(package: &OpcPackage, slide_name: &PackURI) -> Result<Owner> {
    let part = package.get_part(slide_name)?;
    let slide = SlidePart::from_part(part)?;
    load_part(package, &slide)
}

pub(crate) fn load_part(package: &OpcPackage, slide: &SlidePart<'_>) -> Result<Owner> {
    let mut owner = Owner::read(slide.part().blob())?;
    owner.validate_in_package(package, slide.part())?;
    Ok(owner)
}

pub(crate) fn store(
    package: &mut OpcPackage,
    slide_name: &PackURI,
    value: Owner,
) -> Result<Option<Owner>> {
    let (previous, staged) = {
        let part = package.get_part(slide_name)?;
        let slide = SlidePart::from_part(part)?;
        if slide.part().blob() != value.base_xml() {
            return Err(Error::UnsafeEdit {
                operation: "put_zooms",
                reason: "the zoom owner source changed after it was loaded",
            });
        }

        let mut previous = Owner::read(slide.part().blob())?;
        previous.validate_in_package(package, slide.part())?;

        let staged = value.to_xml()?;
        let mut candidate = Owner::read(&staged)?;
        candidate.validate_in_package(package, slide.part())?;
        (previous, staged)
    };

    if previous.xml() == staged {
        return Ok(Some(previous));
    }

    package.get_part_mut(slide_name)?.set_blob(staged);
    package.unsign();
    Ok(Some(previous))
}

pub(crate) fn remove(package: &mut OpcPackage, slide_name: &PackURI) -> Result<Option<Owner>> {
    let current = load(package, slide_name)?;
    if current.is_empty() {
        return Ok(None);
    }
    let previous = current.clone();
    let mut edited = current;
    edited.clear()?;
    let _ = store(package, slide_name, edited)?;
    Ok(Some(previous))
}
