//! Temporary host boundary for the canonical PresentationML structure owner.

use crate::error::{OoxmlError, Result};
use crate::pptx::customshow::CustomShow;
use crate::pptx::sections::Section;
use litchi_opc::OpcPackage;

pub use litchi_pptx::presentation_properties::metadata::structure::{
    Graph as Structure, Reference as SlideReference,
};

fn map<T>(value: litchi_pptx::Result<T>) -> Result<T> {
    value.map_err(OoxmlError::from)
}

pub fn load_presentation_structure(package: &OpcPackage) -> Result<Structure> {
    map(litchi_pptx::presentation_properties::metadata::structure::load(package))
}

pub fn store_presentation_structure(package: &mut OpcPackage, value: &Structure) -> Result<()> {
    map(litchi_pptx::presentation_properties::metadata::structure::store(package, value))
}

pub fn find_custom_show(package: &OpcPackage, id: u32) -> Result<Option<CustomShow>> {
    map(litchi_pptx::presentation_properties::metadata::structure::find_custom_show(package, id))
}

pub fn add_custom_show(package: &mut OpcPackage, show: CustomShow) -> Result<()> {
    map(litchi_pptx::presentation_properties::metadata::structure::add_custom_show(package, show))
}

pub fn update_custom_show(
    package: &mut OpcPackage,
    id: u32,
    replacement: CustomShow,
) -> Result<()> {
    map(
        litchi_pptx::presentation_properties::metadata::structure::update_custom_show(
            package,
            id,
            replacement,
        ),
    )
}

pub fn replace_custom_show(
    package: &mut OpcPackage,
    id: u32,
    replacement: CustomShow,
) -> Result<()> {
    map(
        litchi_pptx::presentation_properties::metadata::structure::replace_custom_show(
            package,
            id,
            replacement,
        ),
    )
}

pub fn remove_custom_show(package: &mut OpcPackage, id: u32) -> Result<bool> {
    map(litchi_pptx::presentation_properties::metadata::structure::remove_custom_show(package, id))
}

pub fn reorder_custom_shows(package: &mut OpcPackage, ordered_ids: &[u32]) -> Result<()> {
    map(
        litchi_pptx::presentation_properties::metadata::structure::reorder_custom_shows(
            package,
            ordered_ids,
        ),
    )
}

pub fn add_custom_show_slide(package: &mut OpcPackage, show_id: u32, slide_id: u32) -> Result<()> {
    map(
        litchi_pptx::presentation_properties::metadata::structure::add_custom_show_slide(
            package, show_id, slide_id,
        ),
    )
}

pub fn remove_custom_show_slide(
    package: &mut OpcPackage,
    show_id: u32,
    slide_id: u32,
) -> Result<bool> {
    map(
        litchi_pptx::presentation_properties::metadata::structure::remove_custom_show_slide(
            package, show_id, slide_id,
        ),
    )
}

pub fn reorder_custom_show_slides(
    package: &mut OpcPackage,
    show_id: u32,
    ordered_slide_ids: &[u32],
) -> Result<()> {
    map(
        litchi_pptx::presentation_properties::metadata::structure::reorder_custom_show_slides(
            package,
            show_id,
            ordered_slide_ids,
        ),
    )
}

pub fn find_section(package: &OpcPackage, id: &str) -> Result<Option<Section>> {
    map(litchi_pptx::presentation_properties::metadata::structure::find_section(package, id))
}

pub fn add_section(package: &mut OpcPackage, section: Section) -> Result<String> {
    map(litchi_pptx::presentation_properties::metadata::structure::add_section(package, section))
}

pub fn update_section(package: &mut OpcPackage, id: &str, replacement: Section) -> Result<()> {
    map(
        litchi_pptx::presentation_properties::metadata::structure::update_section(
            package,
            id,
            replacement,
        ),
    )
}

pub fn replace_section(package: &mut OpcPackage, id: &str, replacement: Section) -> Result<()> {
    map(
        litchi_pptx::presentation_properties::metadata::structure::replace_section(
            package,
            id,
            replacement,
        ),
    )
}

pub fn remove_section(package: &mut OpcPackage, id: &str) -> Result<bool> {
    map(litchi_pptx::presentation_properties::metadata::structure::remove_section(package, id))
}

pub fn reorder_sections(package: &mut OpcPackage, ordered_ids: &[String]) -> Result<()> {
    map(
        litchi_pptx::presentation_properties::metadata::structure::reorder_sections(
            package,
            ordered_ids,
        ),
    )
}

pub fn add_section_slide(package: &mut OpcPackage, section_id: &str, slide_id: u32) -> Result<()> {
    map(
        litchi_pptx::presentation_properties::metadata::structure::add_section_slide(
            package, section_id, slide_id,
        ),
    )
}

pub fn remove_section_slide(
    package: &mut OpcPackage,
    section_id: &str,
    slide_id: u32,
) -> Result<bool> {
    map(
        litchi_pptx::presentation_properties::metadata::structure::remove_section_slide(
            package, section_id, slide_id,
        ),
    )
}

pub fn reorder_section_slides(
    package: &mut OpcPackage,
    section_id: &str,
    ordered_slide_ids: &[u32],
) -> Result<()> {
    map(
        litchi_pptx::presentation_properties::metadata::structure::reorder_section_slides(
            package,
            section_id,
            ordered_slide_ids,
        ),
    )
}

pub fn synchronize_presentation_structure_after_slide_mutation(
    package: &mut OpcPackage,
) -> Result<()> {
    map(
        litchi_pptx::presentation_properties::metadata::structure::synchronize_after_slide_mutation(
            package,
        ),
    )
}
