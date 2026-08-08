#![cfg(feature = "vba-inspection")]

#[cfg(feature = "docx")]
#[test]
fn docx_exposes_inert_vba_inspection() {
    let _ = litchi::docx::vba_project::VbaSupplementalData::new();
}

#[cfg(feature = "ooxml-common")]
#[test]
fn common_exposes_inert_vba_inspection() {
    let _ = litchi::ooxml_common::vba::Host::Excel;
}

#[cfg(feature = "pptx")]
#[test]
fn pptx_exposes_inert_vba_inspection() {
    fn name_project(_: Option<litchi::pptx::presentation::embedded::vba::Project<'_>>) {}
    name_project(None);
}

#[cfg(feature = "xlsb")]
#[test]
fn xlsb_exposes_inert_vba_inspection() {
    fn name_project(_: Option<litchi::xlsb::package::vba_project::VbaProject>) {}
    name_project(None);
}

#[cfg(feature = "xlsx")]
#[test]
fn xlsx_exposes_inert_vba_inspection() {
    fn name_project(_: Option<litchi::xlsx::active_x::Project>) {}
    name_project(None);
}
