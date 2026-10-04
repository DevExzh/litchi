//! OOXML workbook core-properties read shared by the workbook source-path
//! detector and the `litchi` workbook facade.

use litchi_core::Metadata;

/// Core-properties repairs of an OOXML workbook's metadata read, folded into
/// the format report by the workbook facade's `tolerance_report`.
pub type MetadataRepairs = litchi_ooxml_common::leniency::ToleranceReport;

/// Read an OOXML workbook's core properties with their repairs.
pub fn read_source_backed_metadata(
    package: &litchi_opc::SourceBackedPackage,
) -> litchi_ooxml_common::Result<(Metadata, MetadataRepairs)> {
    let mut repairs = MetadataRepairs::default();
    let metadata =
        litchi_ooxml_common::properties::read_source_backed_with_report(package, &mut repairs)?
            .map(Metadata::from)
            .unwrap_or_default();
    Ok((metadata, repairs))
}
