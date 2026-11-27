//! OOXML workbook metadata read shared by the XLSX and XLSB handles: the
//! core properties with their repairs, then the application properties.

use litchi_core::Metadata;

/// Core-properties repairs of an OOXML workbook's metadata read, folded into
/// the format report by the handle's `tolerance_report`.
pub(crate) type MetadataRepairs = litchi_ooxml_common::leniency::ToleranceReport;

/// Read an OOXML workbook's core and application properties, with the core
/// repairs.
pub(crate) fn read_source_backed_metadata(
    package: &litchi_opc::SourceBackedPackage,
) -> litchi_ooxml_common::Result<(Metadata, MetadataRepairs)> {
    let mut repairs = MetadataRepairs::default();
    let metadata =
        litchi_ooxml_common::properties::read_source_backed_with_report(package, &mut repairs)?
            .unwrap_or_default();
    let metadata = crate::format::metadata::with_source_app_props(metadata, package);
    Ok((metadata, repairs))
}
