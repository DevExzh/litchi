//! OOXML workbook core-properties read shared by the XLSX and XLSB handles.

use litchi_core::Metadata;

/// Core-properties repairs of an OOXML workbook's metadata read, folded into
/// the format report by the handle's `tolerance_report`.
pub(crate) type MetadataRepairs = litchi_ooxml_common::leniency::ToleranceReport;

/// Read an OOXML workbook's core properties with their repairs.
pub(crate) fn read_source_backed_metadata(
    package: &litchi_opc::SourceBackedPackage,
) -> litchi_ooxml_common::Result<(Metadata, MetadataRepairs)> {
    let mut repairs = MetadataRepairs::default();
    let metadata =
        litchi_ooxml_common::properties::read_source_backed_with_report(package, &mut repairs)?
            .map(Metadata::from)
            .unwrap_or_default();
    Ok((metadata, repairs))
}
