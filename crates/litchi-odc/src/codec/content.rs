//! Chart content validation.

use litchi_core::Result;
use litchi_odf_common::chart::read;

/// Validate a UTF-8 content part before authoring it into a package.
pub(crate) fn validate(xml: &str) -> Result<()> {
    let chart = read(xml)?;
    let _ = chart.chart_class()?;
    if let Some(plot_area) = chart.plot_area() {
        for axis in plot_area.axes() {
            let _ = axis.dimension()?;
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::validate;

    #[test]
    fn requires_family_body() {
        assert!(validate(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><office:chart><chart:chart/></office:chart></office:body></office:document-content>"#
        )
        .is_err());
        assert!(validate(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><office:chart><chart:chart chart:class="chart:line"><chart:plot-area/></chart:chart></office:chart></office:body></office:document-content>"#
        )
        .is_ok());
        assert!(validate(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><office:chart><chart:chart><chart:plot-area/></chart:chart></office:chart></office:body></office:document-content>"#
        )
        .is_err());
        assert!(validate("<office:text/>").is_err());
    }
}
