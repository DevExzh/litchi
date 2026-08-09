//! Chart content validation.

use litchi_core::Result;
use litchi_odf_common::chart::{Element, read};

const CHART: &str = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

/// Validate a UTF-8 content part before authoring it into a package.
pub(crate) fn validate(xml: &str) -> Result<()> {
    let chart = read(xml)?;
    let _ = chart.chart_class()?;
    if let Some(plot_area) = chart.plot_area() {
        for axis in plot_area.axes() {
            let _ = axis.dimension()?;
        }
    }
    validate_tree(&chart, crate::Limits::default())?;
    Ok(())
}

pub(crate) fn validate_tree(root: &Element, limits: crate::Limits) -> Result<()> {
    let mut stack = vec![root];
    while let Some(element) = stack.pop() {
        for attribute in element.attributes() {
            let is_range = (attribute.namespace_uri() == Some(CHART)
                && matches!(
                    attribute.local_name(),
                    "values-cell-range-address" | "label-cell-address"
                ))
                || (attribute.namespace_uri() == Some(TABLE)
                    && matches!(attribute.local_name(), "cell-range-address" | "cell-range"));
            if is_range {
                if attribute.value().len() > limits.max_scalar_bytes() {
                    return Err(litchi_core::Error::InvalidFormat(
                        "ODC range exceeds the caller-selected scalar limit".into(),
                    ));
                }
                crate::validate_range_list(attribute.value())?;
            }
            if attribute.namespace_uri() == Some(TABLE) && attribute.local_name() == "formula" {
                if attribute.value().len() > limits.max_scalar_bytes() {
                    return Err(litchi_core::Error::InvalidFormat(
                        "ODC formula exceeds the caller-selected scalar limit".into(),
                    ));
                }
                crate::validate_formula(attribute.value())?;
            }
        }
        stack.extend(element.children());
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
