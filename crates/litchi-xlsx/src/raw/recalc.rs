//! Workbook `calcPr` invalidation after semantic cell edits.

use crate::calculation_properties::{Limits, inspect, invalidate_formulas};
use crate::error::Result;

/// Force consumers to recalculate formulas while retaining the workbook's
/// chosen automatic/manual calculation mode.
pub(crate) fn invalidate(content: &[u8]) -> Result<Vec<u8>> {
    let limits = Limits::default();
    let inspection = inspect(content, &limits)?;
    Ok(invalidate_formulas(&inspection, &limits)?.into_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    #[test]
    fn inserts_or_updates_calc_properties_without_changing_other_markup() {
        let source = format!(
            r#"<x:workbook xmlns:x="{S}" xmlns:z="urn:future"><x:sheets/><x:pivotCaches z:keep="yes"/><x:extLst><z:data/></x:extLst></x:workbook>"#
        );
        let updated = invalidate(source.as_bytes()).expect("invalidate");
        let updated = std::str::from_utf8(&updated).expect("UTF-8");
        let calc = updated.find("<x:calcPr").expect("calcPr");
        let pivot = updated.find("<x:pivotCaches").expect("pivot caches");
        assert!(calc < pivot);
        assert!(updated.contains("calcId=\"0\""));
        assert!(updated.contains("<x:extLst><z:data/></x:extLst>"));

        let source = format!(
            r#"<workbook xmlns="{S}"><sheets/><calcPr calcMode="manual" calcId="42" z:future="kept" xmlns:z="urn:future"/></workbook>"#
        );
        let updated = invalidate(source.as_bytes()).expect("update");
        let updated = std::str::from_utf8(&updated).expect("UTF-8");
        assert!(updated.contains("calcMode=\"manual\""));
        assert!(updated.contains("z:future=\"kept\""));
        assert_eq!(updated.matches("calcId=").count(), 1);
    }

    #[test]
    fn refuses_to_rewrite_effective_calc_properties_projected_through_mce() {
        let source = format!(
            r#"<workbook xmlns="{S}" xmlns:mc="{MC}"><sheets/><mc:AlternateContent><mc:Choice Requires="future" xmlns:future="urn:future"><calcPr calcId="7"/></mc:Choice><mc:Fallback><calcPr calcId="42" calcMode="manual"/></mc:Fallback></mc:AlternateContent></workbook>"#
        );

        let error = invalidate(source.as_bytes()).expect_err("projected calcPr is not editable");
        assert!(
            error
                .to_string()
                .contains("cannot rewrite calcPr projected through MCE markup")
        );
    }

    #[test]
    fn rejects_hostile_nesting_at_the_shared_default_limit() {
        let mut source = format!(r#"<workbook xmlns="{S}">"#);
        for _ in 0..Limits::default().max_depth() {
            source.push_str("<future>");
        }
        for _ in 0..Limits::default().max_depth() {
            source.push_str("</future>");
        }
        source.push_str("</workbook>");

        let error = invalidate(source.as_bytes()).expect_err("nesting must be bounded");
        assert!(error.to_string().contains("nesting is too deep"));
    }
}
