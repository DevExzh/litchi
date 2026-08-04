//! Compatibility adapter for the canonical XLSX Named Sheet Views owner.
//!
//! The typed model, bounded MCE/XML codec, inert retained markup, and OPC
//! graph operations live in `litchi_xlsx::named_sheet_view`. This module
//! preserves the historical host path and `OoxmlError` boundary, including
//! the crate-private discovery hook consumed by `worksheet.rs`.

use crate::error::{OoxmlError, Result};
use litchi_opc::{OpcPackage, PackURI, Relationships};

pub use litchi_xlsx::named_sheet_view::{
    ColumnFilter, DifferentialFormat, Extension, Filter, Guid, IconSet, Markup, Range,
    SortCondition, SortConditionKind, SortRule, SortRules, View, Views,
};

fn map_owner_error(error: litchi_xlsx::Error) -> OoxmlError {
    match error {
        litchi_xlsx::Error::Package(error) => OoxmlError::Opc(error),
        litchi_xlsx::Error::MarkupCompatibility(error) => OoxmlError::from(error),
        litchi_xlsx::Error::Xml(error) => OoxmlError::Xml(error.to_string()),
        litchi_xlsx::Error::Common(litchi_ooxml_common::Error::ContentType {
            expected,
            actual,
        }) => OoxmlError::InvalidContentType {
            expected,
            got: actual,
        },
        litchi_xlsx::Error::Common(litchi_ooxml_common::Error::Uri(message)) => {
            OoxmlError::InvalidUri(message)
        },
        litchi_xlsx::Error::Common(error) => OoxmlError::Common(error),
        litchi_xlsx::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
        litchi_xlsx::Error::Allocation { resource, source } => {
            OoxmlError::Allocation { resource, source }
        },
        other => OoxmlError::Xlsx(other),
    }
}

pub(crate) fn discover_named_sheet_views(
    package: &OpcPackage,
    relationships: &Relationships,
) -> Result<Option<Views>> {
    litchi_xlsx::named_sheet_view::discover_named_sheet_views(package, relationships)
        .map_err(map_owner_error)
}

pub fn parse_named_sheet_views(xml: &[u8]) -> Result<Views> {
    litchi_xlsx::named_sheet_view::parse_named_sheet_views(xml).map_err(map_owner_error)
}

pub fn load_worksheet_named_sheet_views(
    package: &OpcPackage,
    worksheet_part: &PackURI,
) -> Result<Option<Views>> {
    litchi_xlsx::named_sheet_view::load_worksheet_named_sheet_views(package, worksheet_part)
        .map_err(map_owner_error)
}

pub fn store_worksheet_named_sheet_views(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    value: &Views,
) -> Result<()> {
    litchi_xlsx::named_sheet_view::store_worksheet_named_sheet_views(package, worksheet_part, value)
        .map_err(map_owner_error)
}

pub fn remove_worksheet_named_sheet_views(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
) -> Result<bool> {
    litchi_xlsx::named_sheet_view::remove_worksheet_named_sheet_views(package, worksheet_part)
        .map_err(map_owner_error)
}

pub fn write_named_sheet_views(value: &Views) -> Result<Vec<u8>> {
    litchi_xlsx::named_sheet_view::write_named_sheet_views(value).map_err(map_owner_error)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::sort::SortBy;
    use litchi_opc::PackURI;

    fn fixture_worksheet() -> PackURI {
        PackURI::new("/xl/worksheets/sheet1.xml").unwrap()
    }

    #[test]
    fn authors_differential_formats_extensions_and_color_sorts() {
        let dxf = DifferentialFormat::from_xml(
            br#"<dxf xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:fill><x:patternFill patternType="solid"><x:fgColor rgb="FFFF0000"/></x:patternFill></x:fill></dxf>"#,
        )
        .unwrap();
        let extension = Extension::new(
            "urn:litchi:named-view",
            br#"<vendor:payload xmlns:vendor="urn:litchi:test" value="inert"/>"#,
        )
        .unwrap();

        let mut condition =
            SortCondition::new(SortConditionKind::Standard, Range::new("C2:C20").unwrap());
        condition.set_color_sort(SortBy::CellColor, 4).unwrap();
        let mut rule = SortRule::new(2).unwrap();
        rule.set_differential_format(Some(dxf.clone()))
            .unwrap()
            .set_condition(Some(condition))
            .unwrap();
        let mut rules = SortRules::new();
        rules
            .add_rule(rule)
            .unwrap()
            .add_extension(extension.clone())
            .unwrap();

        let mut column = ColumnFilter::new(2).unwrap();
        column
            .set_differential_format(Some(dxf))
            .add_extension(extension.clone())
            .unwrap();
        let mut filter = Filter::new(Guid::new("{11111111-2222-3333-4444-555555555555}").unwrap());
        filter
            .set_reference(Some(Range::new("A1:C20").unwrap()))
            .add_column_filter(column)
            .unwrap()
            .set_sort_rules(Some(rules))
            .add_extension(extension.clone())
            .unwrap();
        let mut view = View::with_id(
            "Colors",
            Guid::new("{01234567-89AB-CDEF-0123-456789ABCDEF}").unwrap(),
        )
        .unwrap();
        view.add_filter(filter)
            .unwrap()
            .add_extension(extension.clone())
            .unwrap();
        let mut authored = Views::new(view);
        authored.add_extension(extension).unwrap();

        let xml = authored.to_xml().unwrap();
        let text = std::str::from_utf8(&xml).unwrap();
        assert_eq!(text.matches("<dxf ").count(), 2);
        assert!(text.contains(r#"sortBy="cellColor" dxfId="4""#));
        assert_eq!(text.matches(r#"uri="urn:litchi:named-view""#).count(), 5);
        assert_eq!(parse_named_sheet_views(&xml).unwrap(), authored);

        let mut workbook = crate::xlsx::Workbook::create().unwrap();
        workbook.set_named_sheet_views(0, &authored).unwrap();
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("named-view-dxf-extension.xlsx");
        workbook.save(&path).unwrap();
        let reopened = crate::xlsx::Workbook::open(&path).unwrap();
        assert_eq!(reopened.named_sheet_views(0).unwrap(), Some(authored));
    }

    #[test]
    fn workbook_api_stores_constructed_views_through_materialization() {
        let mut workbook = crate::xlsx::Workbook::create().unwrap();
        let mut views = Views::new(
            View::with_id(
                "Personal",
                Guid::new("{01234567-89AB-CDEF-0123-456789ABCDEF}").unwrap(),
            )
            .unwrap(),
        );
        views.add_view(View::new("Shared").unwrap()).unwrap();

        workbook.set_named_sheet_views(0, &views).unwrap();
        assert_eq!(workbook.named_sheet_views(0).unwrap(), Some(views.clone()));

        workbook
            .worksheet_mut(0)
            .unwrap()
            .set_cell_value(1, 1, "materialized");
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("authored-named-sheet-views.xlsx");
        workbook.save(&path).unwrap();

        let mut reopened = crate::xlsx::Workbook::open(&path).unwrap();
        assert_eq!(reopened.named_sheet_views(0).unwrap(), Some(views));
        assert!(reopened.remove_named_sheet_views(0).unwrap());
        assert!(reopened.named_sheet_views(0).unwrap().is_none());
    }

    #[test]
    fn workbook_materialization_retains_named_sheet_views() {
        let mut workbook = crate::xlsx::Workbook::create().unwrap();
        let worksheet = fixture_worksheet();
        let value = parse_named_sheet_views(
            br#"<namedSheetViews xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews"><namedSheetView name="Retained" id="{01234567-89AB-CDEF-0123-456789ABCDEF}"/></namedSheetViews>"#,
        )
        .unwrap();
        store_worksheet_named_sheet_views(workbook.opc_package_mut(), &worksheet, &value).unwrap();
        workbook
            .worksheet_mut(0)
            .unwrap()
            .set_cell_value(1, 1, "materialized");

        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("materialized-named-sheet-views.xlsx");
        workbook.save(&path).unwrap();
        let reopened = crate::xlsx::Workbook::open(&path).unwrap();
        assert_eq!(
            load_worksheet_named_sheet_views(reopened.opc_package(), &worksheet).unwrap(),
            Some(value)
        );
    }
}
