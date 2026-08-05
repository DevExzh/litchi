//! ODF data-pilot (pivot-table) declarations.
//!
//! The public facade stays intentionally small while the implementation is
//! divided by responsibility: [`model`] owns the typed vocabulary, [`range`]
//! owns spreadsheet-address semantics, [`validation`] owns cross-object
//! invariants, and [`codec`] owns the XML boundary.

mod codec;
mod model;
mod range;
mod validation;

pub use model::{
    DisplayInfo, DisplayMemberMode, Field, FieldReference, GrandTotal, GrandTotalElement,
    GrandTotalOrientation, Group, GroupBoundary, GroupBy, Groups, LayoutInfo, LayoutMode, Level,
    Member, Orientation, ReferenceMemberType, ReferenceType, SortInfo, SortMode, SortOrder, Source,
    Table,
};

pub(crate) fn parse_data_pilot_tables(xml: &str) -> litchi_core::Result<Vec<Table>> {
    codec::parse_data_pilot_tables(xml)
}

pub(crate) fn write_data_pilot_tables(
    output: &mut String,
    tables: &[Table],
) -> litchi_core::Result<()> {
    codec::write_data_pilot_tables(output, tables)
}

pub(crate) fn write_data_pilot_table_fragment(table: &Table) -> litchi_core::Result<String> {
    codec::write_data_pilot_table_fragment(table)
}

pub(crate) fn parse_data_pilot_range(value: &str) -> litchi_core::Result<range::ParsedRange> {
    range::parse_data_pilot_range(value)
}

pub(crate) fn validate_data_pilot_tables(tables: &[Table]) -> litchi_core::Result<()> {
    validation::validate_data_pilot_tables(tables)
}

const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_EXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:office:xmlns:table:1.0";
const CALC_EXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";
const MAX_DATA_PILOT_TABLES: usize = 65_536;
const MAX_DATA_PILOT_FIELDS: usize = 65_536;
const MAX_DATA_PILOT_ITEMS: usize = 1_000_000;
const MAX_DATA_PILOT_STRING: usize = 1024 * 1024;

pub(super) fn invalid(kind: &str, value: &str) -> litchi_core::Error {
    invalid_message(&format!("invalid {kind} '{value}'"))
}

pub(super) fn invalid_message(message: &str) -> litchi_core::Error {
    litchi_core::Error::InvalidFormat(message.to_string())
}

// End-to-end cases require the transactional spreadsheet facade; retained until
// that package owner can be wired without cross-family dependencies.
#[cfg(any())]
mod tests {
    use super::*;

    const XMLNS: &str = r#"xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"#;

    fn complete_xml() -> String {
        format!(
            r#"<o:document-content {XMLNS}><o:body><o:spreadsheet>
            <t:data-pilot-tables><t:data-pilot-table t:name="Pivot &amp; One" t:application-data="app"
              t:grand-total="both" t:ignore-empty-rows="true" t:identify-categories="false"
              t:target-range-address="Result.A1:F20" t:buttons="Result.A1 Result.B1"
              t:show-filter-button="1" t:drill-down-on-double-click="0">
              <t:source-cell-range t:cell-range-address="Source.A1:D100">
                <t:filter t:display-duplicates="false"><t:filter-condition t:field-number="0" t:value="East" t:operator="="/></t:filter>
              </t:source-cell-range>
              <t:data-pilot-field t:source-field-name="Region" t:orientation="row" t:function="auto" t:used-hierarchy="1">
                <t:data-pilot-level t:show-empty="false">
                  <t:data-pilot-subtotals><t:data-pilot-subtotal t:function="sum"></t:data-pilot-subtotal></t:data-pilot-subtotals>
                  <t:data-pilot-members><t:data-pilot-member t:name="East" t:display="true" t:show-details="false"></t:data-pilot-member></t:data-pilot-members>
                  <t:data-pilot-display-info t:enabled="true" t:data-field="Sales" t:member-count="10" t:display-member-mode="from-top"></t:data-pilot-display-info>
                  <t:data-pilot-sort-info t:sort-mode="data" t:data-field="Sales" t:order="descending"></t:data-pilot-sort-info>
                  <t:data-pilot-layout-info t:layout-mode="outline-subtotals-top" t:add-empty-lines="true"></t:data-pilot-layout-info>
                </t:data-pilot-level>
                <t:data-pilot-field-reference t:field-name="Region" t:member-type="named" t:member-name="East" t:type="member-percentage"></t:data-pilot-field-reference>
                <t:data-pilot-groups t:source-field-name="Region" t:start="0" t:end="100" t:step="10" t:grouped-by="days">
                  <t:data-pilot-group t:name="Area"><t:data-pilot-group-member t:name="East"></t:data-pilot-group-member></t:data-pilot-group>
                </t:data-pilot-groups>
              </t:data-pilot-field>
              <t:data-pilot-field t:source-field-name="Page" t:orientation="page" t:selected-page="All"/>
            </t:data-pilot-table></t:data-pilot-tables>
            </o:spreadsheet></o:body></o:document-content>"#
        )
    }

    #[test]
    fn parses_all_standard_metadata_with_namespace_aliases() {
        let tables = parse_data_pilot_tables(&complete_xml()).unwrap();
        assert_eq!(tables.len(), 1);
        let table = &tables[0];
        assert_eq!(table.name, "Pivot & One");
        assert_eq!(table.grand_total, Some(GrandTotal::Both));
        assert_eq!(table.fields.len(), 2);
        assert_eq!(table.fields[0].level.as_ref().unwrap().subtotals, ["sum"]);
        assert_eq!(
            table.fields[0]
                .level
                .as_ref()
                .unwrap()
                .sort
                .as_ref()
                .unwrap()
                .mode,
            SortMode::Data
        );
        assert_eq!(
            table.fields[0].groups.as_ref().unwrap().groups[0].members,
            ["East"]
        );
        assert!(matches!(
            table.source,
            Some(Source::CellRange {
                filter: Some(_),
                ..
            })
        ));
    }

    #[test]
    fn writer_round_trips_complete_declaration() {
        let tables = parse_data_pilot_tables(&complete_xml()).unwrap();
        let mut body = String::new();
        write_data_pilot_tables(&mut body, &tables).unwrap();
        assert!(body.contains("Pivot &amp; One"));
        assert!(body.contains("<table:filter-condition"));
        let wrapped = format!(
            r#"<o:spreadsheet {XMLNS} xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">{body}</o:spreadsheet>"#
        );
        let reparsed = parse_data_pilot_tables(&wrapped).unwrap();
        assert_eq!(reparsed, tables);
    }

    #[test]
    fn rejects_schema_invalid_declarations() {
        for body in [
            r#"<t:data-pilot-tables><t:data-pilot-table t:name="P" t:target-range-address="S.A1"/></t:data-pilot-tables>"#,
            r#"<t:data-pilot-tables><t:data-pilot-table t:name="P" t:target-range-address="S.A1"><t:data-pilot-field t:source-field-name="F" t:orientation="page"/></t:data-pilot-table></t:data-pilot-tables>"#,
            r#"<t:data-pilot-tables><t:data-pilot-table t:name="P" t:target-range-address="S.A1"><t:data-pilot-field t:source-field-name="F" t:orientation="row" t:selected-page="X"/></t:data-pilot-table></t:data-pilot-tables>"#,
            r#"<t:data-pilot-tables><t:data-pilot-table t:name="P" t:target-range-address="S.A1"><t:data-pilot-field t:source-field-name="F" t:orientation="sideways"/></t:data-pilot-table></t:data-pilot-tables>"#,
        ] {
            let xml = format!(r#"<o:spreadsheet {XMLNS}>{body}</o:spreadsheet>"#);
            assert!(parse_data_pilot_tables(&xml).is_err(), "{body}");
        }
    }

    #[test]
    fn round_trips_through_builder_and_mutable_packages() {
        let table = parse_data_pilot_tables(&complete_xml()).unwrap().remove(0);
        let mut builder = crate::Builder::new();
        builder.add_sheet("Source").unwrap();
        builder.add_data_pilot_table(table).unwrap();
        let spreadsheet = crate::Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(spreadsheet.data_pilot_tables().len(), 1);

        let mut mutable = crate::MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
        mutable.data_pilot_tables_mut()[0].name = "Updated".to_string();
        let reparsed = crate::Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(reparsed.data_pilot_tables()[0].name, "Updated");
    }
}
