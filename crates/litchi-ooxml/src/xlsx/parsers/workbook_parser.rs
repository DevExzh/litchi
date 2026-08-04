//! Migration adapter for the canonical `litchi-xlsx` workbook catalog parser.

use litchi_core::sheet::Result as SheetResult;

use crate::xlsx::worksheet::Info;
use crate::xlsx::writer::NamedRange;

pub(crate) use litchi_xlsx::raw::PivotCache as PivotCacheInfo;

pub(crate) struct WorkbookParseResult {
    pub(crate) sheets: Vec<Info>,
    pub(crate) active_sheet_index: usize,
    pub(crate) uses_1904_date_system: bool,
    pub(crate) defined_names: Vec<NamedRange>,
    pub(crate) pivot_caches: Vec<PivotCacheInfo>,
    pub(crate) external_reference_ids: Vec<String>,
}

pub(crate) fn parse_workbook_details(content: &str) -> SheetResult<WorkbookParseResult> {
    litchi_xlsx::raw::parse_catalog(content.as_bytes())
        .map(|catalog| WorkbookParseResult {
            sheets: catalog.sheets.into_iter().map(adapt_sheet).collect(),
            active_sheet_index: catalog.active_sheet_index,
            uses_1904_date_system: catalog.uses_1904_date_system,
            defined_names: catalog
                .defined_names
                .into_iter()
                .map(adapt_defined_name)
                .collect(),
            pivot_caches: catalog.pivot_caches,
            external_reference_ids: catalog.external_reference_ids,
        })
        .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)
}

/// Parse workbook metadata, returning sheets, the active sheet index, and the date system.
pub fn parse_workbook_xml(content: &str) -> SheetResult<(Vec<Info>, usize, bool)> {
    parse_workbook_details(content).map(|details| {
        (
            details.sheets,
            details.active_sheet_index,
            details.uses_1904_date_system,
        )
    })
}

/// Parse a standalone `sheet` element.
pub fn parse_sheet_xml(sheet_xml: &str) -> SheetResult<Option<Info>> {
    litchi_xlsx::raw::parse_sheet(sheet_xml)
        .map(|sheet| sheet.map(adapt_sheet))
        .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)
}

fn adapt_sheet(sheet: litchi_xlsx::raw::Sheet) -> Info {
    Info {
        name: sheet.name,
        relationship_id: sheet.relationship_id,
        sheet_id: sheet.sheet_id,
        is_active: false,
        print_area: None,
        repeating_rows: None,
        repeating_columns: None,
    }
}

fn adapt_defined_name(name: litchi_xlsx::raw::DefinedName) -> NamedRange {
    NamedRange {
        name: name.name,
        reference: name.reference,
        comment: name.comment,
        local_sheet_id: name.local_sheet_id,
        custom_menu: name.custom_menu,
        description: name.description,
        help: name.help,
        status_bar: name.status_bar,
        shortcut_key: name.shortcut_key,
        hidden: name.hidden,
        function: name.function,
        vb_procedure: name.vb_procedure,
        xlm: name.xlm,
        function_group_id: name.function_group_id,
        publish_to_server: name.publish_to_server,
        workbook_parameter: name.workbook_parameter,
    }
}
