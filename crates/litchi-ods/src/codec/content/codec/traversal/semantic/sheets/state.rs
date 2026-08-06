//! Mutable state owned by the ODS sheet streaming traversal.

use super::super::super::{
    BTreeMap, Builder, CellBuilder, CellDetective, CellTextContentBuilder, RowBuilder, SheetBuilder,
};
use super::super::super::model::{
    PendingCalcextRule, PendingConditionalFormat, PendingHyperlink,
    PendingSparklineComplexColor, PendingSparklineGroup, SheetTextField,
};

pub(super) struct TraversalState {
    pub(super) current_sheet: Option<SheetBuilder>,
    pub(super) current_row: Option<RowBuilder>,
    pub(super) current_cell: Option<CellBuilder>,
    pub(super) text_element_depth: usize,
    pub(super) text_content: String,
    pub(super) rich_text_builder: Option<CellTextContentBuilder>,
    pub(super) annotation_builder: Option<Builder>,
    pub(super) annotation_depth: usize,
    pub(super) pending_hyperlink: Option<PendingHyperlink>,
    pub(super) detective_builder: Option<CellDetective>,
    pub(super) detective_child_open: bool,
    pub(super) sheet_text_field: Option<SheetTextField>,
    pub(super) sheet_text: String,
    pub(super) document_namespaces: BTreeMap<String, String>,
    pub(super) namespace_scopes: Vec<Vec<(String, Option<String>)>>,
    pub(super) element_depth: usize,
    pub(super) spreadsheet_depth: Option<usize>,
    pub(super) current_sheet_depth: Option<usize>,
    pub(super) sheet_dde_source_depth: Option<usize>,
    pub(super) conditional_formats_depth: Option<usize>,
    pub(super) pending_conditional_format: Option<PendingConditionalFormat>,
    pub(super) pending_calcext_rule: Option<PendingCalcextRule>,
    pub(super) calcext_leaf_open_depth: Option<usize>,
    pub(super) calcext_skip_depth: Option<usize>,
    pub(super) sparkline_groups_depth: Option<usize>,
    pub(super) pending_sparkline_group: Option<PendingSparklineGroup>,
    pub(super) pending_sparkline_complex_color: Option<PendingSparklineComplexColor>,
    pub(super) sparkline_list_depth: Option<usize>,
}

impl TraversalState {
    pub(super) fn new() -> Self {
        Self {
            current_sheet: None,
            current_row: None,
            current_cell: None,
            text_element_depth: 0,
            text_content: String::new(),
            rich_text_builder: None,
            annotation_builder: None,
            annotation_depth: 0,
            pending_hyperlink: None,
            detective_builder: None,
            detective_child_open: false,
            sheet_text_field: None,
            sheet_text: String::new(),
            document_namespaces: BTreeMap::new(),
            namespace_scopes: Vec::new(),
            element_depth: 0,
            spreadsheet_depth: None,
            current_sheet_depth: None,
            sheet_dde_source_depth: None,
            conditional_formats_depth: None,
            pending_conditional_format: None,
            pending_calcext_rule: None,
            calcext_leaf_open_depth: None,
            calcext_skip_depth: None,
            sparkline_groups_depth: None,
            pending_sparkline_group: None,
            pending_sparkline_complex_color: None,
            sparkline_list_depth: None,
        }
    }
}
