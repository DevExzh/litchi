//! ODS-specific parsing utilities.

use super::{
    Cell, CellDetective, CellHyperlink, CellMatrixSpan, CellMerge, CellRangeSource,
    CellTextContent, CellValue, Column, ConditionalColorScale, ConditionalColorScaleEntry,
    ConditionalDataBar, ConditionalDataBarEntry, ConditionalDateIs, ConditionalDateType,
    ConditionalFormat, ConditionalFormatCondition, ConditionalFormatEntryType,
    ConditionalFormatRule, ConditionalIconSet, ConditionalIconSetEntry, DataBarAxisPosition,
    DetectiveDirection, DetectiveHighlightedRange, DetectiveOperation, DetectiveOperationKind,
    IconSetType, NamedDefinition, NamedDefinitionScope, NamedExpression, NamedRange,
    NamedRangeUsage, Row, Sheet, SheetPrintSettings, SheetScenario, SheetStyle, SheetTableSource,
    Sparkline, SparklineAxisType, SparklineEmptyCells, SparklineGroup, SparklineType, TableGroup, TableRange, TableSourceMode, TableStructure, TableVisibility,
    annotation::{AnnotationBuilder, decode_reference},
    conditional_format::{
        CALCEXT_NAMESPACE_URI, DATA_BAR_ENTRY_COUNT, MAX_CONDITIONAL_FORMATS_PER_SHEET,
        MAX_ENTRIES_PER_RULE, MAX_RULES_PER_FORMAT, validate_color_scale_entry,
        validate_conditional_format, validate_condition, validate_data_bar_attributes,
        validate_data_bar_entry, validate_date_is, validate_icon_set_entry, validate_rule,
    },
    dde::parse_source as parse_dde_source,
    rich_text::CellTextContentBuilder,
    scenario::validate_scenario,
    source::validate_table_source,
    sparkline::{
        MAX_SPARKLINE_GROUPS_PER_SHEET, MAX_SPARKLINES_PER_GROUP, validate_sparkline,
        validate_sparkline_group, validate_sparkline_group_attributes,
    },
    structure::{
        MAX_EXPANDED_COLUMNS_PER_SHEET, MAX_EXPANDED_ROWS_PER_SHEET, MAX_TABLE_STRUCTURE_DEPTH,
        split_cell_range_addresses,
    },
};
use crate::elements::text::{TextHyperlinkActuate, TextHyperlinkShow};
use litchi_core::{Error, Result};
use quick_xml::Reader;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, NamespaceResolver, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::BTreeMap;
use std::num::NonZeroUsize;

const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TABLE_NAMESPACE_URI: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK_NAMESPACE: &str = "http://www.w3.org/1999/xlink";
const MAX_EXPANDED_CELLS_PER_ROW: usize = 1_048_576;
const MAX_EXPANDED_CELLS_PER_SHEET: usize = 4_194_304;
/// Interleaved runs of empty rows kept unmaterialised before the parser gives
/// up and expands them, so deferral cannot grow without bound.
const MAX_DEFERRED_BLANK_ROW_RUNS: usize = 4_096;
/// Longest run of cell-less rows still kept at the end of a table. Anything
/// longer is the full-height grid padding every ODF producer writes.
const MAX_TRAILING_EMPTY_ROWS: usize = 4_096;

/// A `text:a` hyperlink whose text content is still being collected.
struct PendingHyperlink {
    /// The hyperlink parsed from the element's attributes.
    link: CellHyperlink,
    /// Byte offset into the cell text where the link text begins.
    text_start: usize,
    /// The `text_element_depth` value assigned to the `text:a` element.
    depth: usize,
}

/// A `calcext:conditional-format` element whose rules are still being read.
struct PendingConditionalFormat {
    /// Target ranges parsed from `calcext:target-range-address`.
    target_range_addresses: Vec<String>,
    /// Inert rules collected so far, in document order.
    rules: Vec<ConditionalFormatRule>,
    /// The `element_depth` value assigned to the element.
    depth: usize,
}

/// A `calcext:color-scale`, `calcext:data-bar`, or `calcext:icon-set` rule
/// whose threshold entries are still being read.
enum PendingCalcextRule {
    ColorScale {
        entries: Vec<ConditionalColorScaleEntry>,
        depth: usize,
    },
    DataBar {
        data_bar: ConditionalDataBar,
        depth: usize,
    },
    IconSet {
        icon_set: ConditionalIconSet,
        depth: usize,
    },
}

impl PendingCalcextRule {
    fn depth(&self) -> usize {
        match self {
            Self::ColorScale { depth, .. } | Self::DataBar { depth, .. } | Self::IconSet { depth, .. } => *depth,
        }
    }

    fn element_name(&self) -> &'static str {
        match self {
            Self::ColorScale { .. } => "color-scale",
            Self::DataBar { .. } => "data-bar",
            Self::IconSet { .. } => "icon-set",
        }
    }

    fn finish(self) -> ConditionalFormatRule {
        match self {
            Self::ColorScale { entries, .. } => ConditionalColorScale::new(entries).into(),
            Self::DataBar { data_bar, .. } => data_bar.into(),
            Self::IconSet { icon_set, .. } => icon_set.into(),
        }
    }
}

/// A `calcext:sparkline-group` element whose sparklines are still being read.
struct PendingSparklineGroup {
    /// The group parsed from the element's attributes (with no sparklines yet).
    group: SparklineGroup,
    /// The `element_depth` value assigned to the element.
    depth: usize,
}

/// Parser for ODS-specific structures.
///
/// This provides parsing logic specific to spreadsheets,
/// including sheet, row, and cell parsing with proper type detection.
pub(crate) struct OdsParser;

#[derive(Clone, Copy)]
enum SheetTextField {
    Title,
    Description,
}

impl SheetTextField {
    fn local_name(self) -> &'static str {
        match self {
            Self::Title => "title",
            Self::Description => "desc",
        }
    }
}

impl OdsParser {
    /// Parse all sheets from ODS content.xml
    pub fn parse_sheets(xml_content: &str) -> Result<Vec<Sheet>> {
        let mut reader = Reader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut sheets = Vec::new();

        // Parser state
        let mut current_sheet: Option<SheetBuilder> = None;
        let mut current_row: Option<RowBuilder> = None;
        let mut current_cell: Option<CellBuilder> = None;
        let mut text_element_depth = 0usize;
        let mut text_content = String::new();
        let mut rich_text_builder: Option<CellTextContentBuilder> = None;
        let mut annotation_builder: Option<AnnotationBuilder> = None;
        let mut annotation_depth = 0usize;
        let mut pending_hyperlink: Option<PendingHyperlink> = None;
        let mut detective_builder: Option<CellDetective> = None;
        let mut detective_child_open = false;
        let mut sheet_text_field = None;
        let mut sheet_text = String::new();
        let mut document_namespaces = BTreeMap::new();
        let mut namespace_scopes = Vec::new();
        let mut element_depth = 0usize;
        let mut spreadsheet_depth = None;
        let mut current_sheet_depth = None;
        let mut sheet_dde_source_depth = None;
        let mut conditional_formats_depth = None;
        let mut pending_conditional_format: Option<PendingConditionalFormat> = None;
        let mut pending_calcext_rule: Option<PendingCalcextRule> = None;
        let mut calcext_leaf_open_depth = None;
        let mut calcext_skip_depth = None;
        let mut sparkline_groups_depth = None;
        let mut pending_sparkline_group: Option<PendingSparklineGroup> = None;
        let mut sparkline_list_depth = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let namespace_scope =
                        Self::push_namespace_scope(e, reader.decoder(), &mut document_namespaces)?;
                    namespace_scopes.push(namespace_scope);
                    element_depth += 1;
                    if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        OFFICE_NAMESPACE,
                        "spreadsheet",
                    ) {
                        spreadsheet_depth = Some(element_depth);
                    }

                    if sheet_dde_source_depth.is_some() {
                        return Err(Error::InvalidFormat(
                            "office:dde-source must not contain child elements".to_string(),
                        ));
                    } else if let Some(sheet) = current_sheet.as_mut()
                        && current_row.is_none()
                        && current_sheet_depth.is_some_and(|depth| element_depth == depth + 1)
                        && e.local_name().as_ref() == b"dde-source"
                    {
                        if !Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            OFFICE_NAMESPACE,
                            "dde-source",
                        ) {
                            return Err(Error::InvalidFormat(
                                "spoofed office:dde-source namespace".to_string(),
                            ));
                        }
                        let source = parse_dde_source(e, reader.decoder(), &document_namespaces)?;
                        sheet.set_dde_source(source)?;
                        sheet_dde_source_depth = Some(element_depth);
                    } else if calcext_skip_depth.is_some() {
                        // Unknown extension content inside
                        // `calcext:conditional-formats` is skipped entirely.
                    } else if calcext_leaf_open_depth.is_some() {
                        return Err(Error::InvalidFormat(
                            "calcext rule and entry elements must not contain child elements"
                                .to_string(),
                        ));
                    } else if let Some(rule) = pending_calcext_rule.as_mut() {
                        let is_calcext = |local_name: &str| {
                            Self::element_name_is(
                                e.name().as_ref(),
                                &document_namespaces,
                                CALCEXT_NAMESPACE_URI,
                                local_name,
                            )
                        };
                        match rule {
                            PendingCalcextRule::ColorScale { entries, .. }
                                if is_calcext("color-scale-entry") =>
                            {
                                if entries.len() >= MAX_ENTRIES_PER_RULE {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:color-scale exceeds the {MAX_ENTRIES_PER_RULE} entry safety limit"
                                    )));
                                }
                                entries.push(Self::parse_color_scale_entry(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                                calcext_leaf_open_depth = Some(element_depth);
                            },
                            PendingCalcextRule::DataBar { data_bar, .. }
                                if is_calcext("formatting-entry") || is_calcext("data-bar-entry") =>
                            {
                                if data_bar.entries.len() >= DATA_BAR_ENTRY_COUNT {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:data-bar supports exactly {DATA_BAR_ENTRY_COUNT} calcext:formatting-entry elements"
                                    )));
                                }
                                data_bar.entries.push(Self::parse_data_bar_entry(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                                calcext_leaf_open_depth = Some(element_depth);
                            },
                            PendingCalcextRule::IconSet { icon_set, .. }
                                if is_calcext("formatting-entry") =>
                            {
                                if icon_set.entries.len() >= MAX_ENTRIES_PER_RULE {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:icon-set exceeds the {MAX_ENTRIES_PER_RULE} entry safety limit"
                                    )));
                                }
                                icon_set.entries.push(Self::parse_icon_set_entry(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                                calcext_leaf_open_depth = Some(element_depth);
                            },
                            _ => {
                                // Unmodeled content such as
                                // `calcext:custom-iconset` is skipped.
                                calcext_skip_depth = Some(element_depth);
                            },
                        }
                    } else if let Some(pending) = pending_conditional_format.as_mut() {
                        let is_calcext = |local_name: &str| {
                            Self::element_name_is(
                                e.name().as_ref(),
                                &document_namespaces,
                                CALCEXT_NAMESPACE_URI,
                                local_name,
                            )
                        };
                        if is_calcext("condition") {
                            if pending.rules.len() >= MAX_RULES_PER_FORMAT {
                                return Err(Error::InvalidFormat(format!(
                                    "conditional format exceeds the {MAX_RULES_PER_FORMAT} rule safety limit"
                                )));
                            }
                            let condition = Self::parse_calcext_condition(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            pending.rules.push(condition.into());
                            calcext_leaf_open_depth = Some(element_depth);
                        } else if is_calcext("color-scale") {
                            pending_calcext_rule = Some(PendingCalcextRule::ColorScale {
                                entries: Vec::new(),
                                depth: element_depth,
                            });
                        } else if is_calcext("data-bar") {
                            pending_calcext_rule = Some(PendingCalcextRule::DataBar {
                                data_bar: Self::parse_data_bar_attributes(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?,
                                depth: element_depth,
                            });
                        } else if is_calcext("icon-set") {
                            pending_calcext_rule = Some(PendingCalcextRule::IconSet {
                                icon_set: Self::parse_icon_set_attributes(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?,
                                depth: element_depth,
                            });
                        } else if is_calcext("date-is") {
                            if pending.rules.len() >= MAX_RULES_PER_FORMAT {
                                return Err(Error::InvalidFormat(format!(
                                    "conditional format exceeds the {MAX_RULES_PER_FORMAT} rule safety limit"
                                )));
                            }
                            let date_is = Self::parse_date_is(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            pending.rules.push(date_is.into());
                            calcext_leaf_open_depth = Some(element_depth);
                        } else {
                            // Unmodeled rule content is skipped.
                            calcext_skip_depth = Some(element_depth);
                        }
                    } else if conditional_formats_depth.is_some() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "conditional-format",
                        ) {
                            let target_range_addresses = Self::parse_conditional_format_ranges(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            pending_conditional_format = Some(PendingConditionalFormat {
                                target_range_addresses,
                                rules: Vec::new(),
                                depth: element_depth,
                            });
                        } else {
                            calcext_skip_depth = Some(element_depth);
                        }
                    } else if sparkline_list_depth.is_some() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline",
                        ) {
                            let sparkline =
                                Self::parse_sparkline(e, reader.decoder(), &document_namespaces)?;
                            let pending = pending_sparkline_group.as_mut().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "calcext:sparklines must be inside calcext:sparkline-group"
                                        .to_string(),
                                )
                            })?;
                            if pending.group.sparklines.len() >= MAX_SPARKLINES_PER_GROUP {
                                return Err(Error::InvalidFormat(format!(
                                    "sparkline group exceeds the {MAX_SPARKLINES_PER_GROUP} sparkline safety limit"
                                )));
                            }
                            pending.group.sparklines.push(sparkline);
                            calcext_leaf_open_depth = Some(element_depth);
                        } else {
                            calcext_skip_depth = Some(element_depth);
                        }
                    } else if pending_sparkline_group.is_some() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparklines",
                        ) {
                            sparkline_list_depth = Some(element_depth);
                        } else {
                            // `calcext:sparkline-*-complex-color` theme colors
                            // and other unmodeled children are skipped.
                            calcext_skip_depth = Some(element_depth);
                        }
                    } else if sparkline_groups_depth.is_some() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline-group",
                        ) {
                            let group = Self::parse_sparkline_group_attributes(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            pending_sparkline_group = Some(PendingSparklineGroup {
                                group,
                                depth: element_depth,
                            });
                        } else {
                            calcext_skip_depth = Some(element_depth);
                        }
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && current_sheet_depth.is_some_and(|depth| element_depth == depth + 1)
                        && e.local_name().as_ref() == b"sparkline-groups"
                    {
                        if !Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline-groups",
                        ) {
                            return Err(Error::InvalidFormat(
                                "spoofed calcext:sparkline-groups namespace".to_string(),
                            ));
                        }
                        sparkline_groups_depth = Some(element_depth);
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && current_sheet_depth.is_some_and(|depth| element_depth == depth + 1)
                        && e.local_name().as_ref() == b"conditional-formats"
                    {
                        if !Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "conditional-formats",
                        ) {
                            return Err(Error::InvalidFormat(
                                "spoofed calcext:conditional-formats namespace".to_string(),
                            ));
                        }
                        conditional_formats_depth = Some(element_depth);
                    } else if let Some(builder) = detective_builder.as_mut() {
                        if detective_child_open {
                            return Err(Error::InvalidFormat(
                                "table:detective child elements must be empty".to_string(),
                            ));
                        }
                        Self::parse_detective_child(
                            builder,
                            e,
                            reader.decoder(),
                            &document_namespaces,
                        )?;
                        detective_child_open = true;
                    } else if let Some(builder) = annotation_builder.as_mut() {
                        builder.start(e, reader.decoder())?;
                        annotation_depth += 1;
                    } else if current_cell.is_some()
                        && Self::is_office_annotation(e, &document_namespaces)
                    {
                        annotation_builder = Some(AnnotationBuilder::new(
                            e,
                            reader.decoder(),
                            document_namespaces.clone(),
                        )?);
                    } else if text_element_depth > 0 {
                        rich_text_builder
                            .as_mut()
                            .expect("rich-text builder exists inside text:p")
                            .start(e, reader.decoder())?;
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TEXT_NAMESPACE,
                            "a",
                        ) {
                            if pending_hyperlink.is_some() {
                                return Err(Error::InvalidFormat(
                                    "text:a hyperlinks must not be nested".to_string(),
                                ));
                            }
                            pending_hyperlink = Some(PendingHyperlink {
                                link: Self::parse_hyperlink(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?,
                                text_start: text_content.len(),
                                depth: text_element_depth + 1,
                            });
                        }
                        text_element_depth += 1;
                    } else if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "table",
                    ) && spreadsheet_depth.is_some_and(|depth| element_depth == depth + 1)
                    {
                        let name =
                            Self::extract_table_name(e, reader.decoder(), &document_namespaces)?;
                        let (style, print_settings) = Self::parse_sheet_formatting(
                            e,
                            reader.decoder(),
                            &document_namespaces,
                        )?;
                        current_sheet =
                            Some(SheetBuilder::with_formatting(name, style, print_settings));
                        current_sheet_depth = Some(element_depth);
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-source",
                        )
                    {
                        let source =
                            Self::parse_table_source(e, reader.decoder(), &document_namespaces)?;
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.set_table_source(source)?;
                        }
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "title",
                        )
                    {
                        if sheet_text_field.replace(SheetTextField::Title).is_some() {
                            return Err(Error::InvalidFormat(
                                "table title or description elements must not be nested"
                                    .to_string(),
                            ));
                        }
                        sheet_text.clear();
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "desc",
                        )
                    {
                        if sheet_text_field
                            .replace(SheetTextField::Description)
                            .is_some()
                        {
                            return Err(Error::InvalidFormat(
                                "table title or description elements must not be nested"
                                    .to_string(),
                            ));
                        }
                        sheet_text.clear();
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "scenario",
                        )
                    {
                        let scenario =
                            Self::parse_scenario(e, reader.decoder(), &document_namespaces)?;
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.set_scenario(scenario)?;
                        }
                    } else if current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-column-group",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            let display = Self::parse_group_display(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            sheet.column_structure.begin_group(display)?;
                        }
                    } else if current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-header-columns",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.column_structure.begin_header(sheet.columns.len())?;
                        }
                    } else if current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-row-group",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            let display = Self::parse_group_display(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            sheet.begin_row_group(display)?;
                        }
                    } else if current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-header-rows",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.begin_row_header()?;
                        }
                    } else if current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-column",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            let (column, repeated) =
                                Self::parse_column(e, reader.decoder(), &document_namespaces)?;
                            sheet.add_repeated_column(column, repeated)?;
                        }
                    } else if current_sheet.is_some()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-row",
                        )
                    {
                        current_row = Some(RowBuilder::from_element(
                            e,
                            reader.decoder(),
                            &document_namespaces,
                        )?);
                    } else if current_row.is_some()
                        && (Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-cell",
                        ) || Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "covered-table-cell",
                        ))
                    {
                        let covered = Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "covered-table-cell",
                        );
                        let mut cell_builder =
                            Self::parse_cell_attributes(e, reader.decoder(), &document_namespaces)?;
                        if covered {
                            cell_builder.merge = CellMerge::Covered;
                        }
                        current_cell = Some(cell_builder);
                        text_content.clear();
                        rich_text_builder = None;
                        pending_hyperlink = None;
                        text_element_depth = 0;
                    } else if let Some(cell) = current_cell.as_mut()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "cell-range-source",
                        )
                    {
                        let source = Self::parse_cell_range_source(
                            e,
                            reader.decoder(),
                            &document_namespaces,
                        )?;
                        if cell.range_source.replace(source).is_some() {
                            return Err(Error::InvalidFormat(
                                "table cell contains multiple table:cell-range-source elements"
                                    .to_string(),
                            ));
                        }
                    } else if let Some(cell) = current_cell.as_mut()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "detective",
                        )
                    {
                        if cell.detective.is_some() {
                            return Err(Error::InvalidFormat(
                                "table cell contains multiple table:detective elements".to_string(),
                            ));
                        }
                        detective_builder = Some(CellDetective::new());
                    } else if current_cell.is_some()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TEXT_NAMESPACE,
                            "p",
                        )
                    {
                        if !text_content.is_empty() {
                            text_content.push('\n');
                        }
                        let builder = rich_text_builder.get_or_insert_with(|| {
                            CellTextContentBuilder::new(document_namespaces.clone())
                        });
                        builder.start(e, reader.decoder())?;
                        text_element_depth = 1;
                    }
                },
                Ok(Event::Empty(ref e)) => {
                    let empty_scope =
                        Self::push_namespace_scope(e, reader.decoder(), &mut document_namespaces)?;
                    if sheet_dde_source_depth.is_some() {
                        return Err(Error::InvalidFormat(
                            "office:dde-source must not contain child elements".to_string(),
                        ));
                    }
                    if let Some(sheet) = current_sheet.as_mut()
                        && current_row.is_none()
                        && current_sheet_depth.is_some_and(|depth| element_depth == depth)
                        && e.local_name().as_ref() == b"dde-source"
                    {
                        if !Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            OFFICE_NAMESPACE,
                            "dde-source",
                        ) {
                            return Err(Error::InvalidFormat(
                                "spoofed office:dde-source namespace".to_string(),
                            ));
                        }
                        let source = parse_dde_source(e, reader.decoder(), &document_namespaces)?;
                        sheet.set_dde_source(source)?;
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if calcext_skip_depth.is_some() {
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if calcext_leaf_open_depth.is_some() {
                        return Err(Error::InvalidFormat(
                            "calcext rule and entry elements must not contain child elements"
                                .to_string(),
                        ));
                    }
                    if let Some(rule) = pending_calcext_rule.as_mut() {
                        let is_calcext = |local_name: &str| {
                            Self::element_name_is(
                                e.name().as_ref(),
                                &document_namespaces,
                                CALCEXT_NAMESPACE_URI,
                                local_name,
                            )
                        };
                        match rule {
                            PendingCalcextRule::ColorScale { entries, .. }
                                if is_calcext("color-scale-entry") =>
                            {
                                if entries.len() >= MAX_ENTRIES_PER_RULE {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:color-scale exceeds the {MAX_ENTRIES_PER_RULE} entry safety limit"
                                    )));
                                }
                                entries.push(Self::parse_color_scale_entry(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                            },
                            PendingCalcextRule::DataBar { data_bar, .. }
                                if is_calcext("formatting-entry") || is_calcext("data-bar-entry") =>
                            {
                                if data_bar.entries.len() >= DATA_BAR_ENTRY_COUNT {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:data-bar supports exactly {DATA_BAR_ENTRY_COUNT} calcext:formatting-entry elements"
                                    )));
                                }
                                data_bar.entries.push(Self::parse_data_bar_entry(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                            },
                            PendingCalcextRule::IconSet { icon_set, .. }
                                if is_calcext("formatting-entry") =>
                            {
                                if icon_set.entries.len() >= MAX_ENTRIES_PER_RULE {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:icon-set exceeds the {MAX_ENTRIES_PER_RULE} entry safety limit"
                                    )));
                                }
                                icon_set.entries.push(Self::parse_icon_set_entry(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                            },
                            _ => {},
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if let Some(pending) = pending_conditional_format.as_mut() {
                        let is_calcext = |local_name: &str| {
                            Self::element_name_is(
                                e.name().as_ref(),
                                &document_namespaces,
                                CALCEXT_NAMESPACE_URI,
                                local_name,
                            )
                        };
                        if is_calcext("condition") {
                            if pending.rules.len() >= MAX_RULES_PER_FORMAT {
                                return Err(Error::InvalidFormat(format!(
                                    "conditional format exceeds the {MAX_RULES_PER_FORMAT} rule safety limit"
                                )));
                            }
                            let condition = Self::parse_calcext_condition(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            pending.rules.push(condition.into());
                        } else if is_calcext("date-is") {
                            if pending.rules.len() >= MAX_RULES_PER_FORMAT {
                                return Err(Error::InvalidFormat(format!(
                                    "conditional format exceeds the {MAX_RULES_PER_FORMAT} rule safety limit"
                                )));
                            }
                            let date_is = Self::parse_date_is(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            pending.rules.push(date_is.into());
                        } else if is_calcext("color-scale") {
                            let rule: ConditionalFormatRule =
                                ConditionalColorScale::new(Vec::new()).into();
                            validate_rule(&rule)?;
                        } else if is_calcext("data-bar") {
                            let data_bar = Self::parse_data_bar_attributes(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            let rule: ConditionalFormatRule = data_bar.into();
                            validate_rule(&rule)?;
                        } else if is_calcext("icon-set") {
                            let icon_set = Self::parse_icon_set_attributes(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            let rule: ConditionalFormatRule = icon_set.into();
                            validate_rule(&rule)?;
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if conditional_formats_depth.is_some() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "conditional-format",
                        ) {
                            let format = ConditionalFormat {
                                target_range_addresses: Self::parse_conditional_format_ranges(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?,
                                rules: Vec::new(),
                            };
                            validate_conditional_format(&format)?;
                            if let Some(sheet) = current_sheet.as_mut() {
                                sheet.add_conditional_format(format)?;
                            }
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if sparkline_list_depth.is_some() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline",
                        ) {
                            let sparkline =
                                Self::parse_sparkline(e, reader.decoder(), &document_namespaces)?;
                            let pending = pending_sparkline_group.as_mut().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "calcext:sparklines must be inside calcext:sparkline-group"
                                        .to_string(),
                                )
                            })?;
                            if pending.group.sparklines.len() >= MAX_SPARKLINES_PER_GROUP {
                                return Err(Error::InvalidFormat(format!(
                                    "sparkline group exceeds the {MAX_SPARKLINES_PER_GROUP} sparkline safety limit"
                                )));
                            }
                            pending.group.sparklines.push(sparkline);
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if pending_sparkline_group.is_some() {
                        // Empty `calcext:sparklines` containers and unmodeled
                        // complex-color children carry no sparklines.
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if sparkline_groups_depth.is_some() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline-group",
                        ) {
                            let group = Self::parse_sparkline_group_attributes(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            // An empty group has no sparklines and is invalid.
                            validate_sparkline_group(&group)?;
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if current_sheet.is_some()
                        && current_row.is_none()
                        && current_sheet_depth.is_some_and(|depth| element_depth == depth)
                        && e.local_name().as_ref() == b"sparkline-groups"
                    {
                        if !Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline-groups",
                        ) {
                            return Err(Error::InvalidFormat(
                                "spoofed calcext:sparkline-groups namespace".to_string(),
                            ));
                        }
                        // An empty container declares no sparkline groups.
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if current_sheet.is_some()
                        && current_row.is_none()
                        && current_sheet_depth.is_some_and(|depth| element_depth == depth)
                        && e.local_name().as_ref() == b"conditional-formats"
                    {
                        if !Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "conditional-formats",
                        ) {
                            return Err(Error::InvalidFormat(
                                "spoofed calcext:conditional-formats namespace".to_string(),
                            ));
                        }
                        // An empty container declares no conditional formats.
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if spreadsheet_depth == Some(element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table",
                        )
                    {
                        let name =
                            Self::extract_table_name(e, reader.decoder(), &document_namespaces)?;
                        let (style, print_settings) = Self::parse_sheet_formatting(
                            e,
                            reader.decoder(),
                            &document_namespaces,
                        )?;
                        sheets.push(
                            SheetBuilder::with_formatting(name, style, print_settings).build()?,
                        );
                    } else if let Some(builder) = detective_builder.as_mut() {
                        if detective_child_open {
                            return Err(Error::InvalidFormat(
                                "table:detective child elements must be empty".to_string(),
                            ));
                        }
                        Self::parse_detective_child(
                            builder,
                            e,
                            reader.decoder(),
                            &document_namespaces,
                        )?;
                    } else if let Some(builder) = annotation_builder.as_mut() {
                        builder.empty(e, reader.decoder())?;
                    } else if current_cell.is_some()
                        && Self::is_office_annotation(e, &document_namespaces)
                    {
                        let annotation = AnnotationBuilder::new(
                            e,
                            reader.decoder(),
                            document_namespaces.clone(),
                        )?
                        .finish()?;
                        if let Some(cell) = current_cell.as_mut() {
                            cell.annotation = Some(annotation);
                        }
                    } else if text_element_depth > 0 {
                        rich_text_builder
                            .as_mut()
                            .expect("rich-text builder exists inside text:p")
                            .empty(e, reader.decoder())?;
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TEXT_NAMESPACE,
                            "a",
                        ) {
                            if pending_hyperlink.is_some() {
                                return Err(Error::InvalidFormat(
                                    "text:a hyperlinks must not be nested".to_string(),
                                ));
                            }
                            let link =
                                Self::parse_hyperlink(e, reader.decoder(), &document_namespaces)?;
                            let mut link = link;
                            let position = text_content.len();
                            link.set_range(position..position);
                            if let Some(cell) = current_cell.as_mut() {
                                cell.hyperlinks.push(link);
                            }
                        } else {
                            Self::push_text_empty_element(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                                &mut text_content,
                            )?;
                        }
                    } else if current_cell.is_some()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TEXT_NAMESPACE,
                            "p",
                        )
                    {
                        if !text_content.is_empty() {
                            text_content.push('\n');
                        }
                        let builder = rich_text_builder.get_or_insert_with(|| {
                            CellTextContentBuilder::new(document_namespaces.clone())
                        });
                        builder.empty(e, reader.decoder())?;
                    } else if let Some(cell) = current_cell.as_mut()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "cell-range-source",
                        )
                    {
                        let source = Self::parse_cell_range_source(
                            e,
                            reader.decoder(),
                            &document_namespaces,
                        )?;
                        if cell.range_source.replace(source).is_some() {
                            return Err(Error::InvalidFormat(
                                "table cell contains multiple table:cell-range-source elements"
                                    .to_string(),
                            ));
                        }
                    } else if let Some(cell) = current_cell.as_mut()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "detective",
                        )
                    {
                        if cell.detective.replace(CellDetective::new()).is_some() {
                            return Err(Error::InvalidFormat(
                                "table cell contains multiple table:detective elements".to_string(),
                            ));
                        }
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "title",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.set_title(String::new())?;
                        }
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "desc",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.set_description(String::new())?;
                        }
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-source",
                        )
                    {
                        let source =
                            Self::parse_table_source(e, reader.decoder(), &document_namespaces)?;
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.set_table_source(source)?;
                        }
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "scenario",
                        )
                    {
                        let scenario =
                            Self::parse_scenario(e, reader.decoder(), &document_namespaces)?;
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.set_scenario(scenario)?;
                        }
                    } else if current_sheet.is_some()
                        && current_row.is_none()
                        && [
                            "table-column-group",
                            "table-header-columns",
                            "table-row-group",
                            "table-header-rows",
                        ]
                        .iter()
                        .any(|local_name| {
                            Self::element_name_is(
                                e.name().as_ref(),
                                &document_namespaces,
                                TABLE_NAMESPACE_URI,
                                local_name,
                            )
                        })
                    {
                        return Err(Error::InvalidFormat(
                            "empty table groups and header containers are not valid".to_string(),
                        ));
                    } else if current_row.is_none()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-column",
                        )
                    {
                        if let Some(sheet) = current_sheet.as_mut() {
                            let (column, repeated) =
                                Self::parse_column(e, reader.decoder(), &document_namespaces)?;
                            sheet.add_repeated_column(column, repeated)?;
                        }
                    } else if current_row.is_some()
                        && (Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-cell",
                        ) || Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "covered-table-cell",
                        ))
                    {
                        let covered = Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "covered-table-cell",
                        );
                        let mut cell_builder =
                            Self::parse_cell_attributes(e, reader.decoder(), &document_namespaces)?;
                        if covered {
                            cell_builder.merge = CellMerge::Covered;
                        }
                        if let Some(row) = current_row.as_mut() {
                            row.add_repeated_cells(&cell_builder, "", None)?;
                        }
                    } else if current_sheet.is_some()
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table-row",
                        )
                    {
                        let row_builder =
                            RowBuilder::from_element(e, reader.decoder(), &document_namespaces)?;
                        let repeated = row_builder.repeated;
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.add_repeated_row(row_builder.build(), repeated)?;
                        }
                    }
                    Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                },
                Ok(Event::Text(ref t)) if detective_builder.is_some() => {
                    let decoded = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid detective text: {error}"))
                    })?;
                    if !decoded.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "table:detective cannot contain text".to_string(),
                        ));
                    }
                },
                Ok(Event::CData(ref t)) if detective_builder.is_some() => {
                    let decoded = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid detective CDATA: {error}"))
                    })?;
                    if !decoded.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "table:detective cannot contain CDATA".to_string(),
                        ));
                    }
                },
                Ok(Event::GeneralRef(_)) if detective_builder.is_some() => {
                    return Err(Error::InvalidFormat(
                        "table:detective cannot contain entity references".to_string(),
                    ));
                },
                Ok(Event::Text(_)) | Ok(Event::CData(_)) | Ok(Event::GeneralRef(_))
                    if sheet_dde_source_depth.is_some() =>
                {
                    return Err(Error::InvalidFormat(
                        "office:dde-source must be empty".to_string(),
                    ));
                },
                Ok(Event::Text(ref t)) if annotation_builder.is_some() => {
                    if let Some(builder) = annotation_builder.as_mut() {
                        builder.text(t)?;
                    }
                },
                Ok(Event::CData(ref t)) if annotation_builder.is_some() => {
                    if let Some(builder) = annotation_builder.as_mut() {
                        builder.cdata(t)?;
                    }
                },
                Ok(Event::GeneralRef(ref reference)) if annotation_builder.is_some() => {
                    if let Some(builder) = annotation_builder.as_mut() {
                        builder.reference(reference)?;
                    }
                },
                Ok(Event::Text(ref t)) if sheet_text_field.is_some() => {
                    let decoded = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid table text: {error}"))
                    })?;
                    let decoded = quick_xml::escape::unescape(&decoded).map_err(|error| {
                        Error::InvalidFormat(format!("invalid table character reference: {error}"))
                    })?;
                    sheet_text.push_str(&decoded);
                },
                Ok(Event::CData(ref t)) if sheet_text_field.is_some() => {
                    let decoded = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid table CDATA: {error}"))
                    })?;
                    sheet_text.push_str(&decoded);
                },
                Ok(Event::GeneralRef(ref reference)) if sheet_text_field.is_some() => {
                    sheet_text.push_str(&decode_reference(reference)?);
                },
                Ok(Event::Text(ref t)) if text_element_depth > 0 && current_cell.is_some() => {
                    rich_text_builder
                        .as_mut()
                        .expect("rich-text builder exists inside text:p")
                        .text(t)?;
                    let decoded = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid cell text: {error}"))
                    })?;
                    let decoded = quick_xml::escape::unescape(&decoded).map_err(|error| {
                        Error::InvalidFormat(format!("invalid cell character reference: {error}"))
                    })?;
                    text_content.push_str(&decoded);
                },
                Ok(Event::CData(ref t)) if text_element_depth > 0 && current_cell.is_some() => {
                    rich_text_builder
                        .as_mut()
                        .expect("rich-text builder exists inside text:p")
                        .cdata(t)?;
                    let decoded = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid cell CDATA: {error}"))
                    })?;
                    text_content.push_str(&decoded);
                },
                Ok(Event::GeneralRef(ref reference))
                    if text_element_depth > 0 && current_cell.is_some() =>
                {
                    rich_text_builder
                        .as_mut()
                        .expect("rich-text builder exists inside text:p")
                        .reference(reference)?;
                    text_content.push_str(&decode_reference(reference)?);
                },
                Ok(Event::End(ref e)) => {
                    let closes_sheet_dde_source = sheet_dde_source_depth == Some(element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            OFFICE_NAMESPACE,
                            "dde-source",
                        );
                    let closes_current_sheet = current_sheet_depth == Some(element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "table",
                        );
                    let closes_spreadsheet = spreadsheet_depth == Some(element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            OFFICE_NAMESPACE,
                            "spreadsheet",
                        );
                    let closes_calcext_skip = calcext_skip_depth == Some(element_depth);
                    let closes_calcext_leaf = calcext_leaf_open_depth == Some(element_depth);
                    let closes_calcext_rule = pending_calcext_rule
                        .as_ref()
                        .is_some_and(|rule| rule.depth() == element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            pending_calcext_rule
                                .as_ref()
                                .expect("pending calcext rule was checked")
                                .element_name(),
                        );
                    let closes_conditional_format = pending_conditional_format
                        .as_ref()
                        .is_some_and(|pending| pending.depth == element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "conditional-format",
                        );
                    let closes_conditional_formats = conditional_formats_depth == Some(element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "conditional-formats",
                        );
                    let closes_sparkline_list = sparkline_list_depth == Some(element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparklines",
                        );
                    let closes_sparkline_group = pending_sparkline_group
                        .as_ref()
                        .is_some_and(|pending| pending.depth == element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline-group",
                        );
                    let closes_sparkline_groups = sparkline_groups_depth == Some(element_depth)
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            CALCEXT_NAMESPACE_URI,
                            "sparkline-groups",
                        );
                    element_depth = element_depth.saturating_sub(1);
                    if closes_sheet_dde_source {
                        sheet_dde_source_depth = None;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_calcext_skip {
                        calcext_skip_depth = None;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_calcext_leaf {
                        calcext_leaf_open_depth = None;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_calcext_rule {
                        let rule = pending_calcext_rule
                            .take()
                            .expect("pending calcext rule was checked")
                            .finish();
                        validate_rule(&rule)?;
                        let pending = pending_conditional_format.as_mut().ok_or_else(|| {
                            Error::InvalidFormat(
                                "calcext rule closed outside calcext:conditional-format"
                                    .to_string(),
                            )
                        })?;
                        if pending.rules.len() >= MAX_RULES_PER_FORMAT {
                            return Err(Error::InvalidFormat(format!(
                                "conditional format exceeds the {MAX_RULES_PER_FORMAT} rule safety limit"
                            )));
                        }
                        pending.rules.push(rule);
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_conditional_format {
                        let pending = pending_conditional_format
                            .take()
                            .expect("pending conditional format was checked");
                        let format = ConditionalFormat {
                            target_range_addresses: pending.target_range_addresses,
                            rules: pending.rules,
                        };
                        validate_conditional_format(&format)?;
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.add_conditional_format(format)?;
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_conditional_formats {
                        conditional_formats_depth = None;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_sparkline_list {
                        sparkline_list_depth = None;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_sparkline_group {
                        let pending = pending_sparkline_group
                            .take()
                            .expect("pending sparkline group was checked");
                        validate_sparkline_group(&pending.group)?;
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.add_sparkline_group(pending.group)?;
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if closes_sparkline_groups {
                        sparkline_groups_depth = None;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }
                    if detective_builder.is_some() {
                        if detective_child_open {
                            detective_child_open = false;
                        } else if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            "detective",
                        ) {
                            let detective = detective_builder
                                .take()
                                .expect("detective builder was checked");
                            let cell = current_cell.as_mut().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "table:detective must be contained in a table cell".to_string(),
                                )
                            })?;
                            if cell.detective.replace(detective).is_some() {
                                return Err(Error::InvalidFormat(
                                    "table cell contains multiple table:detective elements"
                                        .to_string(),
                                ));
                            }
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }

                    if annotation_builder.is_some() {
                        if annotation_depth == 0 {
                            let annotation = annotation_builder
                                .take()
                                .expect("annotation builder was checked")
                                .finish()?;
                            if let Some(cell) = current_cell.as_mut() {
                                cell.annotation = Some(annotation);
                            }
                        } else {
                            annotation_builder
                                .as_mut()
                                .expect("annotation builder was checked")
                                .end_element()?;
                            annotation_depth -= 1;
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }

                    if text_element_depth > 0 {
                        rich_text_builder
                            .as_mut()
                            .expect("rich-text builder exists inside text:p")
                            .end()?;
                        if pending_hyperlink
                            .as_ref()
                            .is_some_and(|pending| pending.depth == text_element_depth)
                        {
                            let pending = pending_hyperlink
                                .take()
                                .expect("pending hyperlink was checked");
                            let mut link = pending.link;
                            let range = pending.text_start..text_content.len();
                            link.text = text_content
                                .get(range.clone())
                                .unwrap_or_default()
                                .to_string();
                            link.set_range(range);
                            if let Some(cell) = current_cell.as_mut() {
                                cell.hyperlinks.push(link);
                            }
                        }
                        text_element_depth -= 1;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }

                    if let Some(field) = sheet_text_field
                        && Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            TABLE_NAMESPACE_URI,
                            field.local_name(),
                        )
                    {
                        let value = std::mem::take(&mut sheet_text);
                        if let Some(sheet) = current_sheet.as_mut() {
                            match field {
                                SheetTextField::Title => sheet.set_title(value)?,
                                SheetTextField::Description => sheet.set_description(value)?,
                            }
                        }
                        sheet_text_field = None;
                    } else if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "table-cell",
                    ) || Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "covered-table-cell",
                    ) {
                        if let Some(cell_builder) = current_cell.take()
                            && let Some(ref mut row_builder) = current_row
                        {
                            let rich_text = rich_text_builder
                                .take()
                                .map(CellTextContentBuilder::finish)
                                .transpose()?;
                            row_builder.add_repeated_cells(
                                &cell_builder,
                                &text_content,
                                rich_text.as_ref(),
                            )?;
                        }
                    } else if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "table-row",
                    ) {
                        if let Some(row_builder) = current_row.take() {
                            let repeated = row_builder.repeated;
                            let row = row_builder.build();
                            if let Some(ref mut sheet_builder) = current_sheet {
                                sheet_builder.add_repeated_row(row, repeated)?;
                            }
                        }
                    } else if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "table-column-group",
                    ) {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.column_structure.end_group()?;
                        }
                    } else if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "table-header-columns",
                    ) {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.column_structure.end_header(sheet.columns.len())?;
                        }
                    } else if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "table-row-group",
                    ) {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.end_row_group()?;
                        }
                    } else if Self::element_name_is(
                        e.name().as_ref(),
                        &document_namespaces,
                        TABLE_NAMESPACE_URI,
                        "table-header-rows",
                    ) {
                        if let Some(sheet) = current_sheet.as_mut() {
                            sheet.end_row_header()?;
                        }
                    } else if closes_current_sheet && let Some(sheet_builder) = current_sheet.take()
                    {
                        sheets.push(sheet_builder.build()?);
                        current_sheet_depth = None;
                    }
                    if closes_spreadsheet {
                        spreadsheet_depth = None;
                    }
                    Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(Error::InvalidFormat(format!("XML parsing error: {}", e)));
                },
                _ => {},
            }
            buf.clear();
        }

        if sheet_dde_source_depth.is_some() {
            return Err(Error::InvalidFormat(
                "unterminated office:dde-source".to_string(),
            ));
        }
        if conditional_formats_depth.is_some() {
            return Err(Error::InvalidFormat(
                "unterminated calcext:conditional-formats".to_string(),
            ));
        }
        if sparkline_groups_depth.is_some() {
            return Err(Error::InvalidFormat(
                "unterminated calcext:sparkline-groups".to_string(),
            ));
        }

        let images = crate::media::scan_content_images(xml_content)?;
        let mut sheet_indices = std::collections::HashMap::with_capacity(sheets.len());
        for (index, sheet) in sheets.iter().enumerate() {
            if sheet_indices.insert(sheet.name.clone(), index).is_some() {
                return Err(Error::InvalidFormat(format!(
                    "duplicate table name '{}' prevents sheet-image association",
                    sheet.name
                )));
            }
        }
        for image in images {
            let Some(frame) = image.frame.as_ref().filter(|frame| frame.sheet_shape) else {
                continue;
            };
            let sheet_name = frame.sheet_name.as_deref().ok_or_else(|| {
                Error::InvalidFormat("sheet image has no containing table name".to_string())
            })?;
            let index = *sheet_indices.get(sheet_name).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "sheet image references unknown table '{sheet_name}'"
                ))
            })?;
            super::sheet_image::validate_sheet_image(&image)?;
            sheets[index].images.push(image);
        }
        for sheet in &sheets {
            super::sheet_image::validate_sheet_images(&sheet.images)?;
        }

        let shape_tables = crate::odp::OdpParser::parse_sheet_shape_tables(xml_content)?;
        if shape_tables.len() != sheets.len() {
            return Err(Error::InvalidFormat(format!(
                "spreadsheet table structure changed during shape parsing: {} shape container(s) for {} table(s)",
                shape_tables.len(),
                sheets.len()
            )));
        }
        for (sheet, shapes) in sheets.iter_mut().zip(shape_tables) {
            for shape in shapes {
                if let Some(sheet_shape) = super::shape::sheet_shape_from_parsed(shape)? {
                    sheet.shapes.push(sheet_shape);
                }
            }
            super::shape::validate_sheet_shapes(&sheet.shapes)?;
        }
        Ok(sheets)
    }

    fn is_office_annotation(
        element: &BytesStart<'_>,
        namespaces: &BTreeMap<String, String>,
    ) -> bool {
        Self::element_name_is(
            element.name().as_ref(),
            namespaces,
            OFFICE_NAMESPACE,
            "annotation",
        )
    }

    fn element_name_is(
        qualified_name: &[u8],
        namespaces: &BTreeMap<String, String>,
        namespace: &str,
        local_name: &str,
    ) -> bool {
        let Ok(name) = std::str::from_utf8(qualified_name) else {
            return false;
        };
        let (prefix, local) = match name.split_once(':') {
            Some(parts) => parts,
            None => ("", name),
        };
        local == local_name
            && namespaces
                .get(prefix)
                .is_some_and(|candidate| candidate == namespace)
    }

    fn attribute_name_is(
        qualified_name: &[u8],
        namespaces: &BTreeMap<String, String>,
        namespace: &str,
        local_name: &str,
    ) -> bool {
        let Ok(name) = std::str::from_utf8(qualified_name) else {
            return false;
        };
        let Some((prefix, local)) = name.split_once(':') else {
            return false;
        };
        local == local_name
            && namespaces
                .get(prefix)
                .is_some_and(|candidate| candidate == namespace)
    }

    fn push_namespace_scope(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &mut BTreeMap<String, String>,
    ) -> Result<Vec<(String, Option<String>)>> {
        let mut previous_bindings = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| {
                Error::InvalidFormat(format!("invalid XML namespace declaration: {error}"))
            })?;
            let name = std::str::from_utf8(attribute.key.as_ref()).map_err(|_| {
                Error::InvalidFormat("invalid UTF-8 in XML namespace declaration".to_string())
            })?;
            let prefix = if name == "xmlns" {
                ""
            } else if let Some(prefix) = name.strip_prefix("xmlns:") {
                prefix
            } else {
                continue;
            };
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid XML namespace URI: {error}"))
                })?
                .into_owned();
            let prefix = prefix.to_string();
            let previous = namespaces.insert(prefix.clone(), value);
            previous_bindings.push((prefix, previous));
        }
        Ok(previous_bindings)
    }

    fn pop_namespace_scope(
        namespaces: &mut BTreeMap<String, String>,
        scope: Option<Vec<(String, Option<String>)>>,
    ) {
        let Some(scope) = scope else { return };
        for (prefix, previous) in scope.into_iter().rev() {
            if let Some(previous) = previous {
                namespaces.insert(prefix, previous);
            } else {
                namespaces.remove(&prefix);
            }
        }
    }

    fn push_text_empty_element(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
        text: &mut String,
    ) -> Result<()> {
        if Self::element_name_is(
            element.name().as_ref(),
            namespaces,
            TEXT_NAMESPACE,
            "line-break",
        ) {
            text.push('\n');
        } else if Self::element_name_is(element.name().as_ref(), namespaces, TEXT_NAMESPACE, "tab")
        {
            text.push('\t');
        } else if Self::element_name_is(element.name().as_ref(), namespaces, TEXT_NAMESPACE, "s") {
            let mut count = 1usize;
            for attribute in element.attributes() {
                let attribute = attribute.map_err(|error| {
                    Error::InvalidFormat(format!("invalid text:s attribute: {error}"))
                })?;
                if Self::attribute_name_is(attribute.key.as_ref(), namespaces, TEXT_NAMESPACE, "c")
                {
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid text:s count: {error}"))
                        })?;
                    count = value.parse::<usize>().map_err(|_| {
                        Error::InvalidFormat(format!("invalid text:s count '{value}'"))
                    })?;
                }
            }
            if count > 1_000_000 {
                return Err(Error::InvalidFormat(
                    "text:s count exceeds the supported safety limit".to_string(),
                ));
            }
            text.extend(std::iter::repeat_n(' ', count));
        }
        Ok(())
    }

    /// Parse document-global and sheet-local named ranges and expressions.
    ///
    /// Namespace URIs are resolved rather than assuming the conventional
    /// `table` prefix, because XML namespace prefixes are freely replaceable.
    pub fn parse_named_definitions(xml_content: &str) -> Result<Vec<NamedDefinition>> {
        let mut reader = NsReader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut definitions = Vec::new();
        let mut table_stack: Vec<Option<String>> = Vec::new();
        let mut active_scope: Option<NamedDefinitionScope> = None;

        loop {
            let event = reader
                .read_resolved_event_into(&mut buf)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;

            match event {
                (namespace, Event::Start(element)) if Self::is_table_namespace(&namespace) => {
                    match element.local_name().as_ref() {
                        b"table" => {
                            let name = Self::table_attribute(
                                reader.resolver(),
                                reader.decoder(),
                                &element,
                                b"name",
                            )?;
                            table_stack.push(name);
                        },
                        b"named-expressions" => {
                            if active_scope.is_some() {
                                return Err(Error::InvalidFormat(
                                    "nested table:named-expressions element".to_string(),
                                ));
                            }
                            active_scope = Some(match table_stack.last() {
                                Some(Some(sheet)) => NamedDefinitionScope::Sheet(sheet.clone()),
                                Some(None) => {
                                    return Err(Error::InvalidFormat(
                                        "sheet-local named definitions require table:name"
                                            .to_string(),
                                    ));
                                },
                                None => NamedDefinitionScope::Global,
                            });
                        },
                        b"named-range" | b"named-expression" => {
                            if let Some(scope) = &active_scope {
                                definitions.push(Self::parse_named_definition(
                                    reader.resolver(),
                                    reader.decoder(),
                                    &element,
                                    scope.clone(),
                                )?);
                            }
                        },
                        _ => {},
                    }
                },
                (namespace, Event::Empty(element)) if Self::is_table_namespace(&namespace) => {
                    match element.local_name().as_ref() {
                        b"named-range" | b"named-expression" => {
                            if let Some(scope) = &active_scope {
                                definitions.push(Self::parse_named_definition(
                                    reader.resolver(),
                                    reader.decoder(),
                                    &element,
                                    scope.clone(),
                                )?);
                            }
                        },
                        _ => {},
                    }
                },
                (namespace, Event::End(element)) if Self::is_table_namespace(&namespace) => {
                    match element.local_name().as_ref() {
                        b"named-expressions" => active_scope = None,
                        b"table" => {
                            table_stack.pop();
                        },
                        _ => {},
                    }
                },
                (_, Event::Eof) => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(definitions)
    }

    fn is_table_namespace(namespace: &ResolveResult<'_>) -> bool {
        matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == TABLE_NAMESPACE)
    }

    fn parse_named_definition(
        resolver: &NamespaceResolver,
        decoder: Decoder,
        element: &BytesStart<'_>,
        scope: NamedDefinitionScope,
    ) -> Result<NamedDefinition> {
        let name = Self::required_table_attribute(resolver, decoder, element, b"name")?;
        let base_cell_address =
            Self::table_attribute(resolver, decoder, element, b"base-cell-address")?;

        match element.local_name().as_ref() {
            b"named-range" => {
                let cell_range_address = Self::required_table_attribute(
                    resolver,
                    decoder,
                    element,
                    b"cell-range-address",
                )?;
                let mut range = NamedRange::new(name, cell_range_address, scope)?;
                range.base_cell_address = base_cell_address;
                if let Some(usable_as) =
                    Self::table_attribute(resolver, decoder, element, b"range-usable-as")?
                {
                    if usable_as.is_empty() {
                        return Err(Error::InvalidFormat(
                            "table:range-usable-as must not be empty".to_string(),
                        ));
                    }
                    if usable_as != "none" {
                        for token in usable_as.split_whitespace() {
                            let usage = NamedRangeUsage::parse(token)?;
                            if !range.usable_as.contains(&usage) {
                                range.usable_as.push(usage);
                            }
                        }
                    }
                }
                range.validate()?;
                Ok(range.into())
            },
            b"named-expression" => {
                let expression =
                    Self::required_table_attribute(resolver, decoder, element, b"expression")?;
                let formula_namespace_uri = Self::formula_namespace_uri(resolver, &expression)?;
                let mut expression = if let Some(uri) = formula_namespace_uri {
                    NamedExpression::new_with_namespace(name, expression, uri, scope)?
                } else {
                    NamedExpression::new(name, expression, scope)?
                };
                expression.base_cell_address = base_cell_address;
                expression.validate()?;
                Ok(expression.into())
            },
            _ => Err(Error::InvalidFormat(
                "unexpected named definition element".to_string(),
            )),
        }
    }

    fn required_table_attribute(
        resolver: &NamespaceResolver,
        decoder: Decoder,
        element: &BytesStart<'_>,
        local_name: &[u8],
    ) -> Result<String> {
        Self::table_attribute(resolver, decoder, element, local_name)?.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{} is missing required table:{} attribute",
                String::from_utf8_lossy(element.local_name().as_ref()),
                String::from_utf8_lossy(local_name)
            ))
        })
    }

    fn table_attribute(
        resolver: &NamespaceResolver,
        decoder: Decoder,
        element: &BytesStart<'_>,
        local_name: &[u8],
    ) -> Result<Option<String>> {
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let (namespace, local) = resolver.resolve_attribute(attribute.key);
            if Self::is_table_namespace(&namespace) && local.as_ref() == local_name {
                let value = attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid XML attribute value: {error}"))
                    })?;
                return Ok(Some(value.into_owned()));
            }
        }
        Ok(None)
    }

    fn formula_namespace_uri(
        resolver: &NamespaceResolver,
        expression: &str,
    ) -> Result<Option<String>> {
        let Some((prefix, remainder)) = expression.split_once(':') else {
            return Ok(None);
        };
        if prefix.is_empty() || !remainder.starts_with('=') {
            return Ok(None);
        }

        for (declaration, namespace) in resolver.bindings() {
            if let PrefixDeclaration::Named(candidate) = declaration
                && candidate == prefix.as_bytes()
            {
                return String::from_utf8(namespace.as_ref().to_vec())
                    .map(Some)
                    .map_err(|_| {
                        Error::InvalidFormat(format!(
                            "formula namespace for prefix '{prefix}' is not UTF-8"
                        ))
                    });
            }
        }

        Err(Error::InvalidFormat(format!(
            "formula prefix '{prefix}' is not bound to a namespace"
        )))
    }

    /// Extract table name from table:table element
    fn extract_table_name(
        e: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<String> {
        for attr_result in e.attributes() {
            let attr = attr_result
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            if Self::attribute_name_is(attr.key.as_ref(), namespaces, TABLE_NAMESPACE_URI, "name") {
                return attr
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map(|value| value.into_owned())
                    .map_err(|error| Error::InvalidFormat(format!("invalid table name: {error}")));
            }
        }
        Ok("Sheet1".to_string()) // Default name
    }

    fn parse_repeated(
        e: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
        local_name: &str,
    ) -> Result<usize> {
        for attribute in e.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                local_name,
            ) {
                return Self::parse_positive_usize(&attribute, decoder, local_name);
            }
        }
        Ok(1)
    }

    fn parse_structural_attributes(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<(Option<String>, Option<String>, TableVisibility)> {
        let mut style_name = None;
        let mut default_cell_style_name = None;
        let mut visibility = TableVisibility::Visible;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_table = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    TABLE_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_table("style-name") {
                style_name = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:style-name",
                )?);
            } else if is_table("default-cell-style-name") {
                default_cell_style_name = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:default-cell-style-name",
                )?);
            } else if is_table("visibility") {
                visibility = TableVisibility::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:visibility",
                )?)?;
            }
        }
        Ok((style_name, default_cell_style_name, visibility))
    }

    fn parse_group_display(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<bool> {
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "display",
            ) {
                return Self::parse_bool_attribute(&attribute, decoder);
            }
        }
        Ok(true)
    }

    fn parse_sheet_formatting(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<(SheetStyle, SheetPrintSettings)> {
        let mut style = SheetStyle::default();
        let mut print = SheetPrintSettings::default();
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_table = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    TABLE_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_table("style-name") {
                style.style_name = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:style-name",
                )?);
            } else if is_table("template-name") {
                style.template_name = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:template-name",
                )?);
            } else if is_table("use-first-row-styles") {
                style.usage.use_first_row_styles =
                    Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("use-last-row-styles") {
                style.usage.use_last_row_styles =
                    Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("use-first-column-styles") {
                style.usage.use_first_column_styles =
                    Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("use-last-column-styles") {
                style.usage.use_last_column_styles =
                    Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("use-banding-rows-styles") {
                style.usage.use_banding_row_styles =
                    Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("use-banding-columns-styles") {
                style.usage.use_banding_column_styles =
                    Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("print") {
                print.printable = Self::parse_bool_attribute(&attribute, decoder)?;
            } else if is_table("print-ranges") {
                print.ranges = split_cell_range_addresses(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:print-ranges",
                )?)?;
            }
        }
        Ok((style, print))
    }

    fn parse_scenario(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<SheetScenario> {
        let mut ranges = None;
        let mut is_active = None;
        let mut display_border = None;
        let mut border_color = None;
        let mut copy_back = None;
        let mut copy_styles = None;
        let mut copy_formulas = None;
        let mut comment = None;
        let mut protected = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_table = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    TABLE_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_table("scenario-ranges") {
                ranges = Some(split_cell_range_addresses(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:scenario-ranges",
                )?)?);
            } else if is_table("is-active") {
                is_active = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("display-border") {
                display_border = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("border-color") {
                border_color = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:border-color",
                )?);
            } else if is_table("copy-back") {
                copy_back = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("copy-styles") {
                copy_styles = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("copy-formulas") {
                copy_formulas = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("comment") {
                comment = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:comment",
                )?);
            } else if is_table("protected") {
                protected = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            }
        }
        let scenario = SheetScenario {
            ranges: ranges.ok_or_else(|| {
                Error::InvalidFormat("table:scenario requires table:scenario-ranges".to_string())
            })?,
            is_active: is_active.ok_or_else(|| {
                Error::InvalidFormat("table:scenario requires table:is-active".to_string())
            })?,
            display_border,
            border_color,
            copy_back,
            copy_styles,
            copy_formulas,
            comment,
            protected,
        };
        validate_scenario(&scenario)?;
        Ok(scenario)
    }

    /// Parse the attributes of a `text:a` hyperlink inside cell content.
    ///
    /// The visible link text is collected separately while the element's
    /// subtree is read, so the returned hyperlink has empty text.
    fn parse_hyperlink(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<CellHyperlink> {
        let mut href = None;
        let mut name = None;
        let mut title = None;
        let mut target_frame_name = None;
        let mut show = None;
        let mut actuate = None;
        let mut style_name = None;
        let mut visited_style_name = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let decode = |name| Self::decode_attribute(&attribute, decoder, name);
            if Self::attribute_name_is(attribute.key.as_ref(), namespaces, XLINK_NAMESPACE, "href")
            {
                href = Some(decode("xlink:href")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                XLINK_NAMESPACE,
                "type",
            ) {
                let value = decode("xlink:type")?;
                if value != "simple" {
                    return Err(Error::InvalidFormat(format!(
                        "invalid hyperlink xlink:type '{value}'"
                    )));
                }
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                OFFICE_NAMESPACE,
                "name",
            ) {
                name = Some(decode("office:name")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                OFFICE_NAMESPACE,
                "title",
            ) {
                title = Some(decode("office:title")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                OFFICE_NAMESPACE,
                "target-frame-name",
            ) {
                target_frame_name = Some(decode("office:target-frame-name")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                XLINK_NAMESPACE,
                "show",
            ) {
                let value = decode("xlink:show")?;
                show = Some(TextHyperlinkShow::parse(&value).ok_or_else(|| {
                    Error::InvalidFormat(format!("invalid hyperlink xlink:show '{value}'"))
                })?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                XLINK_NAMESPACE,
                "actuate",
            ) {
                let value = decode("xlink:actuate")?;
                actuate = Some(TextHyperlinkActuate::parse(&value).ok_or_else(|| {
                    Error::InvalidFormat(format!("invalid hyperlink xlink:actuate '{value}'"))
                })?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TEXT_NAMESPACE,
                "style-name",
            ) {
                style_name = Some(decode("text:style-name")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TEXT_NAMESPACE,
                "visited-style-name",
            ) {
                visited_style_name = Some(decode("text:visited-style-name")?);
            }
        }
        let href = href.ok_or_else(|| {
            Error::InvalidFormat("text:a hyperlink requires xlink:href".to_string())
        })?;
        Ok(CellHyperlink {
            href,
            text: String::new(),
            range: 0..0,
            name,
            title,
            target_frame_name,
            show,
            actuate,
            style_name,
            visited_style_name,
        })
    }

    fn parse_table_source(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<SheetTableSource> {
        let mut link_type = None;
        let mut href = None;
        let mut actuate_on_request = false;
        let mut mode = None;
        let mut table_name = None;
        let mut filter_name = None;
        let mut filter_options = None;
        let mut refresh_delay = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let decode = |name| Self::decode_attribute(&attribute, decoder, name);
            if Self::attribute_name_is(attribute.key.as_ref(), namespaces, XLINK_NAMESPACE, "type")
            {
                link_type = Some(decode("xlink:type")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                XLINK_NAMESPACE,
                "href",
            ) {
                href = Some(decode("xlink:href")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                XLINK_NAMESPACE,
                "actuate",
            ) {
                let value = decode("xlink:actuate")?;
                if value != "onRequest" {
                    return Err(Error::InvalidFormat(format!(
                        "invalid table source xlink:actuate '{value}'"
                    )));
                }
                actuate_on_request = true;
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "mode",
            ) {
                mode = Some(TableSourceMode::parse(&decode("table:mode")?)?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "table-name",
            ) {
                table_name = Some(decode("table:table-name")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "filter-name",
            ) {
                filter_name = Some(decode("table:filter-name")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "filter-options",
            ) {
                filter_options = Some(decode("table:filter-options")?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "refresh-delay",
            ) {
                refresh_delay = Some(decode("table:refresh-delay")?);
            }
        }
        let link_type = link_type.ok_or_else(|| {
            Error::InvalidFormat("table:table-source requires xlink:type".to_string())
        })?;
        if link_type != "simple" {
            return Err(Error::InvalidFormat(format!(
                "invalid table source xlink:type '{link_type}'"
            )));
        }
        let source = SheetTableSource {
            href: href.ok_or_else(|| {
                Error::InvalidFormat("table:table-source requires xlink:href".to_string())
            })?,
            mode,
            table_name,
            actuate_on_request,
            filter_name,
            filter_options,
            refresh_delay,
        };
        validate_table_source(&source)?;
        Ok(source)
    }

    fn parse_cell_range_source(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<CellRangeSource> {
        let mut name = None;
        let mut href = None;
        let mut link_type = None;
        let mut rows = None;
        let mut columns = None;
        let mut actuate_on_request = false;
        let mut filter_name = None;
        let mut filter_options = None;
        let mut refresh_delay = None;

        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let decode =
                |qualified_name| Self::decode_attribute(&attribute, decoder, qualified_name);
            let is_table = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    TABLE_NAMESPACE_URI,
                    local_name,
                )
            };
            let is_xlink = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    XLINK_NAMESPACE,
                    local_name,
                )
            };
            if is_table("name") {
                name = Some(decode("table:name")?);
            } else if is_table("last-row-spanned") {
                rows = Some(Self::parse_positive_usize(
                    &attribute,
                    decoder,
                    "last-row-spanned",
                )?);
            } else if is_table("last-column-spanned") {
                columns = Some(Self::parse_positive_usize(
                    &attribute,
                    decoder,
                    "last-column-spanned",
                )?);
            } else if is_table("filter-name") {
                filter_name = Some(decode("table:filter-name")?);
            } else if is_table("filter-options") {
                filter_options = Some(decode("table:filter-options")?);
            } else if is_table("refresh-delay") {
                refresh_delay = Some(decode("table:refresh-delay")?);
            } else if is_xlink("type") {
                link_type = Some(decode("xlink:type")?);
            } else if is_xlink("href") {
                href = Some(decode("xlink:href")?);
            } else if is_xlink("actuate") {
                let value = decode("xlink:actuate")?;
                if value != "onRequest" {
                    return Err(Error::InvalidFormat(format!(
                        "invalid cell range source xlink:actuate '{value}'"
                    )));
                }
                actuate_on_request = true;
            }
        }

        let link_type = link_type.ok_or_else(|| {
            Error::InvalidFormat("table:cell-range-source requires xlink:type".to_string())
        })?;
        if link_type != "simple" {
            return Err(Error::InvalidFormat(format!(
                "invalid cell range source xlink:type '{link_type}'"
            )));
        }
        let mut source = CellRangeSource::new(
            name.ok_or_else(|| {
                Error::InvalidFormat("table:cell-range-source requires table:name".to_string())
            })?,
            href.ok_or_else(|| {
                Error::InvalidFormat("table:cell-range-source requires xlink:href".to_string())
            })?,
            rows.ok_or_else(|| {
                Error::InvalidFormat(
                    "table:cell-range-source requires table:last-row-spanned".to_string(),
                )
            })?,
            columns.ok_or_else(|| {
                Error::InvalidFormat(
                    "table:cell-range-source requires table:last-column-spanned".to_string(),
                )
            })?,
        )?;
        source.set_actuate_on_request(actuate_on_request);
        source.set_filter_name(filter_name);
        source.set_filter_options(filter_options);
        source.set_refresh_delay(refresh_delay)?;
        Ok(source)
    }

    fn parse_detective_child(
        detective: &mut CellDetective,
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<()> {
        if Self::element_name_is(
            element.name().as_ref(),
            namespaces,
            TABLE_NAMESPACE_URI,
            "highlighted-range",
        ) {
            if !detective.operations().is_empty() {
                return Err(Error::InvalidFormat(
                    "table:highlighted-range must precede table:operation".to_string(),
                ));
            }
            detective
                .add_highlighted_range(Self::parse_detective_range(element, decoder, namespaces)?);
        } else if Self::element_name_is(
            element.name().as_ref(),
            namespaces,
            TABLE_NAMESPACE_URI,
            "operation",
        ) {
            detective.add_operation(Self::parse_detective_operation(
                element, decoder, namespaces,
            )?);
        } else {
            return Err(Error::InvalidFormat(
                "table:detective contains an unsupported child element".to_string(),
            ));
        }
        Ok(())
    }

    fn parse_detective_range(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<DetectiveHighlightedRange> {
        let mut address = None;
        let mut direction = None;
        let mut contains_error = None;
        let mut marked_invalid = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_table = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    TABLE_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_table("cell-range-address") {
                address = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:cell-range-address",
                )?);
            } else if is_table("direction") {
                direction = Some(DetectiveDirection::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:direction",
                )?)?);
            } else if is_table("contains-error") {
                contains_error = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_table("marked-invalid") {
                marked_invalid = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            }
        }

        if let Some(marked_invalid) = marked_invalid {
            if address.is_some() || direction.is_some() || contains_error.is_some() {
                return Err(Error::InvalidFormat(
                    "invalid detective ranges cannot contain directional range attributes"
                        .to_string(),
                ));
            }
            Ok(DetectiveHighlightedRange::invalid(marked_invalid))
        } else {
            DetectiveHighlightedRange::valid(
                address,
                direction.ok_or_else(|| {
                    Error::InvalidFormat(
                        "table:highlighted-range requires table:direction".to_string(),
                    )
                })?,
                contains_error,
            )
        }
    }

    fn parse_detective_operation(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<DetectiveOperation> {
        let mut kind = None;
        let mut index = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "name",
            ) {
                kind = Some(DetectiveOperationKind::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "table:name",
                )?)?);
            } else if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                TABLE_NAMESPACE_URI,
                "index",
            ) {
                let value = Self::decode_attribute(&attribute, decoder, "table:index")?;
                index = Some(value.parse::<usize>().map_err(|_| {
                    Error::InvalidFormat(format!(
                        "invalid non-negative detective operation index '{value}'"
                    ))
                })?);
            }
        }
        Ok(DetectiveOperation::new(
            kind.ok_or_else(|| {
                Error::InvalidFormat("table:operation requires table:name".to_string())
            })?,
            index.ok_or_else(|| {
                Error::InvalidFormat("table:operation requires table:index".to_string())
            })?,
        ))
    }

    fn parse_column(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<(Column, usize)> {
        let repeated =
            Self::parse_repeated(element, decoder, namespaces, "number-columns-repeated")?;
        let (style_name, default_cell_style_name, visibility) =
            Self::parse_structural_attributes(element, decoder, namespaces)?;
        Ok((
            Column {
                index: 0,
                style_name,
                default_cell_style_name,
                visibility,
            },
            repeated,
        ))
    }

    fn parse_positive_usize(
        attribute: &quick_xml::events::attributes::Attribute<'_>,
        decoder: Decoder,
        attribute_name: &str,
    ) -> Result<usize> {
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid table:{attribute_name}: {error}"))
            })?;
        let parsed = value.parse::<usize>().map_err(|_| {
            Error::InvalidFormat(format!("invalid table:{attribute_name} value '{value}'"))
        })?;
        if parsed == 0 {
            return Err(Error::InvalidFormat(format!(
                "table:{attribute_name} must be positive"
            )));
        }
        Ok(parsed)
    }

    /// Parse cell attributes and create a CellBuilder
    fn parse_cell_attributes(
        e: &quick_xml::events::BytesStart,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<CellBuilder> {
        let mut value_type = None;
        let mut value_str = None;
        let mut currency = None;
        let mut formula = None;
        let mut validation_name = None;
        let mut style_name = None;
        let mut protect = None;
        let mut protected = None;
        let mut repeated = 1;
        let mut row_span = 1;
        let mut column_span = 1;
        let mut matrix_row_span = None;
        let mut matrix_column_span = None;

        for attr_result in e.attributes() {
            let attr = attr_result
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_office = |local_name| {
                Self::attribute_name_is(attr.key.as_ref(), namespaces, OFFICE_NAMESPACE, local_name)
            };
            let is_table = |local_name| {
                Self::attribute_name_is(
                    attr.key.as_ref(),
                    namespaces,
                    TABLE_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_office("value-type") {
                value_type = Some(Self::decode_attribute(&attr, decoder, "office:value-type")?);
            } else if is_office("value") {
                value_str = Some(Self::decode_attribute(&attr, decoder, "office:value")?);
            } else if is_office("currency") {
                currency = Some(Self::decode_attribute(&attr, decoder, "office:currency")?);
            } else if is_table("formula") {
                formula = Some(Self::decode_attribute(&attr, decoder, "table:formula")?);
            } else if is_table("content-validation-name") {
                validation_name = Some(Self::decode_attribute(
                    &attr,
                    decoder,
                    "table:content-validation-name",
                )?);
            } else if is_table("style-name") {
                style_name = Some(Self::decode_attribute(&attr, decoder, "table:style-name")?);
            } else if is_table("protect") {
                protect = Some(Self::parse_bool_attribute(&attr, decoder)?);
            } else if is_table("protected") {
                protected = Some(Self::parse_bool_attribute(&attr, decoder)?);
            } else if is_table("number-columns-repeated") {
                repeated = Self::parse_positive_usize(&attr, decoder, "number-columns-repeated")?;
            } else if is_table("number-rows-spanned") {
                row_span = Self::parse_positive_usize(&attr, decoder, "number-rows-spanned")?;
            } else if is_table("number-columns-spanned") {
                column_span = Self::parse_positive_usize(&attr, decoder, "number-columns-spanned")?;
            } else if is_table("number-matrix-rows-spanned") {
                matrix_row_span = Some(Self::parse_positive_usize(
                    &attr,
                    decoder,
                    "number-matrix-rows-spanned",
                )?);
            } else if is_table("number-matrix-columns-spanned") {
                matrix_column_span = Some(Self::parse_positive_usize(
                    &attr,
                    decoder,
                    "number-matrix-columns-spanned",
                )?);
            }
        }

        Ok(CellBuilder {
            value_type,
            value_str,
            currency,
            formula,
            validation_name,
            style_name,
            matrix_span: if matrix_row_span.is_some() || matrix_column_span.is_some() {
                Some(CellMatrixSpan::new(
                    matrix_row_span.unwrap_or(1),
                    matrix_column_span.unwrap_or(1),
                )?)
            } else {
                None
            },
            protect,
            protected,
            repeated,
            merge: if row_span == 1 && column_span == 1 {
                CellMerge::None
            } else {
                CellMerge::Span {
                    rows: NonZeroUsize::new(row_span).expect("positive row span was checked"),
                    columns: NonZeroUsize::new(column_span)
                        .expect("positive column span was checked"),
                }
            },
            annotation: None,
            hyperlinks: Vec::new(),
            range_source: None,
            detective: None,
        })
    }

    fn parse_bool_attribute(
        attribute: &quick_xml::events::attributes::Attribute<'_>,
        decoder: Decoder,
    ) -> Result<bool> {
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::InvalidFormat(format!("invalid Boolean attribute: {error}")))?;
        match value.as_ref() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(Error::InvalidFormat(format!(
                "invalid Boolean attribute value '{value}'"
            ))),
        }
    }

    /// Parse the `calcext:target-range-address` of a `calcext:conditional-format`.
    fn parse_conditional_format_ranges(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<Vec<String>> {
        let mut ranges = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                CALCEXT_NAMESPACE_URI,
                "target-range-address",
            ) {
                if ranges.is_some() {
                    return Err(Error::InvalidFormat(
                        "duplicate calcext:target-range-address attribute".to_string(),
                    ));
                }
                ranges = Some(split_cell_range_addresses(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:target-range-address",
                )?)?);
            }
        }
        let ranges = ranges.ok_or_else(|| {
            Error::InvalidFormat(
                "calcext:conditional-format requires calcext:target-range-address".to_string(),
            )
        })?;
        if ranges.is_empty() {
            return Err(Error::InvalidFormat(
                "calcext:target-range-address requires at least one range".to_string(),
            ));
        }
        Ok(ranges)
    }

    /// Parse one inert `calcext:condition` rule from its attributes.
    fn parse_calcext_condition(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalFormatCondition> {
        let mut condition = None;
        let mut apply_style_name = None;
        let mut base_cell_address = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("value") {
                condition = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:value",
                )?);
            } else if is_calcext("apply-style-name") {
                apply_style_name = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:apply-style-name",
                )?);
            } else if is_calcext("base-cell-address") {
                base_cell_address = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:base-cell-address",
                )?);
            }
        }
        let rule = ConditionalFormatCondition {
            condition: condition.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:condition requires calcext:value".to_string(),
                )
            })?,
            apply_style_name: apply_style_name.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:condition requires calcext:apply-style-name".to_string(),
                )
            })?,
            base_cell_address,
        };
        validate_condition(&rule)?;
        Ok(rule)
    }

    /// Parse one inert `calcext:color-scale-entry` from its attributes.
    fn parse_color_scale_entry(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalColorScaleEntry> {
        let mut entry_type = None;
        let mut value = None;
        let mut color = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("type") {
                entry_type = Some(ConditionalFormatEntryType::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:type",
                )?)?);
            } else if is_calcext("value") {
                value = Some(Self::decode_attribute(&attribute, decoder, "calcext:value")?);
            } else if is_calcext("color") {
                color = Some(Self::decode_attribute(&attribute, decoder, "calcext:color")?);
            }
        }
        let entry = ConditionalColorScaleEntry {
            entry_type: entry_type.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:color-scale-entry requires calcext:type".to_string(),
                )
            })?,
            value: value.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:color-scale-entry requires calcext:value".to_string(),
                )
            })?,
            color: color.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:color-scale-entry requires calcext:color".to_string(),
                )
            })?,
        };
        validate_color_scale_entry(&entry)?;
        Ok(entry)
    }

    /// Parse one inert data-bar `calcext:formatting-entry` from its attributes.
    fn parse_data_bar_entry(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalDataBarEntry> {
        let (entry_type, value) = Self::parse_formatting_entry(element, decoder, namespaces)?;
        let entry = ConditionalDataBarEntry { entry_type, value };
        validate_data_bar_entry(&entry)?;
        Ok(entry)
    }

    /// Parse one inert icon-set `calcext:formatting-entry` from its attributes.
    fn parse_icon_set_entry(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalIconSetEntry> {
        let mut greater_equal = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            if Self::attribute_name_is(
                attribute.key.as_ref(),
                namespaces,
                CALCEXT_NAMESPACE_URI,
                "greater-equal",
            ) {
                greater_equal = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            }
        }
        let (entry_type, value) = Self::parse_formatting_entry(element, decoder, namespaces)?;
        let entry = ConditionalIconSetEntry {
            entry_type,
            value,
            greater_equal,
        };
        validate_icon_set_entry(&entry)?;
        Ok(entry)
    }

    /// Parse the shared `calcext:type`/`calcext:value` pair of a formatting entry.
    fn parse_formatting_entry(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<(ConditionalFormatEntryType, String)> {
        let mut entry_type = None;
        let mut value = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("type") {
                entry_type = Some(ConditionalFormatEntryType::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:type",
                )?)?);
            } else if is_calcext("value") {
                value = Some(Self::decode_attribute(&attribute, decoder, "calcext:value")?);
            }
        }
        Ok((
            entry_type.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:formatting-entry requires calcext:type".to_string(),
                )
            })?,
            value.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:formatting-entry requires calcext:value".to_string(),
                )
            })?,
        ))
    }

    /// Parse the attributes of an inert `calcext:data-bar` element.
    fn parse_data_bar_attributes(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalDataBar> {
        let mut data_bar = ConditionalDataBar::new(Vec::new());
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("positive-color") {
                data_bar.positive_color = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:positive-color",
                )?);
            } else if is_calcext("negative-color") {
                data_bar.negative_color = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:negative-color",
                )?);
            } else if is_calcext("gradient") {
                data_bar.gradient = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("axis-position") {
                data_bar.axis_position = Some(DataBarAxisPosition::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:axis-position",
                )?)?);
            } else if is_calcext("show-value") {
                data_bar.show_value = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("axis-color") {
                data_bar.axis_color = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:axis-color",
                )?);
            } else if is_calcext("min-length") {
                data_bar.min_length = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:min-length",
                )?);
            } else if is_calcext("max-length") {
                data_bar.max_length = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:max-length",
                )?);
            }
        }
        validate_data_bar_attributes(&data_bar)?;
        Ok(data_bar)
    }

    /// Parse the attributes of an inert `calcext:icon-set` element.
    fn parse_icon_set_attributes(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalIconSet> {
        let mut icon_set_type = None;
        let mut show_value = None;
        let mut custom = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("icon-set-type") {
                icon_set_type = Some(IconSetType::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:icon-set-type",
                )?)?);
            } else if is_calcext("show-value") {
                show_value = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("custom") {
                custom = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            }
        }
        let icon_set = ConditionalIconSet {
            icon_set_type: icon_set_type.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:icon-set requires calcext:icon-set-type".to_string(),
                )
            })?,
            show_value,
            custom,
            entries: Vec::new(),
        };
        Ok(icon_set)
    }

    /// Parse one inert `calcext:date-is` rule from its attributes.
    fn parse_date_is(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalDateIs> {
        let mut date = None;
        let mut style = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("date") {
                date = Some(ConditionalDateType::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:date",
                )?)?);
            } else if is_calcext("style") {
                style = Some(Self::decode_attribute(&attribute, decoder, "calcext:style")?);
            }
        }
        let date_is = ConditionalDateIs {
            date: date.ok_or_else(|| {
                Error::InvalidFormat("calcext:date-is requires calcext:date".to_string())
            })?,
            style: style.ok_or_else(|| {
                Error::InvalidFormat("calcext:date-is requires calcext:style".to_string())
            })?,
        };
        validate_date_is(&date_is)?;
        Ok(date_is)
    }

    /// Parse one inert `calcext:sparkline` cell assignment from its attributes.
    fn parse_sparkline(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<Sparkline> {
        let mut cell_address = None;
        let mut data_ranges = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_calcext("cell-address") {
                cell_address = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:cell-address",
                )?);
            } else if is_calcext("data-range") {
                data_ranges = Some(split_cell_range_addresses(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:data-range",
                )?)?);
            }
        }
        let sparkline = Sparkline {
            cell_address: cell_address.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:sparkline requires calcext:cell-address".to_string(),
                )
            })?,
            data_ranges: data_ranges.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:sparkline requires calcext:data-range".to_string(),
                )
            })?,
        };
        validate_sparkline(&sparkline)?;
        Ok(sparkline)
    }

    /// Parse the attributes of an inert `calcext:sparkline-group` element.
    fn parse_sparkline_group_attributes(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<SparklineGroup> {
        let mut group = SparklineGroup::new(Vec::new());
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_calcext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    CALCEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            let decode = |name| Self::decode_attribute(&attribute, decoder, name);
            if is_calcext("id") {
                group.id = Some(decode("calcext:id")?);
            } else if is_calcext("type") {
                group.sparkline_type = Some(SparklineType::parse(&decode("calcext:type")?)?);
            } else if is_calcext("line-width") {
                group.line_width = Some(decode("calcext:line-width")?);
            } else if is_calcext("date-axis") {
                group.flags.date_axis = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("display-empty-cells-as") {
                group.display_empty_cells_as = Some(SparklineEmptyCells::parse(&decode(
                    "calcext:display-empty-cells-as",
                )?)?);
            } else if is_calcext("markers") {
                group.flags.markers = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("high") {
                group.flags.high = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("low") {
                group.flags.low = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("first") {
                group.flags.first = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("last") {
                group.flags.last = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("negative") {
                group.flags.negative = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("display-x-axis") {
                group.flags.display_x_axis =
                    Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("display-hidden") {
                group.flags.display_hidden = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("min-axis-type") {
                group.min_axis_type =
                    Some(SparklineAxisType::parse(&decode("calcext:min-axis-type")?)?);
            } else if is_calcext("max-axis-type") {
                group.max_axis_type =
                    Some(SparklineAxisType::parse(&decode("calcext:max-axis-type")?)?);
            } else if is_calcext("right-to-left") {
                group.flags.right_to_left =
                    Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("manual-max") {
                group.manual_max = Some(decode("calcext:manual-max")?);
            } else if is_calcext("manual-min") {
                group.manual_min = Some(decode("calcext:manual-min")?);
            } else if is_calcext("color-series") {
                group.colors.series = Some(decode("calcext:color-series")?);
            } else if is_calcext("color-negative") {
                group.colors.negative = Some(decode("calcext:color-negative")?);
            } else if is_calcext("color-axis") {
                group.colors.axis = Some(decode("calcext:color-axis")?);
            } else if is_calcext("color-markers") {
                group.colors.markers = Some(decode("calcext:color-markers")?);
            } else if is_calcext("color-first") {
                group.colors.first = Some(decode("calcext:color-first")?);
            } else if is_calcext("color-last") {
                group.colors.last = Some(decode("calcext:color-last")?);
            } else if is_calcext("color-high") {
                group.colors.high = Some(decode("calcext:color-high")?);
            } else if is_calcext("color-low") {
                group.colors.low = Some(decode("calcext:color-low")?);
            }
        }
        validate_sparkline_group_attributes(&group)?;
        Ok(group)
    }

    fn decode_attribute(
        attribute: &quick_xml::events::attributes::Attribute<'_>,
        decoder: Decoder,
        name: &str,
    ) -> Result<String> {
        attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map(|value| value.into_owned())
            .map_err(|error| Error::InvalidFormat(format!("invalid {name}: {error}")))
    }
}

struct StructureContext {
    display: Option<bool>,
    children: Vec<TableStructure>,
    header_start: Option<usize>,
}

impl StructureContext {
    fn root() -> Self {
        Self {
            display: None,
            children: Vec::new(),
            header_start: None,
        }
    }
}

struct StructureStack {
    contexts: Vec<StructureContext>,
}

impl StructureStack {
    fn new() -> Self {
        Self {
            contexts: vec![StructureContext::root()],
        }
    }

    fn begin_group(&mut self, display: bool) -> Result<()> {
        if self.contexts.len() > MAX_TABLE_STRUCTURE_DEPTH {
            return Err(Error::InvalidFormat(format!(
                "table structure exceeds the {MAX_TABLE_STRUCTURE_DEPTH} level nesting safety limit"
            )));
        }
        if self
            .contexts
            .last()
            .is_some_and(|context| context.header_start.is_some())
        {
            return Err(Error::InvalidFormat(
                "table groups cannot be nested inside a header container".to_string(),
            ));
        }
        self.contexts.push(StructureContext {
            display: Some(display),
            children: Vec::new(),
            header_start: None,
        });
        Ok(())
    }

    fn end_group(&mut self) -> Result<()> {
        if self.contexts.len() <= 1 {
            return Err(Error::InvalidFormat(
                "table group end has no matching start".to_string(),
            ));
        }
        let context = self.contexts.pop().expect("non-root context was checked");
        if context.header_start.is_some() {
            return Err(Error::InvalidFormat(
                "table header container is not closed before its group".to_string(),
            ));
        }
        if context.children.is_empty() {
            return Err(Error::InvalidFormat(
                "table groups must contain at least one row or column".to_string(),
            ));
        }
        self.contexts
            .last_mut()
            .expect("root context is retained")
            .children
            .push(TableStructure::Group(TableGroup {
                display: context.display.expect("group contexts have display state"),
                children: context.children,
            }));
        Ok(())
    }

    fn begin_header(&mut self, position: usize) -> Result<()> {
        let context = self.contexts.last_mut().expect("root context is retained");
        if context.header_start.replace(position).is_some() {
            return Err(Error::InvalidFormat(
                "table header containers cannot be nested".to_string(),
            ));
        }
        Ok(())
    }

    fn end_header(&mut self, position: usize) -> Result<()> {
        let context = self.contexts.last_mut().expect("root context is retained");
        let start = context.header_start.take().ok_or_else(|| {
            Error::InvalidFormat("table header end has no matching start".to_string())
        })?;
        if position <= start {
            return Err(Error::InvalidFormat(
                "table header containers must not be empty".to_string(),
            ));
        }
        Ok(())
    }

    fn add_range(&mut self, start: usize, end: usize) -> Result<()> {
        let range = TableRange::new(start, end)?;
        let context = self.contexts.last_mut().expect("root context is retained");
        let entry = if context.header_start.is_some() {
            TableStructure::Header(range)
        } else {
            TableStructure::Range(range)
        };
        if let Some(previous) = context.children.last_mut() {
            match (previous, &entry) {
                (TableStructure::Range(previous), TableStructure::Range(next))
                | (TableStructure::Header(previous), TableStructure::Header(next))
                    if previous.end == next.start =>
                {
                    previous.end = next.end;
                    return Ok(());
                },
                _ => {},
            }
        }
        context.children.push(entry);
        Ok(())
    }

    fn finish(mut self) -> Result<Vec<TableStructure>> {
        if self.contexts.len() != 1 {
            return Err(Error::InvalidFormat(
                "table group is not closed before the table ends".to_string(),
            ));
        }
        let root = self.contexts.pop().expect("one root context was checked");
        if root.header_start.is_some() {
            return Err(Error::InvalidFormat(
                "table header container is not closed before the table ends".to_string(),
            ));
        }
        Ok(root.children)
    }
}

/// Builder for constructing Sheet during parsing
pub(crate) struct SheetBuilder {
    name: String,
    rows: Vec<Row>,
    columns: Vec<Column>,
    row_structure: StructureStack,
    column_structure: StructureStack,
    style: SheetStyle,
    print_settings: SheetPrintSettings,
    title: Option<String>,
    description: Option<String>,
    table_source: Option<SheetTableSource>,
    dde_source: Option<super::DdeSource>,
    scenario: Option<SheetScenario>,
    conditional_formats: Vec<ConditionalFormat>,
    sparkline_groups: Vec<super::SparklineGroup>,
    images: Vec<crate::OdfImage>,
    cell_count: usize,
    /// Runs of empty rows read but not yet materialised, in document order.
    deferred_rows: Vec<(Row, usize)>,
    /// Total number of rows the deferred runs stand for.
    deferred_row_count: usize,
}

impl SheetBuilder {
    #[cfg(test)]
    pub fn new(name: String) -> Self {
        Self::with_formatting(name, SheetStyle::default(), SheetPrintSettings::default())
    }

    fn with_formatting(
        name: String,
        style: SheetStyle,
        print_settings: SheetPrintSettings,
    ) -> Self {
        Self {
            name,
            rows: Vec::new(),
            columns: Vec::new(),
            row_structure: StructureStack::new(),
            column_structure: StructureStack::new(),
            style,
            print_settings,
            title: None,
            description: None,
            table_source: None,
            dde_source: None,
            scenario: None,
            conditional_formats: Vec::new(),
            sparkline_groups: Vec::new(),
            images: Vec::new(),
            cell_count: 0,
            deferred_rows: Vec::new(),
            deferred_row_count: 0,
        }
    }

    fn set_scenario(&mut self, scenario: SheetScenario) -> Result<()> {
        if self.scenario.replace(scenario).is_some() {
            return Err(Error::InvalidFormat(
                "a table must not contain more than one scenario".to_string(),
            ));
        }
        Ok(())
    }

    fn add_conditional_format(&mut self, format: ConditionalFormat) -> Result<()> {
        if self.conditional_formats.len() >= MAX_CONDITIONAL_FORMATS_PER_SHEET {
            return Err(Error::InvalidFormat(format!(
                "sheet exceeds the {MAX_CONDITIONAL_FORMATS_PER_SHEET} conditional format safety limit"
            )));
        }
        self.conditional_formats.push(format);
        Ok(())
    }

    fn add_sparkline_group(&mut self, group: super::SparklineGroup) -> Result<()> {
        if self.sparkline_groups.len() >= MAX_SPARKLINE_GROUPS_PER_SHEET {
            return Err(Error::InvalidFormat(format!(
                "sheet exceeds the {MAX_SPARKLINE_GROUPS_PER_SHEET} sparkline group safety limit"
            )));
        }
        self.sparkline_groups.push(group);
        Ok(())
    }

    fn set_dde_source(&mut self, source: super::DdeSource) -> Result<()> {
        if self.scenario.is_some() {
            return Err(Error::InvalidFormat(
                "office:dde-source must precede table:scenario".to_string(),
            ));
        }
        if self.dde_source.replace(source).is_some() {
            return Err(Error::InvalidFormat(
                "a table must not contain more than one office:dde-source".to_string(),
            ));
        }
        Ok(())
    }

    fn set_table_source(&mut self, source: SheetTableSource) -> Result<()> {
        if self.dde_source.is_some() || self.scenario.is_some() {
            return Err(Error::InvalidFormat(
                "table:table-source must precede office:dde-source and table:scenario".to_string(),
            ));
        }
        if self.table_source.replace(source).is_some() {
            return Err(Error::InvalidFormat(
                "a table must not contain more than one table source".to_string(),
            ));
        }
        Ok(())
    }

    fn set_title(&mut self, title: String) -> Result<()> {
        if self.title.replace(title).is_some() {
            return Err(Error::InvalidFormat(
                "a table must not contain more than one title".to_string(),
            ));
        }
        Ok(())
    }

    fn set_description(&mut self, description: String) -> Result<()> {
        if self.description.replace(description).is_some() {
            return Err(Error::InvalidFormat(
                "a table must not contain more than one description".to_string(),
            ));
        }
        Ok(())
    }

    pub fn add_row(&mut self, mut row: Row) {
        let row_index = self.rows.len();
        row.index = row_index;
        // Update row index for all cells in this row
        for cell in &mut row.cells {
            cell.row = row_index;
        }
        self.rows.push(row);
    }

    /// Number of rows the sheet logically spans, including runs of empty rows
    /// that have been deferred and may never be materialised.
    fn logical_row_count(&self) -> usize {
        self.rows.len().saturating_add(self.deferred_row_count)
    }

    /// Open a row grouping. Pending empty rows belong to the enclosing context,
    /// so they are materialised before the boundary moves.
    fn begin_row_group(&mut self, display: bool) -> Result<()> {
        self.flush_deferred_rows()?;
        self.row_structure.begin_group(display)
    }

    fn end_row_group(&mut self) -> Result<()> {
        self.flush_deferred_rows()?;
        self.row_structure.end_group()
    }

    fn begin_row_header(&mut self) -> Result<()> {
        self.flush_deferred_rows()?;
        self.row_structure.begin_header(self.rows.len())
    }

    fn end_row_header(&mut self) -> Result<()> {
        self.flush_deferred_rows()?;
        self.row_structure.end_header(self.rows.len())
    }

    fn add_repeated_row(&mut self, row: Row, repeated: usize) -> Result<()> {
        // The logical extent has to stay inside the grid a spreadsheet can
        // address, whether or not the rows are eventually materialised.
        let logical_end = self
            .logical_row_count()
            .checked_add(repeated)
            .ok_or_else(|| {
                Error::InvalidFormat("table row repetition overflows address space".to_string())
            })?;
        if logical_end > MAX_EXPANDED_ROWS_PER_SHEET {
            return Err(Error::InvalidFormat(format!(
                "expanded sheet exceeds the {MAX_EXPANDED_ROWS_PER_SHEET} row safety limit"
            )));
        }

        // Rows with no cell at all are the sheet-height padding producers append
        // after the used range. Defer them: an interior run is expanded again as
        // soon as a row with content follows, while a long trailing run is
        // discarded by `build` instead of costing a million allocations.
        if row.cells.is_empty() && self.deferred_rows.len() < MAX_DEFERRED_BLANK_ROW_RUNS {
            self.deferred_rows.push((row, repeated));
            self.deferred_row_count = self.deferred_row_count.saturating_add(repeated);
            return Ok(());
        }

        self.flush_deferred_rows()?;
        self.materialize_repeated_row(row, repeated)
    }

    /// Materialise every deferred empty-row run because a row with content or a
    /// structure boundary follows, so later rows keep their true index.
    fn flush_deferred_rows(&mut self) -> Result<()> {
        for (row, repeated) in std::mem::take(&mut self.deferred_rows) {
            self.materialize_repeated_row(row, repeated)?;
        }
        self.deferred_row_count = 0;
        Ok(())
    }

    /// Resolve the run of empty rows still pending at the end of a table.
    ///
    /// A short tail is kept, so an authored gap of blank rows survives the round
    /// trip. A long one is producer grid padding — every ODF spreadsheet is
    /// written out to its full addressable height — and is discarded, since it
    /// holds no cell, value, formula, annotation, or text. Discarding it records
    /// no structure range either, so the sheet's row groups never describe rows
    /// that are not there.
    fn finish_deferred_rows(&mut self) -> Result<()> {
        if self.deferred_row_count <= MAX_TRAILING_EMPTY_ROWS {
            return self.flush_deferred_rows();
        }
        self.deferred_rows.clear();
        self.deferred_row_count = 0;
        Ok(())
    }

    /// Expand one run of rows and record the physical range it occupies.
    fn materialize_repeated_row(&mut self, row: Row, repeated: usize) -> Result<()> {
        let start = self.rows.len();
        let added_cells = row.cells.len().checked_mul(repeated).ok_or_else(|| {
            Error::InvalidFormat("table row repetition overflows cell count".to_string())
        })?;
        let expanded_cells = self.cell_count.checked_add(added_cells).ok_or_else(|| {
            Error::InvalidFormat("expanded sheet cell count overflows address space".to_string())
        })?;
        if expanded_cells > MAX_EXPANDED_CELLS_PER_SHEET {
            return Err(Error::InvalidFormat(format!(
                "expanded sheet exceeds the {MAX_EXPANDED_CELLS_PER_SHEET} cell safety limit"
            )));
        }
        self.cell_count = expanded_cells;
        self.rows.reserve(repeated);
        for _ in 0..repeated {
            self.add_row(row.clone());
        }
        self.row_structure.add_range(start, self.rows.len())
    }

    fn add_repeated_column(&mut self, column: Column, repeated: usize) -> Result<()> {
        let start = self.columns.len();
        let expanded = self.columns.len().checked_add(repeated).ok_or_else(|| {
            Error::InvalidFormat("table column repetition overflows address space".to_string())
        })?;
        if expanded > MAX_EXPANDED_COLUMNS_PER_SHEET {
            return Err(Error::InvalidFormat(format!(
                "expanded sheet exceeds the {MAX_EXPANDED_COLUMNS_PER_SHEET} column safety limit"
            )));
        }
        for _ in 0..repeated {
            let mut item = column.clone();
            item.index = self.columns.len();
            self.columns.push(item);
        }
        self.column_structure.add_range(start, self.columns.len())?;
        Ok(())
    }

    pub fn build(mut self) -> Result<Sheet> {
        self.finish_deferred_rows()?;
        Ok(Sheet {
            name: self.name,
            rows: self.rows,
            columns: self.columns,
            column_structure: self.column_structure.finish()?,
            row_structure: self.row_structure.finish()?,
            style: self.style,
            print_settings: self.print_settings,
            title: self.title,
            description: self.description,
            table_source: self.table_source,
            dde_source: self.dde_source,
            scenario: self.scenario,
            conditional_formats: self.conditional_formats,
            sparkline_groups: self.sparkline_groups,
            images: self.images,
            shapes: Vec::new(),
            protection: super::SheetProtection::default(),
        })
    }
}

/// Builder for constructing Row during parsing
pub(crate) struct RowBuilder {
    cells: Vec<Cell>,
    repeated: usize,
    style_name: Option<String>,
    default_cell_style_name: Option<String>,
    visibility: TableVisibility,
    /// Number of attribute-free filler cells read but not yet materialised.
    deferred_blank_cells: usize,
    /// The filler cell to clone when the deferred run has to be materialised.
    deferred_blank_cell: Option<Cell>,
}

impl RowBuilder {
    #[cfg(test)]
    pub fn new() -> Self {
        Self {
            cells: Vec::new(),
            repeated: 1,
            style_name: None,
            default_cell_style_name: None,
            visibility: TableVisibility::Visible,
            deferred_blank_cells: 0,
            deferred_blank_cell: None,
        }
    }

    fn from_element(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<Self> {
        let repeated =
            OdsParser::parse_repeated(element, decoder, namespaces, "number-rows-repeated")?;
        let (style_name, default_cell_style_name, visibility) =
            OdsParser::parse_structural_attributes(element, decoder, namespaces)?;
        Ok(Self {
            cells: Vec::new(),
            repeated,
            style_name,
            default_cell_style_name,
            visibility,
            deferred_blank_cells: 0,
            deferred_blank_cell: None,
        })
    }

    pub fn add_cell(&mut self, mut cell: Cell) {
        cell.col = self.cells.len();
        self.cells.push(cell);
    }

    fn add_repeated_cells(
        &mut self,
        builder: &CellBuilder,
        text: &str,
        rich_text: Option<&CellTextContent>,
    ) -> Result<()> {
        // Producers pad every row out to the full sheet width with attribute-free
        // `<table:table-cell/>` fillers. Defer those instead of materialising them:
        // an interior run is still expanded when real content follows, but a
        // trailing run is dropped by `build`, which is what makes ordinary
        // spreadsheets fit inside the expansion safety limits at all.
        if builder.is_blank(text, rich_text) {
            self.deferred_blank_cells = self
                .deferred_blank_cells
                .checked_add(builder.repeated)
                .ok_or_else(|| {
                    Error::InvalidFormat(
                        "table cell repetition overflows address space".to_string(),
                    )
                })?;
            if self.deferred_blank_cell.is_none() {
                self.deferred_blank_cell = Some(builder.build(text, rich_text));
            }
            return Ok(());
        }
        self.flush_deferred_blank_cells()?;
        let expanded = self
            .cells
            .len()
            .checked_add(builder.repeated)
            .ok_or_else(|| {
                Error::InvalidFormat("table cell repetition overflows address space".to_string())
            })?;
        if expanded > MAX_EXPANDED_CELLS_PER_ROW {
            return Err(Error::InvalidFormat(format!(
                "expanded row exceeds the {MAX_EXPANDED_CELLS_PER_ROW} cell safety limit"
            )));
        }
        for _ in 0..builder.repeated {
            self.add_cell(builder.build(text, rich_text));
        }
        Ok(())
    }

    /// Materialise the deferred blank run because real content follows it, so
    /// the column index of that content stays correct.
    fn flush_deferred_blank_cells(&mut self) -> Result<()> {
        let deferred = std::mem::take(&mut self.deferred_blank_cells);
        let Some(template) = self.deferred_blank_cell.take() else {
            return Ok(());
        };
        if deferred == 0 {
            return Ok(());
        }
        let expanded = self.cells.len().checked_add(deferred).ok_or_else(|| {
            Error::InvalidFormat("table cell repetition overflows address space".to_string())
        })?;
        if expanded > MAX_EXPANDED_CELLS_PER_ROW {
            return Err(Error::InvalidFormat(format!(
                "expanded row exceeds the {MAX_EXPANDED_CELLS_PER_ROW} cell safety limit"
            )));
        }
        self.cells.reserve(deferred);
        for _ in 0..deferred {
            self.add_cell(template.clone());
        }
        Ok(())
    }

    pub fn build(mut self) -> Row {
        // Row index will be set by the parent SheetBuilder
        // For now, set to 0 and update cells
        for cell in &mut self.cells {
            cell.row = 0; // Will be updated by parent
        }

        Row {
            cells: self.cells,
            index: 0, // Will be set by parent
            style_name: self.style_name,
            default_cell_style_name: self.default_cell_style_name,
            visibility: self.visibility,
        }
    }
}

/// Builder for constructing Cell during parsing
pub(crate) struct CellBuilder {
    value_type: Option<String>,
    value_str: Option<String>,
    currency: Option<String>,
    formula: Option<String>,
    validation_name: Option<String>,
    style_name: Option<String>,
    matrix_span: Option<CellMatrixSpan>,
    protect: Option<bool>,
    protected: Option<bool>,
    repeated: usize,
    merge: CellMerge,
    annotation: Option<super::CellAnnotation>,
    hyperlinks: Vec<CellHyperlink>,
    range_source: Option<CellRangeSource>,
    detective: Option<CellDetective>,
}

impl CellBuilder {
    /// Whether this cell carries no user data whatsoever.
    ///
    /// A blank cell is exactly the attribute-free `<table:table-cell/>` filler
    /// producers emit to pad a row out to the full sheet width. Anything that a
    /// user could have authored — a value, formula, style, annotation,
    /// hyperlink, validation, protection flag, merge role, or text — makes the
    /// cell meaningful and therefore not blank.
    fn is_blank(&self, text_content: &str, rich_text: Option<&CellTextContent>) -> bool {
        self.value_type.is_none()
            && self.value_str.is_none()
            && self.currency.is_none()
            && self.formula.is_none()
            && self.validation_name.is_none()
            && self.style_name.is_none()
            && self.matrix_span.is_none()
            && self.protect.is_none()
            && self.protected.is_none()
            && self.annotation.is_none()
            && self.hyperlinks.is_empty()
            && self.range_source.is_none()
            && self.detective.is_none()
            && self.merge == CellMerge::None
            && text_content.is_empty()
            && rich_text.is_none()
    }

    pub fn build(&self, text_content: &str, rich_text: Option<&CellTextContent>) -> Cell {
        let value = self.parse_value(text_content);

        Cell {
            value,
            text: text_content.to_string(),
            // Clone necessary: formula may be reused for repeated cells
            formula: self.formula.clone(),
            annotation: self.annotation.clone(),
            hyperlinks: self.hyperlinks.clone(),
            rich_text: rich_text.cloned(),
            range_source: self.range_source.clone(),
            detective: self.detective.clone(),
            validation_name: self.validation_name.clone(),
            style_name: self.style_name.clone(),
            matrix_span: self.matrix_span,
            merge: self.merge,
            protect: self.protect,
            protected: self.protected,
            row: 0, // Will be set by parent
            col: 0, // Will be set by parent
        }
    }

    fn parse_value(&self, text_content: &str) -> CellValue {
        match self.value_type.as_deref() {
            Some("float") | Some("double") | Some("decimal") => {
                if let Some(ref val_str) = self.value_str {
                    if let Ok(num) = val_str.parse::<f64>() {
                        CellValue::Number(num)
                    } else {
                        CellValue::Text(text_content.to_string())
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("currency") => {
                if let Some(ref val_str) = self.value_str {
                    if let Ok(num) = val_str.parse::<f64>() {
                        let currency_code = self.currency.as_deref().unwrap_or("USD").to_string();
                        CellValue::Currency(num, currency_code)
                    } else {
                        CellValue::Text(text_content.to_string())
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("percentage") => {
                if let Some(ref val_str) = self.value_str {
                    if let Ok(num) = val_str.parse::<f64>() {
                        CellValue::Percentage(num)
                    } else {
                        CellValue::Text(text_content.to_string())
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("boolean") => {
                if let Some(ref val_str) = self.value_str {
                    match val_str.as_str() {
                        "true" => CellValue::Boolean(true),
                        "false" => CellValue::Boolean(false),
                        _ => CellValue::Text(text_content.to_string()),
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("date") => {
                if let Some(ref val_str) = self.value_str {
                    CellValue::Date(val_str.to_string())
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("time") => {
                if let Some(ref val_str) = self.value_str {
                    CellValue::Time(val_str.to_string())
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            _ => {
                if text_content.trim().is_empty() {
                    CellValue::Empty
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const TEST_SHEETS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="Sheet1">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>Hello</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="float" office:value="42">
                        <text:p>42</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_MULTIPLE_SHEETS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="Sheet1">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>First Sheet</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
            <table:table table:name="Sheet2">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>Second Sheet</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_CELL_TYPES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="TypesTest">
                <table:table-row>
                    <table:table-cell office:value-type="string"><text:p>Text</text:p></table:table-cell>
                    <table:table-cell office:value-type="float" office:value="3.14"><text:p>3.14</text:p></table:table-cell>
                    <table:table-cell office:value-type="currency" office:value="100" office:currency="EUR"><text:p>€100</text:p></table:table-cell>
                    <table:table-cell office:value-type="percentage" office:value="0.5"><text:p>50%</text:p></table:table-cell>
                    <table:table-cell office:value-type="boolean" office:value="true"><text:p>TRUE</text:p></table:table-cell>
                    <table:table-cell office:value-type="date" office:value="2024-03-15"><text:p>2024-03-15</text:p></table:table-cell>
                    <table:table-cell office:value-type="time" office:value="PT12H30M00S"><text:p>12:30:00</text:p></table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_FORMULA_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="FormulaTest">
                <table:table-row>
                    <table:table-cell office:value-type="float" office:value="10"><text:p>10</text:p></table:table-cell>
                    <table:table-cell office:value-type="float" office:value="20"><text:p>20</text:p></table:table-cell>
                    <table:table-cell table:formula="=SUM([.A1]:[.B1])" office:value-type="float" office:value="30">
                        <text:p>30</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_REPEATED_CELLS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="RepeatedTest">
                <table:table-row>
                    <table:table-cell table:number-columns-repeated="3" office:value-type="string">
                        <text:p>Repeated</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                        <text:p>Single</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_EMPTY_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="EmptySheet">
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_SPAN_TEXT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="SpanTest">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>Normal text <text:span>spanned text</text:span> more text</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    #[test]
    fn test_parse_sheets_basic() {
        let sheets = OdsParser::parse_sheets(TEST_SHEETS_XML).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Sheet1");
        assert_eq!(sheets[0].rows.len(), 1);
    }

    #[test]
    fn test_parse_multiple_sheets() {
        let sheets = OdsParser::parse_sheets(TEST_MULTIPLE_SHEETS_XML).unwrap();
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0].name, "Sheet1");
        assert_eq!(sheets[1].name, "Sheet2");
    }

    #[test]
    fn test_parse_cell_types() {
        let sheets = OdsParser::parse_sheets(TEST_CELL_TYPES_XML).unwrap();
        assert_eq!(sheets.len(), 1);

        let row = &sheets[0].rows[0];
        assert_eq!(row.cells.len(), 7);

        // Text cell
        match &row.cells[0].value {
            CellValue::Text(t) => assert_eq!(t, "Text"),
            _ => panic!("Expected Text"),
        }

        // Float/Number cell
        match &row.cells[1].value {
            CellValue::Number(n) => {
                let expected = (std::f64::consts::PI * 100.0).trunc() / 100.0;
                assert!((n - expected).abs() < f64::EPSILON);
            },
            _ => panic!("Expected Number"),
        }

        // Currency cell
        match &row.cells[2].value {
            CellValue::Currency(amount, currency) => {
                assert!((amount - 100.0).abs() < f64::EPSILON);
                assert_eq!(currency, "EUR");
            },
            _ => panic!("Expected Currency"),
        }

        // Percentage cell
        match &row.cells[3].value {
            CellValue::Percentage(p) => assert!((p - 0.5).abs() < f64::EPSILON),
            _ => panic!("Expected Percentage"),
        }

        // Boolean cell
        match &row.cells[4].value {
            CellValue::Boolean(b) => assert!(*b),
            _ => panic!("Expected Boolean"),
        }

        // Date cell
        match &row.cells[5].value {
            CellValue::Date(d) => assert_eq!(d, "2024-03-15"),
            _ => panic!("Expected Date"),
        }

        // Time cell
        match &row.cells[6].value {
            CellValue::Time(t) => assert_eq!(t, "PT12H30M00S"),
            _ => panic!("Expected Time"),
        }
    }

    #[test]
    fn test_parse_formula() {
        let sheets = OdsParser::parse_sheets(TEST_FORMULA_XML).unwrap();
        assert_eq!(sheets.len(), 1);

        let row = &sheets[0].rows[0];
        assert_eq!(row.cells.len(), 3);

        // Cell with formula
        assert_eq!(row.cells[2].formula, Some("=SUM([.A1]:[.B1])".to_string()));
        match &row.cells[2].value {
            CellValue::Number(n) => assert!((n - 30.0).abs() < f64::EPSILON),
            _ => panic!("Expected Number for formula result"),
        }
    }

    #[test]
    fn test_parse_repeated_cells() {
        let sheets = OdsParser::parse_sheets(TEST_REPEATED_CELLS_XML).unwrap();
        assert_eq!(sheets.len(), 1);

        let row = &sheets[0].rows[0];
        // 3 repeated cells + 1 single = 4 cells
        assert_eq!(row.cells.len(), 4);

        for i in 0..3 {
            match &row.cells[i].value {
                CellValue::Text(t) => assert_eq!(t, "Repeated"),
                _ => panic!("Expected Text for repeated cell {i}"),
            }
        }

        match &row.cells[3].value {
            CellValue::Text(t) => assert_eq!(t, "Single"),
            _ => panic!("Expected Text for single cell"),
        }
    }

    #[test]
    fn parses_cell_range_sources_with_namespace_aliases_and_repetition() {
        let xml = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
            xmlns:x="http://www.w3.org/1999/xlink">
          <o:body><o:spreadsheet><t:table t:name="Links"><t:table-row>
            <t:table-cell t:number-columns-repeated="2">
              <t:cell-range-source t:name="Named &amp; Range"
                t:last-column-spanned="4" t:last-row-spanned="3"
                t:filter-name="calc8" t:filter-options="A&amp;B"
                t:refresh-delay="PT15M" x:type="simple"
                x:href="../Data&amp;More.ods" x:actuate="onRequest"></t:cell-range-source>
            </t:table-cell>
          </t:table-row></t:table></o:spreadsheet></o:body>
        </o:document-content>"#;

        let sheets = OdsParser::parse_sheets(xml).unwrap();
        let cells = &sheets[0].rows[0].cells;
        assert_eq!(cells.len(), 2);
        for cell in cells {
            let source = cell.range_source().unwrap();
            assert_eq!(source.name(), "Named & Range");
            assert_eq!(source.href(), "../Data&More.ods");
            assert_eq!((source.rows(), source.columns()), (3, 4));
            assert!(source.actuate_on_request());
            assert_eq!(source.filter_name(), Some("calc8"));
            assert_eq!(source.filter_options(), Some("A&B"));
            assert_eq!(source.refresh_delay(), Some("PT15M"));
        }
    }

    #[test]
    fn rejects_incomplete_or_duplicate_cell_range_sources() {
        let missing_type = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
            xmlns:x="http://www.w3.org/1999/xlink">
          <o:body><o:spreadsheet><t:table t:name="Links"><t:table-row><t:table-cell>
            <t:cell-range-source t:name="R" t:last-column-spanned="1"
              t:last-row-spanned="1" x:href="source.ods"/>
          </t:table-cell></t:table-row></t:table></o:spreadsheet></o:body>
        </o:document-content>"#;
        assert!(OdsParser::parse_sheets(missing_type).is_err());

        let duplicate = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
            xmlns:x="http://www.w3.org/1999/xlink">
          <o:body><o:spreadsheet><t:table t:name="Links"><t:table-row><t:table-cell>
            <t:cell-range-source t:name="R1" t:last-column-spanned="1"
              t:last-row-spanned="1" x:type="simple" x:href="one.ods"/>
            <t:cell-range-source t:name="R2" t:last-column-spanned="1"
              t:last-row-spanned="1" x:type="simple" x:href="two.ods"/>
          </t:table-cell></t:table-row></t:table></o:spreadsheet></o:body>
        </o:document-content>"#;
        assert!(OdsParser::parse_sheets(duplicate).is_err());
    }

    #[test]
    fn parses_typed_detective_ranges_and_operations() {
        let xml = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:table t:name="Audit"><t:table-row>
            <t:table-cell t:number-columns-repeated="2"><t:detective>
              <t:highlighted-range t:cell-range-address=".A1:.B2"
                t:direction="from-same-table" t:contains-error="true"/>
              <t:highlighted-range t:marked-invalid="false"></t:highlighted-range>
              <t:operation t:name="trace-precedents" t:index="0"/>
              <t:operation t:name="trace-errors" t:index="7"></t:operation>
            </t:detective></t:table-cell>
          </t:table-row></t:table></o:spreadsheet></o:body>
        </o:document-content>"#;

        let sheets = OdsParser::parse_sheets(xml).unwrap();
        let cells = &sheets[0].rows[0].cells;
        assert_eq!(cells.len(), 2);
        for cell in cells {
            let detective = cell.detective().unwrap();
            assert_eq!(detective.highlighted_ranges().len(), 2);
            assert_eq!(detective.operations().len(), 2);
            let range = &detective.highlighted_ranges()[0];
            assert_eq!(range.cell_range_address(), Some(".A1:.B2"));
            assert_eq!(range.direction(), Some(DetectiveDirection::FromSameTable));
            assert_eq!(range.contains_error(), Some(true));
            assert_eq!(range.marked_invalid(), None);
            assert_eq!(
                detective.highlighted_ranges()[1].marked_invalid(),
                Some(false)
            );
            assert_eq!(
                detective.operations()[1],
                DetectiveOperation::new(DetectiveOperationKind::TraceErrors, 7)
            );
        }
    }

    #[test]
    fn rejects_schema_invalid_detective_metadata() {
        let operation_before_range = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:table t:name="Audit"><t:table-row><t:table-cell>
            <t:detective><t:operation t:name="trace-errors" t:index="0"/>
              <t:highlighted-range t:direction="from-same-table"/></t:detective>
          </t:table-cell></t:table-row></t:table></o:spreadsheet></o:body>
        </o:document-content>"#;
        assert!(OdsParser::parse_sheets(operation_before_range).is_err());

        let mixed_range = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:table t:name="Audit"><t:table-row><t:table-cell>
            <t:detective><t:highlighted-range t:marked-invalid="true"
              t:direction="from-same-table"/></t:detective>
          </t:table-cell></t:table-row></t:table></o:spreadsheet></o:body>
        </o:document-content>"#;
        assert!(OdsParser::parse_sheets(mixed_range).is_err());

        let negative_index = operation_before_range
            .replace(
                r#"t:name="trace-errors" t:index="0""#,
                r#"t:name="trace-errors" t:index="-1""#,
            )
            .replace(
                r#"<t:highlighted-range t:direction="from-same-table"/>"#,
                "",
            );
        assert!(OdsParser::parse_sheets(&negative_index).is_err());

        let nested_child = operation_before_range
            .replace(
                r#"<t:operation t:name="trace-errors" t:index="0"/>"#,
                "",
            )
            .replace(
                r#"<t:highlighted-range t:direction="from-same-table"/>"#,
                r#"<t:highlighted-range t:direction="from-same-table"><t:operation t:name="trace-errors" t:index="1"/></t:highlighted-range>"#,
            );
        assert!(OdsParser::parse_sheets(&nested_child).is_err());
    }

    #[test]
    fn test_parse_empty_sheet() {
        let sheets = OdsParser::parse_sheets(TEST_EMPTY_SHEET_XML).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "EmptySheet");
        assert_eq!(sheets[0].rows.len(), 0);
    }

    #[test]
    fn test_parse_span_text() {
        let sheets = OdsParser::parse_sheets(TEST_SPAN_TEXT_XML).unwrap();
        assert_eq!(sheets.len(), 1);

        let row = &sheets[0].rows[0];
        assert_eq!(row.cells.len(), 1);

        // Text should include content from both text:p and text:span
        assert!(row.cells[0].text.contains("Normal text"));
        assert!(row.cells[0].text.contains("spanned text"));
    }

    #[test]
    fn parses_rich_annotations_without_mixing_them_into_cell_text() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">
  <o:body><o:spreadsheet><table:table table:name="Notes"><table:table-row>
    <table:table-cell table:number-columns-repeated="2" o:value-type="string">
      <o:annotation o:display="true" draw:style-name="gr1" svg:width="3.2cm">
        <dc:creator>A &amp; B</dc:creator><dc:date>2026-07-13T12:34:56Z</dc:date>
        <text:p text:style-name="P1"><text:span text:style-name="T1">first</text:span><text:line-break/>second</text:p>
        <text:list><text:list-item><text:p>item</text:p></text:list-item></text:list>
      </o:annotation>
      <text:p>cell <text:span>value</text:span></text:p><text:p>line two</text:p>
    </table:table-cell>
  </table:table-row></table:table></o:spreadsheet></o:body>
</o:document-content>"#;

        let sheets = OdsParser::parse_sheets(xml).unwrap();
        let cells = &sheets[0].rows[0].cells;
        assert_eq!(cells.len(), 2);
        assert_eq!(cells[0].text, "cell value\nline two");
        assert_eq!(cells[1].text, "cell value\nline two");

        for cell in cells {
            let annotation = cell.annotation().unwrap();
            assert_eq!(annotation.creator().as_deref(), Some("A & B"));
            assert_eq!(annotation.date().as_deref(), Some("2026-07-13T12:34:56Z"));
            assert_eq!(annotation.display(), Some(true));
            assert_eq!(annotation.attribute("draw:style-name"), Some("gr1"));
            assert_eq!(annotation.attribute("svg:width"), Some("3.2cm"));
            assert_eq!(annotation.text(), "first\nsecond\nitem");
            assert_eq!(annotation.children()[2].name(), "text:p");
            assert_eq!(annotation.children()[3].name(), "text:list");
        }
    }

    #[test]
    fn test_extract_table_name_default() {
        // XML without table:name attribute
        let xml = r#"<?xml version="1.0"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
    <office:body><office:spreadsheet><table:table/></office:spreadsheet></office:body>
</office:document-content>"#;

        let sheets = OdsParser::parse_sheets(xml).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Sheet1"); // Default name
    }

    #[test]
    fn parses_repeated_rows_and_merged_cell_coordinates() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:spreadsheet><table:table table:name="Merged"><table:table-row table:number-rows-repeated="2"><table:table-cell office:value-type="string"><text:p>A</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell table:number-rows-spanned="2" table:number-columns-spanned="2" office:value-type="string"><text:p>anchor</text:p></table:table-cell><table:covered-table-cell/><table:table-cell table:number-matrix-rows-spanned="3" table:number-matrix-columns-spanned="2" office:value-type="string"><text:p>C</text:p></table:table-cell></table:table-row><table:table-row><table:covered-table-cell table:number-columns-repeated="2"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;
        let sheets = OdsParser::parse_sheets(xml).unwrap();
        let rows = &sheets[0].rows;
        assert_eq!(rows.len(), 4);
        assert_eq!(rows[0].cells[0].text, "A");
        assert_eq!(rows[1].cells[0].coordinates(), (1, 0));
        assert_eq!(rows[2].cells[0].span(), Some((2, 2)));
        assert_eq!(rows[2].cells[1].merge(), CellMerge::Covered);
        assert_eq!(rows[2].cells[2].text, "C");
        assert_eq!(
            rows[2].cells[2]
                .matrix_span()
                .map(|span| (span.rows(), span.columns())),
            Some((3, 2))
        );
        assert_eq!(rows[2].cells[2].coordinates(), (2, 2));
        assert_eq!(rows[3].cells.len(), 2);
        assert!(
            rows[3]
                .cells
                .iter()
                .all(|cell| cell.merge() == CellMerge::Covered)
        );
    }

    #[test]
    fn parses_sheet_content_with_arbitrary_namespace_prefixes() {
        let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:x="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:spreadsheet><t:table t:name="A&amp;B"><t:table-row t:number-rows-repeated="2"><t:table-cell o:value-type="string" t:style-name="Style&amp;One" t:protected="1"><x:p>one<x:s x:c="2"/>two<x:tab/>three<x:line-break/>four</x:p></t:table-cell><t:covered-table-cell/></t:table-row></t:table></o:spreadsheet></o:body></o:document-content>"#;
        let sheets = OdsParser::parse_sheets(xml).unwrap();
        assert_eq!(sheets[0].name, "A&B");
        assert_eq!(sheets[0].rows.len(), 2);
        for (row_index, row) in sheets[0].rows.iter().enumerate() {
            assert_eq!(row.cells[0].coordinates(), (row_index, 0));
            assert_eq!(row.cells[0].style_name(), Some("Style&One"));
            assert_eq!(row.cells[0].protected(), Some(true));
            assert_eq!(row.cells[0].text, "one  two\tthree\nfour");
            assert_eq!(row.cells[1].merge(), CellMerge::Covered);
        }
    }

    #[test]
    fn parses_repeated_row_and_column_structural_metadata() {
        let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><o:body><o:spreadsheet><t:table t:name="Structure"><t:table-column t:number-columns-repeated="2" t:style-name="Col&amp;Style" t:default-cell-style-name="CellStyle" t:visibility="collapse"/><t:table-column t:visibility="filter"></t:table-column><t:table-row t:number-rows-repeated="2" t:style-name="RowStyle" t:default-cell-style-name="RowCell" t:visibility="filter"><t:table-cell/></t:table-row></t:table></o:spreadsheet></o:body></o:document-content>"#;
        let sheets = OdsParser::parse_sheets(xml).unwrap();
        let sheet = &sheets[0];
        assert_eq!(sheet.columns.len(), 3);
        assert_eq!(sheet.columns[0].index, 0);
        assert_eq!(sheet.columns[1].index, 1);
        assert_eq!(sheet.columns[0].style_name.as_deref(), Some("Col&Style"));
        assert_eq!(
            sheet.columns[0].default_cell_style_name.as_deref(),
            Some("CellStyle")
        );
        assert_eq!(sheet.columns[0].visibility, TableVisibility::Collapse);
        assert_eq!(sheet.columns[2].visibility, TableVisibility::Filter);
        assert_eq!(sheet.rows.len(), 2);
        assert!(sheet.rows.iter().all(|row| {
            row.style_name.as_deref() == Some("RowStyle")
                && row.default_cell_style_name.as_deref() == Some("RowCell")
                && row.visibility == TableVisibility::Filter
        }));
    }

    #[test]
    fn parses_nested_groups_and_header_ranges() {
        let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><o:body><o:spreadsheet><t:table t:name="Outline"><t:table-column-group t:display="false"><t:table-header-columns><t:table-column/></t:table-header-columns><t:table-column-group><t:table-column t:number-columns-repeated="2"/></t:table-column-group></t:table-column-group><t:table-row-group t:display="false"><t:table-header-rows><t:table-row/></t:table-header-rows><t:table-row-group><t:table-row t:number-rows-repeated="2"/></t:table-row-group></t:table-row-group></t:table></o:spreadsheet></o:body></o:document-content>"#;
        let sheets = OdsParser::parse_sheets(xml).unwrap();
        assert_eq!(
            sheets[0].column_structure,
            vec![TableStructure::Group(TableGroup {
                display: false,
                children: vec![
                    TableStructure::Header(TableRange { start: 0, end: 1 }),
                    TableStructure::Group(TableGroup {
                        display: true,
                        children: vec![TableStructure::Range(TableRange { start: 1, end: 3 })],
                    }),
                ],
            })]
        );
        assert_eq!(
            sheets[0].row_structure,
            vec![TableStructure::Group(TableGroup {
                display: false,
                children: vec![
                    TableStructure::Header(TableRange { start: 0, end: 1 }),
                    TableStructure::Group(TableGroup {
                        display: true,
                        children: vec![TableStructure::Range(TableRange { start: 1, end: 3 })],
                    }),
                ],
            })]
        );
    }

    #[test]
    fn parses_sheet_style_and_print_settings() {
        let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><o:body><o:spreadsheet><t:table t:name="Print" t:style-name="Sheet&amp;Style" t:template-name="TemplateOne" t:use-first-row-styles="true" t:use-last-row-styles="0" t:use-first-column-styles="1" t:use-last-column-styles="false" t:use-banding-rows-styles="true" t:use-banding-columns-styles="false" t:print="false" t:print-ranges="$Print.$A$1:$B$2 'Q1 Sales'.$C$3:$D$4"></t:table></o:spreadsheet></o:body></o:document-content>"#;
        let sheets = OdsParser::parse_sheets(xml).unwrap();
        let sheet = &sheets[0];
        assert_eq!(sheet.style.style_name.as_deref(), Some("Sheet&Style"));
        assert_eq!(sheet.style.template_name.as_deref(), Some("TemplateOne"));
        assert_eq!(sheet.style.usage.use_first_row_styles, Some(true));
        assert_eq!(sheet.style.usage.use_last_row_styles, Some(false));
        assert_eq!(sheet.style.usage.use_first_column_styles, Some(true));
        assert_eq!(sheet.style.usage.use_last_column_styles, Some(false));
        assert_eq!(sheet.style.usage.use_banding_row_styles, Some(true));
        assert_eq!(sheet.style.usage.use_banding_column_styles, Some(false));
        assert!(!sheet.print_settings.printable);
        assert_eq!(
            sheet.print_settings.ranges,
            ["$Print.$A$1:$B$2", "'Q1 Sales'.$C$3:$D$4"]
        );
    }

    #[test]
    fn parses_sheet_title_description_and_scenario() {
        let xml = r##"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:l="http://www.w3.org/1999/xlink"><o:body><o:spreadsheet><t:table t:name="Scenario"><t:title>Quarter &amp; Forecast</t:title><t:desc><![CDATA[Best < worst]]></t:desc><t:table-source l:type="simple" l:href="../Q1&amp;Q2.ods" l:actuate="onRequest" t:mode="copy-results-only" t:table-name="Source Sheet" t:filter-name="calc8" t:filter-options="A&amp;B" t:refresh-delay="P1DT2H3.5S"/><t:scenario t:scenario-ranges="$Scenario.$A$1:$B$2 'Q1 Sales'.$C$3:$D$4" t:is-active="true" t:display-border="0" t:border-color="#12AbEF" t:copy-back="1" t:copy-styles="false" t:copy-formulas="true" t:comment="Best &amp; worst" t:protected="false"/></t:table></o:spreadsheet></o:body></o:document-content>"##;
        let sheets = OdsParser::parse_sheets(xml).unwrap();
        let sheet = &sheets[0];
        assert_eq!(sheet.title.as_deref(), Some("Quarter & Forecast"));
        assert_eq!(sheet.description.as_deref(), Some("Best < worst"));
        let source = sheet.table_source.as_ref().unwrap();
        assert_eq!(source.href, "../Q1&Q2.ods");
        assert_eq!(source.mode, Some(TableSourceMode::CopyResultsOnly));
        assert_eq!(source.table_name.as_deref(), Some("Source Sheet"));
        assert!(source.actuate_on_request);
        assert_eq!(source.filter_name.as_deref(), Some("calc8"));
        assert_eq!(source.filter_options.as_deref(), Some("A&B"));
        assert_eq!(source.refresh_delay.as_deref(), Some("P1DT2H3.5S"));
        let scenario = sheet.scenario.as_ref().unwrap();
        assert_eq!(
            scenario.ranges,
            ["$Scenario.$A$1:$B$2", "'Q1 Sales'.$C$3:$D$4"]
        );
        assert!(scenario.is_active);
        assert_eq!(scenario.display_border, Some(false));
        assert_eq!(scenario.border_color.as_deref(), Some("#12AbEF"));
        assert_eq!(scenario.copy_back, Some(true));
        assert_eq!(scenario.copy_styles, Some(false));
        assert_eq!(scenario.copy_formulas, Some(true));
        assert_eq!(scenario.comment.as_deref(), Some("Best & worst"));
        assert_eq!(scenario.protected, Some(false));
    }

    #[test]
    fn rejects_invalid_or_dangerous_repetition_counts() {
        let zero = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:table-row table:number-rows-repeated="0"/></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert!(OdsParser::parse_sheets(zero).is_err());

        let excessive = format!(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:table-row table:number-rows-repeated="{}"/></table:table></office:spreadsheet></office:body></office:document-content>"#,
            MAX_EXPANDED_ROWS_PER_SHEET + 1
        );
        assert!(OdsParser::parse_sheets(&excessive).is_err());

        let excessive_columns = format!(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:table-column table:number-columns-repeated="{}"/></table:table></office:spreadsheet></office:body></office:document-content>"#,
            MAX_EXPANDED_COLUMNS_PER_SHEET + 1
        );
        assert!(OdsParser::parse_sheets(&excessive_columns).is_err());

        let invalid_visibility = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:table-column table:visibility="hidden"/></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert!(OdsParser::parse_sheets(invalid_visibility).is_err());

        let invalid_group_display = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:table-row-group table:display="collapsed"><table:table-row/></table:table-row-group></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert!(OdsParser::parse_sheets(invalid_group_display).is_err());

        let empty_group = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:table-column-group/></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert!(OdsParser::parse_sheets(empty_group).is_err());

        let invalid_print = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table table:print="yes"></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert!(OdsParser::parse_sheets(invalid_print).is_err());

        let invalid_print_ranges = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table table:print-ranges="'Unclosed Sheet.$A$1"></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert!(OdsParser::parse_sheets(invalid_print_ranges).is_err());

        let incomplete_scenario = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:scenario table:scenario-ranges=".A1:.B2"/></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert!(OdsParser::parse_sheets(incomplete_scenario).is_err());

        let invalid_scenario_color = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:scenario table:scenario-ranges=".A1:.B2" table:is-active="false" table:border-color="red"/></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert!(OdsParser::parse_sheets(invalid_scenario_color).is_err());

        let duplicate_scenarios = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:scenario table:scenario-ranges=".A1:.B2" table:is-active="false"/><table:scenario table:scenario-ranges=".C1:.D2" table:is-active="true"/></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert!(OdsParser::parse_sheets(duplicate_scenarios).is_err());

        let duplicate_titles = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:title>First</table:title><table:title>Second</table:title></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert!(OdsParser::parse_sheets(duplicate_titles).is_err());

        let duplicate_descriptions = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:desc>First</table:desc><table:desc>Second</table:desc></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert!(OdsParser::parse_sheets(duplicate_descriptions).is_err());

        let invalid_sources = [
            r#"<table:table-source xlink:href="a.ods"/>"#,
            r#"<table:table-source xlink:type="extended" xlink:href="a.ods"/>"#,
            r#"<table:table-source xlink:type="simple"/>"#,
            r#"<table:table-source xlink:type="simple" xlink:href="a.ods" xlink:actuate="onLoad"/>"#,
            r#"<table:table-source xlink:type="simple" xlink:href="a.ods" table:mode="values"/>"#,
            r#"<table:table-source xlink:type="simple" xlink:href="a.ods" table:refresh-delay="15 minutes"/>"#,
        ];
        for source in invalid_sources {
            let xml = format!(
                r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:spreadsheet><table:table>{source}</table:table></office:spreadsheet></office:body></office:document-content>"#
            );
            assert!(OdsParser::parse_sheets(&xml).is_err(), "{source}");
        }

        let duplicate_sources = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:spreadsheet><table:table><table:table-source xlink:type="simple" xlink:href="a.ods"/><table:table-source xlink:type="simple" xlink:href="b.ods"/></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert!(OdsParser::parse_sheets(duplicate_sources).is_err());
    }

    #[test]
    fn test_sheet_builder() {
        let mut builder = SheetBuilder::new("TestSheet".to_string());

        let row1 = Row {
            cells: vec![],
            index: 0,
            style_name: None,
            default_cell_style_name: None,
            visibility: Default::default(),
        };
        builder.add_row(row1);

        let row2 = Row {
            cells: vec![Cell {
                value: CellValue::Text("A1".to_string()),
                text: "A1".to_string(),
                formula: None,
                annotation: None,
                hyperlinks: Vec::new(),
                rich_text: None,
                range_source: None,
                detective: None,
                validation_name: None,
                style_name: None,
                matrix_span: None,
                merge: Default::default(),
                protect: None,
                protected: None,
                row: 0,
                col: 0,
            }],
            index: 0,
            style_name: None,
            default_cell_style_name: None,
            visibility: Default::default(),
        };
        builder.add_row(row2);

        let sheet = builder.build().unwrap();
        assert_eq!(sheet.name, "TestSheet");
        assert_eq!(sheet.rows.len(), 2);
        assert_eq!(sheet.rows[0].index, 0);
        assert_eq!(sheet.rows[1].index, 1);
    }

    #[test]
    fn test_row_builder() {
        let mut builder = RowBuilder::new();

        let cell1 = Cell {
            value: CellValue::Text("A".to_string()),
            text: "A".to_string(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        builder.add_cell(cell1);

        let cell2 = Cell {
            value: CellValue::Number(42.0),
            text: "42".to_string(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        builder.add_cell(cell2);

        let row = builder.build();
        assert_eq!(row.cells.len(), 2);
        assert_eq!(row.cells[0].col, 0);
        assert_eq!(row.cells[1].col, 1);
    }

    #[test]
    fn test_cell_builder_float_types() {
        // Test "float" value type
        let builder = CellBuilder {
            value_type: Some("float".to_string()),
            value_str: Some("123.45".to_string()),
            currency: None,
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("123.45", None);
        match cell.value {
            CellValue::Number(n) => assert!((n - 123.45).abs() < f64::EPSILON),
            _ => panic!("Expected Number for float"),
        }

        // Test "double" value type
        let builder = CellBuilder {
            value_type: Some("double".to_string()),
            value_str: Some("99.99".to_string()),
            currency: None,
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("99.99", None);
        match cell.value {
            CellValue::Number(n) => assert!((n - 99.99).abs() < f64::EPSILON),
            _ => panic!("Expected Number for double"),
        }

        // Test "decimal" value type
        let builder = CellBuilder {
            value_type: Some("decimal".to_string()),
            value_str: Some("0.001".to_string()),
            currency: None,
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("0.001", None);
        match cell.value {
            CellValue::Number(n) => assert!((n - 0.001).abs() < f64::EPSILON),
            _ => panic!("Expected Number for decimal"),
        }
    }

    #[test]
    fn test_cell_builder_invalid_number_fallback() {
        let builder = CellBuilder {
            value_type: Some("float".to_string()),
            value_str: Some("not-a-number".to_string()),
            currency: None,
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("some text", None);
        match cell.value {
            CellValue::Text(t) => assert_eq!(t, "some text"),
            _ => panic!("Expected Text fallback for invalid number"),
        }
    }

    #[test]
    fn test_cell_builder_boolean_variations() {
        // Test "false" boolean
        let builder = CellBuilder {
            value_type: Some("boolean".to_string()),
            value_str: Some("false".to_string()),
            currency: None,
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("FALSE", None);
        match cell.value {
            CellValue::Boolean(b) => assert!(!b),
            _ => panic!("Expected Boolean false"),
        }

        // Test invalid boolean value (falls back to text)
        let builder = CellBuilder {
            value_type: Some("boolean".to_string()),
            value_str: Some("maybe".to_string()),
            currency: None,
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("maybe", None);
        match cell.value {
            CellValue::Text(t) => assert_eq!(t, "maybe"),
            _ => panic!("Expected Text for invalid boolean"),
        }
    }

    #[test]
    fn test_cell_builder_empty_text() {
        let builder = CellBuilder {
            value_type: None,
            value_str: None,
            currency: None,
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("   ", None);
        match cell.value {
            CellValue::Empty => {},
            _ => panic!("Expected Empty for whitespace-only text"),
        }
    }

    #[test]
    fn test_cell_builder_currency_default() {
        let builder = CellBuilder {
            value_type: Some("currency".to_string()),
            value_str: Some("50".to_string()),
            currency: None, // No currency specified
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("$50", None);
        match cell.value {
            CellValue::Currency(amount, currency) => {
                assert!((amount - 50.0).abs() < f64::EPSILON);
                assert_eq!(currency, "USD"); // Default
            },
            _ => panic!("Expected Currency with default USD"),
        }
    }

    #[test]
    fn test_parse_invalid_xml() {
        let invalid_xml = "<invalid>unclosed tag";
        let result = OdsParser::parse_sheets(invalid_xml);
        // The parser may return Ok with empty sheets or Err depending on implementation
        // Either behavior is acceptable - we just verify it doesn't panic
        match result {
            Ok(sheets) => {
                // If parsing succeeds, we should get 0 sheets
                assert_eq!(sheets.len(), 0);
            },
            Err(_) => {
                // Error is also acceptable
            },
        }
    }

    #[test]
    fn parses_global_and_sheet_local_named_definitions_with_namespace_aliases() {
        let xml = r#"<?xml version="1.0"?>
            <o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
                xmlns:f="urn:example:formula">
              <o:body><o:spreadsheet>
                <t:table t:name="Sales &amp; Tax">
                  <t:table-column/><t:table-row><t:table-cell/></t:table-row>
                  <t:named-expressions>
                    <t:named-range t:name="LocalRange"
                      t:cell-range-address="$'Sales &amp; Tax'.$A$1:.$B$2"
                      t:base-cell-address="$'Sales &amp; Tax'.$A$1"
                      t:range-usable-as="print-range filter repeat-row repeat-column"/>
                  </t:named-expressions>
                </t:table>
                <t:named-expressions>
                  <t:named-expression t:name="TaxRate"
                    t:expression="f:=0.2" t:base-cell-address="$'Sales &amp; Tax'.$A$1"/>
                </t:named-expressions>
              </o:spreadsheet></o:body>
            </o:document-content>"#;

        let definitions = OdsParser::parse_named_definitions(xml).unwrap();
        assert_eq!(definitions.len(), 2);
        let NamedDefinition::Range(range) = &definitions[0] else {
            panic!("expected named range");
        };
        assert_eq!(range.name, "LocalRange");
        assert_eq!(
            range.scope,
            NamedDefinitionScope::Sheet("Sales & Tax".to_string())
        );
        assert_eq!(range.usable_as.len(), 4);
        assert_eq!(
            range.base_cell_address.as_deref(),
            Some("$'Sales & Tax'.$A$1")
        );

        let NamedDefinition::Expression(expression) = &definitions[1] else {
            panic!("expected named expression");
        };
        assert_eq!(expression.name, "TaxRate");
        assert_eq!(expression.expression, "f:=0.2");
        assert_eq!(
            expression.formula_namespace.as_ref().unwrap().uri,
            "urn:example:formula"
        );
        assert_eq!(expression.scope, NamedDefinitionScope::Global);
    }

    #[test]
    fn named_definition_parser_rejects_missing_attributes_and_invalid_usage() {
        let missing_address = r#"<office:spreadsheet
            xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
            <table:named-expressions><table:named-range table:name="Broken"/>
            </table:named-expressions></office:spreadsheet>"#;
        assert!(OdsParser::parse_named_definitions(missing_address).is_err());

        let invalid_usage = r#"<office:spreadsheet
            xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
            <table:named-expressions><table:named-range table:name="Broken"
              table:cell-range-address="$Sheet1.$A$1" table:range-usable-as="chart"/>
            </table:named-expressions></office:spreadsheet>"#;
        assert!(OdsParser::parse_named_definitions(invalid_usage).is_err());
    }

    #[test]
    fn sheet_parser_ignores_dde_cache_tables() {
        let xml = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet>
            <t:dde-links><t:dde-link>
              <o:dde-source o:dde-application="soffice" o:dde-topic="topic" o:dde-item="item"/>
              <t:table t:name="Cached"><t:table-row><t:table-cell o:value-type="string"><text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">cached</text:p></t:table-cell></t:table-row></t:table>
            </t:dde-link></t:dde-links>
            <t:table t:name="Visible"><t:table-row><t:table-cell o:value-type="string"/></t:table-row></t:table>
            <t:table t:name="Empty"/>
          </o:spreadsheet></o:body>
        </o:document-content>"#;

        let sheets = OdsParser::parse_sheets(xml).unwrap();
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0].name, "Visible");
        assert_eq!(sheets[1].name, "Empty");
    }

    const HYPERLINK_DOCUMENT_PREFIX: &str = r#"<office:document-content
        xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
        xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
        xmlns:xlink="http://www.w3.org/1999/xlink"
        xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0">
      <office:body><office:spreadsheet>
        <table:table table:name="Links"><table:table-row>"#;
    const HYPERLINK_DOCUMENT_SUFFIX: &str = "</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>";

    fn hyperlink_document(cells: &str) -> String {
        format!("{HYPERLINK_DOCUMENT_PREFIX}{cells}{HYPERLINK_DOCUMENT_SUFFIX}")
    }

    #[test]
    fn parses_cell_hyperlinks_with_metadata_and_document_order() {
        let attributes = concat!(
            r#"xlink:href="https://example.com/" xlink:type="simple" "#,
            r#"office:name="Example" office:title="Example site" "#,
            r#"office:target-frame-name="_blank" text:style-name="Internet_20_link" "#,
            r#"xlink:show="new" xlink:actuate="onRequest" "#,
            r#"text:visited-style-name="Visited_20_Internet_20_Link""#,
        );
        let xml = hyperlink_document(&format!(
            concat!(
                r#"<table:table-cell office:value-type="string">"#,
                "<text:p>See <text:a {attributes}>the ",
                "<text:span>example</text:span> site</text:a> and ",
                r##"<text:a xlink:href="#Sheet2.B10">an internal target</text:a>.</text:p>"##,
                "</table:table-cell>",
            ),
            attributes = attributes,
        ));

        let sheets = OdsParser::parse_sheets(&xml).unwrap();
        let cell = &sheets[0].rows[0].cells[0];
        assert_eq!(cell.text, "See the example site and an internal target.");
        assert_eq!(cell.hyperlinks().len(), 2);

        let first = cell.hyperlink().unwrap();
        assert_eq!(first.href(), "https://example.com/");
        assert_eq!(first.text(), "the example site");
        assert_eq!(first.range(), 4..20);
        assert_eq!(first.name.as_deref(), Some("Example"));
        assert_eq!(first.title.as_deref(), Some("Example site"));
        assert_eq!(first.target_frame_name.as_deref(), Some("_blank"));
        assert_eq!(first.show, Some(TextHyperlinkShow::New));
        assert_eq!(first.actuate, Some(TextHyperlinkActuate::OnRequest));
        assert_eq!(first.style_name.as_deref(), Some("Internet_20_link"));
        assert_eq!(
            first.visited_style_name.as_deref(),
            Some("Visited_20_Internet_20_Link")
        );

        let second = &cell.hyperlinks()[1];
        assert_eq!(second.href, "#Sheet2.B10");
        assert_eq!(second.text, "an internal target");
        assert_eq!(second.range(), 25..43);
        assert!(second.name.is_none());
        assert!(second.target_frame_name.is_none());
        assert!(second.show.is_none());
        assert!(second.actuate.is_none());
    }

    #[test]
    fn parses_hyperlinks_with_namespace_aliases_and_repeated_cells() {
        let xml = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
            xmlns:tx="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="http://www.w3.org/1999/xlink">
          <o:body><o:spreadsheet><t:table t:name="Links"><t:table-row>
            <t:table-cell t:number-columns-repeated="2" o:value-type="string">
              <tx:p><tx:a x:href="mailto:someone@example.com">mail</tx:a></tx:p>
            </t:table-cell>
          </t:table-row></t:table></o:spreadsheet></o:body></o:document-content>"#;

        let sheets = OdsParser::parse_sheets(xml).unwrap();
        let row = &sheets[0].rows[0];
        assert_eq!(row.cells.len(), 2);
        for cell in &row.cells {
            assert!(cell.has_hyperlinks());
            let link = cell.hyperlink().unwrap();
            assert_eq!(link.href(), "mailto:someone@example.com");
            assert_eq!(link.text(), "mail");
        }
    }

    #[test]
    fn parses_self_closing_hyperlink_with_empty_text() {
        let xml = hyperlink_document(
            r#"<table:table-cell office:value-type="string">
              <text:p>before <text:a xlink:href="https://example.com/"/> after</text:p>
            </table:table-cell>"#,
        );

        let sheets = OdsParser::parse_sheets(&xml).unwrap();
        let cell = &sheets[0].rows[0].cells[0];
        assert_eq!(cell.text, "before  after");
        assert_eq!(cell.hyperlinks().len(), 1);
        assert_eq!(cell.hyperlinks()[0].href, "https://example.com/");
        assert_eq!(cell.hyperlinks()[0].text, "");
        assert_eq!(cell.hyperlinks()[0].range(), 7..7);
    }

    #[test]
    fn preserves_mixed_text_anchor_range_from_libreoffice_fods() {
        let source = include_str!(
            "../../../../test-data/libreoffice-core/sc/qa/unit/data/functions/text/fods/encodeurl.fods"
        );
        let anchor = r#"<text:a xlink:href="http://www.test/libreOffice" xlink:type="simple">"#;
        let anchor_start = source.find(anchor).unwrap();
        let cell_start = source[..anchor_start].rfind("<table:table-cell").unwrap();
        let cell_end = anchor_start
            + source[anchor_start..].find("</table:table-cell>").unwrap()
            + "</table:table-cell>".len();
        let sheets =
            OdsParser::parse_sheets(&hyperlink_document(&source[cell_start..cell_end])).unwrap();
        let cell = sheets
            .iter()
            .flat_map(|sheet| sheet.rows.iter())
            .flat_map(|row| row.cells.iter())
            .find(|cell| {
                cell.hyperlinks()
                    .iter()
                    .any(|link| link.href() == "http://www.test/libreOffice")
            })
            .unwrap();
        let link = cell
            .hyperlinks()
            .iter()
            .find(|link| link.href() == "http://www.test/libreOffice")
            .unwrap();

        assert_eq!(link.range(), 0..link.text().len());
        assert!(cell.text.starts_with(link.text()));
        assert!(cell.text.ends_with("agJohn01Czech Republic"));
    }

    #[test]
    fn hyperlink_text_includes_whitespace_and_break_elements() {
        let xml = hyperlink_document(
            r#"<table:table-cell office:value-type="string">
              <text:p><text:a xlink:href="https://example.com/">a<text:s text:c="2"/>b<text:line-break/>c</text:a></text:p>
            </table:table-cell>"#,
        );

        let sheets = OdsParser::parse_sheets(&xml).unwrap();
        let cell = &sheets[0].rows[0].cells[0];
        assert_eq!(cell.hyperlinks()[0].text, "a  b\nc");
    }

    #[test]
    fn rejects_hyperlink_without_href() {
        let xml = hyperlink_document(
            r#"<table:table-cell office:value-type="string">
              <text:p><text:a office:name="broken">no target</text:a></text:p>
            </table:table-cell>"#,
        );

        let error = OdsParser::parse_sheets(&xml)
            .err()
            .expect("parse must fail");
        assert!(error.to_string().contains("xlink:href"));
    }

    #[test]
    fn rejects_nested_hyperlinks() {
        let xml = hyperlink_document(
            r#"<table:table-cell office:value-type="string">
              <text:p><text:a xlink:href="https://a.example/">outer
                <text:a xlink:href="https://b.example/">inner</text:a></text:a></text:p>
            </table:table-cell>"#,
        );

        let error = OdsParser::parse_sheets(&xml)
            .err()
            .expect("parse must fail");
        assert!(error.to_string().contains("nested"));
    }

    #[test]
    fn rejects_cell_rich_text_beyond_the_depth_limit() {
        let nested = format!(
            "{}x{}",
            "<text:span>".repeat(128),
            "</text:span>".repeat(128)
        );
        let xml = hyperlink_document(&format!(
            r#"<table:table-cell office:value-type="string"><text:p>{nested}</text:p></table:table-cell>"#
        ));

        let error = OdsParser::parse_sheets(&xml)
            .err()
            .expect("overly deep rich text must fail");
        assert!(error.to_string().contains("depth limit"));
    }

    #[test]
    fn rejects_hyperlink_with_invalid_xlink_type() {
        let xml = hyperlink_document(
            r#"<table:table-cell office:value-type="string">
              <text:p><text:a xlink:href="https://example.com/" xlink:type="extended">x</text:a></text:p>
            </table:table-cell>"#,
        );

        let error = OdsParser::parse_sheets(&xml)
            .err()
            .expect("parse must fail");
        assert!(error.to_string().contains("xlink:type"));
    }

    #[test]
    fn rejects_hyperlink_with_invalid_xlink_show_or_actuate() {
        for attributes in [r#"xlink:show="embed""#, r#"xlink:actuate="onLoad""#] {
            let xml = hyperlink_document(&format!(
                r#"<table:table-cell office:value-type="string"><text:p><text:a xlink:href="https://example.com/" {attributes}>x</text:a></text:p></table:table-cell>"#
            ));
            assert!(OdsParser::parse_sheets(&xml).is_err());
        }
    }

    #[test]
    fn annotation_hyperlinks_are_not_reported_as_cell_hyperlinks() {
        let xml = hyperlink_document(
            r#"<table:table-cell office:value-type="string">
              <office:annotation><text:p><text:a xlink:href="https://note.example/">note link</text:a></text:p></office:annotation>
              <text:p>plain</text:p>
            </table:table-cell>"#,
        );

        let sheets = OdsParser::parse_sheets(&xml).unwrap();
        let cell = &sheets[0].rows[0].cells[0];
        assert!(cell.annotation().is_some());
        assert!(!cell.has_hyperlinks());
        assert_eq!(cell.text, "plain");
    }
}
