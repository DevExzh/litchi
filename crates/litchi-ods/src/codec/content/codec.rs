//! Streaming XML codec for ODS `content.xml`.

use super::model::{CellBuilder, RowBuilder, SheetBuilder};

use super::super::{
    Cell, CellDetective, CellMatrixSpan, CellMerge, CellRangeSource, CellTextContent, CellValue,
    ColorTransformationType, Column, ConditionalColorScale, ConditionalColorScaleEntry,
    ConditionalCustomIcon, ConditionalDataBar, ConditionalDataBarEntry, ConditionalDateIs,
    ConditionalDateType, ConditionalFormat, ConditionalFormatCondition, ConditionalFormatEntryType,
    ConditionalFormatRule, ConditionalIconSet, ConditionalIconSetEntry, DataBarAxisPosition,
    DetectiveDirection, DetectiveHighlightedRange, DetectiveOperation, DetectiveOperationKind,
    IconSetType, NamedDefinition, NamedDefinitionScope, NamedExpression, NamedRange,
    NamedRangeUsage, Row, Sheet, SheetPrintSettings, SheetScenario, SheetStyle, SheetTableSource,
    Sparkline, SparklineAxisType, SparklineColorTransformation, SparklineComplexColor,
    SparklineEmptyCells, SparklineGroup, SparklineType, TableGroup, TableRange, TableSourceMode,
    TableStructure, TableVisibility, ThemeColorType,
    annotation::{Builder, decode_reference},
    conditional_format::{
        CALCEXT_NAMESPACE_URI, DATA_BAR_ENTRY_COUNT, MAX_CONDITIONAL_FORMATS_PER_SHEET,
        MAX_ENTRIES_PER_RULE, MAX_RULES_PER_FORMAT, validate_color_scale_entry, validate_condition,
        validate_conditional_format, validate_data_bar_attributes, validate_data_bar_entry,
        validate_date_is, validate_icon_set_entry, validate_rule,
    },
    dde::parse_source as parse_dde_source,
    rich_text::CellTextContentBuilder,
    scenario::validate_scenario,
    source::validate_table_source,
    sparkline::{
        COMPLEX_COLOR_SLOTS, LOEXT_NAMESPACE_URI, MAX_COLOR_TRANSFORMATIONS,
        MAX_SPARKLINE_GROUPS_PER_SHEET, MAX_SPARKLINES_PER_GROUP, validate_sparkline,
        validate_sparkline_group, validate_sparkline_group_attributes,
    },
    structure::{
        MAX_EXPANDED_COLUMNS_PER_SHEET, MAX_EXPANDED_ROWS_PER_SHEET, MAX_TABLE_STRUCTURE_DEPTH,
        split_cell_range_addresses,
    },
};
use crate::elements::text::{TextHyperlinkActuate, TextHyperlinkShow};
use crate::model::hyperlink::Link;
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

/// A `text:a` hyperlink whose text content is still being collected.
struct PendingHyperlink {
    /// The hyperlink parsed from the element's attributes.
    link: Link,
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
            Self::ColorScale { depth, .. }
            | Self::DataBar { depth, .. }
            | Self::IconSet { depth, .. } => *depth,
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

/// A `calcext:sparkline-*-complex-color` element whose `loext:transformation`
/// children are still being read.
struct PendingSparklineComplexColor {
    /// The slot element name (one of `COMPLEX_COLOR_SLOTS`).
    slot: &'static str,
    /// The color parsed from the element's attributes (no transformations yet).
    color: SparklineComplexColor,
    /// The `element_depth` value assigned to the element.
    depth: usize,
}

/// Parser for ODS-specific structures.
///
/// This provides parsing logic specific to spreadsheets,
/// including sheet, row, and cell parsing with proper type detection.
pub(crate) struct Parser;

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

impl Parser {
    /// Parse all sheets from ODS content.xml
    // quick-xml exposes a streaming event source, so the format's nested parser
    // state is intentionally coordinated here without constructing a DOM.
    #[allow(clippy::cognitive_complexity)]
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
        let mut annotation_builder: Option<Builder> = None;
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
        let mut pending_sparkline_complex_color: Option<PendingSparklineComplexColor> = None;
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
                                if is_calcext("formatting-entry")
                                    || is_calcext("data-bar-entry") =>
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
                            PendingCalcextRule::IconSet { icon_set, .. }
                                if is_calcext("custom-iconset") =>
                            {
                                if icon_set.custom_icons.len() >= MAX_ENTRIES_PER_RULE {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:icon-set exceeds the {MAX_ENTRIES_PER_RULE} custom icon safety limit"
                                    )));
                                }
                                icon_set.custom_icons.push(Self::parse_custom_icon(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                                calcext_leaf_open_depth = Some(element_depth);
                            },
                            _ => {
                                // Unmodeled content is skipped.
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
                            let date_is =
                                Self::parse_date_is(e, reader.decoder(), &document_namespaces)?;
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
                    } else if let Some(pending) = pending_sparkline_complex_color.as_mut() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            LOEXT_NAMESPACE_URI,
                            "transformation",
                        ) {
                            if pending.color.transformations.len() >= MAX_COLOR_TRANSFORMATIONS {
                                return Err(Error::InvalidFormat(format!(
                                    "complex color exceeds the {MAX_COLOR_TRANSFORMATIONS} transformation safety limit"
                                )));
                            }
                            pending
                                .color
                                .transformations
                                .push(Self::parse_color_transformation(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
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
                        } else if let Some(slot) = COMPLEX_COLOR_SLOTS.iter().find(|slot| {
                            Self::element_name_is(
                                e.name().as_ref(),
                                &document_namespaces,
                                CALCEXT_NAMESPACE_URI,
                                slot,
                            )
                        }) {
                            let color = Self::parse_complex_color(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            pending_sparkline_complex_color = Some(PendingSparklineComplexColor {
                                slot,
                                color,
                                depth: element_depth,
                            });
                        } else {
                            // Unmodeled children are skipped.
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
                        annotation_builder = Some(Builder::new(
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
                                if is_calcext("formatting-entry")
                                    || is_calcext("data-bar-entry") =>
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
                            PendingCalcextRule::IconSet { icon_set, .. }
                                if is_calcext("custom-iconset") =>
                            {
                                if icon_set.custom_icons.len() >= MAX_ENTRIES_PER_RULE {
                                    return Err(Error::InvalidFormat(format!(
                                        "calcext:icon-set exceeds the {MAX_ENTRIES_PER_RULE} custom icon safety limit"
                                    )));
                                }
                                icon_set.custom_icons.push(Self::parse_custom_icon(
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
                            let date_is =
                                Self::parse_date_is(e, reader.decoder(), &document_namespaces)?;
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
                    if let Some(pending) = pending_sparkline_complex_color.as_mut() {
                        if Self::element_name_is(
                            e.name().as_ref(),
                            &document_namespaces,
                            LOEXT_NAMESPACE_URI,
                            "transformation",
                        ) {
                            if pending.color.transformations.len() >= MAX_COLOR_TRANSFORMATIONS {
                                return Err(Error::InvalidFormat(format!(
                                    "complex color exceeds the {MAX_COLOR_TRANSFORMATIONS} transformation safety limit"
                                )));
                            }
                            pending
                                .color
                                .transformations
                                .push(Self::parse_color_transformation(
                                    e,
                                    reader.decoder(),
                                    &document_namespaces,
                                )?);
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, Some(empty_scope));
                        buf.clear();
                        continue;
                    }
                    if let Some(pending) = pending_sparkline_group.as_mut() {
                        // An empty complex-color element carries no
                        // transformations; an empty `calcext:sparklines`
                        // container and unknown children carry nothing.
                        if let Some(slot) = COMPLEX_COLOR_SLOTS.iter().find(|slot| {
                            Self::element_name_is(
                                e.name().as_ref(),
                                &document_namespaces,
                                CALCEXT_NAMESPACE_URI,
                                slot,
                            )
                        }) {
                            let color = Self::parse_complex_color(
                                e,
                                reader.decoder(),
                                &document_namespaces,
                            )?;
                            pending.group.complex_colors.assign_slot(slot, color)?;
                        }
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
                        let annotation =
                            Builder::new(e, reader.decoder(), document_namespaces.clone())?
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
                    let closes_conditional_formats = conditional_formats_depth
                        == Some(element_depth)
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
                    let closes_sparkline_complex_color = pending_sparkline_complex_color
                        .as_ref()
                        .is_some_and(|pending| {
                            pending.depth == element_depth
                                && Self::element_name_is(
                                    e.name().as_ref(),
                                    &document_namespaces,
                                    CALCEXT_NAMESPACE_URI,
                                    pending.slot,
                                )
                        });
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
                    if closes_sparkline_complex_color {
                        let pending = pending_sparkline_complex_color
                            .take()
                            .expect("pending complex color was checked");
                        let group = pending_sparkline_group.as_mut().ok_or_else(|| {
                            Error::InvalidFormat(
                                "complex color closed outside calcext:sparkline-group".to_string(),
                            )
                        })?;
                        group
                            .group
                            .complex_colors
                            .assign_slot(pending.slot, pending.color)?;
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
            super::super::sheet_image::validate_sheet_image(&image)?;
            sheets[index].images.push(image);
        }
        for sheet in &sheets {
            super::super::sheet_image::validate_sheet_images(&sheet.images)?;
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
                if let Some(sheet_shape) = super::super::shape::sheet_shape_from_parsed(shape)? {
                    sheet.shapes.push(sheet_shape);
                }
            }
            super::super::shape::validate_sheet_shapes(&sheet.shapes)?;
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
    ) -> Result<Link> {
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
        Ok(Link {
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
                Error::InvalidFormat("calcext:condition requires calcext:value".to_string())
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
                value = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:value",
                )?);
            } else if is_calcext("color") {
                color = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:color",
                )?);
            }
        }
        let entry = ConditionalColorScaleEntry {
            entry_type: entry_type.ok_or_else(|| {
                Error::InvalidFormat("calcext:color-scale-entry requires calcext:type".to_string())
            })?,
            value: value.ok_or_else(|| {
                Error::InvalidFormat("calcext:color-scale-entry requires calcext:value".to_string())
            })?,
            color: color.ok_or_else(|| {
                Error::InvalidFormat("calcext:color-scale-entry requires calcext:color".to_string())
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
                value = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:value",
                )?);
            }
        }
        Ok((
            entry_type.ok_or_else(|| {
                Error::InvalidFormat("calcext:formatting-entry requires calcext:type".to_string())
            })?,
            value.ok_or_else(|| {
                Error::InvalidFormat("calcext:formatting-entry requires calcext:value".to_string())
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
                data_bar.axis_position = Some(DataBarAxisPosition::parse(
                    &Self::decode_attribute(&attribute, decoder, "calcext:axis-position")?,
                )?);
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

    /// Parse one inert `calcext:custom-iconset` assignment from its attributes.
    fn parse_custom_icon(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<ConditionalCustomIcon> {
        let mut icon_set_type = None;
        let mut index = None;
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
            if is_calcext("custom-iconset-name") {
                icon_set_type = Some(IconSetType::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:custom-iconset-name",
                )?)?);
            } else if is_calcext("custom-iconset-index") {
                let lexical =
                    Self::decode_attribute(&attribute, decoder, "calcext:custom-iconset-index")?;
                index = Some(lexical.parse::<u32>().map_err(|_| {
                    Error::InvalidFormat(format!(
                        "calcext:custom-iconset-index requires a non-negative integer, found '{lexical}'"
                    ))
                })?);
            }
        }
        Ok(ConditionalCustomIcon {
            icon_set_type: icon_set_type.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:custom-iconset requires calcext:custom-iconset-name".to_string(),
                )
            })?,
            index: index.ok_or_else(|| {
                Error::InvalidFormat(
                    "calcext:custom-iconset requires calcext:custom-iconset-index".to_string(),
                )
            })?,
        })
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
                Error::InvalidFormat("calcext:icon-set requires calcext:icon-set-type".to_string())
            })?,
            show_value,
            custom,
            custom_icons: Vec::new(),
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
                style = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "calcext:style",
                )?);
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
                Error::InvalidFormat("calcext:sparkline requires calcext:cell-address".to_string())
            })?,
            data_ranges: data_ranges.ok_or_else(|| {
                Error::InvalidFormat("calcext:sparkline requires calcext:data-range".to_string())
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
                group.flags.display_x_axis = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("display-hidden") {
                group.flags.display_hidden = Some(Self::parse_bool_attribute(&attribute, decoder)?);
            } else if is_calcext("min-axis-type") {
                group.min_axis_type =
                    Some(SparklineAxisType::parse(&decode("calcext:min-axis-type")?)?);
            } else if is_calcext("max-axis-type") {
                group.max_axis_type =
                    Some(SparklineAxisType::parse(&decode("calcext:max-axis-type")?)?);
            } else if is_calcext("right-to-left") {
                group.flags.right_to_left = Some(Self::parse_bool_attribute(&attribute, decoder)?);
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

    /// Parse one inert `calcext:sparkline-*-complex-color` from its attributes.
    fn parse_complex_color(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<SparklineComplexColor> {
        let mut theme_type = None;
        let mut color_type = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_loext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    LOEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_loext("theme-type") {
                theme_type = Some(ThemeColorType::parse(&Self::decode_attribute(
                    &attribute,
                    decoder,
                    "loext:theme-type",
                )?)?);
            } else if is_loext("color-type") {
                color_type = Some(Self::decode_attribute(
                    &attribute,
                    decoder,
                    "loext:color-type",
                )?);
            }
        }
        if let Some(color_type) = &color_type
            && color_type != "theme"
        {
            return Err(Error::InvalidFormat(format!(
                "unsupported loext:color-type value '{color_type}'"
            )));
        }
        Ok(SparklineComplexColor {
            theme_type: theme_type.ok_or_else(|| {
                Error::InvalidFormat("calcext complex color requires loext:theme-type".to_string())
            })?,
            transformations: Vec::new(),
        })
    }

    /// Parse one inert `loext:transformation` from its attributes.
    fn parse_color_transformation(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<SparklineColorTransformation> {
        let mut transformation_type = None;
        let mut value = None;
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let is_loext = |local_name| {
                Self::attribute_name_is(
                    attribute.key.as_ref(),
                    namespaces,
                    LOEXT_NAMESPACE_URI,
                    local_name,
                )
            };
            if is_loext("type") {
                transformation_type = Some(ColorTransformationType::parse(
                    &Self::decode_attribute(&attribute, decoder, "loext:type")?,
                )?);
            } else if is_loext("value") {
                let lexical = Self::decode_attribute(&attribute, decoder, "loext:value")?;
                let parsed: i32 = lexical.parse().map_err(|_| {
                    Error::InvalidFormat(format!(
                        "loext:value requires an integer value, found '{lexical}'"
                    ))
                })?;
                value = Some(i16::try_from(parsed).map_err(|_| {
                    Error::InvalidFormat(format!(
                        "loext:value '{lexical}' is outside the supported range"
                    ))
                })?);
            }
        }
        Ok(SparklineColorTransformation {
            transformation_type: transformation_type.ok_or_else(|| {
                Error::InvalidFormat("loext:transformation requires loext:type".to_string())
            })?,
            value: value.ok_or_else(|| {
                Error::InvalidFormat("loext:transformation requires loext:value".to_string())
            })?,
        })
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

impl RowBuilder {
    fn from_element(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &BTreeMap<String, String>,
    ) -> Result<Self> {
        let repeated =
            Parser::parse_repeated(element, decoder, namespaces, "number-rows-repeated")?;
        let (style_name, default_cell_style_name, visibility) =
            Parser::parse_structural_attributes(element, decoder, namespaces)?;
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
}
