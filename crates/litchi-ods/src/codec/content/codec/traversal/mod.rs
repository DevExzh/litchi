//! Namespace-aware streaming traversal of ODS `content.xml`.

use super::super::model::{CellBuilder, RowBuilder, SheetBuilder};

use super::super::super::{
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

mod model;
mod semantic;
mod validation;

/// Parser for ODS-specific structures.
///
/// This provides parsing logic specific to spreadsheets,
/// including sheet, row, and cell parsing with proper type detection.
pub(crate) struct Parser;
