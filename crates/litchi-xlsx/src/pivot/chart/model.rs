//! Semantic pivot-chart objects and drop-zone vocabulary.

use crate::error::{Error, Result};
use crate::pivot::PivotTable;
use std::str::FromStr;

use super::invalid;

/// `PivotTable` field-type zone addressed by a pivot-chart drop-zone switch.
///
/// The variants parse from the ECMA-376 axis identifiers (`axisRow`,
/// `axisCol`, `axisPage`, `axisValues`, `dataFields`) and map to the
/// `c14:dropZone*` element names used by the series pivot-options extension.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FieldType {
    /// Row axis fields (`c14:dropZoneCategories`)
    AxisRow,
    /// Column axis fields (`c14:dropZoneSeries`)
    AxisCol,
    /// Page (report filter) fields (`c14:dropZoneAxis`)
    AxisPage,
    /// Value axis fields (`c14:dropZoneValues`)
    AxisValues,
    /// Data fields (`c14:dropZoneData`)
    DataFields,
}

impl FieldType {
    /// ECMA-376 axis identifier for this field type.
    #[must_use]
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::AxisRow => "axisRow",
            Self::AxisCol => "axisCol",
            Self::AxisPage => "axisPage",
            Self::AxisValues => "axisValues",
            Self::DataFields => "dataFields",
        }
    }

    /// Map a `c14:dropZone*` element name to its field type.
    pub(super) fn from_drop_zone_element(local_name: &[u8]) -> Option<Self> {
        match local_name {
            b"dropZoneCategories" => Some(Self::AxisRow),
            b"dropZoneSeries" => Some(Self::AxisCol),
            b"dropZoneAxis" => Some(Self::AxisPage),
            b"dropZoneValues" => Some(Self::AxisValues),
            b"dropZoneData" => Some(Self::DataFields),
            _ => None,
        }
    }

    /// The `c14:dropZone*` element name for this field type.
    #[must_use]
    pub fn drop_zone_element_name(&self) -> &'static str {
        match self {
            Self::AxisRow => "dropZoneCategories",
            Self::AxisCol => "dropZoneSeries",
            Self::AxisPage => "dropZoneAxis",
            Self::AxisValues => "dropZoneValues",
            Self::DataFields => "dropZoneData",
        }
    }
}

impl FromStr for FieldType {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        Ok(match value {
            "axisRow" => Self::AxisRow,
            "axisCol" => Self::AxisCol,
            "axisPage" => Self::AxisPage,
            "axisValues" => Self::AxisValues,
            "dataFields" => Self::DataFields,
            _ => {
                return Err(invalid(format!("unknown pivot-chart field type '{value}'")));
            },
        })
    }
}

/// Visibility of one drop zone in a pivot chart.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DropZoneVisibility {
    /// `PivotTable` field type occupying the drop zone
    pub field_type: FieldType,
    /// Whether the drop zone's field buttons are visible
    pub visible: bool,
}

/// Drop-zone metadata parsed from one series' `c14:pivotOptions` extension.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Options {
    /// `c14:dropZoneVisible` master switch; `None` when omitted
    pub drop_zone_visible: Option<bool>,
    /// Per field-type drop-zone switches, in document order
    pub drop_zones: Vec<DropZoneVisibility>,
}

impl Options {
    /// Visibility recorded for one field-type zone, if present.
    #[must_use]
    pub fn visibility(&self, field_type: FieldType) -> Option<bool> {
        self.drop_zones
            .iter()
            .find(|zone| zone.field_type == field_type)
            .map(|zone| zone.visible)
    }

    /// Drop-zone defaults with every field button visible, matching what
    /// Excel writes for newly created pivot charts.
    #[must_use]
    pub fn all_visible() -> Self {
        // Emission order follows the CT_PivotOptions element sequence.
        const DEFAULT_ORDER: [FieldType; 5] = [
            FieldType::AxisRow,
            FieldType::DataFields,
            FieldType::AxisCol,
            FieldType::AxisPage,
            FieldType::AxisValues,
        ];
        Self {
            drop_zone_visible: Some(true),
            drop_zones: DEFAULT_ORDER
                .iter()
                .map(|&field_type| DropZoneVisibility {
                    field_type,
                    visible: true,
                })
                .collect(),
        }
    }
}

/// Typed `c:pivotSource` metadata binding a chart to a pivot table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Source {
    /// Pivot-table reference as written in `c:name`, optionally qualified
    /// with a `[workbook]` prefix and a sheet name (for example
    /// `[Book1.xlsx]Sheet1!PivotTable1`)
    pub name: String,
    /// Pivot format identifier from `c:fmtId`
    pub format_id: u32,
    /// Extension URIs recorded under `c:pivotSource/c:extLst`, retained
    /// inertly in document order
    pub extension_uris: Vec<String>,
}

/// Pivot metadata for one chart series.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Series {
    /// Series index from the series' `c:idx` element
    pub index: u32,
    /// Drop-zone options parsed from the series' `c14:pivotOptions`
    /// extension; `None` when the series carries no such extension
    pub options: Option<Options>,
}

/// Unresolved pivot-chart payload of one chart part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Binding {
    /// Pivot-table source metadata
    pub source: Source,
    /// Per-series pivot options, keyed by series index as written
    pub series: Vec<Series>,
}

/// A pivot chart anchored on a worksheet, with its pivot-table binding
/// resolved against the workbook's typed pivot-table models.
#[derive(Debug, Clone)]
pub struct Chart {
    /// Relationship ID from the worksheet drawing part to the chart part
    pub relationship_id: String,
    /// Chart part name (for example `/xl/charts/chart1.xml`)
    pub part_name: String,
    /// Pivot-table source metadata from the chart part
    pub source: Source,
    /// Per-series pivot options parsed from chart extensions
    pub series: Vec<Series>,
    /// Resolved typed pivot-table model named by `source`
    pub pivot_table: PivotTable,
}

/// Kind of sheet hosting one or more pivot charts.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SheetKind {
    /// An ordinary worksheet anchoring charts through its drawing part
    Worksheet,
    /// A chartsheet whose drawing part holds its chart directly
    Chartsheet,
}

/// Pivot charts anchored on one worksheet or chartsheet.
#[derive(Debug, Clone)]
pub struct SheetCharts {
    /// Sheet name from the workbook
    pub sheet_name: String,
    /// Sheet part name (for example `/xl/worksheets/sheet1.xml`)
    pub sheet_part_name: String,
    /// Kind of the hosting sheet
    pub sheet_kind: SheetKind,
    /// Pivot charts anchored on the sheet, in drawing order
    pub charts: Vec<Chart>,
}
