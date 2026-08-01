//! Planned native chart authoring model for PPT slides.
//!
//! A native chart is placed on a slide as an OfficeArt picture-frame shape
//! (MSOSPT 75) whose `ClientData` record carries an `ExObjRefAtom` pointing
//! at an `ExOleEmbedContainer` in the document's `ExObjList` ([MS-PPT] 2.10).
//! The container declares the `Excel.Chart.8` ProgID and references a
//! persisted `ExOleObjStg` record whose payload is a BIFF8 chart workbook.
//!
//! Everything stays inert: chart data links are never evaluated, no external
//! workbook is opened, and the embedded payload is never activated. Public
//! chart creation currently returns
//! [`litchi_ograph::Error::UnsupportedAuthoring`] through
//! [`PptWriteError::Graph`] before mutating the presentation; the complete
//! Office-compatible BIFF chart grammar must land before emission is enabled.
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi_ole::ppt::writer::{Chart, ChartKind, PptWriter};
//!
//! let mut writer = PptWriter::new();
//! let slide = writer.add_slide()?;
//!
//! let mut chart = Chart::new(ChartKind::Bar);
//! chart.set_title("Quarterly sales");
//! chart.set_categories(["Q1", "Q2", "Q3", "Q4"]);
//! chart.add_series(Some("2024"), vec![1.5, 2.0, 2.5, 3.0])?;
//! let error = writer
//!     .add_chart(slide, 50, 50, 400, 300, chart)
//!     .expect_err("binary chart authoring is not enabled yet");
//! assert!(matches!(
//!     error,
//!     litchi_ole::ppt::writer::PptWriteError::Graph(
//!         litchi_ograph::Error::UnsupportedAuthoring { .. }
//!     )
//! ));
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use litchi_core::unit::emu_i32_to_ppt_master_i16_round;
use zerocopy::IntoBytes;

use super::core::PptWriteError;
use super::escher::{
    EscherBuilder, EscherSpData, PptError, ShapeFlags, header_version, record_type,
};
use super::hyperlink::{HyperlinkCollection, record_type as hyperlink_record_type};
use super::records::RecordBuilder;
use crate::ppt::ole_object::{
    PowerPointOleColorFollow, PowerPointOleContainerKind, PowerPointOleDimensionPolicy,
    PowerPointOleDrawAspect, PowerPointOleEmbedPreferences, PowerPointOleObjectDefinition,
    PowerPointOleObjectMetadata, PowerPointOleObjectSubtype, PowerPointOleObjectType,
};
use crate::ppt::ole_storage::{
    PowerPointOleStorage, PowerPointOleStorageCompression, PowerPointOleStorageKind,
};
#[cfg(test)]
use crate::xls::chart::{
    Cache, CellRef, Chart as Model, DataKind, DataLink, Editor, Group, GroupKind,
    Kind as ModelKind, Limits, Location, Role, Series, Source, Value, build_workbook_fixture,
};

/// MSOSPT value of the OfficeArt frame used for OLE objects ([MS-ODRAW]).
const MSOSPT_PICTURE_FRAME: u16 = 75;
/// `RT_ExternalObjectRefAtom` ([MS-PPT] 2.13).
const EX_OBJ_REF_ATOM: u16 = 3009;
/// ProgID declared for authored chart objects.
const EXCEL_CHART_PROG_ID: &str = "Excel.Chart.8";
/// Maximum categories or values per series (BIFF8 `SERIES` count bound).
const MAX_DATA_POINT_COUNT: usize = 32_767;
/// Maximum generated value columns after the shared category column.
const MAX_SERIES_COUNT: usize = 255;
/// Maximum UTF-16 code units in a category label or title (BIFF8 short string).
const MAX_STRING_UNITS: usize = 255;
/// Default gap width between bar clusters, in percent of bar width.
#[cfg(test)]
const BAR_GAP_WIDTH_PERCENT: u16 = 150;
/// ptgArea token with reference class (BIFF8 chart formula).
#[cfg(test)]
const PTG_AREA: u8 = 0x3b;
/// ExternSheet index used by generated data links (the single local sheet).
#[cfg(test)]
const LOCAL_SHEET_INDEX: u16 = 0;
/// Worksheet column holding the generated category range.
#[cfg(test)]
const CATEGORY_COLUMN: u16 = 0;
/// First worksheet row of generated data ranges (row 0 holds series names).
#[cfg(test)]
const FIRST_DATA_ROW: u16 = 1;

/// Maximum number of chart objects authored into one presentation.
///
/// Matches the read-side inventory bound in [`crate::ppt::chart`].
pub(crate) const MAX_CHART_OBJECTS: usize = 512;

fn invalid(message: impl Into<String>) -> PptWriteError {
    PptWriteError::InvalidData(message.into())
}

/// Requested native-chart family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ChartKind {
    /// Clustered column chart (BIFF `Bar` group).
    #[default]
    Bar,
    /// Line chart (BIFF `Line` group).
    Line,
    /// Pie chart (BIFF `Pie` group).
    Pie,
}

/// One requested data series: an optional name and its values.
#[derive(Debug)]
pub struct ChartSeries {
    /// Series name shown in the legend.
    pub name: Option<String>,
    /// Data points, one per category.
    pub values: Vec<f64>,
}

/// A native-chart request: type, title, shared categories, and data series.
#[derive(Debug)]
pub struct Chart {
    kind: ChartKind,
    title: Option<String>,
    categories: Vec<String>,
    series: Vec<ChartSeries>,
}

impl Chart {
    /// Create an empty chart of the given kind.
    pub fn new(kind: ChartKind) -> Self {
        Self {
            kind,
            title: None,
            categories: Vec::new(),
            series: Vec::new(),
        }
    }

    /// Returns this chart's requested chart kind.
    pub fn kind(&self) -> ChartKind {
        self.kind
    }

    /// Set the chart title.
    pub fn set_title(&mut self, title: impl Into<String>) {
        self.title = Some(title.into());
    }

    /// Set the category labels shared by every series.
    pub fn set_categories<I, S>(&mut self, categories: I)
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        self.categories = categories.into_iter().map(Into::into).collect();
    }

    /// Append a data series. Values must be finite and within BIFF8 bounds.
    pub fn add_series(
        &mut self,
        name: Option<impl Into<String>>,
        values: Vec<f64>,
    ) -> Result<(), PptWriteError> {
        if values.is_empty() {
            return Err(invalid("chart series must contain at least one value"));
        }
        if values.len() > MAX_DATA_POINT_COUNT {
            return Err(invalid(format!(
                "chart series exceeds {MAX_DATA_POINT_COUNT} values"
            )));
        }
        if values.iter().any(|value| !value.is_finite()) {
            return Err(invalid("chart series values must be finite"));
        }
        let name = name.map(Into::into);
        if let Some(name) = &name {
            check_string_units(name, "chart series name")?;
        }
        self.series.push(ChartSeries { name, values });
        Ok(())
    }

    pub(super) fn validate(&self) -> Result<(), PptWriteError> {
        if self.series.is_empty() {
            return Err(invalid("chart must contain at least one series"));
        }
        if self.series.len() > MAX_SERIES_COUNT {
            return Err(invalid(format!("chart exceeds {MAX_SERIES_COUNT} series")));
        }
        if self.categories.len() > MAX_DATA_POINT_COUNT {
            return Err(invalid(format!(
                "chart exceeds {MAX_DATA_POINT_COUNT} categories"
            )));
        }
        for category in &self.categories {
            check_string_units(category, "chart category")?;
        }
        if let Some(title) = &self.title {
            check_string_units(title, "chart title")?;
        }
        Ok(())
    }

    /// Converts a validated chart into the abbreviated parser fixture model.
    #[cfg(test)]
    fn model(self) -> Result<Model, PptWriteError> {
        self.validate()?;
        let Self {
            kind,
            title,
            categories,
            series,
        } = self;

        let group_kind = match kind {
            ChartKind::Bar => GroupKind::Bar {
                overlap: 0,
                gap: BAR_GAP_WIDTH_PERCENT,
                flags: 0,
            },
            ChartKind::Line => GroupKind::Line { flags: 0 },
            ChartKind::Pie => GroupKind::Pie {
                rotation: 0,
                hole_size: 0,
                flags: 0,
            },
        };
        let mut chart = Model {
            title,
            groups: vec![Group {
                order: 0,
                vary_colors: matches!(kind, ChartKind::Pie),
                kind: group_kind,
                lines: Vec::new(),
                drop_bars: Vec::new(),
            }],
            ..Default::default()
        };
        let category_count = checked_point_count(categories.len())?;
        let category_column = u8::try_from(CATEGORY_COLUMN)
            .map_err(|_| invalid("category column exceeds BIFF8 grid"))?;
        for (index, series) in series.into_iter().enumerate() {
            let value_count = checked_point_count(series.values.len())?;
            let value_column = u16::try_from(index)
                .ok()
                .and_then(|index| index.checked_add(1))
                .and_then(|index| CATEGORY_COLUMN.checked_add(index))
                .ok_or_else(|| invalid("chart series count exceeds worksheet columns"))?;
            let cache_column = u8::try_from(value_column)
                .map_err(|_| invalid("chart series column exceeds BIFF8 grid"))?;
            let last_category_row = FIRST_DATA_ROW + category_count.saturating_sub(1);
            let last_value_row = FIRST_DATA_ROW + value_count.saturating_sub(1);
            let categories_link = cell_range_link(
                Role::Categories,
                FIRST_DATA_ROW,
                last_category_row,
                CATEGORY_COLUMN,
            );
            let values_link =
                cell_range_link(Role::Values, FIRST_DATA_ROW, last_value_row, value_column);
            chart.series.push(Series {
                category_kind: DataKind::Text,
                category_count,
                value_count,
                bubble_count: 0,
                chart_group: 0,
                name: series.name,
                links: vec![categories_link, values_link],
            });
            let categories_cache = cache_index(index, Role::Categories)?;
            for (point, category) in categories.iter().enumerate() {
                chart.cached_values.push(Cache {
                    cache_index: categories_cache,
                    row: point as u16,
                    column: category_column,
                    format: 0,
                    value: Value::Text(category.clone()),
                });
            }
            let values_cache = cache_index(index, Role::Values)?;
            for (point, value) in series.values.into_iter().enumerate() {
                chart.cached_values.push(Cache {
                    cache_index: values_cache,
                    row: point as u16,
                    column: cache_column,
                    format: 0,
                    value: Value::Number(value),
                });
            }
        }
        Ok(chart)
    }

    /// Generates the abbreviated workbook used only by private parser tests.
    #[cfg(test)]
    fn build_workbook_fixture(self) -> Result<Vec<u8>, PptWriteError> {
        build_workbook_fixture(self.model()?, Limits::default())
            .map_err(|error| invalid(format!("chart workbook generation failed: {error}")))
    }
}

fn check_string_units(value: &str, context: &str) -> Result<(), PptWriteError> {
    if value.encode_utf16().count() > MAX_STRING_UNITS {
        return Err(invalid(format!(
            "{context} exceeds {MAX_STRING_UNITS} UTF-16 code units"
        )));
    }
    Ok(())
}

#[cfg(test)]
fn checked_point_count(count: usize) -> Result<u16, PptWriteError> {
    if count > MAX_DATA_POINT_COUNT {
        return Err(invalid(format!(
            "chart data range exceeds {MAX_DATA_POINT_COUNT} points"
        )));
    }
    u16::try_from(count).map_err(|_| invalid("chart data point count exceeds u16"))
}

/// Cache indexes alternate per series: even for categories, odd for values.
#[cfg(test)]
fn cache_index(series_index: usize, role: Role) -> Result<u16, PptWriteError> {
    let value_offset = u16::from(role == Role::Values);
    u16::try_from(series_index)
        .ok()
        .and_then(|index| index.checked_mul(2))
        .and_then(|base| base.checked_add(value_offset))
        .ok_or_else(|| invalid("chart cache index exceeds BIFF8 bounds"))
}

/// A BRAI cell-range link to the generated single-sheet workbook.
#[cfg(test)]
fn cell_range_link(role: Role, first_row: u16, last_row: u16, column: u16) -> DataLink {
    let mut formula_tokens = Vec::with_capacity(11);
    formula_tokens.push(PTG_AREA);
    formula_tokens.extend_from_slice(&LOCAL_SHEET_INDEX.to_le_bytes());
    formula_tokens.extend_from_slice(&first_row.to_le_bytes());
    formula_tokens.extend_from_slice(&last_row.to_le_bytes());
    formula_tokens.extend_from_slice(&column.to_le_bytes());
    formula_tokens.extend_from_slice(&column.to_le_bytes());
    DataLink {
        role,
        source: Source::Cells,
        unlinked_number_format: false,
        number_format: 0,
        formula_tokens,
        references: vec![CellRef {
            extern_sheet_index: LOCAL_SHEET_INDEX,
            first_row,
            last_row,
            first_column: column,
            last_column: column,
        }],
    }
}

/// A chart pinned to a slide rectangle, with its generated workbook payload.
#[derive(Debug, Clone)]
pub(crate) struct PositionedChart {
    /// X position in EMUs.
    pub x: i32,
    /// Y position in EMUs.
    pub y: i32,
    /// Width in EMUs.
    pub width: i32,
    /// Height in EMUs.
    pub height: i32,
    /// Generated BIFF8 chart workbook (the `ExOleObjStg` payload).
    pub workbook: Vec<u8>,
}

/// Save-time identifiers assigned to one positioned chart.
#[derive(Debug, Clone, Copy)]
pub(crate) struct ChartPlan {
    /// Slide index owning the chart.
    pub slide: usize,
    /// Index of the chart within that slide.
    pub chart: usize,
    /// External-object identifier (`ExOleObjAtom.exObjId`).
    pub ex_obj_id: u32,
    /// Persist identifier of the chart's `ExOleObjStg` record.
    pub persist_id: u32,
}

impl ChartPlan {
    /// The `ExOleEmbedContainer` record declaring this chart object.
    fn embed_container_bytes(&self) -> Result<Vec<u8>, PptWriteError> {
        PowerPointOleObjectDefinition {
            kind: PowerPointOleContainerKind::Embedded(PowerPointOleEmbedPreferences {
                color_follow: PowerPointOleColorFollow::EntireScheme,
                cannot_lock_server: false,
                dimension_policy: PowerPointOleDimensionPolicy::Send,
                is_word_table: false,
                unused: 0,
            }),
            object: PowerPointOleObjectMetadata {
                draw_aspect: PowerPointOleDrawAspect::Content,
                object_type: PowerPointOleObjectType::Embedded,
                id: self.ex_obj_id,
                subtype: PowerPointOleObjectSubtype::ExcelChart,
                persist_id: self.persist_id,
                unused: [0; 4],
            },
            menu_name: None,
            program_id: Some(EXCEL_CHART_PROG_ID.to_string()),
            clipboard_name: None,
            metafile: None,
        }
        .to_record_bytes()
        .map_err(|error| invalid(format!("chart object container is invalid: {error}")))
    }
}

/// Build the document `ExObjList` combining hyperlink and chart objects.
///
/// Returns empty bytes when the presentation has neither, matching the
/// previous hyperlink-only behavior.
pub(crate) fn build_ex_obj_list(
    hyperlinks: &HyperlinkCollection,
    plans: &[ChartPlan],
) -> Result<Vec<u8>, PptWriteError> {
    if hyperlinks.is_empty() && plans.is_empty() {
        return Ok(Vec::new());
    }
    let id_seed = plans
        .last()
        .map_or_else(|| hyperlinks.id_seed(), |plan| plan.ex_obj_id);
    let mut container = RecordBuilder::new(0x0F, 0, hyperlink_record_type::EX_OBJ_LIST);
    let mut list_atom = RecordBuilder::new(0x00, 0, hyperlink_record_type::EX_OBJ_LIST_ATOM);
    list_atom.write_data(&id_seed.to_le_bytes());
    container.write_child(&list_atom.build()?);
    for record in hyperlinks.build_ex_hyperlink_records()? {
        container.write_child(&record);
    }
    for plan in plans {
        container.write_child(&plan.embed_container_bytes()?);
    }
    Ok(container.build()?)
}

/// Build the uncompressed `ExOleObjStg` record persisting a chart workbook.
pub(crate) fn chart_storage_record(workbook: &[u8]) -> Result<Vec<u8>, PptWriteError> {
    PowerPointOleStorage {
        kind: PowerPointOleStorageKind::OleObject,
        compression: PowerPointOleStorageCompression::Uncompressed,
        data: workbook.to_vec(),
    }
    .to_record_bytes()
    .map_err(|error| invalid(format!("chart storage record is invalid: {error}")))
}

/// A chart frame shape placed into a slide's drawing at save time.
#[derive(Debug, Clone, Copy)]
pub(crate) struct ChartFrame {
    /// Left edge in EMUs.
    left: i32,
    /// Top edge in EMUs.
    top: i32,
    /// Right edge in EMUs.
    right: i32,
    /// Bottom edge in EMUs.
    bottom: i32,
    /// External-object identifier referenced by the frame's `ExObjRefAtom`.
    ex_obj_id: u32,
}

impl ChartFrame {
    /// Creates a frame only when both far-edge calculations are representable.
    pub(crate) fn new(
        x: i32,
        y: i32,
        width: i32,
        height: i32,
        ex_obj_id: u32,
    ) -> Result<Self, PptError> {
        let right = x.checked_add(width).ok_or_else(|| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "chart frame horizontal extent exceeds i32",
            )
        })?;
        let bottom = y.checked_add(height).ok_or_else(|| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "chart frame vertical extent exceeds i32",
            )
        })?;
        Ok(Self {
            left: x,
            top: y,
            right,
            bottom,
            ex_obj_id,
        })
    }
}

/// Build the OLE object frame `SpContainer` for one chart.
///
/// Layout per [MS-ODRAW] and POI `HSLFOLEShape`: an `Sp` record with the
/// picture-frame shape type, a `ClientAnchor`, and a `ClientData` record
/// holding only the `ExObjRefAtom`.
pub(crate) fn build_chart_sp_container(
    frame: &ChartFrame,
    shape_id: u32,
) -> Result<Vec<u8>, PptError> {
    let mut container = EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    let mut sp = EscherBuilder::new(header_version::SP, MSOSPT_PICTURE_FRAME, record_type::SP);
    sp.add_data(
        EscherSpData::with_flags(shape_id, ShapeFlags::HAVE_ANCHOR | ShapeFlags::HAVE_SPT)
            .as_bytes(),
    );
    container.add_data(&sp.build()?);

    // ClientAnchor with position/size (8-byte short format for PPT shapes).
    let mut anchor = EscherBuilder::new(header_version::SIMPLE, 0, record_type::CLIENT_ANCHOR);
    anchor.add_data(&emu_i32_to_ppt_master_i16_round(frame.top).to_le_bytes());
    anchor.add_data(&emu_i32_to_ppt_master_i16_round(frame.left).to_le_bytes());
    anchor.add_data(&emu_i32_to_ppt_master_i16_round(frame.right).to_le_bytes());
    anchor.add_data(&emu_i32_to_ppt_master_i16_round(frame.bottom).to_le_bytes());
    container.add_data(&anchor.build()?);

    // ClientData wrapping the ExObjRefAtom that ties the frame to the chart.
    let mut reference = RecordBuilder::new(0x00, 0, EX_OBJ_REF_ATOM);
    reference.write_data(&frame.ex_obj_id.to_le_bytes());
    let mut client_data =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::CLIENT_DATA);
    client_data.add_data(&reference.build()?);
    container.add_data(&client_data.build()?);

    container.build()
}

#[cfg(test)]
mod tests {
    use super::*;
    fn sample_chart(kind: ChartKind) -> Chart {
        let mut chart = Chart::new(kind);
        chart.set_title("Quarterly sales");
        chart.set_categories(["Q1", "Q2", "Q3"]);
        chart.add_series(Some("2024"), vec![1.5, 2.0, 2.5]).unwrap();
        chart.add_series(None::<&str>, vec![3.0, 2.5, 2.0]).unwrap();
        chart
    }

    #[test]
    fn abbreviated_test_model_keeps_links_and_caches_for_parser_coverage() {
        let chart = sample_chart(ChartKind::Bar).model().unwrap();
        assert!(matches!(chart.kind(), ModelKind::Bar));
        assert_eq!(chart.title.as_deref(), Some("Quarterly sales"));
        assert_eq!(chart.series.len(), 2);
        let first = &chart.series[0];
        assert_eq!(first.category_count, 3);
        assert_eq!(first.value_count, 3);
        assert_eq!(first.name.as_deref(), Some("2024"));
        assert_eq!(first.links.len(), 2);
        assert_eq!(first.links[0].role, Role::Categories);
        assert_eq!(first.links[1].role, Role::Values);
        assert_eq!(first.links[1].references[0].first_column, 1);
        assert!(
            chart
                .cached_values
                .iter()
                .any(|entry| entry.value == Value::Text("Q2".to_string()))
        );
        assert!(
            chart
                .cached_values
                .iter()
                .any(|entry| entry.value == Value::Number(2.5))
        );
    }

    #[test]
    fn chart_validation_rejects_empty_and_non_finite_data() {
        assert!(Chart::new(ChartKind::Line).model().is_err());
        let mut chart = Chart::new(ChartKind::Line);
        assert!(chart.add_series(None::<String>, Vec::new()).is_err());
        assert!(chart.add_series(None::<String>, vec![f64::NAN]).is_err());
        chart.set_categories(std::iter::repeat_n("x", MAX_DATA_POINT_COUNT + 1));
        chart.add_series(None::<String>, vec![1.0]).unwrap();
        assert!(chart.model().is_err());

        let mut chart = Chart::new(ChartKind::Line);
        for _ in 0..=MAX_SERIES_COUNT {
            chart
                .add_series(None::<String>, vec![1.0])
                .expect("individual series is valid");
        }
        assert!(chart.model().is_err());
    }

    #[test]
    fn abbreviated_test_workbook_exercises_chart_editor() {
        let bytes = sample_chart(ChartKind::Pie)
            .build_workbook_fixture()
            .unwrap();
        let editor = Editor::open(bytes, Limits::default()).unwrap();
        let mut charts = editor.charts();
        assert_eq!(charts.len(), 1);
        let chart = charts.next().expect("one chart");
        assert!(matches!(
            chart.location,
            Location::Embedded { sheet_index: 0, .. }
        ));
        assert!(matches!(chart.chart.kind(), ModelKind::Pie));
        assert_eq!(chart.chart.series.len(), 2);
        assert_eq!(chart.chart.title.as_deref(), Some("Quarterly sales"));
    }

    #[test]
    fn ex_obj_list_combines_hyperlinks_and_charts() {
        let hyperlinks = HyperlinkCollection::new();
        assert!(build_ex_obj_list(&hyperlinks, &[]).unwrap().is_empty());

        let plans = [ChartPlan {
            slide: 0,
            chart: 0,
            ex_obj_id: 1,
            persist_id: 7,
        }];
        let bytes = build_ex_obj_list(&hyperlinks, &plans).unwrap();
        let (record, _) = crate::ppt::records::PptRecord::parse(&bytes, 0).unwrap();
        let collection = crate::ppt::ole_object::PowerPointOleObjectCollection::parse(&record)
            .unwrap()
            .expect("chart object list");
        assert_eq!(collection.objects.len(), 1);
        let crate::ppt::ole_object::PowerPointOleExternalObject::Object(definition) =
            &collection.objects[0]
        else {
            panic!("chart plan produces an embedded object");
        };
        assert_eq!(definition.object.id, 1);
        assert_eq!(definition.object.persist_id, 7);
        assert_eq!(
            definition.object.subtype,
            PowerPointOleObjectSubtype::ExcelChart
        );
        assert_eq!(definition.program_id.as_deref(), Some(EXCEL_CHART_PROG_ID));
    }

    #[test]
    fn chart_frame_container_exposes_object_reference() {
        let frame =
            ChartFrame::new(914_400, 914_400, 3_657_600, 2_743_200, 42).expect("valid frame");
        let bytes = build_chart_sp_container(&frame, 1027).unwrap();
        let (record, consumed) = litchi_odraw::Record::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        let shape = litchi_odraw::shape::Shape::try_from(record).unwrap();
        use crate::ppt::odraw::ShapeExt as _;
        assert_eq!(shape.external_object_id().unwrap(), Some(42));
    }

    #[test]
    fn chart_frame_rejects_unrepresentable_extents() {
        assert!(ChartFrame::new(i32::MAX, 0, 1, 1, 1).is_err());
        assert!(ChartFrame::new(0, i32::MAX, 1, 1, 1).is_err());
    }
}
