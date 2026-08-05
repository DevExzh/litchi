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
//! use litchi_ppt::writer::{Chart, ChartKind, PptWriter};
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
//!     litchi_ppt::writer::PptWriteError::Graph(
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
use crate::embedded::object::{
    ColorFollow, ContainerKind, Definition, DimensionPolicy, DrawAspect, EmbedPreferences,
    Metadata, ObjectSubtype, ObjectType,
};
use crate::embedded::reference::Reference;
use crate::embedded::storage::{Kind as StorageKind, Storage};

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
/// Maximum number of chart objects authored into one presentation.
///
/// Matches the read-side inventory bound in [`crate::chart`].
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
}

fn check_string_units(value: &str, context: &str) -> Result<(), PptWriteError> {
    if value.encode_utf16().count() > MAX_STRING_UNITS {
        return Err(invalid(format!(
            "{context} exceeds {MAX_STRING_UNITS} UTF-16 code units"
        )));
    }
    Ok(())
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
        Definition {
            kind: ContainerKind::Embedded(EmbedPreferences {
                color_follow: ColorFollow::EntireScheme,
                cannot_lock_server: false,
                dimension_policy: DimensionPolicy::Send,
                is_word_table: false,
                unused: 0,
            }),
            object: Metadata {
                draw_aspect: DrawAspect::Content,
                object_type: ObjectType::Embedded,
                id: self.ex_obj_id,
                subtype: ObjectSubtype::ExcelChart,
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
    Storage::uncompressed(StorageKind::OleObject, workbook.to_vec())
        .map_err(|error| invalid(format!("chart storage is invalid: {error}")))?
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
        if ex_obj_id == 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "chart frame external-object ID must be positive",
            ));
        }
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
    let object_ref = Reference::new(frame.ex_obj_id).map_err(|error| {
        std::io::Error::new(std::io::ErrorKind::InvalidInput, error.to_string())
    })?;
    let payload = object_ref.to_payload_bytes().map_err(|error| {
        std::io::Error::new(std::io::ErrorKind::InvalidInput, error.to_string())
    })?;
    reference.write_data(&payload);
    let mut client_data =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::CLIENT_DATA);
    client_data.add_data(&reference.build()?);
    container.add_data(&client_data.build()?);

    container.build()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn chart_validation_rejects_empty_and_non_finite_data() {
        assert!(Chart::new(ChartKind::Line).validate().is_err());
        let mut chart = Chart::new(ChartKind::Line);
        assert!(chart.add_series(None::<String>, Vec::new()).is_err());
        assert!(chart.add_series(None::<String>, vec![f64::NAN]).is_err());
        chart.set_categories(std::iter::repeat_n("x", MAX_DATA_POINT_COUNT + 1));
        chart.add_series(None::<String>, vec![1.0]).unwrap();
        assert!(chart.validate().is_err());

        let mut chart = Chart::new(ChartKind::Line);
        for _ in 0..=MAX_SERIES_COUNT {
            chart
                .add_series(None::<String>, vec![1.0])
                .expect("individual series is valid");
        }
        assert!(chart.validate().is_err());
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
        let (record, _) = crate::records::PptRecord::parse(&bytes, 0).unwrap();
        let collection = crate::embedded::object::Collection::parse(&record)
            .unwrap()
            .expect("chart object list");
        assert_eq!(collection.objects.len(), 1);
        let crate::embedded::object::ExternalObject::Object(definition) = &collection.objects[0]
        else {
            panic!("chart plan produces an embedded object");
        };
        assert_eq!(definition.object.id, 1);
        assert_eq!(definition.object.persist_id, 7);
        assert_eq!(definition.object.subtype, ObjectSubtype::ExcelChart);
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
        use crate::odraw::ShapeExt as _;
        assert_eq!(shape.external_object_id().unwrap(), Some(42));
    }

    #[test]
    fn chart_frame_rejects_unrepresentable_extents() {
        assert!(ChartFrame::new(i32::MAX, 0, 1, 1, 1).is_err());
        assert!(ChartFrame::new(0, i32::MAX, 1, 1, 1).is_err());
    }
}
