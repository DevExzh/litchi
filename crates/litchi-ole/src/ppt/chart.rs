//! Typed, inert chart model for native charts embedded in legacy PowerPoint.
//!
//! A native chart is an embedded OLE object whose `ExOleObjAtom` subtype or
//! ProgID identifies Microsoft Graph (`MSGraph.Chart`) or Excel
//! (`Excel.Chart`) ([MS-PPT] 2.13.11). Its `ExOleObjStg` payload is an OLE2
//! compound storage holding a BIFF8 `Workbook` stream whose chart substreams
//! ([MS-OGRAPH]) are parsed by the shared `xls::chart` machinery.
//!
//! Everything is inert: no formula evaluation, no rendering, no OLE
//! activation, and linked external workbooks are never opened.

use std::collections::HashMap;
use std::io::Read;

use super::ole_object::{
    PowerPointOleContainerKind, PowerPointOleExternalObject, PowerPointOleObjectCollection,
    PowerPointOleObjectSubtype,
};
use super::ole_storage::{
    PowerPointOleStorage, PowerPointOleStorageCompression, PowerPointOleStorageKind,
};
use super::package::{PptError, Result};
use super::presentation::Presentation;
use super::shapes::{PictureFrameKind, ShapeEnum};
use crate::xls::{XlsChartEditor, XlsChartEntry, XlsChartLimits};

/// Maximum number of chart-bearing OLE objects enumerated per presentation.
const MAX_CHART_OBJECTS: usize = 512;

/// The chart-producing application behind an embedded OLE object.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PowerPointChartKind {
    /// Microsoft Graph chart (`MSGraph.Chart`, subtype `ExOleSub_Graph`).
    Graph,
    /// Microsoft Excel chart (`Excel.Chart`, subtype `ExOleSub_ExcelChart`).
    ExcelChart,
}

/// The slide shape that displays an embedded chart object, when known.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PowerPointChartFrame {
    /// One-based slide number.
    pub slide_number: usize,
    /// OfficeArt shape identifier of the OLE object frame.
    pub shape_id: u32,
}

/// One parsed native chart object. Read-only and fully inert.
#[derive(Debug)]
pub struct PowerPointChart {
    /// The `ExOleObjAtom` object identifier.
    pub object_id: u32,
    /// The persist identifier of the object's `ExOleObjStg` storage.
    pub persist_id: u32,
    /// The chart-producing application.
    pub kind: PowerPointChartKind,
    /// The declared ProgID, when present.
    pub program_id: Option<String>,
    /// The slide frame displaying this chart, when one was found.
    pub frame: Option<PowerPointChartFrame>,
    /// Chart substreams of the embedded BIFF8 workbook. Never empty.
    pub charts: Vec<XlsChartEntry>,
}

/// A chart-bearing OLE object whose payload could not be parsed.
///
/// One corrupt chart never aborts enumeration of the remaining charts.
#[derive(Debug)]
pub struct PowerPointChartFailure {
    /// The `ExOleObjAtom` object identifier.
    pub object_id: u32,
    /// The persist identifier of the object's `ExOleObjStg` storage.
    pub persist_id: u32,
    /// The slide frame displaying this chart, when one was found.
    pub frame: Option<PowerPointChartFrame>,
    /// Why the payload could not be decoded or parsed.
    pub error: PptError,
}

/// Read-only inventory of a presentation's embedded native charts.
#[derive(Debug, Default)]
pub struct PowerPointChartInventory {
    /// Successfully parsed chart objects in external-object order.
    pub charts: Vec<PowerPointChart>,
    /// Chart objects whose payloads failed to decode or parse.
    pub failures: Vec<PowerPointChartFailure>,
}

impl PowerPointChartInventory {
    /// Whether the presentation contains no chart objects at all.
    pub fn is_empty(&self) -> bool {
        self.charts.is_empty() && self.failures.is_empty()
    }
}

/// Identify a chart-bearing OLE object from its subtype and ProgID.
fn classify_chart_object(
    subtype: PowerPointOleObjectSubtype,
    program_id: Option<&str>,
) -> Option<PowerPointChartKind> {
    match subtype {
        PowerPointOleObjectSubtype::Graph => return Some(PowerPointChartKind::Graph),
        PowerPointOleObjectSubtype::ExcelChart => return Some(PowerPointChartKind::ExcelChart),
        _ => {},
    }
    let base = prog_id_base(program_id?);
    if base.eq_ignore_ascii_case("MSGraph.Chart") || base.eq_ignore_ascii_case("MSGraph") {
        Some(PowerPointChartKind::Graph)
    } else if base.eq_ignore_ascii_case("Excel.Chart") {
        Some(PowerPointChartKind::ExcelChart)
    } else {
        None
    }
}

/// Strip a trailing numeric version (`"Excel.Chart.8"` -> `"Excel.Chart"`).
fn prog_id_base(prog_id: &str) -> &str {
    match prog_id.rsplit_once('.') {
        Some((base, version))
            if !version.is_empty() && version.bytes().all(|byte| byte.is_ascii_digit()) =>
        {
            base
        },
        _ => prog_id,
    }
}

/// Decode an inert `ExOleObjStg` payload into compound-file bytes.
fn decode_chart_payload(storage: PowerPointOleStorage, limits: XlsChartLimits) -> Result<Vec<u8>> {
    if storage.kind != PowerPointOleStorageKind::OleObject {
        return corrupted("chart persist ID does not reference an OLE object storage");
    }
    match storage.compression {
        PowerPointOleStorageCompression::Uncompressed => Ok(storage.data),
        PowerPointOleStorageCompression::Zlib { uncompressed_len } => {
            if uncompressed_len as usize > limits.max_workbook_bytes {
                return corrupted("compressed chart storage declares an excessive size");
            }
            let mut bytes = Vec::new();
            flate2::read::ZlibDecoder::new(storage.data.as_slice())
                .take(limits.max_workbook_bytes as u64 + 1)
                .read_to_end(&mut bytes)?;
            if bytes.len() > limits.max_workbook_bytes || bytes.len() != uncompressed_len as usize {
                return corrupted("compressed chart storage size mismatch or limit exceeded");
            }
            Ok(bytes)
        },
    }
}

/// Enumerate the presentation's embedded native charts, degrading per object.
pub(crate) fn enumerate(
    presentation: &Presentation,
    limits: XlsChartLimits,
) -> Result<PowerPointChartInventory> {
    let document = presentation.live_document_record()?;
    let Some(objects) = PowerPointOleObjectCollection::parse(&document)? else {
        return Ok(PowerPointChartInventory::default());
    };
    objects.validate_persist_mapping(&presentation.persist_mapping)?;
    let frames = chart_frames(presentation);
    let mut inventory = PowerPointChartInventory::default();
    for object in &objects.objects {
        let PowerPointOleExternalObject::Object(definition) = object else {
            continue;
        };
        // Linked charts are never opened: their payloads live in external files.
        if !matches!(definition.kind, PowerPointOleContainerKind::Embedded(_)) {
            continue;
        }
        let Some(kind) =
            classify_chart_object(definition.object.subtype, definition.program_id.as_deref())
        else {
            continue;
        };
        if inventory.charts.len() + inventory.failures.len() >= MAX_CHART_OBJECTS {
            return corrupted(format!(
                "presentation exceeds {MAX_CHART_OBJECTS} chart objects"
            ));
        }
        let frame = frames.get(&definition.object.id).copied();
        match parse_chart_object(presentation, definition.object.persist_id, limits) {
            Ok(charts) => inventory.charts.push(PowerPointChart {
                object_id: definition.object.id,
                persist_id: definition.object.persist_id,
                kind,
                program_id: definition.program_id.clone(),
                frame,
                charts,
            }),
            Err(error) => inventory.failures.push(PowerPointChartFailure {
                object_id: definition.object.id,
                persist_id: definition.object.persist_id,
                frame,
                error,
            }),
        }
    }
    Ok(inventory)
}

/// Open one chart object's storage and parse its BIFF8 chart substreams.
fn parse_chart_object(
    presentation: &Presentation,
    persist_id: u32,
    limits: XlsChartLimits,
) -> Result<Vec<XlsChartEntry>> {
    let storage = presentation.ole_storage(persist_id)?.ok_or_else(|| {
        PptError::Corrupted(format!(
            "chart object persist ID {persist_id} has no storage"
        ))
    })?;
    let bytes = decode_chart_payload(storage, limits)?;
    let charts = XlsChartEditor::open(bytes, limits)
        .map(XlsChartEditor::into_charts)
        .map_err(|error| {
            PptError::Corrupted(format!(
                "chart payload is not a BIFF8 chart workbook: {error}"
            ))
        })?;
    if charts.is_empty() {
        return corrupted("chart workbook contains no chart substream");
    }
    Ok(charts)
}

/// Map external-object IDs to the slide OLE frames that display them.
///
/// Slide-shape failures only lose frame attribution; chart enumeration is
/// document-level and continues regardless.
fn chart_frames(presentation: &Presentation) -> HashMap<u32, PowerPointChartFrame> {
    let mut frames = HashMap::new();
    let Ok(slides) = presentation.slides() else {
        return frames;
    };
    for slide in &slides {
        let Ok(shapes) = slide.shapes() else {
            continue;
        };
        collect_chart_frames(shapes, slide.slide_number(), &mut frames);
    }
    frames
}

fn collect_chart_frames(
    shapes: &[ShapeEnum],
    slide_number: usize,
    frames: &mut HashMap<u32, PowerPointChartFrame>,
) {
    for shape in shapes {
        match shape {
            ShapeEnum::Picture(picture) => {
                if picture.frame_kind() == PictureFrameKind::OleObject
                    && let Some(object_id) = picture.external_object_id()
                {
                    frames.entry(object_id).or_insert(PowerPointChartFrame {
                        slide_number,
                        shape_id: picture.properties.id,
                    });
                }
            },
            ShapeEnum::Group(group) => collect_chart_frames(group.children(), slide_number, frames),
            _ => {},
        }
    }
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(PptError::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ppt::shapes::PictureShape;
    use std::io::Write;

    #[test]
    fn subtype_and_prog_id_identify_chart_objects() {
        assert_eq!(
            classify_chart_object(PowerPointOleObjectSubtype::Graph, None),
            Some(PowerPointChartKind::Graph)
        );
        assert_eq!(
            classify_chart_object(PowerPointOleObjectSubtype::ExcelChart, None),
            Some(PowerPointChartKind::ExcelChart)
        );
        for (prog_id, kind) in [
            ("MSGraph.Chart.8", PowerPointChartKind::Graph),
            ("MSGraph.Chart", PowerPointChartKind::Graph),
            ("MSGraph", PowerPointChartKind::Graph),
            ("msgraph.chart.8", PowerPointChartKind::Graph),
            ("Excel.Chart.8", PowerPointChartKind::ExcelChart),
            ("Excel.Chart", PowerPointChartKind::ExcelChart),
            ("EXCEL.CHART.8", PowerPointChartKind::ExcelChart),
        ] {
            assert_eq!(
                classify_chart_object(PowerPointOleObjectSubtype::Default, Some(prog_id)),
                Some(kind),
                "{prog_id}"
            );
        }
        for prog_id in [
            "Excel.Sheet.8",
            "Excel.SheetMacroEnabled.12",
            "Word.Document.8",
            "PowerPoint.Show.8",
            "Equation.3",
            "Excel.ChartTool.8",
        ] {
            assert_eq!(
                classify_chart_object(PowerPointOleObjectSubtype::Default, Some(prog_id)),
                None,
                "{prog_id}"
            );
        }
        assert_eq!(
            classify_chart_object(PowerPointOleObjectSubtype::Excel, Some("Excel.Sheet.8")),
            None
        );
        assert_eq!(
            classify_chart_object(PowerPointOleObjectSubtype::Default, None),
            None
        );
    }

    #[test]
    fn prog_id_base_strips_only_trailing_numeric_versions() {
        assert_eq!(prog_id_base("Excel.Chart.8"), "Excel.Chart");
        assert_eq!(prog_id_base("Excel.Chart"), "Excel.Chart");
        assert_eq!(prog_id_base("MSGraph"), "MSGraph");
        assert_eq!(prog_id_base("Excel.Chart."), "Excel.Chart.");
    }

    fn storage(
        compression: PowerPointOleStorageCompression,
        data: Vec<u8>,
    ) -> PowerPointOleStorage {
        PowerPointOleStorage {
            kind: PowerPointOleStorageKind::OleObject,
            compression,
            data,
        }
    }

    #[test]
    fn uncompressed_payload_passes_through_and_wrong_kind_is_rejected() {
        let bytes = b"compound".to_vec();
        let decoded = decode_chart_payload(
            storage(PowerPointOleStorageCompression::Uncompressed, bytes.clone()),
            XlsChartLimits::default(),
        )
        .unwrap();
        assert_eq!(decoded, bytes);
        let vba = PowerPointOleStorage {
            kind: PowerPointOleStorageKind::VbaProject,
            compression: PowerPointOleStorageCompression::Uncompressed,
            data: bytes,
        };
        assert!(decode_chart_payload(vba, XlsChartLimits::default()).is_err());
    }

    #[test]
    fn zlib_payload_roundtrips_and_bombs_are_rejected() {
        let raw = vec![7u8; 65_536];
        let mut encoder =
            flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
        encoder.write_all(&raw).unwrap();
        let compressed = encoder.finish().unwrap();
        let decoded = decode_chart_payload(
            storage(
                PowerPointOleStorageCompression::Zlib {
                    uncompressed_len: raw.len() as u32,
                },
                compressed,
            ),
            XlsChartLimits::default(),
        )
        .unwrap();
        assert_eq!(decoded, raw);

        let mut encoder =
            flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
        encoder.write_all(&[1u8; 4096]).unwrap();
        let compressed = encoder.finish().unwrap();
        let bomb = storage(
            PowerPointOleStorageCompression::Zlib {
                uncompressed_len: 4096,
            },
            compressed,
        );
        let limits = XlsChartLimits {
            max_workbook_bytes: 1024,
            ..Default::default()
        };
        assert!(decode_chart_payload(bomb, limits).is_err());

        let declared_bomb = storage(
            PowerPointOleStorageCompression::Zlib {
                uncompressed_len: 2048,
            },
            vec![0x78, 0x9c],
        );
        assert!(decode_chart_payload(declared_bomb, XlsChartLimits::default()).is_err());

        let truncated = storage(
            PowerPointOleStorageCompression::Zlib {
                uncompressed_len: 10,
            },
            vec![0x78, 0x9c, 1, 2],
        );
        assert!(decode_chart_payload(truncated, XlsChartLimits::default()).is_err());
    }

    #[test]
    fn ole_frames_are_attributed_recursively_first_wins() {
        let mut chart_frame = PictureShape::new(7);
        chart_frame.set_frame_kind(PictureFrameKind::OleObject);
        chart_frame.set_external_object_id(42);
        let mut plain = PictureShape::new(8);
        plain.set_frame_kind(PictureFrameKind::Picture);
        let mut nested = PictureShape::new(9);
        nested.set_frame_kind(PictureFrameKind::OleObject);
        nested.set_external_object_id(77);
        let mut group = crate::ppt::shapes::shape_enum::GroupShape::new(10);
        group.add_child(ShapeEnum::Picture(nested));
        let mut duplicate = PictureShape::new(11);
        duplicate.set_frame_kind(PictureFrameKind::OleObject);
        duplicate.set_external_object_id(42);
        let shapes = vec![
            ShapeEnum::Picture(chart_frame),
            ShapeEnum::Picture(plain),
            ShapeEnum::Group(group),
            ShapeEnum::Picture(duplicate),
        ];
        let mut frames = HashMap::new();
        collect_chart_frames(&shapes, 3, &mut frames);
        assert_eq!(frames.len(), 2);
        assert_eq!(
            frames[&42],
            PowerPointChartFrame {
                slide_number: 3,
                shape_id: 7
            }
        );
        assert_eq!(
            frames[&77],
            PowerPointChartFrame {
                slide_number: 3,
                shape_id: 9
            }
        );
    }
}
