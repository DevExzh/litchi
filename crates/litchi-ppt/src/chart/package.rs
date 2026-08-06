//! PowerPoint package integration for embedded OLE2 chart objects.

use litchi_ograph::Limits;

use super::codec::{Parsed, corrupted, parse};
use super::model::{Chart, Excel, Failure, Frame, Graph, Info, Inventory, Kind};
use crate::embedded::object::{Collection, ContainerKind, ExternalObject, ObjectSubtype};
use crate::package::Result;
use crate::presentation::Presentation;
use crate::shapes::{PictureFrameKind, ShapeEnum};

/// Maximum chart-bearing OLE objects enumerated per presentation.
pub(crate) const MAX_CHART_OBJECTS: usize = 512;

/// Enumerate embedded native charts, degrading malformed payloads per object.
pub(crate) fn enumerate(presentation: &Presentation, limits: Limits) -> Result<Inventory> {
    let limits = limits.validate()?;
    let document = presentation.live_document_record()?;
    let Some(objects) = Collection::parse(&document)? else {
        return Ok(Inventory::default());
    };
    objects.validate_persist_mapping(&presentation.persist_mapping)?;

    let mut wanted = Vec::new();
    wanted
        .try_reserve(objects.objects.len().min(MAX_CHART_OBJECTS))
        .map_err(|_| litchi_ograph::Error::Allocation {
            resource: "PPT chart object identifiers",
        })?;
    for object in &objects.objects {
        let ExternalObject::Object(definition) = object else {
            continue;
        };
        if matches!(definition.kind, ContainerKind::Embedded(_))
            && classify(definition.object.subtype, definition.program_id.as_deref()).is_some()
        {
            if wanted.len() >= MAX_CHART_OBJECTS {
                return corrupted(format!(
                    "presentation exceeds {MAX_CHART_OBJECTS} chart objects"
                ));
            }
            wanted.push(definition.object.id);
        }
    }

    let frames = chart_frames(presentation, &wanted)?;
    let mut inventory = Inventory::default();
    inventory
        .try_reserve(wanted.len())
        .map_err(|_| litchi_ograph::Error::Allocation {
            resource: "PPT chart inventory",
        })?;
    for object in objects.objects {
        let ExternalObject::Object(definition) = object else {
            continue;
        };
        if !matches!(definition.kind, ContainerKind::Embedded(_)) {
            continue;
        }
        let Some(kind) = classify(definition.object.subtype, definition.program_id.as_deref())
        else {
            continue;
        };
        let info = Info::new(
            definition.object.id,
            definition.object.persist_id,
            definition.program_id,
            frames
                .iter()
                .find_map(|(object, frame)| (*object == definition.object.id).then_some(*frame)),
        );
        match parse(presentation, info.persist_id(), kind, limits) {
            Ok(Parsed::Graph {
                package,
                book,
                compression,
            }) => {
                inventory.push_chart(Chart::Graph(Graph::new(info, package, book, compression)));
            },
            Ok(Parsed::Excel { book }) => {
                inventory.push_chart(Chart::Excel(Excel::new(info, book)));
            },
            Err(error) => inventory.push_failure(Failure::new(info, kind, error)),
        }
    }
    Ok(inventory)
}

pub(super) fn classify(subtype: ObjectSubtype, program: Option<&str>) -> Option<Kind> {
    match subtype {
        ObjectSubtype::Graph => return Some(Kind::Graph),
        ObjectSubtype::ExcelChart => return Some(Kind::Excel),
        _ => {},
    }
    let base = program_base(program?);
    if base.eq_ignore_ascii_case("MSGraph.Chart") || base.eq_ignore_ascii_case("MSGraph") {
        Some(Kind::Graph)
    } else if base.eq_ignore_ascii_case("Excel.Chart") {
        Some(Kind::Excel)
    } else {
        None
    }
}

pub(super) fn program_base(program: &str) -> &str {
    match program.rsplit_once('.') {
        Some((base, version))
            if !version.is_empty() && version.bytes().all(|byte| byte.is_ascii_digit()) =>
        {
            base
        },
        _ => program,
    }
}

fn chart_frames(presentation: &Presentation, wanted: &[u32]) -> Result<Vec<(u32, Frame)>> {
    let mut frames = Vec::new();
    frames
        .try_reserve(wanted.len())
        .map_err(|_| litchi_ograph::Error::Allocation {
            resource: "PPT chart frames",
        })?;
    let Ok(slides) = presentation.slides() else {
        return Ok(frames);
    };
    for slide in &slides {
        let Ok(shapes) = slide.shapes() else {
            continue;
        };
        collect_frames(shapes, slide.slide_number(), wanted, &mut frames);
    }
    Ok(frames)
}

pub(super) fn collect_frames(
    shapes: &[ShapeEnum<'_>],
    slide: usize,
    wanted: &[u32],
    frames: &mut Vec<(u32, Frame)>,
) {
    for shape in shapes {
        match shape {
            ShapeEnum::Picture(picture) => {
                if picture.frame_kind() == PictureFrameKind::OleObject
                    && let Some(object) = picture.external_object_id()
                    && wanted.contains(&object)
                    && !frames.iter().any(|(existing, _)| *existing == object)
                    && let Some(frame) = Frame::new(slide, picture.properties.id)
                {
                    frames.push((object, frame));
                }
            },
            ShapeEnum::Group(group) => collect_frames(group.children(), slide, wanted, frames),
            _ => {},
        }
    }
}
