//! Worksheet chart lifecycle: add, replace and remove ordinary charts.
//!
//! [`Package::edit_charts`] starts an [`Edit`] over an immutable snapshot.
//! [`Edit::add`] places a chart-space part on a worksheet through a
//! two-cell anchor of the worksheet's drawing (creating the drawing part and
//! the worksheet's `drawing` element when the sheet has none),
//! [`Edit::set`] replaces a chart's XML in place and [`Edit::remove`]
//! deletes its anchor and its chart part. [`Edit::commit`] returns a
//! [`Commit`] holding the new [`Package`], the reversible [`Patch`] of the
//! parts it changed and the [`Diagnostics`]; [`Package::apply_charts`]
//! replays a patch onto an equal package. The source package never changes.
//!
//! Charts are numbered per worksheet in drawing order, counting the ordinary
//! `c:chart` frames only ([`Package::chart_count`], [`Package::chart_xml`]).
//! Chart XML is validated by the `DrawingML` chart reader before it is
//! staged. A changed edit of a signed workbook is refused with
//! [`Error::Signed`]; an exact no-op keeps every byte.
//!
//! ```
//! use litchi_drawingml::chart::authoring::{Kind, Series, Spec, Variant};
//! use litchi_xlsx::Package;
//! use litchi_xlsx::chart::Anchor;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::new()?;
//! let xml = Spec::new(Variant::new(Kind::Column))
//!     .with_title("Sales")
//!     .with_categories(["Q1", "Q2"])
//!     .with_series(Series::new("East", vec![1.0, 2.0]))
//!     .to_xml(None)?;
//! let mut edit = package.edit_charts()?;
//! let index = edit.add(0, &Anchor::new(1, 1, 8, 15), "Chart 1", &xml)?;
//! let after = edit.commit()?.into_package();
//! assert_eq!(index, 0);
//! assert_eq!(after.chart_count(0)?, 1);
//! assert_eq!(package.chart_count(0)?, 0);
//! # Ok(())
//! # }
//! ```

use std::collections::BTreeSet;
use std::sync::Arc;

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

pub use crate::Error;
use crate::Package;
use crate::error::invalid;

use super::Anchor;

type Result<T> = std::result::Result<T, Error>;

const CHART_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chart";
const STRICT_CHART_NS: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/chart";
const RELATIONSHIP_NS: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIP_NS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";

/// The most ordinary charts one worksheet drawing may carry.
pub const MAX_CHARTS: usize = 4_096;
/// The largest chart-space part, in bytes.
pub const MAX_CHART_BYTES: usize = 64 * 1024 * 1024;

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Ordinary charts on one worksheet ([`MAX_CHARTS`]).
    Charts,
    /// Bytes of one chart-space part ([`MAX_CHART_BYTES`]).
    ChartBytes,
}

impl LimitKind {
    const fn maximum(self) -> usize {
        match self {
            Self::Charts => MAX_CHARTS,
            Self::ChartBytes => MAX_CHART_BYTES,
        }
    }

    fn check(self, observed: usize) -> Result<()> {
        if observed > self.maximum() {
            return Err(invalid(format!(
                "worksheet chart {} limit exceeded: {observed} exceeds {}",
                match self {
                    Self::Charts => "count",
                    Self::ChartBytes => "byte",
                },
                self.maximum()
            )));
        }
        Ok(())
    }
}

/// What a committed chart edit changed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    parts: usize,
}

impl Diagnostics {
    /// Whether the edit changed the package.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.parts > 0
    }

    /// The number of package parts the edit added, changed or removed.
    #[must_use]
    pub const fn parts(&self) -> usize {
        self.parts
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Relationship {
    id: String,
    kind: String,
    target: String,
    external: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct State {
    content_type: String,
    blob: Arc<Vec<u8>>,
    relationships: Vec<Relationship>,
}

impl State {
    fn read(part: &dyn Part) -> Self {
        let mut relationships: Vec<_> = part
            .rels()
            .iter()
            .map(|relationship| Relationship {
                id: relationship.r_id().to_owned(),
                kind: relationship.reltype().to_owned(),
                target: relationship.target_ref().to_owned(),
                external: relationship.is_external(),
            })
            .collect();
        relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
        Self {
            content_type: part.content_type().to_owned(),
            blob: part.blob_arc(),
            relationships,
        }
    }

    fn of(package: &OpcPackage, name: &PackURI) -> Option<Self> {
        package.get_part(name).ok().map(Self::read)
    }

    fn to_part(&self, name: PackURI) -> Result<BlobPart> {
        let mut part =
            BlobPart::new_shared(name, self.content_type.clone(), Arc::clone(&self.blob));
        for relationship in &self.relationships {
            part.rels_mut().try_add_relationship(
                relationship.kind.clone(),
                relationship.target.clone(),
                relationship.id.clone(),
                if relationship.external {
                    TargetMode::External
                } else {
                    TargetMode::Internal
                },
            )?;
        }
        Ok(part)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Delta {
    name: PackURI,
    before: Option<State>,
    after: Option<State>,
}

/// The exact, reversible part-level change of a chart edit.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Patch {
    deltas: Vec<Delta>,
}

impl Patch {
    /// Whether the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.deltas.is_empty()
    }

    /// The reverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            deltas: self
                .deltas
                .iter()
                .map(|delta| Delta {
                    name: delta.name.clone(),
                    before: delta.after.clone(),
                    after: delta.before.clone(),
                })
                .collect(),
        }
    }

    fn capture(before: &OpcPackage, after: &OpcPackage, names: &BTreeSet<String>) -> Result<Self> {
        let names = names
            .iter()
            .map(|name| PackURI::new(name.as_str()).map_err(|error| invalid(error.to_string())))
            .collect::<Result<Vec<_>>>()?;
        let deltas = names
            .into_iter()
            .filter_map(|name| {
                let before = State::of(before, &name);
                let after = State::of(after, &name);
                (before != after).then_some(Delta {
                    name,
                    before,
                    after,
                })
            })
            .collect();
        Ok(Self { deltas })
    }

    fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        for delta in &self.deltas {
            if State::of(target, &delta.name) != delta.before {
                return Err(Error::PatchConflict {
                    part: delta.name.as_str().to_owned(),
                });
            }
        }
        for delta in &self.deltas {
            target.remove_part(&delta.name);
            if let Some(after) = &delta.after {
                target.add_part(Box::new(after.to_part(delta.name.clone())?));
            }
        }
        Ok(())
    }
}

/// The result of [`Edit::commit`].
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// The package after the edit.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit into the package after the edit.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// The exact, reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// What the edit changed.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// Consume the commit into the package and the patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

/// A staged chart edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    package: &'a Package,
    staged: OpcPackage,
    touched: BTreeSet<String>,
}

/// One ordinary chart frame of a worksheet drawing.
#[derive(Debug, Clone)]
struct Frame {
    /// The byte span of the frame's anchor element in the drawing XML.
    anchor: std::ops::Range<usize>,
    /// The chart relationship the frame names.
    relationship_id: String,
}

/// The ordinary chart frames of a drawing, and its largest `cNvPr` id.
struct Inventory {
    frames: Vec<Frame>,
    max_id: u32,
    root_close: usize,
}

fn is_chart_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(ns) if ns.as_ref() == CHART_NS || ns.as_ref() == STRICT_CHART_NS)
}

fn is_relationship_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(ns) if ns.as_ref() == RELATIONSHIP_NS || ns.as_ref() == STRICT_RELATIONSHIP_NS)
}

fn inventory(xml: &[u8]) -> Result<Inventory> {
    use litchi_core::xml::ReaderOrigin;
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let position = |reader: &NsReader<&[u8]>| {
        origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("worksheet drawing position overflow"))
    };
    let mut depth = 0_usize;
    let mut anchor_start = None;
    let mut anchor_chart: Option<String> = None;
    let mut frames = Vec::new();
    let mut max_id = 1_u32;
    let mut root_close = None;
    loop {
        let start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("worksheet drawing XML: {error}")))?;
        let in_chart_namespace = is_chart_namespace(&namespace);
        let event = event.into_owned();
        let end = position(&reader)?;
        let (element, empty) = match &event {
            Event::Start(element) => (Some(element.clone()), false),
            Event::Empty(element) => (Some(element.clone()), true),
            _ => (None, false),
        };
        if let Some(element) = element {
            if depth == 1 {
                anchor_start = Some(start);
                anchor_chart = None;
            }
            let local = element.local_name();
            if local.as_ref() == b"cNvPr" {
                for attribute in element.checked_attributes().flatten() {
                    if attribute.key.as_ref() == b"id"
                        && let Ok(value) = std::str::from_utf8(&attribute.value)
                        && let Ok(id) = value.trim().parse::<u32>()
                    {
                        max_id = max_id.max(id);
                    }
                }
            }
            if depth >= 2 && local.as_ref() == b"chart" && in_chart_namespace {
                for attribute in element.checked_attributes().flatten() {
                    let (attribute_ns, local_name) =
                        reader.resolver().resolve_attribute(attribute.key);
                    if local_name.as_ref() == b"id" && is_relationship_namespace(&attribute_ns) {
                        let value = attribute
                            .decoded_and_normalized_value(
                                quick_xml::XmlVersion::Implicit1_0,
                                reader.decoder(),
                            )
                            .map_err(|error| invalid(format!("chart relationship id: {error}")))?;
                        anchor_chart = Some(value.into_owned());
                    }
                }
            }
            if empty {
                if depth == 1 {
                    anchor_start = None;
                }
            } else {
                depth += 1;
            }
            continue;
        }
        match event {
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 1
                    && let Some(anchor) = anchor_start.take()
                    && let Some(relationship_id) = anchor_chart.take()
                {
                    frames.push(Frame {
                        anchor: anchor..end,
                        relationship_id,
                    });
                }
                if depth == 0 {
                    root_close = Some(start);
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(Inventory {
        frames,
        max_id,
        root_close: root_close.ok_or_else(|| invalid("worksheet drawing has no root element"))?,
    })
}

fn validate_anchor(anchor: &Anchor) -> Result<()> {
    if anchor.to_row < anchor.from_row || anchor.to_col < anchor.from_col {
        return Err(invalid("chart anchor cannot be descending"));
    }
    if anchor.to_row >= 1_048_576 || anchor.to_col >= 16_384 {
        return Err(invalid("chart anchor exceeds the worksheet bounds"));
    }
    if [
        anchor.from_col_offset,
        anchor.from_row_offset,
        anchor.to_col_offset,
        anchor.to_row_offset,
    ]
    .iter()
    .any(|offset| *offset < 0)
    {
        return Err(invalid("chart anchor offsets cannot be negative"));
    }
    Ok(())
}

fn worksheet_uri(package: &Package, sheet: usize) -> Result<PackURI> {
    let worksheet = package
        .workbook_ref()
        .sheet(sheet)?
        .ok_or_else(|| invalid(format!("worksheet {sheet} does not exist")))?;
    if worksheet.kind() != crate::workbook::WorksheetKind::Worksheet {
        return Err(Error::NotWorksheet {
            sheet: worksheet.name().to_owned(),
        });
    }
    Ok(worksheet.part_uri().clone())
}

/// The drawing part of a worksheet part, if it has one.
fn drawing_uri(package: &OpcPackage, worksheet: &PackURI) -> Result<Option<PackURI>> {
    drawing_in(package.get_part(worksheet)?.rels())
}

/// The chart part a drawing frame's relationship names.
fn chart_uri(package: &OpcPackage, drawing: &PackURI, relationship_id: &str) -> Result<PackURI> {
    let part = package.get_part(drawing)?;
    let relationship = part
        .rels()
        .get(relationship_id)
        .ok_or_else(|| invalid("drawing chart relationship is missing"))?;
    if relationship.is_external() || !matches!(relationship.reltype(), rt::CHART | rt::STRICT_CHART)
    {
        return Err(invalid(
            "drawing chart relationship is not an internal chart",
        ));
    }
    let target = relationship.target_partname()?;
    if package.get_part(&target)?.content_type() != ct::DML_CHART {
        return Err(invalid("drawing chart target is not a chart part"));
    }
    Ok(target)
}

fn free_name(package: &OpcPackage, prefix: &str, suffix: &str) -> Result<PackURI> {
    for index in 1_u32..=1_000_000 {
        let candidate = PackURI::new(format!("{prefix}{index}{suffix}"))
            .map_err(|error| invalid(error.to_string()))?;
        if !package.contains_part(&candidate) {
            return Ok(candidate);
        }
    }
    Err(invalid("no free part name is left"))
}

fn check_chart(xml: &[u8]) -> Result<()> {
    LimitKind::ChartBytes.check(xml.len())?;
    litchi_drawingml::chart::reader::read(xml)?;
    Ok(())
}

/// The drawing part named by a worksheet's relationships, if any.
fn drawing_in(relationships: &litchi_opc::Relationships) -> Result<Option<PackURI>> {
    let mut found = None;
    for relationship in relationships.iter() {
        if !matches!(relationship.reltype(), rt::DRAWING | rt::STRICT_DRAWING) {
            continue;
        }
        if relationship.is_external() {
            return Err(invalid("worksheet drawing relationship cannot be external"));
        }
        if found.is_some() {
            return Err(invalid("worksheet has more than one drawing part"));
        }
        found = Some(relationship.target_partname()?);
    }
    Ok(found)
}

fn source_bytes(package: &litchi_opc::SourceBackedPackage, uri: &PackURI) -> Result<Vec<u8>> {
    let part = package.part(uri)?;
    let declared = part.declared_uncompressed_size()?;
    if declared > MAX_CHART_BYTES as u64 {
        LimitKind::ChartBytes.check(usize::try_from(declared).unwrap_or(usize::MAX))?;
    }
    let data = part.data()?;
    LimitKind::ChartBytes.check(data.as_bytes().len())?;
    Ok(data.as_bytes().to_vec())
}

/// The ordinary chart frames of a source-backed worksheet, with the drawing
/// part that holds them.
fn source_frames(
    package: &litchi_opc::SourceBackedPackage,
    worksheet: &PackURI,
) -> Result<Option<(PackURI, Vec<Frame>)>> {
    package.check_execution()?;
    let Some(drawing) = drawing_in(package.part(worksheet)?.rels())? else {
        return Ok(None);
    };
    let frames = inventory(&source_bytes(package, &drawing)?)?.frames;
    Ok(Some((drawing, frames)))
}

/// The number of ordinary charts on a source-backed worksheet part.
pub(crate) fn source_chart_count(
    package: &litchi_opc::SourceBackedPackage,
    worksheet: &PackURI,
) -> Result<usize> {
    Ok(source_frames(package, worksheet)?.map_or(0, |(_, frames)| frames.len()))
}

/// The chart-space XML of the ordinary chart at `index` on a source-backed
/// worksheet part. Only the worksheet's relationships, its drawing and the
/// selected chart part are read.
pub(crate) fn source_chart_xml(
    package: &litchi_opc::SourceBackedPackage,
    worksheet: &PackURI,
    index: usize,
) -> Result<Option<Vec<u8>>> {
    let Some((drawing, frames)) = source_frames(package, worksheet)? else {
        return Ok(None);
    };
    let Some(frame) = frames.get(index) else {
        return Ok(None);
    };
    let drawing_part = package.part(&drawing)?;
    let relationship = drawing_part
        .rels()
        .get(&frame.relationship_id)
        .ok_or_else(|| invalid("drawing chart relationship is missing"))?;
    if relationship.is_external() || !matches!(relationship.reltype(), rt::CHART | rt::STRICT_CHART)
    {
        return Err(invalid(
            "drawing chart relationship is not an internal chart",
        ));
    }
    let chart = relationship.target_partname()?;
    if package.part(&chart)?.content_type() != ct::DML_CHART {
        return Err(invalid("drawing chart target is not a chart part"));
    }
    source_bytes(package, &chart).map(Some)
}

impl Package {
    /// The number of ordinary charts on a worksheet.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet, a chartsheet, or a malformed
    /// drawing graph.
    pub fn chart_count(&self, sheet: usize) -> Result<usize> {
        let worksheet = worksheet_uri(self, sheet)?;
        let Some(drawing) = drawing_uri(self.opc(), &worksheet)? else {
            return Ok(0);
        };
        Ok(inventory(self.opc().get_part(&drawing)?.blob())?
            .frames
            .len())
    }

    /// The chart-space XML of the ordinary chart at `index` (drawing order)
    /// on a worksheet, or `None` past the last chart.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet, a chartsheet, or a malformed
    /// drawing graph.
    pub fn chart_xml(&self, sheet: usize, index: usize) -> Result<Option<&[u8]>> {
        let worksheet = worksheet_uri(self, sheet)?;
        let Some(drawing) = drawing_uri(self.opc(), &worksheet)? else {
            return Ok(None);
        };
        let frames = inventory(self.opc().get_part(&drawing)?.blob())?.frames;
        let Some(frame) = frames.get(index) else {
            return Ok(None);
        };
        let chart = chart_uri(self.opc(), &drawing, &frame.relationship_id)?;
        Ok(Some(self.opc().get_part(&chart)?.blob()))
    }

    /// Start a worksheet chart edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused.
    pub fn edit_charts(&self) -> Result<Edit<'_>> {
        self.workbook_ref().ensure_mutation_allowed("edit_charts")?;
        Ok(Edit {
            package: self,
            staged: self.opc().clone(),
            touched: BTreeSet::new(),
        })
    }

    /// Apply a chart patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the
    /// patch's before state and [`Error::Signed`] for a changed patch of a
    /// signed workbook.
    pub fn apply_charts(&self, patch: &Patch) -> Result<Self> {
        if patch.is_noop() {
            return Ok(self.clone());
        }
        self.workbook_ref()
            .ensure_mutation_allowed("apply_charts")?;
        crate::signature_policy::ensure_unsigned(self.opc())?;
        let mut candidate = self.opc().clone();
        patch.apply(&mut candidate)?;
        self.workbook_ref()
            .adopt_published_package(candidate)
            .map(Self::from_workbook)
    }
}

impl Edit<'_> {
    fn touch(&mut self, name: &PackURI) {
        self.touched.insert(name.as_str().to_owned());
    }

    fn frames(&self, sheet: usize) -> Result<(PackURI, PackURI, Inventory)> {
        let worksheet = worksheet_uri(self.package, sheet)?;
        let drawing = drawing_uri(&self.staged, &worksheet)?
            .ok_or_else(|| invalid("the worksheet has no charts"))?;
        let inventory = inventory(self.staged.get_part(&drawing)?.blob())?;
        Ok((worksheet, drawing, inventory))
    }

    /// Place a chart-space part on a worksheet (by position), anchored from
    /// `anchor.from` to `anchor.to`, as a frame named `name`. Returns the
    /// chart's index on the sheet.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or chartsheet, an invalid
    /// anchor, XML that is not a chart, a full drawing, or a signed
    /// workbook.
    pub fn add(&mut self, sheet: usize, anchor: &Anchor, name: &str, xml: &[u8]) -> Result<usize> {
        validate_anchor(anchor)?;
        check_chart(xml)?;
        if name.contains('\0') {
            return Err(invalid("chart frame name contains a NUL character"));
        }
        crate::signature_policy::ensure_unsigned(&self.staged)?;
        let worksheet = worksheet_uri(self.package, sheet)?;
        let mut staged = self.staged.clone();
        let drawing = match drawing_uri(&staged, &worksheet)? {
            Some(drawing) => drawing,
            None => {
                let drawing = free_name(&staged, "/xl/drawings/drawing", ".xml")?;
                staged.add_part(Box::new(BlobPart::new(
                    drawing.clone(),
                    ct::OFC_DRAWING.to_owned(),
                    concat!(
                        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" "#,
                        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>"#
                    )
                    .as_bytes()
                    .to_vec(),
                )));
                let target = drawing.relative_ref(worksheet.base_uri());
                let part = staged.get_part_mut(&worksheet)?;
                let relationship_id = part.relate_to(&target, rt::DRAWING);
                let xml =
                    crate::workbook::edit::attach_worksheet_drawing(part.blob(), &relationship_id)?;
                part.set_blob(xml);
                self.touch(&drawing);
                drawing
            },
        };
        let inventory = inventory(staged.get_part(&drawing)?.blob())?;
        LimitKind::Charts.check(inventory.frames.len() + 1)?;
        let chart = free_name(&staged, "/xl/charts/chart", ".xml")?;
        staged.add_part(Box::new(BlobPart::new(
            chart.clone(),
            ct::DML_CHART.to_owned(),
            xml.to_vec(),
        )));
        let target = chart.relative_ref(drawing.base_uri());
        let part = staged.get_part_mut(&drawing)?;
        let relationship_id = part.relate_to(&target, rt::CHART);
        let id = inventory
            .max_id
            .checked_add(1)
            .ok_or_else(|| invalid("drawing shape ids are exhausted"))?;
        let fragment = anchor_xml(anchor, id, name, &relationship_id);
        let mut bytes = part.blob().to_vec();
        bytes.splice(inventory.root_close..inventory.root_close, fragment.bytes());
        part.set_blob(bytes);
        let after = staged.get_part(&drawing)?.blob();
        let frames = inventory_of(after)?;
        let index = frames
            .iter()
            .position(|frame| frame.relationship_id == relationship_id)
            .ok_or_else(|| invalid("the added chart frame did not read back"))?;
        self.touch(&worksheet);
        self.touch(&drawing);
        self.touch(&chart);
        self.staged = staged;
        Ok(index)
    }

    /// The chart-space XML of the chart at `index` on a staged worksheet.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or chart.
    pub fn chart(&self, sheet: usize, index: usize) -> Result<Vec<u8>> {
        let (_, drawing, inventory) = self.frames(sheet)?;
        let frame = inventory
            .frames
            .get(index)
            .ok_or_else(|| invalid(format!("chart {index} does not exist")))?;
        let chart = chart_uri(&self.staged, &drawing, &frame.relationship_id)?;
        Ok(self.staged.get_part(&chart)?.blob().to_vec())
    }

    /// Replace the chart-space XML of the chart at `index` on a worksheet.
    /// Returns whether anything changed; identical bytes are a no-op.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or chart, XML that is not a
    /// chart, or a changed edit of a signed workbook.
    pub fn set(&mut self, sheet: usize, index: usize, xml: Vec<u8>) -> Result<bool> {
        check_chart(&xml)?;
        let (_, drawing, inventory) = self.frames(sheet)?;
        let frame = inventory
            .frames
            .get(index)
            .ok_or_else(|| invalid(format!("chart {index} does not exist")))?;
        let chart = chart_uri(&self.staged, &drawing, &frame.relationship_id)?;
        if self.staged.get_part(&chart)?.blob() == xml.as_slice() {
            return Ok(false);
        }
        crate::signature_policy::ensure_unsigned(&self.staged)?;
        self.staged.get_part_mut(&chart)?.set_blob(xml);
        self.touch(&chart);
        Ok(true)
    }

    /// Remove the chart at `index` on a worksheet: its anchor, its drawing
    /// relationship and its chart part (with the parts only it refers to).
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or chart, or a signed workbook.
    pub fn remove(&mut self, sheet: usize, index: usize) -> Result<()> {
        let (_, drawing, inventory) = self.frames(sheet)?;
        let frame = inventory
            .frames
            .get(index)
            .ok_or_else(|| invalid(format!("chart {index} does not exist")))?
            .clone();
        crate::signature_policy::ensure_unsigned(&self.staged)?;
        let chart = chart_uri(&self.staged, &drawing, &frame.relationship_id)?;
        let mut staged = self.staged.clone();
        let part = staged.get_part_mut(&drawing)?;
        let mut bytes = part.blob().to_vec();
        bytes.drain(frame.anchor.clone());
        part.set_blob(bytes);
        part.rels_mut().remove(&frame.relationship_id);
        let mut removed = vec![chart];
        let mut position = 0;
        while let Some(name) = removed.get(position).cloned() {
            position += 1;
            if is_referenced(&staged, &name) {
                continue;
            }
            let targets: Vec<PackURI> = staged
                .get_part(&name)?
                .rels()
                .iter()
                .filter(|relationship| !relationship.is_external())
                .filter_map(|relationship| relationship.target_partname().ok())
                .collect();
            staged.remove_part(&name);
            self.touch(&name);
            removed.extend(targets);
        }
        self.touch(&drawing);
        self.staged = staged;
        Ok(())
    }

    /// Validate the staged package and publish it.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        let patch = Patch::capture(self.package.opc(), &self.staged, &self.touched)?;
        let diagnostics = Diagnostics {
            parts: patch.deltas.len(),
        };
        if patch.is_noop() {
            return Ok(Commit {
                package: self.package.clone(),
                patch,
                diagnostics,
            });
        }
        let package = self
            .package
            .workbook_ref()
            .adopt_published_package(self.staged)
            .map(Package::from_workbook)?;
        Ok(Commit {
            package,
            patch,
            diagnostics,
        })
    }
}

fn inventory_of(xml: &[u8]) -> Result<Vec<Frame>> {
    Ok(inventory(xml)?.frames)
}

/// Whether any part of the package still relates to `name`.
fn is_referenced(package: &OpcPackage, name: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.partname() != name
            && part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship
                        .target_partname()
                        .is_ok_and(|target| &target == name)
            })
    }) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|target| &target == name)
    })
}

fn anchor_xml(anchor: &Anchor, id: u32, name: &str, relationship_id: &str) -> String {
    use litchi_core::xml::escape_xml;
    format!(
        concat!(
            r#"<xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
            "<xdr:from><xdr:col>{fc}</xdr:col><xdr:colOff>{fco}</xdr:colOff><xdr:row>{fr}</xdr:row><xdr:rowOff>{fro}</xdr:rowOff></xdr:from>",
            "<xdr:to><xdr:col>{tc}</xdr:col><xdr:colOff>{tco}</xdr:colOff><xdr:row>{tr}</xdr:row><xdr:rowOff>{tro}</xdr:rowOff></xdr:to>",
            r#"<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{id}" name="{name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"#,
            r#"<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>"#,
            r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">"#,
            r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rel}"/>"#,
            "</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>"
        ),
        fc = anchor.from_col,
        fco = anchor.from_col_offset,
        fr = anchor.from_row,
        fro = anchor.from_row_offset,
        tc = anchor.to_col,
        tco = anchor.to_col_offset,
        tr = anchor.to_row,
        tro = anchor.to_row_offset,
        id = id,
        name = escape_xml(name),
        rel = escape_xml(relationship_id),
    )
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use litchi_drawingml::chart::authoring::{Kind, Series, Spec, Variant};

    fn chart(title: &str) -> Vec<u8> {
        Spec::new(Variant::new(Kind::Line))
            .with_title(title)
            .with_categories(["A", "B"])
            .with_series(Series::new("S", vec![1.0, 2.0]))
            .to_xml(None)
            .unwrap()
    }

    #[test]
    fn add_set_remove_round_trip_and_replay() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_charts().unwrap();
        assert_eq!(
            edit.add(0, &Anchor::new(0, 0, 5, 10), "One", &chart("One"))
                .unwrap(),
            0
        );
        assert_eq!(
            edit.add(0, &Anchor::new(6, 0, 11, 10), "Two", &chart("Two"))
                .unwrap(),
            1
        );
        let commit = edit.commit().unwrap();
        assert!(commit.diagnostics().changed());
        let (after, patch) = commit.into_parts();
        let bytes = after.to_bytes().unwrap();
        let reopened = Package::from_bytes(&bytes).unwrap();
        assert_eq!(reopened.chart_count(0).unwrap(), 2);
        let second = reopened.chart_xml(0, 1).unwrap().unwrap();
        assert!(String::from_utf8_lossy(second).contains("Two"));
        // Deterministic.
        let mut again = package.edit_charts().unwrap();
        again
            .add(0, &Anchor::new(0, 0, 5, 10), "One", &chart("One"))
            .unwrap();
        again
            .add(0, &Anchor::new(6, 0, 11, 10), "Two", &chart("Two"))
            .unwrap();
        assert_eq!(
            again.commit().unwrap().into_package().to_bytes().unwrap(),
            bytes
        );
        // Replay and inverse.
        let replayed = package.apply_charts(&patch).unwrap();
        assert!(replayed.same_content(&after).unwrap());
        let restored = after.apply_charts(&patch.inverse()).unwrap();
        assert!(restored.same_content(&package).unwrap());
        assert!(matches!(
            after.apply_charts(&patch),
            Err(Error::PatchConflict { .. })
        ));

        let mut edit = reopened.edit_charts().unwrap();
        assert!(
            !edit
                .set(0, 0, reopened.chart_xml(0, 0).unwrap().unwrap().to_vec())
                .unwrap()
        );
        assert!(edit.set(0, 0, chart("Renamed")).unwrap());
        assert!(edit.set(0, 0, b"<x/>".to_vec()).is_err());
        edit.remove(0, 1).unwrap();
        let after = edit.commit().unwrap().into_package();
        assert_eq!(after.chart_count(0).unwrap(), 1);
        assert!(
            String::from_utf8_lossy(after.chart_xml(0, 0).unwrap().unwrap()).contains("Renamed")
        );
        assert!(after.chart_xml(0, 1).unwrap().is_none());
    }

    #[test]
    fn a_no_op_edit_keeps_the_package() {
        let package = Package::new().unwrap();
        let edit = package.edit_charts().unwrap();
        let commit = edit.commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
    }

    #[test]
    fn invalid_input_is_refused() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_charts().unwrap();
        assert!(
            edit.add(0, &Anchor::new(5, 5, 1, 1), "x", &chart("x"))
                .is_err()
        );
        assert!(edit.add(0, &Anchor::new(0, 0, 1, 1), "x", b"<x/>").is_err());
        assert!(
            edit.add(9, &Anchor::new(0, 0, 1, 1), "x", &chart("x"))
                .is_err()
        );
        assert!(edit.remove(0, 0).is_err());
    }
}
