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
//! Charts are numbered per worksheet in drawing order, counting the
//! ordinary `c:chart` frames and the extended `cx:chart` frames
//! ([`Package::chart_count`], [`Package::chart_xml`],
//! [`Package::chart_anchor`]). An extended chart part
//! (`litchi_drawingml::chart::chartex`) is staged with its chart-style and
//! color-style companions; [`Edit::set`] keeps a part's family, and
//! [`Edit::set_anchor`] moves or resizes a frame. Chart XML is validated by
//! the `DrawingML` chart (or `ChartEx`) reader before it is staged. A changed edit of a signed workbook is refused with
//! [`Error::Signed`]; an exact no-op keeps every byte.
//!
//! ```
//! use litchi_drawingml::chart::authoring::{Kind, Series, Spec, Variant};
//! use litchi_xlsx::Package;
//! use litchi_xlsx::builder::chart::Anchor;
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
use std::sync::{Arc, Mutex, MutexGuard, Weak};

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

pub use crate::Error;
use crate::Package;
use crate::error::invalid;

use crate::builder::chart::Anchor;
use litchi_drawingml::chart::chartex;

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
            return Err(Error::limit(self, observed, self.maximum()));
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

    /// The semantic payload retained by this reversible chart patch.
    ///
    /// Package endpoints are charged by the caller. Parser and derived-cache
    /// allocations are outside this semantic retention budget.
    #[must_use]
    pub fn retained_bytes(&self) -> usize {
        // State strings are canonical `str::to_owned` fields; charge their
        // logical bytes so metadata-only admission uses the same definition.
        // Native blob buffers retain their actual Vec capacity.
        fn state_bytes(state: &State) -> usize {
            state
                .content_type
                .len()
                .saturating_add(state.blob.capacity())
                .saturating_add(
                    state
                        .relationships
                        .iter()
                        .map(|relationship| {
                            relationship
                                .id
                                .len()
                                .saturating_add(relationship.kind.len())
                                .saturating_add(relationship.target.len())
                        })
                        .sum::<usize>(),
                )
        }

        self.deltas.iter().fold(0usize, |total, delta| {
            total
                .saturating_add(delta.name.as_str().len())
                .saturating_add(delta.before.as_ref().map_or(0, state_bytes))
                .saturating_add(delta.after.as_ref().map_or(0, state_bytes))
        })
    }

    fn estimate_retained_bytes(
        before: &OpcPackage,
        after: &OpcPackage,
        names: &BTreeSet<String>,
    ) -> Result<usize> {
        let mut observed = 0usize;
        for name in names {
            let name = PackURI::new(name.as_str()).map_err(|error| invalid(error.to_string()))?;
            let before = before.get_part(&name).ok();
            let after = after.get_part(&name).ok();
            let changed = match (before, after) {
                (Some(before), Some(after)) => {
                    before.content_type() != after.content_type()
                        || before.blob() != after.blob()
                        || before.rels().len() != after.rels().len()
                        || before.rels().iter().any(|relationship| {
                            after
                                .rels()
                                .get(relationship.r_id())
                                .is_none_or(|candidate| {
                                    candidate.reltype() != relationship.reltype()
                                        || candidate.target_ref() != relationship.target_ref()
                                        || candidate.is_external() != relationship.is_external()
                                })
                        })
                },
                (None, None) => false,
                _ => true,
            };
            if changed {
                observed = observed
                    .saturating_add(name.as_str().len())
                    .saturating_add(before.map_or(0, |part| {
                        part.content_type()
                            .len()
                            .saturating_add(part.blob_arc().capacity())
                    }))
                    .saturating_add(after.map_or(0, |part| {
                        part.content_type()
                            .len()
                            .saturating_add(part.blob_arc().capacity())
                    }));
                if let Some(part) = before {
                    observed = observed.saturating_add(
                        part.rels()
                            .iter()
                            .map(|relationship| {
                                relationship.r_id().len()
                                    + relationship.reltype().len()
                                    + relationship.target_ref().len()
                            })
                            .sum::<usize>(),
                    );
                }
                if let Some(part) = after {
                    observed = observed.saturating_add(
                        part.rels()
                            .iter()
                            .map(|relationship| {
                                relationship.r_id().len()
                                    + relationship.reltype().len()
                                    + relationship.target_ref().len()
                            })
                            .sum::<usize>(),
                    );
                }
            }
        }
        Ok(observed)
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

    /// The semantic payload retained by this commit's reversible patch.
    #[must_use]
    pub fn retained_bytes(&self) -> usize {
        self.patch.retained_bytes()
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
pub struct Edit {
    package: Package,
    staged: OpcPackage,
    touched: BTreeSet<String>,
    max_retained_bytes: Option<usize>,
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

/// Owner-scoped, single-entry memo of the chart inventory of the drawing
/// payload read last (record 0286).
///
/// A path resolution, [`Edit::chart`] and [`Edit::set`] each need the
/// inventory of the same worksheet drawing; this lets them scan it once. It
/// follows the format-table cache (`style::cell_format::cache`): the owner
/// holds it and its lineage shares it, there is no process-wide state; the
/// key is the drawing payload's allocation, held through a [`Weak`]
/// reference, so a changed or replaced drawing never reuses an inventory and
/// the cache never keeps a payload alive; a failed scan is never cached; and
/// an inventory of more than [`MAX_CHARTS`] frames is returned but not
/// retained, which bounds what one cache keeps.
#[derive(Default)]
pub(crate) struct InventoryCache {
    entry: Mutex<Option<(Weak<Vec<u8>>, Arc<Inventory>)>>,
}

impl std::fmt::Debug for InventoryCache {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let frames = self
            .lock()
            .as_ref()
            .map(|(_, inventory)| inventory.frames.len());
        formatter
            .debug_struct("InventoryCache")
            .field("retained_frames", &frames)
            .finish()
    }
}

impl InventoryCache {
    fn lock(&self) -> MutexGuard<'_, Option<(Weak<Vec<u8>>, Arc<Inventory>)>> {
        self.entry
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
    }

    /// The inventory of `drawing`'s payload, scanned on a miss.
    fn get(&self, drawing: &dyn Part) -> Result<Arc<Inventory>> {
        let payload = drawing.blob_arc();
        let mut entry = self.lock();
        // The retained weak reference keeps the allocation's block reserved,
        // so an equal pointer is the same live allocation.
        if let Some((key, inventory)) = entry.as_ref()
            && std::ptr::eq(key.as_ptr(), Arc::as_ptr(&payload))
        {
            return Ok(Arc::clone(inventory));
        }
        *entry = None;
        let inventory = Arc::new(inventory(&payload)?);
        if inventory.frames.len() <= MAX_CHARTS {
            *entry = Some((Arc::downgrade(&payload), Arc::clone(&inventory)));
        }
        Ok(inventory)
    }
}

fn is_chart_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(ns) if ns.as_ref() == CHART_NS
        || ns.as_ref() == STRICT_CHART_NS
        || ns.as_ref() == chartex::NAMESPACE.as_bytes())
}

/// Whether a drawing relationship names a chart: an ordinary one, or an
/// extended (`cx:chartSpace`) one.
fn is_chart_relationship(reltype: &str) -> bool {
    matches!(reltype, rt::CHART | rt::STRICT_CHART) || reltype == chartex::RELATIONSHIP_TYPE
}

/// The content type a chart relationship's target must have.
fn chart_content_type(reltype: &str) -> &'static str {
    if reltype == chartex::RELATIONSHIP_TYPE {
        chartex::CONTENT_TYPE
    } else {
        ct::DML_CHART
    }
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
        // The event borrows the input, not the reader: no copy is needed.
        let (in_chart_namespace, event) = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| invalid(format!("worksheet drawing XML: {error}")))?;
            (is_chart_namespace(&namespace), event)
        };
        let end = position(&reader)?;
        let (element, empty) = match &event {
            Event::Start(element) => (Some(element), false),
            Event::Empty(element) => (Some(element), true),
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

/// The two-cell anchor of a frame: its value, and the byte spans and
/// qualified names of its `from` and `to` markers within the fragment.
struct Marker {
    anchor: Anchor,
    from: std::ops::Range<usize>,
    to: std::ops::Range<usize>,
    prefix: String,
}

/// Read the two-cell anchor of one anchor fragment; `None` for a one-cell
/// or absolute anchor.
fn marker_of(fragment: &[u8]) -> Result<Option<Marker>> {
    use litchi_core::xml::ReaderOrigin;
    let mut reader = NsReader::from_reader(fragment);
    let origin = ReaderOrigin::of(fragment);
    reader.config_mut().trim_text(true);
    let position = |reader: &NsReader<&[u8]>| {
        origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("chart anchor position overflow"))
    };
    let mut depth = 0_usize;
    let mut root_is_two_cell = false;
    let mut marker: Option<(bool, usize)> = None;
    let mut field: Option<String> = None;
    let mut values = [[0_i64; 4]; 2];
    let mut spans = [None, None];
    let mut prefix = String::new();
    loop {
        let start = position(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("chart anchor XML: {error}")))?;
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                let local = element.local_name();
                if depth == 0 {
                    root_is_two_cell = local.as_ref() == b"twoCellAnchor";
                } else if depth == 1 && matches!(local.as_ref(), b"from" | b"to") {
                    let name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
                    prefix = name
                        .split_once(':')
                        .map(|(p, _)| p.to_owned())
                        .unwrap_or_default();
                    marker = Some((local.as_ref() == b"to", start));
                } else if depth == 2 && marker.is_some() {
                    field = Some(String::from_utf8_lossy(local.as_ref()).into_owned());
                }
                depth += 1;
            },
            Event::Text(text) => {
                if let (Some((to, _)), Some(name)) = (marker, &field) {
                    let text = text
                        .decode()
                        .map_err(|error| invalid(format!("chart anchor text: {error}")))?;
                    let slot = match name.as_str() {
                        "col" => 0,
                        "colOff" => 1,
                        "row" => 2,
                        "rowOff" => 3,
                        _ => continue,
                    };
                    values[usize::from(to)][slot] = text
                        .trim()
                        .parse::<i64>()
                        .map_err(|_| invalid("chart anchor marker is not a number"))?;
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 2 {
                    field = None;
                } else if depth == 1
                    && let Some((to, begin)) = marker.take()
                {
                    spans[usize::from(to)] = Some(begin..end);
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let (Some(from), Some(to)) = (spans[0].clone(), spans[1].clone()) else {
        return Ok(None);
    };
    if !root_is_two_cell {
        return Ok(None);
    }
    let cell =
        |value: i64| u32::try_from(value).map_err(|_| invalid("chart anchor cell is out of range"));
    let [f, t] = values;
    Ok(Some(Marker {
        anchor: Anchor {
            from_col: cell(f[0])?,
            from_col_offset: f[1],
            from_row: cell(f[2])?,
            from_row_offset: f[3],
            to_col: cell(t[0])?,
            to_col_offset: t[1],
            to_row: cell(t[2])?,
            to_row_offset: t[3],
        },
        from,
        to,
        prefix,
    }))
}

fn marker_xml(
    prefix: &str,
    element: &str,
    col: u32,
    col_offset: i64,
    row: u32,
    row_offset: i64,
) -> String {
    let q = |local: &str| {
        if prefix.is_empty() {
            local.to_owned()
        } else {
            format!("{prefix}:{local}")
        }
    };
    format!(
        "<{e}><{c}>{col}</{c}><{co}>{col_offset}</{co}><{r}>{row}</{r}><{ro}>{row_offset}</{ro}></{e}>",
        e = q(element),
        c = q("col"),
        co = q("colOff"),
        r = q("row"),
        ro = q("rowOff"),
    )
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
    if relationship.is_external() || !is_chart_relationship(relationship.reltype()) {
        return Err(invalid(
            "drawing chart relationship is not an internal chart",
        ));
    }
    let target = relationship.target_partname()?;
    if package.get_part(&target)?.content_type() != chart_content_type(relationship.reltype()) {
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

/// Validate a chart part; `true` for an extended (`cx:chartSpace`) one.
fn check_chart(xml: &[u8]) -> Result<bool> {
    LimitKind::ChartBytes.check(xml.len())?;
    let mut drawing_limits = litchi_drawingml::chart::Limits::DEFAULT;
    drawing_limits.max_input_bytes = MAX_CHART_BYTES.min(drawing_limits.max_input_bytes);
    drawing_limits.max_output_bytes = MAX_CHART_BYTES.min(drawing_limits.max_output_bytes);
    if chartex::is_chartex(xml) {
        chartex::read(xml, &drawing_limits)?;
        return Ok(true);
    }
    litchi_drawingml::chart::reader::read(xml, &drawing_limits)?;
    Ok(false)
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

/// The two-cell anchor of the chart at `index` on a source-backed
/// worksheet part.
pub(crate) fn source_chart_anchor(
    package: &litchi_opc::SourceBackedPackage,
    worksheet: &PackURI,
    index: usize,
) -> Result<Option<Anchor>> {
    package.check_execution()?;
    let Some(drawing) = drawing_in(package.part(worksheet)?.rels())? else {
        return Ok(None);
    };
    let xml = source_bytes(package, &drawing)?;
    let frames = inventory(&xml)?.frames;
    let Some(frame) = frames.get(index) else {
        return Ok(None);
    };
    Ok(marker_of(&xml[frame.anchor.clone()])?.map(|marker| marker.anchor))
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
    if relationship.is_external() || !is_chart_relationship(relationship.reltype()) {
        return Err(invalid(
            "drawing chart relationship is not an internal chart",
        ));
    }
    let chart = relationship.target_partname()?;
    if package.part(&chart)?.content_type() != chart_content_type(relationship.reltype()) {
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
        Ok(self.chart_inventory(&drawing)?.frames.len())
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
        let inventory = self.chart_inventory(&drawing)?;
        let Some(frame) = inventory.frames.get(index) else {
            return Ok(None);
        };
        let chart = chart_uri(self.opc(), &drawing, &frame.relationship_id)?;
        Ok(Some(self.opc().get_part(&chart)?.blob()))
    }

    /// The two-cell anchor of the chart at `index` (drawing order) on a
    /// worksheet; `None` past the last chart or for a chart placed by a
    /// one-cell or absolute anchor.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet, a chartsheet, or a malformed
    /// drawing.
    pub fn chart_anchor(&self, sheet: usize, index: usize) -> Result<Option<Anchor>> {
        let worksheet = worksheet_uri(self, sheet)?;
        let Some(drawing) = drawing_uri(self.opc(), &worksheet)? else {
            return Ok(None);
        };
        let xml = self.opc().get_part(&drawing)?.blob();
        let inventory = self.chart_inventory(&drawing)?;
        let Some(frame) = inventory.frames.get(index) else {
            return Ok(None);
        };
        Ok(marker_of(&xml[frame.anchor.clone()])?.map(|marker| marker.anchor))
    }

    /// The chart inventory of the drawing part `drawing`, through the
    /// workbook lineage's memo.
    fn chart_inventory(&self, drawing: &PackURI) -> Result<Arc<Inventory>> {
        self.workbook_ref()
            .chart_inventories()
            .get(self.opc().get_part(drawing)?)
    }

    /// Start a worksheet chart edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused.
    pub fn edit_charts(&self) -> Result<Edit> {
        self.edit_charts_with_options(crate::EditOptions::default())
    }

    /// Start a chart edit with an explicit retained-patch budget.
    pub fn edit_charts_with_options(&self, options: crate::EditOptions) -> Result<Edit> {
        self.workbook_ref().ensure_mutation_allowed("edit_charts")?;
        Ok(Edit {
            package: self.clone(),
            staged: self.opc().clone(),
            touched: BTreeSet::new(),
            max_retained_bytes: options.max_retained_bytes_value(),
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
    pub fn apply_charts(&self, patch: &Patch) -> Result<Commit> {
        let diagnostics = Diagnostics {
            parts: patch.deltas.len(),
        };
        let package = if patch.is_noop() {
            self.clone()
        } else {
            self.workbook_ref()
                .ensure_mutation_allowed("apply_charts")?;
            crate::signature_policy::ensure_unsigned(self.opc())?;
            let mut candidate = self.opc().clone();
            patch.apply(&mut candidate)?;
            self.derive(candidate)?
        };
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics,
        })
    }
}

impl Edit {
    fn touch(&mut self, name: &PackURI) {
        self.touched.insert(name.as_str().to_owned());
    }

    fn frames(&self, sheet: usize) -> Result<(PackURI, PackURI, Arc<Inventory>)> {
        let worksheet = worksheet_uri(&self.package, sheet)?;
        let drawing = drawing_uri(&self.staged, &worksheet)?
            .ok_or_else(|| invalid("the worksheet has no charts"))?;
        let inventory = self
            .package
            .workbook_ref()
            .chart_inventories()
            .get(self.staged.get_part(&drawing)?)?;
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
        let extended = check_chart(xml)?;
        if name.contains('\0') {
            return Err(invalid("chart frame name contains a NUL character"));
        }
        crate::signature_policy::ensure_unsigned(&self.staged)?;
        let worksheet = worksheet_uri(&self.package, sheet)?;
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
        let (stem, content_type, relationship_type) = if extended {
            (
                "/xl/charts/chartEx",
                chartex::CONTENT_TYPE,
                chartex::RELATIONSHIP_TYPE,
            )
        } else {
            ("/xl/charts/chart", ct::DML_CHART, rt::CHART)
        };
        let chart = free_name(&staged, stem, ".xml")?;
        let mut chart_part = BlobPart::new(chart.clone(), content_type.to_owned(), xml.to_vec());
        if extended {
            // An extended chart needs its style and color companions.
            for (stem, content_type, relationship_type, bytes) in [
                (
                    "/xl/charts/style",
                    chartex::STYLE_CONTENT_TYPE,
                    chartex::STYLE_RELATIONSHIP_TYPE,
                    chartex::style_xml(),
                ),
                (
                    "/xl/charts/colors",
                    chartex::COLORS_CONTENT_TYPE,
                    chartex::COLORS_RELATIONSHIP_TYPE,
                    chartex::colors_xml(),
                ),
            ] {
                let companion = free_name(&staged, stem, ".xml")?;
                chart_part.relate_to(&companion.relative_ref(chart.base_uri()), relationship_type);
                staged.add_part(Box::new(BlobPart::new(
                    companion.clone(),
                    content_type.to_owned(),
                    bytes,
                )));
                self.touch(&companion);
            }
        }
        staged.add_part(Box::new(chart_part));
        let target = chart.relative_ref(drawing.base_uri());
        let part = staged.get_part_mut(&drawing)?;
        let relationship_id = part.relate_to(&target, relationship_type);
        let id = inventory
            .max_id
            .checked_add(1)
            .ok_or_else(|| invalid("drawing shape ids are exhausted"))?;
        let fragment = anchor_xml(anchor, id, name, &relationship_id, extended);
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
        let extended = check_chart(&xml)?;
        let (_, drawing, inventory) = self.frames(sheet)?;
        let frame = inventory
            .frames
            .get(index)
            .ok_or_else(|| invalid(format!("chart {index} does not exist")))?;
        let chart = chart_uri(&self.staged, &drawing, &frame.relationship_id)?;
        if (self.staged.get_part(&chart)?.content_type() == chartex::CONTENT_TYPE) != extended {
            return Err(invalid(
                "a chart part keeps its family: ordinary and extended charts do not replace each other",
            ));
        }
        if self.staged.get_part(&chart)?.blob() == xml.as_slice() {
            return Ok(false);
        }
        crate::signature_policy::ensure_unsigned(&self.staged)?;
        self.staged.get_part_mut(&chart)?.set_blob(xml);
        self.touch(&chart);
        Ok(true)
    }

    /// Move or resize the chart at `index` on a worksheet: replace the cells
    /// and offsets of its two-cell anchor. Returns whether anything changed;
    /// the frame, its name and its chart part are kept.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or chart, an invalid anchor, a
    /// chart placed by a one-cell or absolute anchor, or a changed edit of
    /// a signed workbook.
    pub fn set_anchor(&mut self, sheet: usize, index: usize, anchor: &Anchor) -> Result<bool> {
        validate_anchor(anchor)?;
        let (_, drawing, inventory) = self.frames(sheet)?;
        let frame = inventory
            .frames
            .get(index)
            .ok_or_else(|| invalid(format!("chart {index} does not exist")))?;
        let xml = self.staged.get_part(&drawing)?.blob();
        let fragment = &xml[frame.anchor.clone()];
        let marker = marker_of(fragment)?
            .ok_or_else(|| invalid("only a two-cell chart anchor can be moved"))?;
        if marker.anchor == *anchor {
            return Ok(false);
        }
        crate::signature_policy::ensure_unsigned(&self.staged)?;
        let from = marker_xml(
            &marker.prefix,
            "from",
            anchor.from_col,
            anchor.from_col_offset,
            anchor.from_row,
            anchor.from_row_offset,
        );
        let to = marker_xml(
            &marker.prefix,
            "to",
            anchor.to_col,
            anchor.to_col_offset,
            anchor.to_row,
            anchor.to_row_offset,
        );
        let base = frame.anchor.start;
        let mut bytes = xml.to_vec();
        // The `to` marker follows the `from` one: replace it first.
        bytes.splice(base + marker.to.start..base + marker.to.end, to.bytes());
        bytes.splice(
            base + marker.from.start..base + marker.from.end,
            from.bytes(),
        );
        self.staged.get_part_mut(&drawing)?.set_blob(bytes);
        self.touch(&drawing);
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
        let mut endpoint_charge = None;
        let target = if let Some(maximum) = self.max_retained_bytes {
            let estimated_patch =
                Patch::estimate_retained_bytes(self.package.opc(), &self.staged, &self.touched)?;
            let target = if estimated_patch == 0 {
                self.package.clone()
            } else {
                self.package
                    .workbook_ref()
                    .adopt_published_package(self.staged.clone())
                    .map(Package::from_workbook)?
            };
            let source_bytes = self.package.retained_bytes();
            let target_bytes = target.retained_bytes();
            let observed = source_bytes
                .saturating_add(target_bytes)
                .saturating_add(estimated_patch);
            if observed > maximum {
                return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
                    resource: litchi_core::Resource::Memory,
                    observed: u64::try_from(observed).unwrap_or(u64::MAX),
                    limit: u64::try_from(maximum).unwrap_or(u64::MAX),
                    scope: "XLSX chart retained package".into(),
                }));
            }
            endpoint_charge = Some((source_bytes, target_bytes));
            Some(target)
        } else {
            None
        };
        let patch = Patch::capture(self.package.opc(), &self.staged, &self.touched)?;
        let diagnostics = Diagnostics {
            parts: patch.deltas.len(),
        };
        if patch.is_noop() {
            return Ok(Commit {
                package: self.package,
                patch,
                diagnostics,
            });
        }
        let package = match target {
            Some(target) => target,
            None => self
                .package
                .workbook_ref()
                .adopt_published_package(self.staged)
                .map(Package::from_workbook)?,
        };
        if let Some(maximum) = self.max_retained_bytes {
            let (source_bytes, target_bytes) =
                endpoint_charge.expect("retained endpoint charge established with the budget");
            let observed = source_bytes
                .saturating_add(target_bytes)
                .saturating_add(patch.retained_bytes());
            if observed > maximum {
                return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
                    resource: litchi_core::Resource::Memory,
                    observed: u64::try_from(observed).unwrap_or(u64::MAX),
                    limit: u64::try_from(maximum).unwrap_or(u64::MAX),
                    scope: "XLSX chart retained package".into(),
                }));
            }
        }
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

fn anchor_xml(
    anchor: &Anchor,
    id: u32,
    name: &str,
    relationship_id: &str,
    extended: bool,
) -> String {
    use litchi_core::xml::escape_xml;
    let (uri, prefix) = if extended {
        (chartex::GRAPHIC_DATA_URI, "cx")
    } else {
        (
            "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "c",
        )
    };
    format!(
        concat!(
            r#"<xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
            "<xdr:from><xdr:col>{fc}</xdr:col><xdr:colOff>{fco}</xdr:colOff><xdr:row>{fr}</xdr:row><xdr:rowOff>{fro}</xdr:rowOff></xdr:from>",
            "<xdr:to><xdr:col>{tc}</xdr:col><xdr:colOff>{tco}</xdr:colOff><xdr:row>{tr}</xdr:row><xdr:rowOff>{tro}</xdr:rowOff></xdr:to>",
            r#"<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{id}" name="{name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"#,
            r#"<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>"#,
            r#"<a:graphic><a:graphicData uri="{uri}">"#,
            r#"<{prefix}:chart xmlns:{prefix}="{uri}" "#,
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
        uri = uri,
        prefix = prefix,
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
        let replayed = package.apply_charts(&patch).unwrap().into_package();
        assert!(replayed.same_content(&after).unwrap());
        let restored = after.apply_charts(&patch.inverse()).unwrap().into_package();
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
    fn retained_patch_budget_rejects_one_below_and_accepts_exact() {
        let package = Package::new().unwrap();
        let xml = chart("Budget");

        let mut probe = package.edit_charts().unwrap();
        probe
            .add(0, &Anchor::new(0, 0, 5, 10), "Budget", &xml)
            .unwrap();
        let (after, patch) = probe.commit().unwrap().into_parts();
        let required = package
            .retained_bytes()
            .saturating_add(after.retained_bytes())
            .saturating_add(patch.retained_bytes());
        assert!(required > 0);

        let mut rejected = package
            .edit_charts_with_options(
                crate::EditOptions::default().with_max_retained_bytes(required - 1),
            )
            .unwrap();
        rejected
            .add(0, &Anchor::new(0, 0, 5, 10), "Budget", &xml)
            .unwrap();
        assert!(matches!(
            rejected.commit(),
            Err(Error::ResourceLimit(limit)) if limit.limit == (required - 1) as u64
        ));

        let mut exact = package
            .edit_charts_with_options(
                crate::EditOptions::default().with_max_retained_bytes(required),
            )
            .unwrap();
        exact
            .add(0, &Anchor::new(0, 0, 5, 10), "Budget", &xml)
            .unwrap();
        let commit = exact.commit().unwrap();
        assert_eq!(
            package
                .retained_bytes()
                .saturating_add(commit.package().retained_bytes())
                .saturating_add(commit.patch().retained_bytes()),
            required
        );
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

    #[test]
    fn extended_charts_and_anchor_moves_round_trip() {
        let package = Package::new().unwrap();
        let xml = Spec::new(Variant::new(Kind::Funnel))
            .with_title("Pipeline")
            .with_categories(["Leads", "Won"])
            .with_series(Series::new("Count", vec![10.0, 2.0]))
            .to_xml(None)
            .unwrap();
        let mut edit = package.edit_charts().unwrap();
        let classic = edit
            .add(0, &Anchor::new(0, 0, 4, 8), "Classic", &chart("A"))
            .unwrap();
        let extended = edit
            .add(0, &Anchor::new(5, 0, 9, 8), "Funnel", &xml)
            .unwrap();
        assert_eq!((classic, extended), (0, 1));
        let after = edit.commit().unwrap().into_package();
        let reopened = Package::from_bytes(&after.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.chart_count(0).unwrap(), 2);
        assert_eq!(reopened.chart_xml(0, 1).unwrap().unwrap(), xml.as_slice());
        let types: Vec<String> = reopened
            .opc()
            .iter_parts()
            .map(|part| part.content_type().to_owned())
            .collect();
        for expected in [
            chartex::CONTENT_TYPE,
            chartex::STYLE_CONTENT_TYPE,
            chartex::COLORS_CONTENT_TYPE,
        ] {
            assert!(types.iter().any(|ct| ct == expected), "{expected}");
        }
        assert_eq!(
            reopened.chart_anchor(0, 1).unwrap(),
            Some(Anchor::new(5, 0, 9, 8))
        );
        // A part keeps its family.
        let mut edit = reopened.edit_charts().unwrap();
        assert!(edit.set(0, 1, chart("B")).is_err());
        assert!(edit.set(0, 0, xml.clone()).is_err());
        // Moving keeps the frame and its chart.
        assert!(!edit.set_anchor(0, 1, &Anchor::new(5, 0, 9, 8)).unwrap());
        assert!(edit.set_anchor(0, 1, &Anchor::new(2, 3, 12, 20)).unwrap());
        assert!(edit.set_anchor(0, 1, &Anchor::new(9, 0, 2, 8)).is_err());
        let moved = edit.commit().unwrap().into_package();
        assert_eq!(
            moved.chart_anchor(0, 1).unwrap(),
            Some(Anchor::new(2, 3, 12, 20))
        );
        assert_eq!(moved.chart_xml(0, 1).unwrap().unwrap(), xml.as_slice());
        let source = crate::SourceBackedWorkbook::from_reader(std::io::Cursor::new(
            moved.to_bytes().unwrap(),
        ))
        .unwrap();
        let sheet = source.sheet(0).unwrap().unwrap();
        assert_eq!(sheet.chart_count().unwrap(), 2);
        assert_eq!(
            sheet.chart_anchor(1).unwrap(),
            Some(Anchor::new(2, 3, 12, 20))
        );
        assert_eq!(sheet.chart_xml(1).unwrap().unwrap(), xml);
        // Removing the extended chart drops its companions.
        let mut edit = moved.edit_charts().unwrap();
        edit.remove(0, 1).unwrap();
        let removed = edit.commit().unwrap().into_package();
        assert!(!removed.opc().iter_parts().any(|part| {
            [
                chartex::CONTENT_TYPE,
                chartex::STYLE_CONTENT_TYPE,
                chartex::COLORS_CONTENT_TYPE,
            ]
            .contains(&part.content_type())
        }));
        assert_eq!(removed.chart_count(0).unwrap(), 1);
    }
}
