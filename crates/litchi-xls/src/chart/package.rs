//! Workbook/container hosting and transactional chart edits.

#[cfg(test)]
use std::collections::HashMap;
use std::collections::HashSet;
use std::sync::Arc;

use litchi_biff::{Limits as BiffLimits, Records};
use litchi_ograph::chart::{Kind as GraphChartKind, Ref as GraphChartRef, Refs as GraphCharts};
use litchi_ole_common::object::{Editor as ObjectEditor, Limits as ObjectLimits, Targets};

use super::codec::parse_chart;
#[cfg(test)]
use super::codec::serialize_chart;
use super::model::*;
use super::wire::*;
use crate::{Error, Result};

/// Chart plus its current workbook host location.
#[derive(Clone, Debug, PartialEq)]
pub struct Entry {
    /// Stable location within the current workbook revision.
    pub location: Location,
    /// Owned chart model.
    pub chart: Chart,
}

#[derive(Clone)]
struct StoredChart {
    entry: Entry,
    #[cfg(test)]
    start: usize,
    #[cfg(test)]
    end: usize,
    #[cfg(test)]
    object: Option<(usize, usize)>,
}

const UNSUPPORTED_AUTHORING_REASON: &str = "fresh and replacement XLS chart authoring requires the complete Office-compatible BIFF chart grammar";
const UNSUPPORTED_EMBEDDED_MUTATION_REASON: &str = "embedded XLS chart mutation requires complete MsoDrawing/Continue, Obj/Continue, chart-substream, and OfficeArt drawing-group ownership";

pub(crate) fn unsupported_authoring<T>() -> Result<T> {
    Err(litchi_ograph::Error::UnsupportedAuthoring {
        reason: UNSUPPORTED_AUTHORING_REASON,
    }
    .into())
}

pub(crate) fn unsupported_embedded_mutation<T>() -> Result<T> {
    Err(litchi_ograph::Error::UnsupportedMutation {
        operation: "embedded XLS chart drawing",
        reason: UNSUPPORTED_EMBEDDED_MUTATION_REASON,
    }
    .into())
}

/// Transactional editor for existing BIFF8 chart substreams.
pub struct Editor {
    pub(super) package: ObjectEditor,
    pub(super) workbook_path: Vec<String>,
    pub(super) workbook: Arc<[u8]>,
    limits: Limits,
    charts: Vec<StoredChart>,
}

impl Editor {
    /// Takes ownership of an XLS compound file and validates its chart inventory.
    pub fn open(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        validate_limits(limits)?;
        let package = ObjectEditor::open(bytes, Targets::default(), ObjectLimits::default())?;
        let workbook_path = [vec!["Workbook".into()], vec!["Book".into()]]
            .into_iter()
            .find(|path| package.stream(path).is_some())
            .ok_or_else(|| Error::InvalidData("Workbook stream not found".into()))?;
        let workbook = package
            .stream_shared(&workbook_path)
            .ok_or_else(|| Error::InvalidData("selected Workbook stream disappeared".into()))?;
        if workbook.len() > limits.max_workbook_bytes {
            return invalid(CHART, "Workbook stream exceeds chart editor limit");
        }
        let charts = parse_workbook_charts(&workbook, limits)?;
        Ok(Self {
            package,
            workbook_path,
            workbook,
            limits,
            charts,
        })
    }

    /// Iterates borrowed chart entries in workbook drawing order.
    pub fn charts(&self) -> impl ExactSizeIterator<Item = &Entry> {
        self.charts.iter().map(|value| &value.entry)
    }

    /// Consume the editor and return the parsed chart inventory without persisting.
    pub fn into_charts(self) -> Vec<Entry> {
        self.charts.into_iter().map(|value| value.entry).collect()
    }

    /// Looks up a chart by worksheet name and semantic position.
    pub fn get(&self, selector: Selector<'_>) -> Result<Option<&Chart>> {
        let Some(location) = self.resolve(selector)? else {
            return Ok(None);
        };
        Ok(self.at(&location))
    }

    /// Looks up a chart using a checked low-level host location.
    pub fn at(&self, location: &Location) -> Option<&Chart> {
        self.charts
            .iter()
            .find(|value| &value.entry.location == location)
            .map(|value| &value.entry.chart)
    }

    /// Refuses fresh embedded-chart authoring until its complete BIFF grammar is available.
    ///
    /// The editor is unchanged when this returns
    /// [`litchi_ograph::Error::UnsupportedAuthoring`] through [`Error::Graph`].
    pub fn add(&mut self, _sheet: &str, _chart: Chart) -> Result<Location> {
        unsupported_authoring()
    }

    /// Refuses fresh embedded-chart authoring at a checked raw host location.
    ///
    /// The editor is unchanged when this returns
    /// [`litchi_ograph::Error::UnsupportedAuthoring`] through [`Error::Graph`].
    pub fn insert_at(
        &mut self,
        _sheet_index: usize,
        _object_id: u16,
        _index: usize,
        _chart: Chart,
    ) -> Result<()> {
        unsupported_authoring()
    }

    #[cfg(test)]
    pub(crate) fn insert_fixture_at(
        &mut self,
        sheet_index: usize,
        object_id: u16,
        index: usize,
        chart: Chart,
    ) -> Result<()> {
        let (_, sheets) = bindings(&self.workbook)?;
        let sheet = sheets
            .iter()
            .find(|value| value.index == sheet_index)
            .ok_or_else(|| invalid_error(BOUNDSHEET, "worksheet index was not found"))?;
        if sheet.kind != 0 {
            return invalid(BOUNDSHEET, "embedded charts require a worksheet tab");
        }
        if object_id == 0 || sheet_object_ids(&self.workbook, sheet)?.contains(&object_id) {
            return invalid(OBJ, "embedded chart object ID is zero or duplicated");
        }
        chart.validate(self.limits)?;
        let original = self.charts.clone();
        let mut desired = original.iter().map(|v| v.entry.clone()).collect::<Vec<_>>();
        let positions = desired
            .iter()
            .enumerate()
            .filter_map(|(i, value)| (value.location.sheet_index() == sheet_index).then_some(i))
            .collect::<Vec<_>>();
        if index > positions.len() {
            return invalid(CHART, "embedded chart insertion index is out of range");
        }
        let insert = positions
            .get(index)
            .copied()
            .unwrap_or_else(|| positions.last().map_or(desired.len(), |v| v + 1));
        desired.insert(
            insert,
            Entry {
                location: Location::Embedded {
                    sheet_index,
                    object_id,
                },
                chart,
            },
        );
        self.commit_fixture(&original, desired)
    }

    /// Refuses fresh chart-sheet authoring until its complete BIFF grammar is available.
    ///
    /// The editor is unchanged when this returns
    /// [`litchi_ograph::Error::UnsupportedAuthoring`] through [`Error::Graph`].
    pub fn add_sheet(&mut self, _name: impl Into<String>, _chart: Chart) -> Result<()> {
        unsupported_authoring()
    }

    /// Refuses fresh chart-sheet authoring at a checked raw tab index.
    ///
    /// The editor is unchanged when this returns
    /// [`litchi_ograph::Error::UnsupportedAuthoring`] through [`Error::Graph`].
    pub fn insert_sheet_at(
        &mut self,
        _index: usize,
        _name: impl Into<String>,
        _chart: Chart,
    ) -> Result<()> {
        unsupported_authoring()
    }

    /// Remove a chart-sheet tab. References to that tab cause atomic failure.
    pub fn remove_sheet_at(&mut self, sheet_index: usize) -> Result<Chart> {
        let (_, sheets) = bindings(&self.workbook)?;
        let sheet = sheets
            .iter()
            .find(|value| value.index == sheet_index)
            .ok_or_else(|| invalid_error(BOUNDSHEET, "sheet index was not found"))?;
        if sheet.kind != 2 {
            return invalid(BOUNDSHEET, "selected tab is not a chart sheet");
        }
        let chart_index = self
            .charts
            .iter()
            .position(|value| value.entry.location == Location::ChartSheet { sheet_index })
            .ok_or_else(|| invalid_error(CHART, "chart sheet has no chart"))?;
        let order = (0..sheets.len())
            .filter(|value| *value != sheet_index)
            .map(Some)
            .collect::<Vec<_>>();
        let workbook = rewrite_sheet_directory(&self.workbook, &order, None)?;
        let mut previous = self.install_workbook(workbook)?;
        if chart_index >= previous.len() {
            return invalid(CHART, "removed chart-sheet inventory changed unexpectedly");
        }
        Ok(previous.swap_remove(chart_index).entry.chart)
    }

    /// Reorders workbook tabs by Unicode case-insensitive tab names.
    pub fn reorder_sheets(&mut self, order: &[&str]) -> Result<()> {
        let (_, sheets) = bindings(&self.workbook)?;
        if order.len() != sheets.len() {
            return invalid(BOUNDSHEET, "sheet reorder must contain every tab");
        }
        let mut indexes = Vec::new();
        indexes
            .try_reserve(order.len())
            .map_err(|_| Error::InvalidData("could not allocate sheet reorder".into()))?;
        let mut seen = HashSet::new();
        for name in order {
            let sheet = sheets
                .iter()
                .find(|value| names_equal(&value.name, name))
                .ok_or_else(|| invalid_error(BOUNDSHEET, "sheet reorder name was not found"))?;
            if !seen.insert(sheet.index) {
                return invalid(BOUNDSHEET, "sheet reorder repeats a tab name");
            }
            indexes.push(sheet.index);
        }
        self.reorder_sheets_at(&indexes)
    }

    /// Reorders all workbook tabs by checked previous zero-based indexes.
    pub fn reorder_sheets_at(&mut self, order: &[usize]) -> Result<()> {
        let count = bindings(&self.workbook)?.1.len();
        if order.len() != count {
            return invalid(BOUNDSHEET, "sheet reorder must contain every tab");
        }
        let mut seen = HashSet::new();
        if order
            .iter()
            .any(|value| *value >= count || !seen.insert(*value))
        {
            return invalid(
                BOUNDSHEET,
                "sheet reorder contains an invalid or repeated tab",
            );
        }
        let workbook = rewrite_sheet_directory(
            &self.workbook,
            &order.iter().copied().map(Some).collect::<Vec<_>>(),
            None,
        )?;
        self.install_workbook(workbook).map(drop)
    }

    /// Refuses replacement authoring until its complete BIFF grammar is available.
    ///
    /// The editor is unchanged when this returns
    /// [`litchi_ograph::Error::UnsupportedAuthoring`] through [`Error::Graph`].
    pub fn replace(&mut self, _selector: Selector<'_>, _chart: Chart) -> Result<()> {
        unsupported_authoring()
    }

    /// Refuses replacement authoring at a checked low-level host location.
    ///
    /// The editor is unchanged when this returns
    /// [`litchi_ograph::Error::UnsupportedAuthoring`] through [`Error::Graph`].
    pub fn replace_at(&mut self, _location: &Location, _chart: Chart) -> Result<()> {
        unsupported_authoring()
    }

    /// Removes a chart sheet transactionally and refuses embedded-chart removal.
    ///
    /// Embedded charts participate in the worksheet OfficeArt drawing graph;
    /// until that complete ownership is modeled, the editor returns
    /// [`litchi_ograph::Error::UnsupportedMutation`] without mutation.
    pub fn remove(&mut self, selector: Selector<'_>) -> Result<Chart> {
        let location = self
            .resolve(selector)?
            .ok_or_else(|| invalid_error(CHART, "chart selector was not found"))?;
        self.remove_at(&location)
    }

    /// Removes a chart sheet using a checked low-level host location.
    ///
    /// An existing embedded location is validated and then refused atomically
    /// until its complete OfficeArt drawing ownership can be rewritten.
    pub fn remove_at(&mut self, location: &Location) -> Result<Chart> {
        if let Location::ChartSheet { sheet_index } = location {
            return self.remove_sheet_at(*sheet_index);
        }
        if self
            .charts
            .iter()
            .all(|value| &value.entry.location != location)
        {
            return Err(invalid_error(CHART, "chart location was not found"));
        }
        unsupported_embedded_mutation()
    }

    /// Validates embedded-chart order on a named worksheet.
    ///
    /// The current identity order is a no-op. A structural reorder is refused
    /// atomically until complete OfficeArt drawing ownership is modeled.
    pub fn reorder(&mut self, sheet: &str, order: &[usize]) -> Result<()> {
        let (_, sheets) = bindings(&self.workbook)?;
        let sheet = sheets
            .iter()
            .find(|value| names_equal(&value.name, sheet))
            .ok_or_else(|| invalid_error(BOUNDSHEET, "worksheet name was not found"))?;
        let ids = self
            .charts
            .iter()
            .filter_map(|value| match value.entry.location {
                Location::Embedded {
                    sheet_index,
                    object_id,
                } if sheet_index == sheet.index => Some(object_id),
                _ => None,
            })
            .collect::<Vec<_>>();
        if order.len() != ids.len() {
            return invalid(CHART, "chart reorder must contain every chart");
        }
        let mut seen = HashSet::new();
        let mut object_ids = Vec::new();
        object_ids
            .try_reserve(order.len())
            .map_err(|_| Error::InvalidData("could not allocate chart reorder".into()))?;
        for index in order {
            let id = ids
                .get(*index)
                .copied()
                .filter(|_| seen.insert(*index))
                .ok_or_else(|| {
                    invalid_error(CHART, "chart reorder index is invalid or repeated")
                })?;
            object_ids.push(id);
        }
        self.reorder_at(sheet.index, &object_ids)
    }

    /// Validates embedded-chart order using checked worksheet and Obj identifiers.
    ///
    /// The current identity order is a no-op. A structural reorder is refused
    /// atomically until complete OfficeArt drawing ownership is modeled.
    pub fn reorder_at(&mut self, sheet_index: usize, object_ids: &[u16]) -> Result<()> {
        let (_, sheets) = bindings(&self.workbook)?;
        let sheet = sheets
            .iter()
            .find(|value| value.index == sheet_index)
            .ok_or_else(|| invalid_error(BOUNDSHEET, "worksheet index was not found"))?;
        if sheet.kind != 0 {
            return invalid(BOUNDSHEET, "embedded charts require a worksheet tab");
        }
        let slots = self
            .charts
            .iter()
            .enumerate()
            .filter_map(|(i, value)| match value.entry.location {
                Location::Embedded {
                    sheet_index: sheet, ..
                } if sheet == sheet_index => Some(i),
                _ => None,
            })
            .collect::<Vec<_>>();
        if slots.len() != object_ids.len() {
            return invalid(
                CHART,
                "reorder must include every embedded chart on the worksheet",
            );
        }
        let mut available = slots.clone();
        let mut current = Vec::new();
        current
            .try_reserve_exact(slots.len())
            .map_err(|_| Error::InvalidData("could not allocate chart reorder".into()))?;
        for index in &slots {
            let object_id = self
                .charts
                .get(*index)
                .and_then(|value| match value.entry.location {
                    Location::Embedded { object_id, .. } => Some(object_id),
                    Location::ChartSheet { .. } => None,
                })
                .ok_or_else(|| invalid_error(CHART, "chart reorder slot is invalid"))?;
            current.push(object_id);
        }
        for id in object_ids {
            let position = available
                .iter()
                .position(|index| {
                    self.charts.get(*index).is_some_and(|value| {
                        matches!(
                            value.entry.location,
                            Location::Embedded { object_id, .. } if object_id == *id
                        )
                    })
                })
                .ok_or_else(|| {
                    invalid_error(CHART, "reorder contains an unknown or repeated object ID")
                })?;
            available.remove(position);
        }
        if current == object_ids {
            return Ok(());
        }
        unsupported_embedded_mutation()
    }

    /// Consumes the editor and returns the rewritten compound-file allocation.
    pub fn finish(self) -> Result<Vec<u8>> {
        self.package.finish().map_err(Into::into)
    }

    #[cfg(test)]
    fn commit_fixture(&mut self, original: &[StoredChart], desired: Vec<Entry>) -> Result<()> {
        if desired.len() > self.limits.max_charts {
            return invalid(CHART, "chart count exceeds limit");
        }
        let workbook = rewrite_workbook_charts(&self.workbook, original, &desired, self.limits)?;
        let reparsed = parse_workbook_charts(&workbook, self.limits)?;
        let actual = reparsed.iter().map(|v| v.entry.clone()).collect::<Vec<_>>();
        if actual != desired {
            return invalid(
                CHART,
                "rewritten chart substreams failed typed round-trip validation",
            );
        }
        let workbook: Arc<[u8]> = workbook.into();
        self.package
            .put_stream_shared(&self.workbook_path, Arc::clone(&workbook))?;
        self.workbook = workbook;
        self.charts = reparsed;
        Ok(())
    }

    fn install_workbook(&mut self, workbook: Vec<u8>) -> Result<Vec<StoredChart>> {
        if workbook.len() > self.limits.max_workbook_bytes {
            return invalid(CHART, "rewritten Workbook exceeds limit");
        }
        let charts = parse_workbook_charts(&workbook, self.limits)?;
        let workbook: Arc<[u8]> = workbook.into();
        self.package
            .put_stream_shared(&self.workbook_path, Arc::clone(&workbook))?;
        self.workbook = workbook;
        Ok(std::mem::replace(&mut self.charts, charts))
    }

    pub(crate) fn resolve(&self, selector: Selector<'_>) -> Result<Option<Location>> {
        let (_, sheets) = bindings(&self.workbook)?;
        match selector {
            Selector::Sheet(name) => Ok(sheets
                .iter()
                .find(|sheet| sheet.kind == 2 && names_equal(&sheet.name, name))
                .map(|sheet| Location::ChartSheet {
                    sheet_index: sheet.index,
                })),
            Selector::Embedded { sheet, index } => {
                let Some(sheet) = sheets
                    .iter()
                    .find(|value| value.kind == 0 && names_equal(&value.name, sheet))
                else {
                    return Ok(None);
                };
                Ok(self
                    .charts
                    .iter()
                    .filter(|value| value.entry.location.sheet_index() == sheet.index)
                    .filter_map(|value| match value.entry.location {
                        Location::Embedded { .. } => Some(value.entry.location.clone()),
                        Location::ChartSheet { .. } => None,
                    })
                    .nth(index))
            },
        }
    }
}

/// The Obj identifier assigned to the chart embedded in a test-only workbook.
#[cfg(test)]
pub(super) const GENERATED_CHART_OBJECT_ID: u16 = 1;
/// BIFF record type marking the workbook-globals substream.
#[cfg(test)]
const BOF_WORKBOOK_GLOBALS: u16 = 0x0005;
/// BIFF record type marking a worksheet substream.
#[cfg(test)]
const BOF_WORKSHEET: u16 = 0x0010;
/// Sheet name of the single worksheet hosting a test-only embedded chart.
#[cfg(test)]
const GENERATED_SHEET_NAME: &str = "Sheet1";

/// Refuses fresh standalone BIFF8 chart-workbook authoring.
///
/// This public entry point returns [`litchi_ograph::Error::UnsupportedAuthoring`]
/// through [`Error::Graph`] until the complete Office-compatible chart
/// grammar is implemented.
pub fn build_workbook(_chart: Chart, _limits: Limits) -> Result<Vec<u8>> {
    unsupported_authoring()
}

/// Builds the abbreviated workbook used only to exercise the private parser.
#[cfg(test)]
pub(super) fn build_workbook_fixture(chart: Chart, limits: Limits) -> Result<Vec<u8>> {
    validate_limits(limits)?;
    let mut package = litchi_cfb::OleWriter::new();
    package.create_stream(&["Workbook"], &minimal_workbook_stream()?)?;
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.write_to(&mut bytes)?;
    let mut editor = Editor::open(bytes.into_inner(), limits)?;
    editor.insert_fixture_at(0, GENERATED_CHART_OBJECT_ID, 0, chart)?;
    editor.finish()
}

/// A minimal one-worksheet BIFF8 `Workbook` stream accepted by the chart
/// editor: workbook globals with a single `BoundSheet` directory entry
/// followed by an empty worksheet substream.
#[cfg(test)]
fn minimal_workbook_stream() -> Result<Vec<u8>> {
    let mut output = record(BOF, &bof_body(BOF_WORKBOOK_GLOBALS))?;
    let bound_offset_position = output.len() + 4;
    output.extend(record(
        BOUNDSHEET,
        &bound_sheet_body(GENERATED_SHEET_NAME, 0)?,
    )?);
    output.extend(record(EOF, &[])?);
    let sheet_offset = u32::try_from(output.len())
        .map_err(|_| Error::InvalidData("BoundSheet offset exceeds u32".into()))?;
    output[bound_offset_position..bound_offset_position + 4]
        .copy_from_slice(&sheet_offset.to_le_bytes());
    output.extend(record(BOF, &bof_body(BOF_WORKSHEET))?);
    output.extend(record(EOF, &[])?);
    Ok(output)
}

fn rewrite_sheet_directory(
    input: &[u8],
    order: &[Option<usize>],
    insert: Option<(usize, Vec<u8>, Vec<u8>)>,
) -> Result<Vec<u8>> {
    let (_, sheets) = bindings(input)?;
    let old_count = sheets.len();
    let mut logical = vec![None; old_count];
    for sheet in &sheets {
        *logical
            .get_mut(sheet.index)
            .ok_or_else(|| invalid_error(BOUNDSHEET, "sheet directory index is invalid"))? =
            Some((sheet.kind, input[sheet.start..sheet.end].to_vec()));
    }
    let globals_end = sheets
        .iter()
        .map(|value| value.start)
        .min()
        .ok_or_else(|| invalid_error(BOUNDSHEET, "workbook has no sheet substreams"))?;
    let global_ranges = ranges(&input[..globals_end])?;
    let old_bounds = global_ranges
        .iter()
        .filter(|value| value.kind == BOUNDSHEET)
        .map(|value| input[value.body_start..value.body_end].to_vec())
        .collect::<Vec<_>>();
    if old_bounds.len() != old_count {
        return invalid(BOUNDSHEET, "BoundSheet directory count mismatch");
    }
    let mut tabs = Vec::new();
    for old in order {
        let index = old
            .ok_or_else(|| invalid_error(BOUNDSHEET, "unexpected empty sheet permutation slot"))?;
        let (_, stream) = logical
            .get(index)
            .and_then(|value| value.clone())
            .ok_or_else(|| invalid_error(BOUNDSHEET, "sheet permutation target is missing"))?;
        let bound = old_bounds
            .get(index)
            .cloned()
            .ok_or_else(|| invalid_error(BOUNDSHEET, "BoundSheet index is invalid"))?;
        tabs.push((Some(index), bound, stream));
    }
    if let Some((index, bound, stream)) = insert {
        if index > tabs.len() {
            return invalid(BOUNDSHEET, "inserted chart-sheet index is out of range");
        }
        tabs.insert(index, (None, bound, stream));
    }
    if tabs.is_empty() {
        return invalid(BOUNDSHEET, "workbook must retain at least one sheet");
    }
    let mut old_to_new = vec![None; old_count];
    for (new, (old, _, _)) in tabs.iter().enumerate() {
        if let Some(old) = old {
            *old_to_new
                .get_mut(*old)
                .ok_or_else(|| invalid_error(BOUNDSHEET, "sheet permutation index is invalid"))? =
                Some(new);
        }
    }
    let insert_index = tabs.iter().position(|value| value.0.is_none());
    let globals = rewrite_chart_globals(&input[..globals_end], &tabs, &old_to_new, insert_index)?;
    let mut output = globals.bytes;
    let mut offsets = Vec::with_capacity(tabs.len());
    for (_, _, stream) in &tabs {
        offsets.push(output.len());
        output.extend_from_slice(stream);
    }
    for (position, offset) in globals.bound_positions.into_iter().zip(offsets) {
        output[position..position + 4].copy_from_slice(
            &u32::try_from(offset)
                .map_err(|_| Error::InvalidData("BoundSheet offset exceeds u32".into()))?
                .to_le_bytes(),
        );
    }
    Ok(output)
}

struct RewrittenGlobals {
    bytes: Vec<u8>,
    bound_positions: Vec<usize>,
}

fn rewrite_chart_globals(
    input: &[u8],
    tabs: &[(Option<usize>, Vec<u8>, Vec<u8>)],
    old_to_new: &[Option<usize>],
    insert_index: Option<usize>,
) -> Result<RewrittenGlobals> {
    let records = ranges(input)?;
    let mut output = Vec::new();
    let mut bound_positions = Vec::new();
    let mut bounds_written = false;
    let internal_books = internal_sup_books(input, &records)?;
    let rr_ids = records
        .iter()
        .find(|value| value.kind == RR_TAB_ID)
        .map(|value| parse_rr_tab_ids(&input[value.body_start..value.body_end], old_to_new.len()))
        .transpose()?;
    for value in records {
        let data = &input[value.body_start..value.body_end];
        if value.kind == BOUNDSHEET {
            if bounds_written {
                continue;
            }
            bounds_written = true;
            for (_, body, _) in tabs {
                let mut body = body.clone();
                body[..4].fill(0);
                bound_positions.push(output.len() + 4);
                output.extend(record(BOUNDSHEET, &body)?);
            }
            continue;
        }
        let rewritten = match value.kind {
            WINDOW1 => remap_window1(data, old_to_new)?,
            RR_TAB_ID => write_rr_tab_ids(
                rr_ids.as_ref().ok_or_else(|| {
                    invalid_error(RR_TAB_ID, "RRTabId record inventory is missing")
                })?,
                tabs,
            )?,
            SUP_BOOK if data.len() == 4 && u16_at(data, 2)? == 0x0401 => {
                let mut value = data.to_vec();
                value[..2].copy_from_slice(
                    &u16::try_from(tabs.len())
                        .map_err(|_| Error::InvalidData("sheet count exceeds u16".into()))?
                        .to_le_bytes(),
                );
                value
            },
            EXTERN_SHEET => remap_extern_sheet(data, &internal_books, old_to_new, insert_index)?,
            LBL => remap_lbl(data, old_to_new)?,
            _ => data.to_vec(),
        };
        output.extend(record(value.kind, &rewritten)?);
    }
    if !bounds_written {
        return invalid(BOUNDSHEET, "workbook globals contain no BoundSheet records");
    }
    Ok(RewrittenGlobals {
        bytes: output,
        bound_positions,
    })
}

pub(crate) fn internal_sup_books(input: &[u8], records: &[Range]) -> Result<HashSet<u16>> {
    let mut result = HashSet::new();
    let mut ordinal = 0u16;
    for value in records {
        if value.kind != SUP_BOOK {
            continue;
        }
        let data = &input[value.body_start..value.body_end];
        if data.len() == 4 && u16_at(data, 2)? == 0x0401 {
            result.insert(ordinal);
        }
        ordinal = ordinal
            .checked_add(1)
            .ok_or_else(|| Error::InvalidData("SupBook count overflow".into()))?;
    }
    Ok(result)
}

pub(super) fn remap_extern_sheet(
    data: &[u8],
    internal: &HashSet<u16>,
    old_to_new: &[Option<usize>],
    insert_index: Option<usize>,
) -> Result<Vec<u8>> {
    if data.len() < 2 {
        return invalid(EXTERN_SHEET, "ExternSheet is truncated");
    }
    let count = usize::from(u16_at(data, 0)?);
    if data.len() != 2 + count * 6 {
        return invalid(EXTERN_SHEET, "ExternSheet count does not match payload");
    }
    let mut output = data.to_vec();
    for index in 0..count {
        let offset = 2 + index * 6;
        if !internal.contains(&u16_at(data, offset)?) {
            continue;
        }
        let first = u16_at(data, offset + 2)?;
        let last = u16_at(data, offset + 4)?;
        if matches!(first, 0xfffe | 0xffff) || matches!(last, 0xfffe | 0xffff) {
            continue;
        }
        let first = usize::from(first);
        let last = usize::from(last);
        if first > last || last >= old_to_new.len() {
            return invalid(EXTERN_SHEET, "internal ExternSheet range is invalid");
        }
        if insert_index.is_some_and(|insert| first < insert && insert <= last) {
            return invalid(
                EXTERN_SHEET,
                "cannot insert a sheet inside an existing 3-D formula range",
            );
        }
        let mapped = (first..=last)
            .map(|old| {
                old_to_new.get(old).copied().flatten().ok_or_else(|| {
                    invalid_error(
                        EXTERN_SHEET,
                        "cannot remove a sheet referenced by a formula",
                    )
                })
            })
            .collect::<Result<Vec<_>>>()?;
        let minimum = *mapped
            .iter()
            .min()
            .ok_or_else(|| invalid_error(EXTERN_SHEET, "empty 3-D formula range"))?;
        let maximum = *mapped
            .iter()
            .max()
            .ok_or_else(|| invalid_error(EXTERN_SHEET, "empty 3-D formula range"))?;
        if maximum - minimum + 1 != mapped.len() {
            return invalid(
                EXTERN_SHEET,
                "sheet reorder would make a 3-D formula range noncontiguous",
            );
        }
        output[offset + 2..offset + 4].copy_from_slice(&(minimum as u16).to_le_bytes());
        output[offset + 4..offset + 6].copy_from_slice(&(maximum as u16).to_le_bytes());
    }
    Ok(output)
}

pub(crate) fn remap_lbl(data: &[u8], old_to_new: &[Option<usize>]) -> Result<Vec<u8>> {
    if data.len() < 10 {
        return invalid(LBL, "Lbl is truncated");
    }
    let scope = usize::from(u16_at(data, 8)?);
    if scope == 0 {
        return Ok(data.to_vec());
    }
    let old = scope - 1;
    let new =
        old_to_new.get(old).copied().flatten().ok_or_else(|| {
            invalid_error(LBL, "cannot remove a sheet owning a scoped defined name")
        })?;
    let mut output = data.to_vec();
    output[8..10].copy_from_slice(
        &u16::try_from(new + 1)
            .map_err(|_| Error::InvalidData("Lbl sheet scope exceeds u16".into()))?
            .to_le_bytes(),
    );
    Ok(output)
}

pub(crate) fn remap_window1(data: &[u8], old_to_new: &[Option<usize>]) -> Result<Vec<u8>> {
    if data.len() != 18 {
        return invalid(WINDOW1, "Window1 must contain 18 bytes");
    }
    let mut output = data.to_vec();
    for offset in [10usize, 12] {
        let old = usize::from(u16_at(data, offset)?);
        let new = old_to_new
            .get(old)
            .copied()
            .flatten()
            .unwrap_or_else(|| old_to_new.iter().flatten().copied().min().unwrap_or(0));
        output[offset..offset + 2].copy_from_slice(&(new as u16).to_le_bytes());
    }
    let selected = usize::from(u16_at(data, 14)?).clamp(1, old_to_new.iter().flatten().count());
    output[14..16].copy_from_slice(&(selected as u16).to_le_bytes());
    Ok(output)
}

pub(crate) fn parse_rr_tab_ids(data: &[u8], count: usize) -> Result<Vec<u16>> {
    if data.len() != count * 2 {
        return invalid(RR_TAB_ID, "RRTabId count does not match BoundSheet count");
    }
    (0..count).map(|index| u16_at(data, index * 2)).collect()
}
pub(crate) fn write_rr_tab_ids(
    old: &[u16],
    tabs: &[(Option<usize>, Vec<u8>, Vec<u8>)],
) -> Result<Vec<u8>> {
    let next = old
        .iter()
        .copied()
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| Error::InvalidData("RRTabId identifier overflow".into()))?;
    let mut output = Vec::new();
    for (old_index, _, _) in tabs {
        let value = match old_index {
            Some(index) => old
                .get(*index)
                .copied()
                .ok_or_else(|| invalid_error(RR_TAB_ID, "RRTabId index is invalid"))?,
            None => next,
        };
        output.extend(value.to_le_bytes());
    }
    Ok(output)
}
pub(crate) fn names_equal(left: &str, right: &str) -> bool {
    left.chars()
        .flat_map(char::to_lowercase)
        .eq(right.chars().flat_map(char::to_lowercase))
}
#[cfg(test)]
pub(crate) fn validate_sheet_name(name: &str) -> Result<()> {
    let count = name.encode_utf16().count();
    if !(1..=31).contains(&count)
        || name.chars().any(|value| {
            matches!(
                value,
                '\0' | '\u{0003}' | ':' | '\\' | '*' | '?' | '/' | '[' | ']'
            )
        })
        || name.starts_with('\'')
        || name.ends_with('\'')
    {
        return invalid(BOUNDSHEET, "invalid chart-sheet name");
    }
    Ok(())
}
#[cfg(test)]
pub(crate) fn bound_sheet_body(name: &str, kind: u8) -> Result<Vec<u8>> {
    validate_sheet_name(name)?;
    let units = name.encode_utf16().collect::<Vec<_>>();
    let wide = units.iter().any(|v| *v > 255);
    let mut output = vec![0; 6];
    output[5] = kind;
    output.push(units.len() as u8);
    output.push(u8::from(wide));
    if wide {
        for value in units {
            output.extend(value.to_le_bytes());
        }
    } else {
        output.extend(units.into_iter().map(|v| v as u8));
    }
    Ok(output)
}
pub(crate) fn bound_sheet_name(data: &[u8]) -> Result<String> {
    if data.len() < 8 {
        return invalid(BOUNDSHEET, "BoundSheet is truncated");
    }
    parse_biff8_string(&data[6..]).map_err(|_| Error::InvalidRecord {
        record_type: BOUNDSHEET,
        message: "invalid BoundSheet name".into(),
    })
}

fn parse_workbook_charts(input: &[u8], limits: Limits) -> Result<Vec<StoredChart>> {
    let (_, sheets) = bindings(input)?;
    let mut output = Vec::new();
    for sheet in &sheets {
        if sheet.kind == 2 {
            let bytes = &input[sheet.start..sheet.end];
            let chart_ref = GraphChartRef::with_limits(bytes, chart_scan_limits(limits))
                .map_err(|error| graph_error(BOF, error))?;
            if chart_ref.kind() != GraphChartKind::Excel {
                return invalid(BOF, "chart sheet uses a non-Excel chart grammar");
            }
            let chart = parse_chart(chart_ref.as_bytes(), limits)?;
            output.push(StoredChart {
                entry: Entry {
                    location: Location::ChartSheet {
                        sheet_index: sheet.index,
                    },
                    chart,
                },
                #[cfg(test)]
                start: sheet.start,
                #[cfg(test)]
                end: sheet.end,
                #[cfg(test)]
                object: None,
            });
            continue;
        }
        if sheet.kind != 0 {
            continue;
        }
        let bytes = &input[sheet.start..sheet.end];
        let records = ranges(bytes)?;
        let mut chart_objects = Vec::new();
        let mut used = HashSet::new();
        for value in records {
            if value.kind == OBJ
                && let Some(id) = parse_chart_object(&bytes[value.body_start..value.body_end])?
            {
                let start = sheet
                    .start
                    .checked_add(value.start)
                    .ok_or_else(|| Error::InvalidData("object start offset overflow".into()))?;
                let end = sheet
                    .start
                    .checked_add(value.end)
                    .ok_or_else(|| Error::InvalidData("object end offset overflow".into()))?;
                chart_objects.push((id, start, end));
            }
        }
        let charts = GraphCharts::with_limits(bytes, chart_scan_limits(limits))
            .map_err(|error| graph_error(BOF, error))?;
        for chart_ref in charts {
            let chart_ref = chart_ref.map_err(|error| graph_error(BOF, error))?;
            if chart_ref.kind() != GraphChartKind::Excel {
                return invalid(BOF, "embedded chart uses a non-Excel chart grammar");
            }
            let start = sheet
                .start
                .checked_add(chart_ref.offset())
                .ok_or_else(|| Error::InvalidData("chart start offset overflow".into()))?;
            let end = start
                .checked_add(chart_ref.as_bytes().len())
                .ok_or_else(|| Error::InvalidData("chart end offset overflow".into()))?;
            #[cfg(not(test))]
            let _ = end;
            let object = chart_objects
                .iter()
                .rev()
                .find(|(id, _, object_end)| *object_end <= start && !used.contains(id))
                .copied()
                .ok_or_else(|| {
                    invalid_error(OBJ, "embedded chart BOF has no preceding chart Obj/FtCmo")
                })?;
            used.insert(object.0);
            let chart = parse_chart(chart_ref.as_bytes(), limits)?;
            output.push(StoredChart {
                entry: Entry {
                    location: Location::Embedded {
                        sheet_index: sheet.index,
                        object_id: object.0,
                    },
                    chart,
                },
                #[cfg(test)]
                start,
                #[cfg(test)]
                end,
                #[cfg(test)]
                object: Some((object.1, object.2)),
            });
        }
    }
    if output.len() > limits.max_charts {
        return invalid(CHART, "chart count exceeds limit");
    }
    Ok(output)
}

#[cfg(test)]
fn rewrite_workbook_charts(
    input: &[u8],
    original: &[StoredChart],
    desired: &[Entry],
    limits: Limits,
) -> Result<Vec<u8>> {
    let (refs, sheets) = bindings(input)?;
    let mut output = input[..sheets.first().map_or(input.len(), |v| v.start)].to_vec();
    let mut new_offsets = HashMap::new();
    for sheet in &sheets {
        new_offsets.insert(sheet.start, output.len());
        let existing = original
            .iter()
            .filter(|v| v.entry.location.sheet_index() == sheet.index)
            .collect::<Vec<_>>();
        let wanted = desired
            .iter()
            .filter(|v| v.location.sheet_index() == sheet.index)
            .collect::<Vec<_>>();
        let changed = existing.iter().map(|v| &v.entry).collect::<Vec<_>>() != wanted;
        if sheet.kind == 2 {
            if !changed {
                output.extend_from_slice(&input[sheet.start..sheet.end]);
                continue;
            }
            let current = existing
                .first()
                .ok_or_else(|| invalid_error(CHART, "chart sheet has no parsed chart"))?;
            let replacement = wanted
                .iter()
                .find(|v| v.location == current.entry.location)
                .ok_or_else(|| invalid_error(CHART, "chart sheet removal is not supported"))?;
            output.extend(serialize_chart(&replacement.chart, limits)?);
            continue;
        }
        if !changed {
            output.extend_from_slice(&input[sheet.start..sheet.end]);
            continue;
        }
        if sheet.kind != 0 {
            return invalid(CHART, "embedded charts can only be written to worksheets");
        }
        let mut remove = Vec::new();
        for value in &existing {
            remove.push((value.start, value.end));
            if let Some(range) = value.object {
                remove.push(range);
            }
        }
        remove.sort_unstable();
        let segment = &input[sheet.start..sheet.end];
        let records = ranges(segment)?;
        let eof = records
            .iter()
            .rfind(|v| v.kind == EOF)
            .ok_or_else(|| invalid_error(EOF, "worksheet has no EOF"))?
            .start
            + sheet.start;
        let mut cursor = sheet.start;
        for (start, end) in remove {
            if start < cursor || end > sheet.end {
                return invalid(CHART, "overlapping chart/object ranges");
            }
            if start >= eof {
                break;
            }
            output.extend_from_slice(&input[cursor..start]);
            cursor = end;
        }
        if cursor > eof {
            return invalid(CHART, "chart range crosses worksheet EOF");
        }
        output.extend_from_slice(&input[cursor..eof]);
        for value in wanted {
            let object_id = match value.location {
                Location::Embedded { object_id, .. } => object_id,
                _ => return invalid(CHART, "chart sheet cannot be embedded"),
            };
            output.extend(chart_object_record(object_id)?);
            output.extend(serialize_chart(&value.chart, limits)?);
        }
        output.extend_from_slice(&input[eof..sheet.end]);
    }
    for (reference, old) in refs {
        let new = *new_offsets
            .get(&old)
            .ok_or_else(|| invalid_error(BOUNDSHEET, "BoundSheet target is missing"))?;
        output[reference..reference + 4].copy_from_slice(
            &u32::try_from(new)
                .map_err(|_| Error::InvalidData("BoundSheet offset exceeds u32".into()))?
                .to_le_bytes(),
        );
    }
    if output.len() > limits.max_workbook_bytes {
        return invalid(CHART, "rewritten Workbook exceeds limit");
    }
    Ok(output)
}

#[derive(Clone, Copy)]
pub(super) struct Range {
    pub(super) start: usize,
    pub(super) end: usize,
    pub(super) kind: u16,
    pub(super) body_start: usize,
    pub(super) body_end: usize,
}
#[derive(Clone)]
struct Sheet {
    index: usize,
    start: usize,
    end: usize,
    kind: u8,
    name: String,
}
#[allow(clippy::type_complexity)]
fn bindings(input: &[u8]) -> Result<(Vec<(usize, usize)>, Vec<Sheet>)> {
    let mut refs = Vec::new();
    for value in ranges(input)? {
        if value.kind == BOUNDSHEET {
            let data = &input[value.body_start..value.body_end];
            if data.len() < 8 {
                return invalid(BOUNDSHEET, "BoundSheet is truncated");
            }
            refs.push((
                value.start + 4,
                u32_at(data, 0)? as usize,
                data[5],
                bound_sheet_name(data)?,
            ));
        }
    }
    let mut physical = refs
        .iter()
        .enumerate()
        .map(|(index, (_, start, kind, name))| (index, *start, *kind, name.clone()))
        .collect::<Vec<_>>();
    physical.sort_by_key(|v| v.1);
    if physical.is_empty()
        || physical.windows(2).any(|v| v[0].1 >= v[1].1)
        || physical.iter().any(|v| v.1 >= input.len())
    {
        return invalid(BOUNDSHEET, "invalid or missing BoundSheet offsets");
    }
    let sheets = physical
        .iter()
        .enumerate()
        .map(|(slot, (index, start, kind, name))| Sheet {
            index: *index,
            start: *start,
            end: physical.get(slot + 1).map_or(input.len(), |v| v.1),
            kind: *kind,
            name: name.clone(),
        })
        .collect();
    Ok((
        refs.into_iter().map(|(p, o, _, _)| (p, o)).collect(),
        sheets,
    ))
}
pub(super) fn ranges(input: &[u8]) -> Result<Vec<Range>> {
    ranges_with(input, BiffLimits::default().max_records)
}
pub(super) fn ranges_with(input: &[u8], max_records: usize) -> Result<Vec<Range>> {
    let mut out = Vec::new();
    let biff_limits = BiffLimits {
        max_records,
        max_input_bytes: input.len().max(1),
        ..BiffLimits::default()
    };
    let records = Records::with_limits(input, biff_limits)
        .map_err(|error| Error::InvalidData(error.to_string()))?;
    for record in records {
        let record = record.map_err(|error| Error::InvalidData(error.to_string()))?;
        let start = record.offset();
        let body_start = start
            .checked_add(4)
            .ok_or_else(|| Error::InvalidData("BIFF body offset overflow".into()))?;
        let end = start
            .checked_add(record.encoded().len())
            .ok_or_else(|| Error::InvalidData("BIFF record length overflow".into()))?;
        out.try_reserve(1)
            .map_err(|_| Error::InvalidData("could not allocate BIFF ranges".into()))?;
        out.push(Range {
            start,
            end,
            kind: record.kind().get(),
            body_start,
            body_end: end,
        });
    }
    Ok(out)
}
#[cfg(test)]
fn sheet_object_ids(input: &[u8], sheet: &Sheet) -> Result<HashSet<u16>> {
    let bytes = input
        .get(sheet.start..sheet.end)
        .ok_or_else(|| invalid_error(BOUNDSHEET, "worksheet range is out of bounds"))?;
    let mut ids = HashSet::new();
    for value in ranges(bytes)? {
        if value.kind != OBJ {
            continue;
        }
        if let Some((_, id)) = parse_object(&bytes[value.body_start..value.body_end])? {
            ids.insert(id);
        }
    }
    Ok(ids)
}
pub(crate) fn parse_chart_object(data: &[u8]) -> Result<Option<u16>> {
    Ok(parse_object(data)?.and_then(|(kind, id)| (kind == 5).then_some(id)))
}
pub(crate) fn parse_object(data: &[u8]) -> Result<Option<(u16, u16)>> {
    let mut offset = 0;
    while offset < data.len() {
        let h = data
            .get(offset..offset + 4)
            .ok_or_else(|| invalid_error(OBJ, "truncated Obj subrecord"))?;
        let kind = u16::from_le_bytes([h[0], h[1]]);
        let len = usize::from(u16::from_le_bytes([h[2], h[3]]));
        offset += 4;
        let end = offset
            .checked_add(len)
            .ok_or_else(|| Error::InvalidData("Obj length overflow".into()))?;
        let body = data
            .get(offset..end)
            .ok_or_else(|| invalid_error(OBJ, "truncated Obj subrecord body"))?;
        if kind == 0x15 {
            if len != 18 {
                return invalid(OBJ, "FtCmo must contain 18 bytes");
            }
            return Ok(Some((u16_at(body, 0)?, u16_at(body, 2)?)));
        }
        offset = end;
    }
    Ok(None)
}
#[cfg(test)]
pub(crate) fn chart_object_record(id: u16) -> Result<Vec<u8>> {
    let mut body = Vec::new();
    body.extend(0x15u16.to_le_bytes());
    body.extend(18u16.to_le_bytes());
    body.extend(5u16.to_le_bytes());
    body.extend(id.to_le_bytes());
    body.extend(0x6011u16.to_le_bytes());
    body.extend([0; 12]);
    body.extend(0u16.to_le_bytes());
    body.extend(0u16.to_le_bytes());
    record(OBJ, &body)
}
pub(super) fn is_chart_bof(data: &[u8]) -> bool {
    data.len() >= 4
        && u16::from_le_bytes([data[0], data[1]]) == 0x0600
        && u16::from_le_bytes([data[2], data[3]]) == 0x0020
}
#[cfg(test)]
pub(super) fn chart_bof() -> Vec<u8> {
    bof_body(0x0020)
}
#[cfg(test)]
pub(crate) fn bof_body(kind: u16) -> Vec<u8> {
    let mut d = Vec::new();
    d.extend(0x0600u16.to_le_bytes());
    d.extend(kind.to_le_bytes());
    d.extend(0x0dbbu16.to_le_bytes());
    d.extend(0x07ccu16.to_le_bytes());
    d.extend(0u32.to_le_bytes());
    d.extend(6u32.to_le_bytes());
    d
}
