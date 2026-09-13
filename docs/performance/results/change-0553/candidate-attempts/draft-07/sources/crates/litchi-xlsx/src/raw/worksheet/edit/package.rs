//! Worksheet-package orchestration and merge-container patching.

use std::collections::BTreeMap;

use litchi_sheet::{Cell as Address, Rect};

use super::codec::{
    CompactLayout, ExtensionNames, Layout, MergeCellsSlot, Span, Tag, scan, sibling_name,
    write_close, write_columns, write_compact_dimension, write_defaults, write_new_columns,
    write_new_defaults, write_root, write_sheet_data, write_sheet_data_with_compact_provenance,
    write_sheet_data_with_provenance, write_tag,
};
use super::model::{MergePlan, Payload, Plan};
use super::validation::{
    expanded_dimension, plan_sets_descent, validate_actions, validate_column_actions,
    validate_defaults_action, validate_row_actions,
};
use crate::cell::Stored;
use crate::error::{Error, MergeEditBlock, Result, allocation, invalid};
use crate::merge;

#[cfg(test)]
mod compact_action_domain;
#[cfg(test)]
mod compact_dimension_expansion;
#[cfg(test)]
mod compact_layout_order_fallback;

/// One contiguous range of cell records omitted from the independent
/// readback representation. The offsets always refer to the actual complete
/// rewritten output, not to a reduced buffer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct OmittedCells {
    pub(crate) row: u32,
    pub(crate) first_column: u32,
    pub(crate) last_column: u32,
    pub(crate) start: usize,
    pub(crate) end: usize,
}

/// Complete value-only output plus its private readback provenance.
///
/// Borrowing the exact input slice ties the omission metadata to the source
/// store from which it was derived. The snapshot constructor checks that
/// identity before merging retained cells into the reduced parse.
#[derive(Debug)]
pub(crate) struct ValueOnlyRewrite<'a> {
    pub(crate) source: &'a [u8],
    pub(crate) bytes: Vec<u8>,
    pub(crate) omitted: Box<[OmittedCells]>,
}
pub(crate) fn rewrite(content: &[u8], sheet: &str, plan: impl Into<Plan>) -> Result<Vec<u8>> {
    let plan = plan.into();
    if plan.is_empty() {
        return Ok(content.to_vec());
    }
    let layout = scan(content)?;
    validate_actions(&layout, sheet, &plan.cells)?;
    validate_row_actions(&layout, sheet, &plan.rows)?;
    validate_column_actions(&layout, sheet, &plan.columns)?;
    validate_defaults_action(&layout, sheet, plan.defaults)?;
    rewrite_with_layout(content, sheet, layout, plan)
}

fn rewrite_with_layout(content: &[u8], sheet: &str, layout: Layout, plan: Plan) -> Result<Vec<u8>> {
    let dimension = expanded_dimension(&layout, &plan.cells);
    let extension_names = ExtensionNames::plan(&layout, plan_sets_descent(&plan))?;

    let effects = plan
        .cells
        .len()
        .checked_add(plan.rows.len())
        .and_then(|count| count.checked_add(plan.columns.len()))
        .and_then(|count| count.checked_add(usize::from(plan.defaults.is_some())))
        .ok_or_else(|| invalid("worksheet edit effect count overflow"))?;
    let extra = effects
        .checked_mul(128)
        .and_then(|value| content.len().checked_add(value))
        .ok_or_else(|| invalid("worksheet edit output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve(extra)
        .map_err(|source| allocation("worksheet edit output", source))?;
    let Plan {
        defaults,
        cells,
        rows,
        columns,
    } = plan;
    let mut cursor = 0usize;
    if let Some(effect) = &extension_names.root {
        output.extend_from_slice(&content[cursor..layout.root.span.start]);
        write_root(&mut output, &layout.root, effect);
        cursor = layout.root.span.end;
    }
    if let Some((tag, range)) = dimension {
        output.extend_from_slice(&content[cursor..tag.span.start]);
        write_tag(
            &mut output,
            &tag.tag,
            tag.empty,
            &["ref"],
            &[("ref", range.a1())],
        );
        cursor = tag.span.end;
    }
    if let Some(action) = defaults {
        match layout.defaults.as_ref() {
            Some(stored) => {
                output.extend_from_slice(&content[cursor..stored.span.start]);
                if !action.is_remove() {
                    write_defaults(
                        &mut output,
                        content,
                        stored,
                        action.effects(),
                        &extension_names.descent,
                    );
                }
                cursor = stored.span.end;
            },
            None if action.materializes() => {
                let insertion = layout
                    .columns
                    .as_ref()
                    .map_or(layout.sheet_data.span.start, |columns| columns.span.start);
                output.extend_from_slice(&content[cursor..insertion]);
                write_new_defaults(
                    &mut output,
                    &layout.sheet_data.tag.name,
                    action.effects(),
                    &extension_names.descent,
                );
                cursor = insertion;
            },
            None => {},
        }
    }
    if !columns.is_empty() {
        if let Some(stored) = layout.columns.as_ref() {
            output.extend_from_slice(&content[cursor..stored.span.start]);
            write_columns(&mut output, content, stored, columns, sheet)?;
            cursor = stored.span.end;
        } else {
            output.extend_from_slice(&content[cursor..layout.sheet_data.span.start]);
            write_new_columns(&mut output, &layout.sheet_data.tag.name, columns);
            cursor = layout.sheet_data.span.start;
        }
    }
    output.extend_from_slice(&content[cursor..layout.sheet_data.span.start]);
    if cells.is_empty() && rows.is_empty() {
        output
            .extend_from_slice(&content[layout.sheet_data.span.start..layout.sheet_data.span.end]);
    } else {
        write_sheet_data(
            &mut output,
            content,
            &layout.sheet_data,
            cells,
            rows,
            &extension_names.descent,
        )?;
    }
    output.extend_from_slice(&content[layout.sheet_data.span.end..]);
    Ok(output)
}

/// Rewrite a value-only plan while recording spans that can be omitted from
/// the independent semantic readback. All bytes in `bytes` remain the exact
/// ordinary rewrite output. Unsupported eligibility falls back to that same
/// output with no reuse metadata.
pub(crate) fn rewrite_value_only_with_provenance<'a>(
    content: &'a [u8],
    sheet: &str,
    cells: BTreeMap<Address, super::model::Action>,
) -> Result<ValueOnlyRewrite<'a>> {
    rewrite_value_only_with_complete_provenance(content, sheet, cells)
}

/// Try the compact source proof before falling back to the complete scanner.
/// The compact route is intentionally restricted to existing-cell scalar
/// updates; every refusal is silent and leaves the established implementation
/// responsible for validation, diagnostics, and output.
pub(crate) fn rewrite_value_only_with_compact_proof<'a>(
    content: &'a [u8],
    sheet: &str,
    cells: BTreeMap<Address, super::model::Action>,
    proof: Option<CompactLayout>,
    entries: &[Stored],
) -> Result<ValueOnlyRewrite<'a>> {
    let compact = proof
        .as_ref()
        .and_then(|proof| try_compact_value_rewrite(content, proof, entries, &cells));
    drop(proof);
    if let Some(rewrite) = compact {
        return Ok(rewrite);
    }
    rewrite_value_only_with_complete_provenance(content, sheet, cells)
}

/// Attempt only the proven compact route. Keeping the same entry point for
/// production and exact differential tests makes fallback observable in tests.
pub(crate) fn try_compact_value_rewrite<'a>(
    content: &'a [u8],
    proof: &CompactLayout,
    entries: &[Stored],
    cells: &BTreeMap<Address, super::model::Action>,
) -> Option<ValueOnlyRewrite<'a>> {
    if !proof.matches_source(content)
        || !proof.matches_entries(entries)
        || !compact_layout_offsets_valid(content, proof)
        || !compact_plan_eligible(proof, entries, cells)
    {
        return None;
    }
    try_compact_rewrite(content, proof, entries, cells)
}

fn compact_layout_offsets_valid(content: &[u8], proof: &CompactLayout) -> bool {
    fn valid_span(content_len: usize, span: super::codec::CompactSpan) -> bool {
        span.start() <= span.end() && span.end() <= content_len
    }

    let content_len = content.len();
    let data = &proof.sheet_data;
    if !valid_span(content_len, data.span)
        || data.span.start() > data.tag_end as usize
        || data.tag_end as usize > data.close_start as usize
        || data.close_start as usize > data.span.end()
    {
        return false;
    }
    if data.empty {
        return data.rows.is_empty() && data.tag_end == data.close_start;
    }
    let mut previous_row_end = data.tag_end as usize;
    for row in &data.rows {
        if !valid_span(content_len, row.span)
            || row.span.start() < previous_row_end
            || row.span.start() > row.tag_end as usize
            || row.tag_end as usize > row.close_start as usize
            || row.close_start as usize > row.span.end()
            || row.span.end() > data.close_start as usize
        {
            return false;
        }
        let mut previous_cell_end = row.tag_end as usize;
        for cell in &row.cells {
            if !valid_span(content_len, cell.span)
                || cell.span.start() < previous_cell_end
                || cell.span.end() > row.close_start as usize
            {
                return false;
            }
            previous_cell_end = cell.span.end();
        }
        previous_row_end = row.span.end();
    }
    if let Some(dimension) = proof.dimension
        && (!valid_span(content_len, dimension.span) || dimension.span.end() > data.span.start())
    {
        return false;
    }
    true
}

fn compact_plan_eligible(
    proof: &CompactLayout,
    entries: &[Stored],
    cells: &BTreeMap<Address, super::model::Action>,
) -> bool {
    if cells.is_empty() || usize::try_from(proof.cell_count).ok() != Some(entries.len()) {
        return false;
    }
    let mut entry_index = 0usize;
    let mut matched = 0usize;
    let mut previous_row = None;
    for row in &proof.sheet_data.rows {
        if previous_row.is_some_and(|previous| row.number <= previous) {
            return false;
        }
        previous_row = Some(row.number);
        let mut previous_column = None;
        for _slot in &row.cells {
            let Some(entry) = entries.get(entry_index) else {
                return false;
            };
            if entry.address.row().get().checked_add(1) != Some(row.number)
                || previous_column.is_some_and(|previous| entry.address.column().get() <= previous)
            {
                return false;
            }
            previous_column = Some(entry.address.column().get());
            if let Some(action) = cells.get(&entry.address) {
                if !matches!(
                    action,
                    super::model::Action::Update {
                        payload: None
                            | Some(Payload::Set(_) | Payload::Clear | Payload::ClearIfPresent),
                        ..
                    }
                ) {
                    return false;
                }
                let Some(next) = matched.checked_add(1) else {
                    return false;
                };
                matched = next;
            }
            let Some(next) = entry_index.checked_add(1) else {
                return false;
            };
            entry_index = next;
        }
    }
    entry_index == entries.len() && matched == cells.len()
}

fn compact_dimension(
    proof: &CompactLayout,
    entries: &[Stored],
    cells: &BTreeMap<Address, super::model::Action>,
) -> Option<(super::codec::CompactDimensionTag, Rect)> {
    let dimension = proof.dimension?;
    let mut bounds = None::<Rect>;
    for entry in entries {
        if !matches!(
            cells.get(&entry.address),
            Some(super::model::Action::Remove)
        ) {
            let cell = Rect::single(entry.address);
            bounds = Some(bounds.map_or(cell, |range| range.union(cell)));
        }
    }
    for (address, action) in cells {
        if action.creates_missing() {
            let cell = Rect::single(*address);
            bounds = Some(bounds.map_or(cell, |range| range.union(cell)));
        }
    }
    let expanded = dimension.declared.union(bounds?);
    (expanded != dimension.declared).then_some((dimension, expanded))
}

fn try_compact_rewrite<'a>(
    content: &'a [u8],
    proof: &CompactLayout,
    entries: &[Stored],
    cells: &BTreeMap<Address, super::model::Action>,
) -> Option<ValueOnlyRewrite<'a>> {
    let dimension = compact_dimension(proof, entries, cells);
    let effects = cells.len();
    let extra = effects.checked_mul(128)?.checked_add(content.len())?;
    let mut output = Vec::new();
    output.try_reserve(extra).ok()?;
    let mut cursor = 0usize;
    if let Some((dimension, range)) = dimension {
        output.extend_from_slice(&content[cursor..dimension.span.start()]);
        write_compact_dimension(&mut output, content, &dimension, range).ok()?;
        cursor = dimension.span.end();
    }
    let data = &proof.sheet_data;
    output.extend_from_slice(&content[cursor..data.span.start()]);
    let omitted =
        write_sheet_data_with_compact_provenance(&mut output, content, data, entries, cells)
            .ok()?;
    output.extend_from_slice(&content[data.span.end()..]);
    Some(ValueOnlyRewrite {
        source: content,
        bytes: output,
        omitted,
    })
}

fn rewrite_value_only_with_complete_provenance<'a>(
    content: &'a [u8],
    sheet: &str,
    cells: BTreeMap<Address, super::model::Action>,
) -> Result<ValueOnlyRewrite<'a>> {
    let plan = Plan::cells(cells);
    if plan.is_empty() {
        return Ok(ValueOnlyRewrite {
            source: content,
            bytes: content.to_vec(),
            omitted: Box::new([]),
        });
    }

    let layout = scan(content)?;
    validate_actions(&layout, sheet, &plan.cells)?;
    validate_row_actions(&layout, sheet, &plan.rows)?;
    validate_column_actions(&layout, sheet, &plan.columns)?;
    validate_defaults_action(&layout, sheet, plan.defaults)?;
    let eligible = !layout.has_shared_formulas
        && !plan
            .cells
            .values()
            .any(|action| matches!(action.payload(), Some(Payload::SharedFormula { .. })))
        && plan.cells.keys().all(|address| {
            let Some(number) = address.row().get().checked_add(1) else {
                return false;
            };
            layout
                .sheet_data
                .rows
                .binary_search_by_key(&number, |row| row.number)
                .is_ok()
        });
    if !eligible {
        return Ok(ValueOnlyRewrite {
            source: content,
            bytes: rewrite_with_layout(content, sheet, layout, plan)?,
            omitted: Box::new([]),
        });
    }

    let dimension = expanded_dimension(&layout, &plan.cells);
    let extension_names = ExtensionNames::plan(&layout, plan_sets_descent(&plan))?;
    let effects = plan
        .cells
        .len()
        .checked_add(plan.rows.len())
        .and_then(|count| count.checked_add(plan.columns.len()))
        .and_then(|count| count.checked_add(usize::from(plan.defaults.is_some())))
        .ok_or_else(|| invalid("worksheet edit effect count overflow"))?;
    let extra = effects
        .checked_mul(128)
        .and_then(|value| content.len().checked_add(value))
        .ok_or_else(|| invalid("worksheet edit output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve(extra)
        .map_err(|source| allocation("worksheet edit output", source))?;
    // Plan::cells has no row, column, or default effects; retain only the
    // dimension and sheet-data envelope for this value-only route.
    let Plan { cells, .. } = plan;
    let mut cursor = 0usize;
    if let Some((tag, range)) = dimension {
        output.extend_from_slice(&content[cursor..tag.span.start]);
        write_tag(
            &mut output,
            &tag.tag,
            tag.empty,
            &["ref"],
            &[("ref", range.a1())],
        );
        cursor = tag.span.end;
    }
    output.extend_from_slice(&content[cursor..layout.sheet_data.span.start]);
    let omitted = if cells.is_empty() {
        output
            .extend_from_slice(&content[layout.sheet_data.span.start..layout.sheet_data.span.end]);
        Box::new([])
    } else {
        write_sheet_data_with_provenance(
            &mut output,
            content,
            &layout.sheet_data,
            cells,
            BTreeMap::new(),
            &extension_names.descent,
        )?
    };
    output.extend_from_slice(&content[layout.sheet_data.span.end..]);
    Ok(ValueOnlyRewrite {
        source: content,
        bytes: output,
        omitted,
    })
}

/// Build the XML used only for the reduced semantic readback. Every span is
/// checked against the complete output and must be sorted and disjoint.
pub(crate) fn reduced_readback(content: &[u8], omitted: &[OmittedCells]) -> Result<Vec<u8>> {
    let mut removed = 0usize;
    let mut cursor = 0usize;
    for span in omitted {
        if span.start < cursor || span.start >= span.end || span.end > content.len() {
            return Err(invalid("worksheet omission provenance is invalid"));
        }
        removed = removed
            .checked_add(span.end - span.start)
            .ok_or_else(|| invalid("worksheet omission size overflows usize"))?;
        cursor = span.end;
    }
    let capacity = content
        .len()
        .checked_sub(removed)
        .ok_or_else(|| invalid("worksheet omission exceeds rewritten output"))?;
    let mut reduced = Vec::new();
    reduced
        .try_reserve_exact(capacity)
        .map_err(|source| allocation("worksheet reduced readback", source))?;
    cursor = 0;
    for span in omitted {
        reduced.extend_from_slice(&content[cursor..span.start]);
        cursor = span.end;
    }
    reduced.extend_from_slice(&content[cursor..]);
    Ok(reduced)
}

#[derive(Debug)]
struct MergeReplacement {
    span: Span,
    bytes: Vec<u8>,
}

/// Losslessly add and remove direct worksheet merge records.
pub(crate) fn rewrite_merges(content: &[u8], sheet: &str, plan: MergePlan) -> Result<Vec<u8>> {
    if plan.is_empty() {
        return Ok(content.to_vec());
    }
    let layout = scan(content)?;
    let requested = plan
        .add
        .first()
        .or_else(|| plan.remove.first())
        .copied()
        .ok_or_else(|| invalid("merged-range edit lost its requested range"))?;
    if layout.protected {
        return Err(merge_block(
            sheet,
            requested,
            MergeEditBlock::ProtectedSheet,
        ));
    }
    if layout.merge_compatibility {
        return Err(merge_block(
            sheet,
            requested,
            MergeEditBlock::MarkupCompatibility,
        ));
    }
    if layout
        .merge_cells
        .as_ref()
        .is_some_and(|container| container.payload)
    {
        return Err(merge_block(
            sheet,
            requested,
            MergeEditBlock::UnmodeledPayload,
        ));
    }

    let merge_count = layout
        .merge_cells
        .as_ref()
        .map_or(0, |container| container.merges.len());
    let mut base = Vec::new();
    base.try_reserve_exact(merge_count)
        .map_err(|source| allocation("source merged ranges", source))?;
    if let Some(container) = layout.merge_cells.as_ref() {
        base.extend(container.merges.iter().map(|merge| merge.range));
    }
    let mut projected = Vec::new();
    projected
        .try_reserve_exact(base.len().saturating_add(plan.add.len()))
        .map_err(|source| allocation("projected merged ranges", source))?;
    projected.extend_from_slice(&base);
    for range in &plan.remove {
        projected.retain(|candidate| candidate != range);
    }
    for range in plan.add {
        if range.rows() == 1 && range.columns() == 1 {
            return Err(merge_block(sheet, range, MergeEditBlock::SingleCell));
        }
        if layout
            .formula_ranges
            .iter()
            .any(|formula| formula.overlaps(range))
        {
            return Err(merge_block(sheet, range, MergeEditBlock::GroupFormula));
        }
        if projected.contains(&range) {
            continue;
        }
        if let Some(existing) = projected
            .iter()
            .copied()
            .find(|existing| merge::overlaps(*existing, range))
        {
            return Err(merge_block(
                sheet,
                range,
                MergeEditBlock::Overlap { existing },
            ));
        }
        projected.push(range);
    }
    if projected == base {
        return Ok(content.to_vec());
    }
    let projected = merge::Index::new(projected)?;
    let projected = projected.as_slice();

    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(2)
        .map_err(|source| allocation("merged-range replacements", source))?;
    if let Some(dimension) = layout.dimension.as_ref() {
        let expanded = projected
            .iter()
            .copied()
            .filter(|range| !base.contains(range))
            .fold(dimension.declared, Rect::union);
        if expanded != dimension.declared {
            let mut bytes = Vec::new();
            write_tag(
                &mut bytes,
                &dimension.tag,
                dimension.empty,
                &["ref"],
                &[("ref", expanded.a1())],
            );
            replacements.push(MergeReplacement {
                span: dimension.span,
                bytes,
            });
        }
    }

    match layout.merge_cells.as_ref() {
        Some(container) => replacements.push(MergeReplacement {
            span: container.span,
            bytes: write_merge_cells(content, container, projected),
        }),
        None => replacements.push(MergeReplacement {
            span: Span {
                start: layout.merge_insertion,
                end: layout.merge_insertion,
            },
            bytes: write_new_merge_cells(&layout.sheet_data.tag.name, projected),
        }),
    }
    apply_merge_replacements(content, replacements)
}

fn merge_block(sheet: &str, range: Rect, reason: MergeEditBlock) -> Error {
    Error::MergeEditBlocked {
        sheet: sheet.to_owned(),
        range,
        reason,
    }
}

fn write_merge_cells(content: &[u8], container: &MergeCellsSlot, projected: &[Rect]) -> Vec<u8> {
    if projected.is_empty() {
        return Vec::new();
    }
    let mut output = Vec::new();
    write_tag(
        &mut output,
        &container.tag,
        false,
        &["count"],
        &[("count", projected.len().to_string())],
    );
    if !container.empty {
        let mut cursor = container.tag_end;
        for stored in &container.merges {
            output.extend_from_slice(&content[cursor..stored.span.start]);
            if projected.contains(&stored.range) {
                output.extend_from_slice(&content[stored.span.start..stored.span.end]);
            }
            cursor = stored.span.end;
        }
        output.extend_from_slice(&content[cursor..container.close_start]);
    }
    let child_name = sibling_name(&container.tag.name, "mergeCell");
    let child = Tag {
        name: child_name.into_boxed_str(),
        attributes: Box::new([]),
    };
    for range in projected
        .iter()
        .copied()
        .filter(|range| !container.merges.iter().any(|stored| stored.range == *range))
    {
        write_tag(&mut output, &child, true, &[], &[("ref", range.a1())]);
    }
    write_close(&mut output, &container.tag.name);
    output
}

fn write_new_merge_cells(sheet_data_name: &str, projected: &[Rect]) -> Vec<u8> {
    let name = sibling_name(sheet_data_name, "mergeCells");
    let child_name = sibling_name(sheet_data_name, "mergeCell");
    let tag = Tag {
        name: name.into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut output = Vec::new();
    write_tag(
        &mut output,
        &tag,
        false,
        &[],
        &[("count", projected.len().to_string())],
    );
    let child = Tag {
        name: child_name.into_boxed_str(),
        attributes: Box::new([]),
    };
    for range in projected {
        write_tag(&mut output, &child, true, &[], &[("ref", range.a1())]);
    }
    write_close(&mut output, &tag.name);
    output
}

fn apply_merge_replacements(
    content: &[u8],
    mut replacements: Vec<MergeReplacement>,
) -> Result<Vec<u8>> {
    replacements.sort_unstable_by_key(|replacement| replacement.span.start);
    if replacements
        .windows(2)
        .any(|pair| pair[0].span.end > pair[1].span.start)
    {
        return Err(invalid("overlapping merged-range replacements"));
    }
    let size = replacements
        .iter()
        .try_fold(content.len(), |size, replacement| {
            size.checked_sub(replacement.span.end - replacement.span.start)?
                .checked_add(replacement.bytes.len())
        })
        .ok_or_else(|| invalid("merged-range output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(size)
        .map_err(|source| allocation("merged-range output", source))?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&content[cursor..replacement.span.start]);
        output.extend_from_slice(&replacement.bytes);
        cursor = replacement.span.end;
    }
    output.extend_from_slice(&content[cursor..]);
    Ok(output)
}
