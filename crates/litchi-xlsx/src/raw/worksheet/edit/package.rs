//! Worksheet-package orchestration and merge-container patching.

use std::collections::BTreeMap;

use litchi_sheet::{Cell as Address, Rect};

use super::codec::{
    CellFact, CellSlot, DimensionFact, ExtensionNames, Layout, MergeCellsSlot, RowFact,
    SourceFacts, Span, Tag, materialize_cell, materialize_tag, scan, sibling_name, write_close,
    write_columns, write_defaults, write_new_columns, write_new_defaults, write_root,
    write_sheet_data, write_sheet_data_from_facts, write_sheet_data_with_provenance, write_tag,
};
use super::model::{MergePlan, Payload, Plan};
use super::validation::{
    expanded_dimension, plan_sets_descent, validate_actions, validate_column_actions,
    validate_defaults_action, validate_row_actions,
};
use crate::error::{Error, MergeEditBlock, Result, allocation, invalid};
use crate::merge;

/// One contiguous range of cell records omitted from the independent
/// readback representation. The offsets always refer to the actual complete
/// rewritten output, not to a reduced buffer.
#[derive(Debug, Clone, Copy)]
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
    facts: Option<&SourceFacts>,
) -> Result<ValueOnlyRewrite<'a>> {
    let plan = Plan::cells(cells);
    if plan.is_empty() {
        return Ok(ValueOnlyRewrite {
            source: content,
            bytes: content.to_vec(),
            omitted: Box::new([]),
        });
    }

    let plan = match facts {
        Some(facts) => match rewrite_from_facts(content, facts, plan)? {
            FactRewrite::Written(rewrite) => return Ok(rewrite),
            FactRewrite::Declined(plan) => plan,
        },
        None => plan,
    };

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

/// Test-only accounting of which commit route ran.
///
/// The oracle needs to prove that its byte comparison is not vacuous: a
/// declined fact route and a scan route produce the same bytes trivially.
#[cfg(test)]
pub(crate) mod route {
    use std::cell::Cell as CellCounter;

    thread_local! {
        static ACCEPTED: CellCounter<usize> = const { CellCounter::new(0) };
        static DECLINED: CellCounter<usize> = const { CellCounter::new(0) };
    }

    pub(crate) fn note_accepted() {
        ACCEPTED.with(|count| count.set(count.get().saturating_add(1)));
    }

    pub(crate) fn note_declined() {
        DECLINED.with(|count| count.set(count.get().saturating_add(1)));
    }

    pub(crate) fn reset() {
        ACCEPTED.with(|count| count.set(0));
        DECLINED.with(|count| count.set(0));
    }

    pub(crate) fn accepted() -> usize {
        ACCEPTED.with(CellCounter::get)
    }

    pub(crate) fn declined() -> usize {
        DECLINED.with(CellCounter::get)
    }
}

/// Outcome of the compact-fact commit route.
enum FactRewrite<'a> {
    Written(ValueOnlyRewrite<'a>),
    Declined(Plan),
}

/// Rewrite a value-only plan from the compact planning facts, or decline.
///
/// The facts prove that the complete layout scan would have succeeded and
/// what it would have produced for the `<sheetData>` body. This route is
/// admitted only for the narrow shape in which no row or cell is created,
/// removed or promoted, so every guard the scan-derived route consults is
/// either provably empty (protection, data validations, merged ranges,
/// formula ranges and shared formula groups cannot survive the value-only
/// validator's element allow-list) or unreachable. Every decline happens
/// before the first output byte is written and falls back to today's scan.
fn rewrite_from_facts<'a>(
    content: &'a [u8],
    facts: &SourceFacts,
    plan: Plan,
) -> Result<FactRewrite<'a>> {
    if !facts.describes(content) || !facts_admit_plan(facts, &plan) {
        #[cfg(test)]
        route::note_declined();
        return Ok(FactRewrite::Declined(plan));
    }
    let Some(slots) = materialize_changed_cells(content, facts, &plan)? else {
        #[cfg(test)]
        route::note_declined();
        return Ok(FactRewrite::Declined(plan));
    };
    for action in plan.cells.values() {
        if let Some(Payload::Set(content)) = action.payload() {
            content.validate_for_write()?;
        }
        if let Some(Payload::SharedString { text, .. }) = action.payload() {
            crate::Content::Value(crate::Value::Text(text.clone())).validate_for_write()?;
        }
    }

    let dimension = match expanded_dimension_from_facts(facts, &plan.cells) {
        Some((fact, range)) => match materialize_tag(content, fact.span)? {
            Some(tag) => Some((fact, tag, range)),
            None => {
                #[cfg(test)]
                route::note_declined();
                return Ok(FactRewrite::Declined(plan));
            },
        },
        None => None,
    };

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
    let Plan { cells, .. } = plan;
    let mut cursor = 0usize;
    if let Some((fact, tag, range)) = dimension.as_ref() {
        output.extend_from_slice(&content[cursor..fact.span.start]);
        write_tag(
            &mut output,
            tag,
            fact.empty,
            &["ref"],
            &[("ref", range.a1())],
        );
        cursor = fact.span.end;
    }
    output.extend_from_slice(&content[cursor..facts.sheet_data.start]);
    let omitted = write_sheet_data_from_facts(&mut output, content, facts, cells, &slots)?;
    output.extend_from_slice(&content[facts.sheet_data.end..]);
    #[cfg(test)]
    route::note_accepted();
    Ok(FactRewrite::Written(ValueOnlyRewrite {
        source: content,
        bytes: output,
        omitted,
    }))
}

/// Locate the retained row record for a one-based row number.
fn facts_row(facts: &SourceFacts, number: u32) -> Option<&RowFact> {
    facts
        .rows
        .binary_search_by_key(&number, |row| row.number)
        .ok()
        .and_then(|index| facts.rows.get(index))
}

/// Whether every staged action only replaces the value of a cell the facts
/// already recorded, which is the one shape this route reproduces exactly.
fn facts_admit_plan(facts: &SourceFacts, plan: &Plan) -> bool {
    if !plan.rows.is_empty() || !plan.columns.is_empty() || plan.defaults.is_some() {
        return false;
    }
    plan.cells.iter().all(|(address, action)| {
        if matches!(action, super::model::Action::Remove)
            || matches!(action.payload(), Some(Payload::SharedFormula { .. }))
        {
            return false;
        }
        let Some(number) = address.row().get().checked_add(1) else {
            return false;
        };
        let Some(row) = facts_row(facts, number) else {
            return false;
        };
        if row.empty {
            return false;
        }
        facts.row_cells(row).is_some_and(|cells| {
            cells
                .binary_search_by_key(address, |cell| cell.address)
                .is_ok()
        })
    })
}

/// Materialize the scanner slot of every staged cell before any output byte
/// is written, so a slot this route cannot reproduce declines cleanly.
fn materialize_changed_cells(
    content: &[u8],
    facts: &SourceFacts,
    plan: &Plan,
) -> Result<Option<BTreeMap<Address, CellSlot>>> {
    let mut slots = BTreeMap::new();
    for address in plan.cells.keys() {
        let Some(fact) = facts_cell(facts, *address) else {
            return Ok(None);
        };
        let Some(slot) = materialize_cell(content, &fact)? else {
            return Ok(None);
        };
        if slot.address != *address || slot.mce_payload {
            return Ok(None);
        }
        slots.insert(*address, slot);
    }
    if slots.len() != plan.cells.len() {
        return Ok(None);
    }
    Ok(Some(slots))
}

/// The retained cell record at one address.
fn facts_cell(facts: &SourceFacts, address: Address) -> Option<CellFact> {
    let number = address.row().get().checked_add(1)?;
    let row = facts_row(facts, number)?;
    let cells = facts.row_cells(row)?;
    let index = cells
        .binary_search_by_key(&address, |cell| cell.address)
        .ok()?;
    cells.get(index).copied()
}

/// The dimension expansion the scan-derived route would compute, read from
/// the retained cell addresses instead of a rebuilt layout.
fn expanded_dimension_from_facts(
    facts: &SourceFacts,
    actions: &BTreeMap<Address, super::model::Action>,
) -> Option<(DimensionFact, Rect)> {
    let dimension = facts.dimension?;
    let mut bounds = None::<Rect>;
    for cell in &*facts.cells {
        if matches!(
            actions.get(&cell.address),
            Some(super::model::Action::Remove)
        ) {
            continue;
        }
        let single = Rect::single(cell.address);
        bounds = Some(bounds.map_or(single, |range| range.union(single)));
    }
    for (address, action) in actions {
        if action.creates_missing() {
            let single = Rect::single(*address);
            bounds = Some(bounds.map_or(single, |range| range.union(single)));
        }
    }
    let result = bounds?;
    let expanded = dimension.declared.union(result);
    (expanded != dimension.declared).then_some((dimension, expanded))
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
