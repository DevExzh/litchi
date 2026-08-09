//! Worksheet-package orchestration and merge-container patching.

use litchi_sheet::Rect;

use super::codec::{
    ExtensionNames, MergeCellsSlot, Span, Tag, scan, sibling_name, write_close, write_columns,
    write_defaults, write_new_columns, write_new_defaults, write_root, write_sheet_data, write_tag,
};
use super::model::{MergePlan, Plan};
use super::validation::{
    expanded_dimension, plan_sets_descent, validate_actions, validate_column_actions,
    validate_defaults_action, validate_row_actions,
};
use crate::error::{Error, MergeEditBlock, Result, allocation, invalid};
use crate::merge;
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
