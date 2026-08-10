//! Raw package and worksheet XML rewrite helpers for edit commits.

use super::{
    Added, Arc, BTreeMap, BlobPart, Change, ColumnState, CreatedSheet, Error, GraphAction,
    GraphChange, HashSet, MergePlan, OpcPackage, OptionalAction, PackURI, Part, PartChange, Plan,
    Relationship, Result, RowState, SheetActions, State, TabAction, TargetMode, WebBindings,
    Workbook, Worksheet, WorksheetKind, allocation, chain, defaults_after, invalid, project_merges,
    raw,
};

pub(super) fn compose_part(
    parts: &mut Vec<PartChange>,
    workbook: &Workbook,
    uri: &PackURI,
    rewrite: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
) -> Result<()> {
    if let Some(part) = parts.iter_mut().find(|part| &part.uri == uri) {
        let after = rewrite(&part.after)?;
        if after.as_slice() != part.after.as_slice() {
            part.after = Arc::new(after);
        }
        return Ok(());
    }
    let before = workbook.inner.package.get_part(uri)?.blob_arc();
    let after = rewrite(&before)?;
    if after.as_slice() != before.as_slice() {
        parts.push(PartChange {
            uri: uri.clone(),
            before,
            after: Arc::new(after),
        });
    }
    Ok(())
}

pub(super) fn compose_part_optional(
    parts: &mut Vec<PartChange>,
    workbook: &Workbook,
    uri: &PackURI,
    rewrite: impl FnOnce(&[u8]) -> Result<Option<Vec<u8>>>,
) -> Result<()> {
    if let Some(part) = parts.iter_mut().find(|part| &part.uri == uri) {
        if let Some(after) = rewrite(&part.after)?
            && after.as_slice() != part.after.as_slice()
        {
            part.after = Arc::new(after);
        }
        return Ok(());
    }
    let before = workbook.inner.package.get_part(uri)?.blob_arc();
    let Some(after) = rewrite(&before)? else {
        return Ok(());
    };
    if after.as_slice() != before.as_slice() {
        parts.push(PartChange {
            uri: uri.clone(),
            before,
            after: Arc::new(after),
        });
    }
    Ok(())
}

pub(super) fn reference_part(part: &dyn Part) -> bool {
    let uri = part.partname().as_str();
    if uri.starts_with("/xl/externalLinks/")
        || part.content_type() == litchi_opc::constants::content_type::SML_EXTERNAL_LINK
    {
        return false;
    }
    (uri.starts_with("/xl/")
        && (part.content_type().ends_with("+xml")
            || part.content_type().ends_with("/xml")
            || part.content_type() == litchi_opc::constants::content_type::OFC_VML_DRAWING))
        || part.content_type() == litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES
}

pub(super) fn removal_reference_part(part: &dyn Part) -> bool {
    let uri = part.partname().as_str();
    if uri.starts_with("/xl/externalLinks/")
        || part.content_type() == litchi_opc::constants::content_type::SML_EXTERNAL_LINK
    {
        return false;
    }
    part.content_type().ends_with("+xml")
        || part.content_type().ends_with("/xml")
        || part.content_type() == litchi_opc::constants::content_type::OFC_VML_DRAWING
}

pub(super) fn ensure_unsigned(workbook: &Workbook) -> Result<()> {
    if workbook.inner.package.is_signed() {
        Err(Error::Signed)
    } else {
        Ok(())
    }
}

pub(super) fn validate_web_integrity(workbook: &Workbook) -> Result<()> {
    let refs = match workbook.task_panes()? {
        Some(panes) => crate::web::Refs::from_panes(panes)?,
        None => crate::web::Refs::new(std::iter::empty::<&str>())?,
    };
    for data in &workbook.inner.sheets {
        if data.kind != WorksheetKind::Worksheet {
            continue;
        }
        let sheet = Worksheet {
            owner: Arc::clone(&workbook.inner),
            data: Arc::clone(data),
        };
        check_web_bindings(&refs, &data.name, sheet.web_bindings()?)?;
    }
    Ok(())
}

pub(super) fn check_web_bindings(
    refs: &crate::web::Refs<'_>,
    sheet: &str,
    bindings: &WebBindings,
) -> Result<()> {
    bindings.validate_all()?;
    if let Some(binding) = bindings
        .iter()
        .find(|binding| !refs.contains(binding.app_ref()))
    {
        return Err(Error::DanglingWebBinding {
            sheet: sheet.to_owned(),
            app_ref: binding.app_ref().to_owned(),
        });
    }
    Ok(())
}

pub(super) fn create_sheets(
    workbook: &Workbook,
    added: Vec<Added>,
    positions: &[usize],
    active: Option<usize>,
    changes: &mut Vec<Change>,
    needs_recalculation: &mut bool,
) -> Result<Vec<CreatedSheet>> {
    if added.is_empty() {
        return Ok(Vec::new());
    }
    if positions.len() != added.len() {
        return Err(invalid(
            "created worksheet positions do not match the creation plan",
        ));
    }
    let main = workbook
        .inner
        .package
        .get_part(&workbook.inner.workbook_uri)?;
    let dialect = raw::catalog_edit::dialect(main.blob())?;
    let relationship_type = match dialect {
        raw::catalog_edit::Dialect::Transitional => {
            litchi_opc::constants::relationship_type::WORKSHEET
        },
        raw::catalog_edit::Dialect::Strict => {
            litchi_opc::constants::relationship_type::STRICT_WORKSHEET
        },
    };
    let namespace = dialect.worksheet_namespace();

    let mut used_sheet_ids = HashSet::new();
    used_sheet_ids
        .try_reserve(workbook.inner.sheets.len().saturating_add(added.len()))
        .map_err(|source| allocation("native sheet-ID index", source))?;
    used_sheet_ids.extend(workbook.inner.sheets.iter().map(|sheet| sheet.native_id));

    let mut used_relationship_ids = HashSet::<String>::new();
    used_relationship_ids
        .try_reserve(main.rels().len().saturating_add(added.len()))
        .map_err(|source| allocation("relationship-ID index", source))?;
    used_relationship_ids.extend(
        main.rels()
            .iter()
            .map(|relationship| relationship.r_id().to_owned()),
    );

    let mut reserved_parts = Vec::<PackURI>::new();
    reserved_parts
        .try_reserve_exact(added.len())
        .map_err(|source| allocation("worksheet part names", source))?;
    let mut created = Vec::new();
    created
        .try_reserve_exact(added.len())
        .map_err(|source| allocation("worksheet graph changes", source))?;

    let mut next_sheet_id = 1u32;
    let mut next_relationship_id = 1u32;
    let mut next_part = 1u32;
    for (index, added) in added.into_iter().enumerate() {
        while used_sheet_ids.contains(&next_sheet_id) {
            next_sheet_id = next_sheet_id
                .checked_add(1)
                .ok_or_else(|| invalid("native worksheet ID space is exhausted"))?;
        }
        if next_sheet_id > raw::catalog_edit::MAX_SHEET_ID {
            return Err(invalid("native worksheet ID space is exhausted"));
        }
        let sheet_id = next_sheet_id;
        used_sheet_ids.insert(sheet_id);
        next_sheet_id = next_sheet_id.saturating_add(1);

        let relationship_id = loop {
            let candidate = format!("rId{next_relationship_id}");
            next_relationship_id = next_relationship_id
                .checked_add(1)
                .ok_or_else(|| invalid("workbook relationship-ID space is exhausted"))?;
            if used_relationship_ids.insert(candidate.clone()) {
                break candidate;
            }
        };

        let part_uri = loop {
            let base_uri = workbook.inner.workbook_uri.base_uri();
            let candidate_path = if base_uri == "/" {
                format!("/worksheets/sheet{next_part}.xml")
            } else {
                format!("{base_uri}/worksheets/sheet{next_part}.xml")
            };
            let candidate = PackURI::new(candidate_path).map_err(invalid)?;
            next_part = next_part
                .checked_add(1)
                .ok_or_else(|| invalid("worksheet part-name space is exhausted"))?;
            if workbook
                .inner
                .package
                .validate_new_part_name(&candidate)
                .is_ok()
                && !reserved_parts
                    .iter()
                    .any(|reserved| reserved.is_equivalent_to(&candidate))
            {
                reserved_parts.push(candidate.clone());
                break candidate;
            }
        };

        let position = positions
            .get(index)
            .copied()
            .ok_or_else(|| invalid("created worksheet has no checked position"))?;
        let Added {
            name,
            actions,
            placement: _,
        } = added;
        let visibility = actions.visibility.unwrap_or(TabAction::Show);
        changes.push(Change::Create {
            sheet: name.as_str().into(),
            position,
            visibility: visibility.visibility(),
        });

        let SheetActions {
            rename: _,
            visibility: _,
            defaults,
            web,
            cells,
            rows,
            columns,
            merges,
            page_breaks,
            page_margins,
            page_setup,
            print_options,
        } = actions;
        let change_start = changes.len();
        if let Some(after) = &web
            && !after.is_empty()
        {
            changes.push(Change::Web {
                sheet: name.as_str().into(),
                before: WebBindings::new(),
                after: after.clone(),
            });
        }
        let merge_projection = project_merges(name.as_str(), None, merges, &cells)?;
        for (range, change) in &merge_projection.changes {
            changes.push(Change::Merge {
                sheet: name.as_str().into(),
                range: *range,
                change: *change,
            });
        }
        let mut effective_defaults = None;
        if let Some(action) = defaults {
            let after =
                defaults_after(None, action).map_err(|reason| Error::DefaultsEditBlocked {
                    sheet: name.as_str().to_owned(),
                    reason,
                })?;
            if after.is_some() {
                effective_defaults = Some(action);
                changes.push(Change::Defaults {
                    sheet: name.as_str().into(),
                    before: None,
                    after,
                });
            }
        }
        let mut effective_cells = BTreeMap::new();
        for (address, action) in cells {
            let before = State::Missing;
            let after = State::after(None, &action, workbook);
            if before == after {
                continue;
            }
            *needs_recalculation |=
                State::calculation_content(&before) != State::calculation_content(&after);
            effective_cells.insert(address, action);
            changes.push(Change::Cell {
                sheet: name.as_str().into(),
                address,
                before,
                after,
            });
        }
        let mut effective_rows = BTreeMap::new();
        for (row, action) in rows {
            let before = RowState::Missing;
            let after = RowState::after(None, action, workbook);
            if before == after {
                continue;
            }
            effective_rows.insert(row, action);
            changes.push(Change::Row {
                sheet: name.as_str().into(),
                row,
                before,
                after,
            });
        }
        let mut effective_columns = BTreeMap::new();
        for (column, action) in columns {
            let before = ColumnState::Missing;
            let after = ColumnState::after(None, action, workbook);
            if before == after {
                continue;
            }
            effective_columns.insert(column, action);
            changes.push(Change::Column {
                sheet: name.as_str().into(),
                column,
                before,
                after,
            });
        }

        let template = format!(
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<worksheet xmlns="{}"><dimension ref="A1"/><sheetData/></worksheet>"#
            ),
            namespace
        )
        .into_bytes();
        let mut content = raw::worksheet::edit::rewrite(
            &template,
            name.as_str(),
            Plan {
                defaults: effective_defaults,
                cells: effective_cells,
                rows: effective_rows,
                columns: effective_columns,
            },
        )?;
        let MergePlan { add, remove } = merge_projection.plan;
        if !remove.is_empty() {
            return Err(invalid(
                "new worksheet merge projection unexpectedly removed a range",
            ));
        }
        if !add.is_empty() {
            content = raw::worksheet::edit::rewrite_merges(
                &content,
                name.as_str(),
                MergePlan {
                    add,
                    remove: Vec::new(),
                },
            )?;
        }
        if let Some(bindings) = &web {
            content = raw::web::replace(&content, bindings)?;
        }
        if let Some(page_breaks) = &page_breaks
            && page_breaks != &crate::page_breaks::PageBreaks::new()
        {
            changes.push(Change::PageBreaks {
                sheet: name.as_str().into(),
                before: crate::page_breaks::PageBreaks::new(),
                after: page_breaks.clone(),
            });
            content = crate::page_breaks::replace(&content, page_breaks)?;
        }
        if let Some(OptionalAction::Put(page_margins)) = &page_margins {
            changes.push(Change::PageMargins {
                sheet: name.as_str().into(),
                before: None,
                after: Some(*page_margins),
            });
            content = crate::page_margins::replace_page_margins(&content, Some(page_margins))?;
        }
        if let Some(OptionalAction::Put(page_setup)) = &page_setup {
            changes.push(Change::PageSetup {
                sheet: name.as_str().into(),
                before: None,
                after: Some(page_setup.clone()),
            });
            content = crate::page_setup::replace_worksheet_page_setup(&content, Some(page_setup))?;
        }
        if let Some(OptionalAction::Put(print_options)) = &print_options {
            changes.push(Change::PrintOptions {
                sheet: name.as_str().into(),
                before: None,
                after: Some(*print_options),
            });
            content = crate::print_options::replace_print_options(&content, Some(print_options))?;
        }
        if active == Some(index) {
            content = raw::sheet_view_edit::rewrite(
                &content,
                true,
                raw::sheet_view_edit::Context {
                    sheet: name.as_str(),
                    position,
                },
            )?;
        }
        let parsed = raw::worksheet::parse(&content, || workbook.inner.shared_strings())?;
        workbook.inner.validate_styles(&parsed)?;
        for change in &changes[change_start..] {
            match change {
                Change::Merge {
                    sheet,
                    range,
                    change,
                    ..
                } => {
                    if parsed.merge_ranges().contains(range) != change.after() {
                        return Err(invalid(format!(
                            "new worksheet merged-range verification failed at {sheet}!{range}"
                        )));
                    }
                },
                Change::Cell {
                    sheet,
                    address,
                    after,
                    ..
                } => {
                    if State::read(parsed.entry(*address), workbook) != *after {
                        return Err(invalid(format!(
                            "new worksheet verification failed at {sheet}!{address}"
                        )));
                    }
                },
                Change::Row {
                    sheet, row, after, ..
                } => {
                    if RowState::read(parsed.row_entry(*row), workbook) != *after {
                        return Err(invalid(format!(
                            "new worksheet row verification failed at {sheet}!row {}",
                            row.get()
                        )));
                    }
                },
                Change::Column {
                    sheet,
                    column,
                    after,
                    ..
                } => {
                    if ColumnState::read(parsed.column_entry(*column), workbook) != *after {
                        return Err(invalid(format!(
                            "new worksheet column verification failed at {sheet}!column {}",
                            column.get()
                        )));
                    }
                },
                Change::PageBreaks {
                    sheet,
                    after: expected,
                    ..
                } => {
                    let actual = crate::page_breaks::parse(&content)?;
                    if &actual != expected {
                        return Err(invalid(format!(
                            "new worksheet page-break verification failed at {sheet}"
                        )));
                    }
                },
                Change::PageMargins {
                    sheet,
                    after: expected,
                    ..
                } => {
                    let actual = crate::page_margins::parse_page_margins(&content)?;
                    if &actual != expected {
                        return Err(invalid(format!(
                            "new worksheet page-margin verification failed at {sheet}"
                        )));
                    }
                },
                Change::PageSetup {
                    sheet,
                    after: expected,
                    ..
                } => {
                    let actual = crate::page_setup::parse_worksheet_page_setup(&content)?;
                    if &actual != expected {
                        return Err(invalid(format!(
                            "new worksheet page-setup verification failed at {sheet}"
                        )));
                    }
                },
                Change::PrintOptions {
                    sheet,
                    after: expected,
                    ..
                } => {
                    let actual = crate::print_options::parse_print_options(&content)?;
                    if &actual != expected {
                        return Err(invalid(format!(
                            "new worksheet print-options verification failed at {sheet}"
                        )));
                    }
                },
                Change::Defaults { sheet, after, .. } => {
                    if parsed.defaults() != after.as_ref() {
                        return Err(invalid(format!(
                            "new worksheet defaults verification failed at {sheet}"
                        )));
                    }
                },
                Change::Web {
                    sheet,
                    after: expected,
                    ..
                } => {
                    if raw::web::read(&content)? != *expected {
                        return Err(invalid(format!(
                            "new worksheet web-binding verification failed at {sheet}"
                        )));
                    }
                },
                Change::Create { .. }
                | Change::Remove { .. }
                | Change::Rename { .. }
                | Change::Move { .. }
                | Change::Active { .. }
                | Change::Visibility { .. } => {},
            }
        }

        let target_ref = part_uri.relative_ref(workbook.inner.workbook_uri.base_uri());
        let relationship = Relationship::new_with_mode(
            relationship_id.clone(),
            relationship_type.to_owned(),
            target_ref,
            workbook.inner.workbook_uri.base_uri().to_owned(),
            TargetMode::Internal,
        );
        let part = BlobPart::new(
            part_uri,
            litchi_opc::constants::content_type::SML_WORKSHEET.to_owned(),
            content,
        );
        created.push(CreatedSheet {
            name,
            position,
            sheet_id,
            relationship_id,
            visibility,
            graph: GraphChange {
                action: GraphAction::Add,
                source: workbook.inner.workbook_uri.clone(),
                relationship,
                part: Box::new(part),
            },
        });
    }
    Ok(created)
}

pub(super) fn calculation_chain_removal(workbook: &Workbook) -> Result<Vec<GraphChange>> {
    chain::validate_package(&workbook.inner.package)?;
    let main = workbook
        .inner
        .package
        .get_part(&workbook.inner.workbook_uri)?;
    let mut matching = main.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::CALC_CHAIN
                | litchi_opc::constants::relationship_type::STRICT_CALC_CHAIN
        )
    });
    let Some(relationship) = matching.next() else {
        return Ok(Vec::new());
    };
    if matching.next().is_some() {
        return Err(invalid(
            "workbook has multiple calculation-chain relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid("calculation-chain relationship cannot be external"));
    }
    let target = relationship.target_partname()?;
    let part = workbook.inner.package.get_part(&target)?;
    ensure_exclusive_incoming_relationship(
        &workbook.inner.package,
        part.partname(),
        &workbook.inner.workbook_uri,
        relationship.r_id(),
    )?;
    Ok(vec![GraphChange {
        action: GraphAction::Remove,
        source: workbook.inner.workbook_uri.clone(),
        relationship: relationship.clone(),
        part: part.clone_part(),
    }])
}

pub(super) fn ensure_exclusive_incoming_relationship(
    package: &OpcPackage,
    target: &PackURI,
    expected_source: &PackURI,
    expected_id: &str,
) -> Result<()> {
    let targets = |relationship: &Relationship| -> Result<bool> {
        if relationship.is_external() {
            return Ok(false);
        }
        relationship
            .target_partname()
            .map(|candidate| candidate.as_str().eq_ignore_ascii_case(target.as_str()))
            .map_err(Into::into)
    };
    for relationship in package.rels().iter() {
        if targets(relationship)? {
            return Err(invalid(format!(
                "calculation-chain part '{target}' has another incoming relationship"
            )));
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if targets(relationship)?
                && !(source.partname() == expected_source && relationship.r_id() == expected_id)
            {
                return Err(invalid(format!(
                    "calculation-chain part '{target}' has another incoming relationship"
                )));
            }
        }
    }
    Ok(())
}

pub(super) fn same_relationship(left: &Relationship, right: &Relationship) -> bool {
    left.r_id() == right.r_id()
        && left.reltype() == right.reltype()
        && left.target_ref() == right.target_ref()
        && left.target_mode() == right.target_mode()
}

pub(super) fn same_part(left: &dyn Part, right: &dyn Part) -> bool {
    left.partname() == right.partname()
        && left.content_type() == right.content_type()
        && left.blob() == right.blob()
        && left.rels().len() == right.rels().len()
        && left.rels().iter().all(|relationship| {
            right
                .rels()
                .get(relationship.r_id())
                .is_some_and(|other| same_relationship(relationship, other))
        })
}
