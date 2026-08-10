//! Checked transfer of ordinary RTF owners and their bounded dependencies.

use super::{Commit, Edit, Error, HeaderFooterParagraph, Snapshot};
use crate::{RtfWriter, TableCellPath};
use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet};

/// A checked, non-applying transfer into one immutable target snapshot.
pub struct TransferPlan {
    edit: Edit,
    dependency_count: usize,
}

impl TransferPlan {
    /// Plans insertion of one source paragraph as a plain target paragraph.
    ///
    /// # Errors
    /// Returns an error when the source contains an inline line break or when
    /// the checked target cannot accept the structural insertion.
    pub fn plain_paragraph(
        source: &Snapshot,
        source_position: usize,
        target: &Snapshot,
        insert_after: usize,
    ) -> Result<Self, Error> {
        let paragraphs = source.body().paragraphs().collect::<Vec<_>>();
        let count = paragraphs.len();
        let text = paragraphs
            .get(source_position)
            .ok_or(Error::ParagraphOutOfRange {
                position: source_position,
                count,
            })?
            .to_text();
        if text.contains('\n') {
            return Err(Error::UnsupportedSource(
                "plain paragraph transfer refuses inline line breaks",
            ));
        }
        let mut edit = target.edit();
        edit.insert_paragraph_after(insert_after, text)?;
        Ok(Self {
            edit,
            dependency_count: 0,
        })
    }

    /// Plans copying plain text between checked table-cell destinations.
    ///
    /// # Errors
    /// Returns an error when either path is invalid or the target cell's
    /// dependent positional content cannot survive the replacement.
    pub fn table_cell_text(
        source: &Snapshot,
        source_path: &TableCellPath,
        target: &Snapshot,
        target_path: TableCellPath,
    ) -> Result<Self, Error> {
        let text = super::table_cell(source, source_path)?.text().to_string();
        let mut edit = target.edit();
        edit.set_table_cell_text(target_path, text)?;
        Ok(Self {
            edit,
            dependency_count: 0,
        })
    }

    /// Plans copying one plain header/footer paragraph into another.
    ///
    /// # Errors
    /// Returns an error when either selector is invalid or the target story
    /// owns positioned content that would acquire stale offsets.
    pub fn header_footer_text(
        source: &Snapshot,
        source_target: HeaderFooterParagraph,
        target: &Snapshot,
        target_destination: HeaderFooterParagraph,
    ) -> Result<Self, Error> {
        let text = super::header_footer(source, source_target)?
            .paragraphs
            .get(source_target.paragraph())
            .ok_or(Error::DestinationOutOfRange("header/footer paragraph"))?
            .text
            .to_string();
        let mut edit = target.edit();
        edit.set_header_footer_text(target_destination, text)?;
        Ok(Self {
            edit,
            dependency_count: 0,
        })
    }

    /// Plans transfer of one inert, plain-result body field at the target body end.
    ///
    /// Active/external field kinds and nested generic fields are refused.
    /// Self-contained result drawings and explicit breaks are retained.
    ///
    /// # Errors
    /// Returns an error for an invalid selector, active field, unresolved
    /// result-story dependency, opaque syntax, or invalid candidate readback.
    pub fn field(source: &Snapshot, source_index: usize, target: &Snapshot) -> Result<Self, Error> {
        let source_field = source
            .fields()
            .get(source_index)
            .ok_or(Error::DestinationOutOfRange("field"))?;
        refuse_feature_opaque(source, crate::opaque::Context::Field, source_field.position)?;
        if active_field(source_field.field_type) {
            return Err(Error::UnsupportedSource(
                "field transfer refuses active or external field instructions",
            ));
        }
        if source_field
            .result_events
            .iter()
            .any(|event| matches!(event, crate::StoryEvent::Field(_)))
        {
            return Err(Error::UnsupportedSource(
                "field transfer refuses nested generic-field dependencies",
            ));
        }
        if !source_field.shapes.is_empty() || !source_field.shape_groups.is_empty() {
            refuse_feature_opaque(source, crate::opaque::Context::Drawing, usize::MAX)?;
            refuse_drawing_story_fields(&source_field.shapes, &source_field.shape_groups)?;
        }
        let mut field = owned_field(source_field);
        let dependency_count = field.shapes.len().saturating_add(field.shape_groups.len());
        field.owner = crate::FieldOwner::Body;
        field.position = target.text().len();
        field.range_end = field.position;
        let after = canonical_candidate(target, |model| model.push_field(field))?;
        root_plan(
            target,
            "field.transfer",
            "body:field:append",
            after,
            dependency_count,
        )
    }

    /// Plans transfer of one inert root body text frame at the target body end.
    ///
    /// Geometry, properties, text formatting, fallbacks, and self-contained
    /// nested drawings are retained when font/color resources match. Active
    /// hyperlinks, nested generic fields, and unknown drawing syntax are refused.
    ///
    /// # Errors
    /// Returns an error for an invalid selector, active or unresolved drawing
    /// dependency, opaque syntax, or invalid candidate readback.
    pub fn shape(source: &Snapshot, source_index: usize, target: &Snapshot) -> Result<Self, Error> {
        require_equal_text_resources(source, target)?;
        let shape = source
            .shapes()
            .get(source_index)
            .ok_or(Error::DestinationOutOfRange("shape"))?;
        if !shape.text_destination_present
            || matches!(shape.shape_type, crate::ShapeType::PictureFrame)
        {
            return Err(Error::UnsupportedSource(
                "shape transfer supports retained non-picture text frames",
            ));
        }
        refuse_feature_opaque(source, crate::opaque::Context::Drawing, shape.position)?;
        refuse_drawing_story_fields(std::slice::from_ref(shape), &[])?;
        if super::shape_has_active_link(shape) {
            return Err(Error::UnsupportedSource(
                "shape transfer refuses active hyperlink metadata",
            ));
        }
        let dependency_count = shape_dependency_count(shape);
        let mut transferred_shape = shape.clone().into_owned();
        transferred_shape.position = target.text().len();
        let effect = format!("body:shape:append:{source_index}");
        let after = canonical_candidate(target, |model| model.push_shape(transferred_shape))?;
        root_plan(target, "shape.transfer", &effect, after, dependency_count)
    }

    /// Plans copying one complete nested table tree into a target cell end.
    ///
    /// Nested table text, geometry, borders, and recursively nested tables are
    /// retained. Self-contained drawings, explicit breaks, and passive generic
    /// fields are copied; navigation/revision handles and incompatible
    /// table-style dependencies are refused.
    ///
    /// # Errors
    /// Returns an error for invalid paths, unresolved dependencies, opaque
    /// syntax, or invalid candidate readback.
    pub fn nested_table(
        source: &Snapshot,
        source_cell: &TableCellPath,
        source_nested_index: usize,
        target: &Snapshot,
        target_cell: TableCellPath,
    ) -> Result<Self, Error> {
        let nested = super::table_cell(source, source_cell)?
            .nested_tables()
            .get(source_nested_index)
            .ok_or(Error::DestinationOutOfRange("nested table"))?;
        let source_parent_depth = table_cell_depth(source_cell)?;
        let target_parent_depth = table_cell_depth(&target_cell)?;
        refuse_feature_opaque(source, crate::opaque::Context::Table, usize::MAX)?;
        let field_indices = table_story_field_indices(&nested.table)?;
        let field_count = field_indices.len();
        let mut field_drawing_count = 0usize;
        let mut transferred_fields = Vec::new();
        for index in field_indices {
            let field = source.fields().get(index).ok_or(Error::UnsupportedSource(
                "nested-table story references a missing field",
            ))?;
            field_drawing_count =
                field_drawing_count.saturating_add(validate_transferable_field(source, field)?);
            let mut transferred_field = owned_field(field);
            let crate::FieldOwner::TableCell(source_depth) = transferred_field.owner else {
                return Err(Error::UnsupportedSource(
                    "nested-table story field has an incompatible owner",
                ));
            };
            let relative_depth =
                source_depth
                    .checked_sub(source_parent_depth)
                    .ok_or(Error::UnsupportedSource(
                        "nested-table field owner precedes its selected table",
                    ))?;
            if relative_depth == 0 {
                return Err(Error::UnsupportedSource(
                    "nested-table field owner does not belong to the selected table",
                ));
            }
            let target_depth =
                target_parent_depth
                    .checked_add(relative_depth)
                    .ok_or(Error::UnsupportedSource(
                        "nested-table target field depth overflows",
                    ))?;
            transferred_field.owner = crate::FieldOwner::TableCell(target_depth);
            transferred_fields.push((index, transferred_field));
        }
        let drawing_count = validate_table_transfer(source, &nested.table)?;
        let styled_rows = count_styled_rows(&nested.table);
        if styled_rows != 0 && source.styles() != target.styles() {
            return Err(Error::UnsupportedSource(
                "nested-table transfer requires identical table-style dependencies",
            ));
        }
        let mut table = crate::document::owned_table(&nested.table)?;
        let target_offset = super::table_cell(target, &target_cell)?.text().len();
        let effect = format!(
            "{}:nested-table:append",
            super::table_cell_effect(&target_cell)
        );
        let after = canonical_candidate(target, move |model| {
            let mut nested_field_indices = BTreeMap::new();
            for (source_index, field) in transferred_fields {
                nested_field_indices.insert(source_index, model.push_story_field_metadata(field)?);
            }
            remap_table_field_indices(&mut table, &nested_field_indices)?;
            model
                .table_cell_mut(&target_cell)?
                .add_nested_table(target_offset, table)
        })?;
        let dependency_count = field_count
            .saturating_add(field_drawing_count)
            .saturating_add(drawing_count)
            .saturating_add(styled_rows);
        root_plan(
            target,
            "nested-table.transfer",
            &effect,
            after,
            dependency_count,
        )
    }

    /// Plans copying one style and its based-on/next/linked style closure.
    ///
    /// Font and color tables must be identical so copied numeric references
    /// retain their meaning. Existing equal definitions are reused; conflicting
    /// typed IDs are refused.
    ///
    /// # Errors
    /// Returns an error for missing or conflicting dependencies, opaque syntax,
    /// or invalid candidate readback.
    pub fn style(
        source: &Snapshot,
        kind: crate::style::Kind,
        id: u16,
        target: &Snapshot,
    ) -> Result<Self, Error> {
        refuse_feature_opaque(source, crate::opaque::Context::Metadata, usize::MAX)?;
        require_equal_text_resources(source, target)?;
        let styles = style_closure(source, kind, id)?;
        let mut additions = Vec::new();
        for style in styles {
            match target
                .model()
                .stylesheet()
                .get_typed(style.style_type, style.id)
            {
                Some(existing) if existing == style => {},
                Some(_) => {
                    return Err(Error::UnsupportedSource(
                        "style transfer found a conflicting typed style ID",
                    ));
                },
                None => additions.push(owned_style(style)),
            }
        }
        let dependency_count = additions.len().saturating_sub(1);
        let effect = format!("stylesheet:{kind:?}:{id}");
        let after = canonical_candidate(target, |model| {
            for style in additions {
                model.stylesheet_mut().add(style);
            }
            Ok(())
        })?;
        root_plan(target, "style.transfer", &effect, after, dependency_count)
    }

    /// Plans copying one list definition and all overrides that reference it.
    ///
    /// List fonts must resolve identically in both documents. Referenced
    /// picture-bullet slots and their shared pictures are copied and remapped.
    /// Existing equal definitions are reused; ID/index collisions are refused.
    ///
    /// # Errors
    /// Returns an error for missing or conflicting dependencies, opaque syntax,
    /// or invalid candidate readback.
    pub fn list(source: &Snapshot, id: i32, target: &Snapshot) -> Result<Self, Error> {
        refuse_feature_opaque(source, crate::opaque::Context::Metadata, usize::MAX)?;
        require_equal_text_resources(source, target)?;
        let list = source
            .model()
            .list_table()
            .get(id)
            .ok_or(Error::DestinationOutOfRange("list definition"))?;
        let add_list = match target.model().list_table().get(id) {
            Some(existing) if existing == list => false,
            Some(_) => {
                return Err(Error::UnsupportedSource(
                    "list transfer found a conflicting list ID",
                ));
            },
            None => true,
        };
        let mut overrides = Vec::new();
        for entry in source
            .model()
            .list_override_table()
            .overrides()
            .iter()
            .filter(|entry| entry.list_id == id)
        {
            match target.model().list_override_table().get(entry.index) {
                Some(existing) if existing == entry => {},
                Some(_) => {
                    return Err(Error::UnsupportedSource(
                        "list transfer found a conflicting override index",
                    ));
                },
                None => overrides.push(entry.clone()),
            }
        }
        let mut owned_list = owned_list(list);
        let mut picture_slots = Vec::new();
        if add_list {
            let source_slots = list
                .levels
                .iter()
                .filter_map(|level| level.picture_index)
                .collect::<BTreeSet<_>>();
            if !source_slots.is_empty() {
                refuse_feature_opaque(source, crate::opaque::Context::Drawing, usize::MAX)?;
            }
            let first_target_slot = target.model().list_table().picture_bullet_count;
            let mut slot_mapping = BTreeMap::new();
            for source_slot in source_slots {
                let source_slot_index = usize::try_from(source_slot).map_err(|_error| {
                    Error::UnsupportedSource("picture-bullet source index overflows")
                })?;
                let offset = u32::try_from(picture_slots.len()).map_err(|_error| {
                    Error::UnsupportedSource("picture-bullet dependency count overflows")
                })?;
                let target_slot =
                    first_target_slot
                        .checked_add(offset)
                        .ok_or(Error::UnsupportedSource(
                            "picture-bullet target index overflows",
                        ))?;
                let source_picture = source
                    .model()
                    .list_table()
                    .picture_bullet_picture_indices()
                    .get(source_slot_index)
                    .copied()
                    .flatten()
                    .and_then(|picture_index| source.pictures().get(picture_index))
                    .cloned()
                    .map(crate::Picture::into_owned)
                    .ok_or(Error::UnsupportedSource(
                        "picture-bullet slot references a missing picture",
                    ))?;
                slot_mapping.insert(source_slot, target_slot);
                picture_slots.push(source_picture);
            }
            for level in &mut owned_list.levels {
                if let Some(source_slot) = level.picture_index {
                    level.picture_index = slot_mapping.get(&source_slot).copied();
                }
            }
        }
        let picture_dependency_count = picture_slots.len().saturating_mul(2);
        let dependency_count = overrides.len().saturating_add(picture_dependency_count);
        let effect = format!("list-table:list:{id}");
        let after = canonical_candidate(target, |model| {
            if add_list {
                let mut target_slots = model.list_table().picture_bullet_picture_indices().to_vec();
                for picture in picture_slots {
                    target_slots.push(Some(model.push_picture(picture)));
                }
                model.set_list_picture_bullet_indices(target_slots)?;
                model.list_table_mut().add(owned_list);
            }
            for entry in overrides {
                model.list_override_table_mut().add(entry);
            }
            Ok(())
        })?;
        root_plan(target, "list.transfer", &effect, after, dependency_count)
    }

    /// Plans transfer of one inert embedded object and its result pictures.
    ///
    /// External and automatically updated links are refused. Referenced result
    /// pictures are copied and their indices remapped into the target store.
    ///
    /// # Errors
    /// Returns an error for invalid selectors, active/external dependencies,
    /// opaque syntax, or invalid candidate readback.
    pub fn object(
        source: &Snapshot,
        source_index: usize,
        target: &Snapshot,
    ) -> Result<Self, Error> {
        let object = source
            .model()
            .objects()
            .get(source_index)
            .ok_or(Error::DestinationOutOfRange("embedded object"))?;
        refuse_feature_opaque(source, crate::opaque::Context::Drawing, object.position)?;
        if matches!(
            object.kind,
            crate::ObjectKind::Link
                | crate::ObjectKind::AutoLink
                | crate::ObjectKind::Subscriber
                | crate::ObjectKind::Publisher
        ) || object.update_requested
        {
            return Err(Error::UnsupportedSource(
                "object transfer refuses external or automatically updated dependencies",
            ));
        }
        let mut pictures = Vec::new();
        for index in &object.result_picture_indices {
            let picture = source
                .pictures()
                .get(*index)
                .ok_or(Error::UnsupportedSource(
                    "object result references a missing picture",
                ))?;
            pictures.push(picture.clone().into_owned());
        }
        let dependency_count = pictures.len();
        let mut transferred_object = owned_object(object);
        transferred_object.position = target.text().len();
        transferred_object.result_picture_indices.clear();
        let effect = format!("body:object:append:{source_index}");
        let after = canonical_candidate(target, |model| {
            for picture in pictures {
                transferred_object
                    .result_picture_indices
                    .push(model.push_picture(picture));
            }
            model.push_object(transferred_object)
        })?;
        root_plan(target, "object.transfer", &effect, after, dependency_count)
    }

    /// Number of resource or owner dependencies carried by this plan.
    #[must_use]
    pub const fn dependency_count(&self) -> usize {
        self.dependency_count
    }

    /// Whether this plan imports no format resource handles.
    #[must_use]
    pub const fn is_dependency_free(&self) -> bool {
        self.dependency_count == 0
    }

    /// Returns the still-uncommitted target edit.
    #[must_use]
    pub fn into_edit(self) -> Edit {
        self.edit
    }

    /// Validates and publishes the target transaction atomically.
    ///
    /// # Errors
    /// Returns the ordinary transaction refusal without mutating either input.
    pub fn commit(self) -> Result<Commit, Error> {
        self.edit.commit()
    }
}

fn root_plan(
    target: &Snapshot,
    vocabulary: &'static str,
    effect: &str,
    after: Vec<u8>,
    dependency_count: usize,
) -> Result<TransferPlan, Error> {
    let mut edit = target.edit();
    edit.stage_root_transfer(vocabulary, effect.to_string(), after)?;
    Ok(TransferPlan {
        edit,
        dependency_count,
    })
}

fn canonical_candidate(
    target: &Snapshot,
    mutation: impl FnOnce(&mut crate::document::RtfDocument<'static>) -> crate::RtfResult<()>,
) -> Result<Vec<u8>, Error> {
    if !target.opaque().is_empty() {
        return Err(Error::UnsupportedSource(
            "ordinary-root transfer refuses unknown target destinations",
        ));
    }
    let bytes = target
        .source_bytes()
        .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
    if crate::compressed::is_compressed_rtf(bytes) {
        return Err(Error::UnsupportedSource(
            "compressed RTF needs a transport-aware rewrite",
        ));
    }
    let mut model = crate::document::RtfDocument::parse_bytes_with_limits(bytes, target.limits())?;
    mutation(&mut model)?;
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&model)
        .map_err(|error| Error::Write(error.to_string()))?;
    let limit = target.limits().max_source_bytes();
    if output.len() > limit {
        return Err(Error::InputTooLarge {
            observed: output.len(),
            limit,
        });
    }
    let reopened = Snapshot::from_bytes_with_limits(&output, target.limits())?;
    if !reopened.opaque().is_empty() {
        return Err(Error::UnsupportedSource(
            "ordinary-root transfer produced unknown destinations",
        ));
    }
    Ok(output)
}

fn refuse_feature_opaque(
    source: &Snapshot,
    context: crate::opaque::Context,
    body_position: usize,
) -> Result<(), Error> {
    if source.opaque().iter().any(|node| match node.anchor() {
        crate::opaque::Anchor::Body(position) => position == body_position,
        crate::opaque::Anchor::Structural { context: owner, .. } => owner == context,
    }) {
        return Err(Error::UnsupportedSource(
            "selected feature owns an unknown destination",
        ));
    }
    Ok(())
}

fn require_equal_text_resources(source: &Snapshot, target: &Snapshot) -> Result<(), Error> {
    if source.model().font_table() != target.model().font_table()
        || !equal_color_tables(source.model().color_table(), target.model().color_table())
    {
        return Err(Error::UnsupportedSource(
            "transfer requires identical font and color dependencies",
        ));
    }
    Ok(())
}

fn equal_color_tables(left: &crate::ColorTable, right: &crate::ColorTable) -> bool {
    left.colors() == right.colors()
        && (0..left.colors().len()).all(|index| {
            let reference = u16::try_from(index).unwrap_or(u16::MAX);
            left.is_automatic(reference) == right.is_automatic(reference)
        })
}

fn active_field(kind: crate::FieldType) -> bool {
    matches!(
        kind,
        crate::FieldType::MacroButton
            | crate::FieldType::GoToButton
            | crate::FieldType::Print
            | crate::FieldType::Embed
            | crate::FieldType::AddIn
            | crate::FieldType::Control
            | crate::FieldType::HtmlControl
            | crate::FieldType::Dde
            | crate::FieldType::DdeAuto
            | crate::FieldType::Link
            | crate::FieldType::Include
            | crate::FieldType::Import
            | crate::FieldType::IncludeText
            | crate::FieldType::IncludePicture
            | crate::FieldType::Database
            | crate::FieldType::Ask
            | crate::FieldType::FillIn
    )
}

fn owned_field(field: &crate::Field<'_>) -> crate::Field<'static> {
    crate::Field {
        field_type: field.field_type,
        instruction: Cow::Owned(field.instruction.to_string()),
        result: Cow::Owned(field.result.to_string()),
        status: field.status,
        shapes: field
            .shapes
            .iter()
            .cloned()
            .map(crate::Shape::into_owned)
            .collect(),
        shape_groups: field
            .shape_groups
            .iter()
            .cloned()
            .map(crate::ShapeGroup::into_owned)
            .collect(),
        drawing_order: field.drawing_order.clone(),
        result_events: field.result_events.clone(),
        owner: field.owner,
        position: field.position,
        range_end: field.range_end,
    }
}

fn validate_table_transfer(source: &Snapshot, table: &crate::Table<'_>) -> Result<usize, Error> {
    let mut drawing_count = 0usize;
    for row in table.rows() {
        for cell in row.cells() {
            refuse_drawing_story_fields(cell.shapes(), cell.shape_groups())?;
            drawing_count = drawing_count
                .saturating_add(cell.shapes().len())
                .saturating_add(cell.shape_groups().len());
            if cell.story_events().iter().any(|event| {
                matches!(
                    event,
                    crate::CellStoryEvent::NavigationEntry(_)
                        | crate::CellStoryEvent::RevisionStart(_)
                        | crate::CellStoryEvent::RevisionEnd(_)
                        | crate::CellStoryEvent::RevisionDeletion(_)
                )
            }) {
                return Err(Error::UnsupportedSource(
                    "nested-table transfer refuses unresolved navigation or revision handles",
                ));
            }
            for nested in cell.nested_tables() {
                drawing_count =
                    drawing_count.saturating_add(validate_table_transfer(source, &nested.table)?);
            }
        }
    }
    if drawing_count != 0 {
        refuse_feature_opaque(source, crate::opaque::Context::Drawing, usize::MAX)?;
    }
    Ok(drawing_count)
}

fn table_story_field_indices(table: &crate::Table<'_>) -> Result<BTreeSet<usize>, Error> {
    let mut indices = BTreeSet::new();
    collect_table_story_field_indices(table, &mut indices)?;
    Ok(indices)
}

fn collect_table_story_field_indices(
    table: &crate::Table<'_>,
    indices: &mut BTreeSet<usize>,
) -> Result<(), Error> {
    for row in table.rows() {
        for cell in row.cells() {
            for event in cell.story_events() {
                if let crate::CellStoryEvent::Field(field) = event
                    && !indices.insert(field.field_index)
                {
                    return Err(Error::UnsupportedSource(
                        "nested-table field metadata has multiple story owners",
                    ));
                }
            }
            for nested in cell.nested_tables() {
                collect_table_story_field_indices(&nested.table, indices)?;
            }
        }
    }
    Ok(())
}

fn table_cell_depth(path: &TableCellPath) -> Result<u8, Error> {
    let depth = path
        .nested
        .len()
        .checked_add(1)
        .ok_or(Error::UnsupportedSource("table-cell path depth overflows"))?;
    u8::try_from(depth)
        .map_err(|_error| Error::UnsupportedSource("table-cell path depth exceeds RTF limits"))
}

fn remap_table_field_indices(
    table: &mut crate::Table<'static>,
    indices: &BTreeMap<usize, usize>,
) -> crate::RtfResult<()> {
    for row in table.rows_mut() {
        for cell in row.cells_mut() {
            for nested in cell.nested_tables_mut() {
                remap_table_field_indices(&mut nested.table, indices)?;
            }
            let mut events = cell.story_events().to_vec();
            for event in &mut events {
                if let crate::CellStoryEvent::Field(field) = event {
                    field.field_index = *indices.get(&field.field_index).ok_or_else(|| {
                        crate::RtfError::MalformedDocument(
                            "nested-table field remapping is incomplete".to_string(),
                        )
                    })?;
                }
            }
            cell.set_story_content(
                cell.shapes().to_vec(),
                cell.shape_groups().to_vec(),
                cell.drawing_order().to_vec(),
                events,
            )?;
        }
    }
    Ok(())
}

fn validate_transferable_field(
    source: &Snapshot,
    field: &crate::Field<'_>,
) -> Result<usize, Error> {
    if !matches!(field.owner, crate::FieldOwner::TableCell(_)) {
        return Err(Error::UnsupportedSource(
            "nested-table story field has an incompatible owner",
        ));
    }
    if active_field(field.field_type) {
        return Err(Error::UnsupportedSource(
            "nested-table field transfer refuses active or external instructions",
        ));
    }
    if field
        .result_events
        .iter()
        .any(|event| matches!(event, crate::StoryEvent::Field(_)))
    {
        return Err(Error::UnsupportedSource(
            "nested-table field transfer refuses nested generic fields",
        ));
    }
    refuse_feature_opaque(source, crate::opaque::Context::Field, field.position)?;
    if !field.shapes.is_empty() || !field.shape_groups.is_empty() {
        refuse_feature_opaque(source, crate::opaque::Context::Drawing, usize::MAX)?;
        refuse_drawing_story_fields(&field.shapes, &field.shape_groups)?;
    }
    Ok(field.shapes.len().saturating_add(field.shape_groups.len()))
}

fn refuse_drawing_story_fields(
    shapes: &[crate::Shape<'_>],
    groups: &[crate::ShapeGroup<'_>],
) -> Result<(), Error> {
    if shapes.iter().any(shape_has_story_fields) || groups.iter().any(shape_group_has_story_fields)
    {
        return Err(Error::UnsupportedSource(
            "drawing transfer refuses unresolved shape-text field handles",
        ));
    }
    Ok(())
}

fn shape_has_story_fields(shape: &crate::Shape<'_>) -> bool {
    shape
        .text_story_events
        .iter()
        .any(|event| matches!(event, crate::StoryEvent::Field(_)))
        || shape.text_shapes.iter().any(shape_has_story_fields)
        || shape
            .text_shape_groups
            .iter()
            .any(shape_group_has_story_fields)
}

fn shape_group_has_story_fields(group: &crate::ShapeGroup<'_>) -> bool {
    group.shapes.iter().any(shape_has_story_fields)
        || group.groups.iter().any(shape_group_has_story_fields)
}

fn shape_dependency_count(shape: &crate::Shape<'_>) -> usize {
    let mut count = shape
        .text_shapes
        .len()
        .saturating_add(shape.text_shape_groups.len());
    for nested in &shape.text_shapes {
        count = count.saturating_add(shape_dependency_count(nested));
    }
    for group in &shape.text_shape_groups {
        count = count.saturating_add(shape_group_dependency_count(group));
    }
    count
}

fn shape_group_dependency_count(group: &crate::ShapeGroup<'_>) -> usize {
    let mut count = group.shapes.len().saturating_add(group.groups.len());
    for shape in &group.shapes {
        count = count.saturating_add(shape_dependency_count(shape));
    }
    for nested in &group.groups {
        count = count.saturating_add(shape_group_dependency_count(nested));
    }
    count
}

fn count_styled_rows(table: &crate::Table<'_>) -> usize {
    table
        .rows()
        .iter()
        .map(|row| usize::from(row.table_style().is_some()))
        .sum()
}

fn style_closure(
    source: &Snapshot,
    kind: crate::style::Kind,
    id: u16,
) -> Result<Vec<&crate::style::Style<'_>>, Error> {
    let stylesheet = source.model().stylesheet();
    if stylesheet.get_typed(kind, id).is_none() {
        return Err(Error::DestinationOutOfRange("style definition"));
    }
    let mut pending = vec![(kind, id)];
    let mut seen = BTreeSet::new();
    let mut output = Vec::new();
    while let Some((style_kind, style_id)) = pending.pop() {
        if !seen.insert((style_kind as u8, style_id)) {
            continue;
        }
        let style = stylesheet
            .get_typed(style_kind, style_id)
            .ok_or(Error::UnsupportedSource("style dependency is missing"))?;
        if let Some(parent) = style.based_on {
            pending.push((style_kind, parent));
        }
        if let Some(next) = style.next_style {
            pending.push((style_kind, next));
        }
        if let Some(linked) = style.linked_style {
            let linked_kind = match style_kind {
                crate::style::Kind::Paragraph => crate::style::Kind::Character,
                crate::style::Kind::Character => crate::style::Kind::Paragraph,
                crate::style::Kind::Section => crate::style::Kind::Section,
                crate::style::Kind::Table => crate::style::Kind::Table,
            };
            pending.push((linked_kind, linked));
        }
        output.push(style);
    }
    output.reverse();
    Ok(output)
}

fn owned_style(style: &crate::style::Style<'_>) -> crate::style::Style<'static> {
    crate::style::Style {
        id: style.id,
        name: Cow::Owned(style.name.to_string()),
        style_type: style.style_type,
        based_on: style.based_on,
        next_style: style.next_style,
        linked_style: style.linked_style,
        formatting: style.formatting,
        paragraph: style.paragraph,
        table_conditional: style.table_conditional,
        builtin: style.builtin,
        hidden: style.hidden,
        additive: style.additive,
        auto_update: style.auto_update,
        locked: style.locked,
        semi_hidden: style.semi_hidden,
        unhide_when_used: style.unhide_when_used,
        quick_format: style.quick_format,
        priority: style.priority,
        revision_id: style.revision_id,
        personal: style.personal,
        compose: style.compose,
        reply: style.reply,
    }
}

fn owned_list(list: &crate::list::List<'_>) -> crate::list::List<'static> {
    crate::list::List {
        id: list.id,
        template_id: list.template_id,
        simple: list.simple,
        hybrid: list.hybrid,
        name: Cow::Owned(list.name.to_string()),
        style_name: Cow::Owned(list.style_name.to_string()),
        style_priority: list.style_priority,
        levels: list
            .levels
            .iter()
            .map(|level| crate::ListLevel {
                level: level.level,
                level_type: level.level_type,
                number_text: Cow::Owned(level.number_text.to_string()),
                number_positions: Cow::Owned(level.number_positions.to_string()),
                start_at: level.start_at,
                justification: level.justification,
                follow_previous: level.follow_previous,
                follow: level.follow,
                font_ref: level.font_ref,
                indent: level.indent,
                space: level.space,
                left_indent: level.left_indent,
                first_line_indent: level.first_line_indent,
                tabs: level.tabs.clone(),
                picture_index: level.picture_index,
                tentative: level.tentative,
                legal_format: level.legal_format,
                no_restart: level.no_restart,
                legacy: level.legacy,
                include_previous: level.include_previous,
                include_previous_space: level.include_previous_space,
                template_id: level.template_id,
            })
            .collect(),
    }
}

fn owned_object(object: &crate::EmbeddedObject<'_>) -> crate::EmbeddedObject<'static> {
    crate::EmbeddedObject {
        position: object.position,
        kind: object.kind,
        link_self: object.link_self,
        class_name: Cow::Owned(object.class_name.to_string()),
        name: Cow::Owned(object.name.to_string()),
        alias: object
            .alias
            .as_ref()
            .map(|value| Cow::Owned(value.to_string())),
        section: object
            .section
            .as_ref()
            .map(|value| Cow::Owned(value.to_string())),
        time: object.time,
        class_id: Cow::Owned(object.class_id.to_string()),
        width: object.width,
        height: object.height,
        alignment: object.alignment,
        translation_y: object.translation_y,
        crop_top: object.crop_top,
        crop_bottom: object.crop_bottom,
        crop_left: object.crop_left,
        crop_right: object.crop_right,
        scale_x: object.scale_x,
        scale_y: object.scale_y,
        locked: object.locked,
        update_requested: object.update_requested,
        set_size: object.set_size,
        merge_result: object.merge_result,
        result_kind: object.result_kind,
        result_text: Cow::Owned(object.result_text.to_string()),
        result_picture_indices: object.result_picture_indices.clone(),
        data: object.data.clone(),
    }
}
