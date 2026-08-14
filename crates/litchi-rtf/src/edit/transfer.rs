//! Checked transfer of ordinary RTF owners and their bounded dependencies.

use super::{Commit, Edit, Error, HeaderFooterParagraph, Snapshot};
use crate::{RtfWriter, TableCellPath};
use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet};
use std::io;

const MAX_TRANSFER_TABLE_NODES: usize = 65_536;

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
        let count = source.paragraph_count();
        let text = source
            .body()
            .paragraphs()
            .nth(source_position)
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
    /// Active/external field kinds, legacy shape-result fallbacks, and nested
    /// generic fields are refused. Self-contained instruction-backed drawings
    /// and explicit breaks are retained.
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
        validate_transferable_field_kind(source_field)?;
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
            require_equal_text_resources(source, target)?;
        }
        let mut field = owned_field(source_field, source.limits())?;
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
    /// Geometry, properties, text formatting, and self-contained nested
    /// drawings are retained when font/color resources match. Legacy
    /// `shprslt` fallbacks, active hyperlinks, nested generic fields, and
    /// unknown drawing syntax are refused because their recursive ownership
    /// is not transfer-safe.
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
        refuse_active_shape_links(std::slice::from_ref(shape), &[])?;
        if shape.result.is_some() {
            return Err(Error::UnsupportedSource(
                "shape transfer refuses legacy shape-result ownership",
            ));
        }
        let dependency_count = shape_dependency_count(shape)?;
        let mut transferred_shape = owned_shape(shape, source.limits())?;
        transferred_shape.position = target.text().len();
        let effect = format!("body:shape:append:{source_index}");
        let after = canonical_candidate(target, |model| model.push_shape(transferred_shape))?;
        root_plan(target, "shape.transfer", &effect, after, dependency_count)
    }

    /// Plans copying one complete nested table tree into a target cell end.
    ///
    /// Nested table text, geometry, borders, and recursively nested tables are
    /// retained. Self-contained instruction-backed drawings, explicit breaks,
    /// and passive generic fields are copied; legacy shape-result fallbacks,
    /// navigation/revision handles, and incompatible table-style dependencies
    /// are refused.
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
        // Table borders, shading, cell text, and nested drawings can all
        // carry numeric font/color references.  The table owner has no safe
        // remapping seam, so require the complete resource tables before
        // copying any part of the tree.
        require_equal_text_resources(source, target)?;
        let field_indices = table_story_field_indices(&nested.table)?;
        let field_count = field_indices.len();
        let mut field_drawing_count = 0usize;
        let mut transferred_fields = Vec::new();
        transferred_fields
            .try_reserve(field_count)
            .map_err(|_error| {
                Error::Write("could not reserve nested-table field transfers".to_string())
            })?;
        for index in field_indices {
            let field = source.fields().get(index).ok_or(Error::UnsupportedSource(
                "nested-table story references a missing field",
            ))?;
            field_drawing_count =
                field_drawing_count.saturating_add(validate_transferable_field(source, field)?);
            let mut transferred_field = owned_field(field, source.limits())?;
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
        let drawing_count = validate_table_transfer(source, &nested.table, 0)?;
        let styled_rows = count_styled_rows(&nested.table);
        if styled_rows != 0 && source.styles() != target.styles() {
            return Err(Error::UnsupportedSource(
                "nested-table transfer requires identical table-style dependencies",
            ));
        }
        let mut table = owned_transfer_table(&nested.table, source.limits(), 0)?;
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
        pictures
            .try_reserve(object.result_picture_indices.len())
            .map_err(|_error| {
                Error::Write("could not reserve embedded-object pictures".to_string())
            })?;
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
    let limit = target.limits().max_source_bytes();
    let mut output = BoundedOutput::new(limit);
    RtfWriter::new(&mut output)
        .write_document(&model)
        .map_err(|error| Error::Write(error.to_string()))?;
    let output = output.into_inner();
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

/// A fallible, source-size-bounded sink for canonical transfer output.
///
/// `RtfWriter` writes incrementally, so retaining the output in a plain
/// `Vec<u8>` would otherwise permit infallible growth until the writer's
/// final size check.  This sink enforces the document limit before every
/// append and maps allocation failure back into the writer's `io::Result`.
struct BoundedOutput {
    bytes: Vec<u8>,
    limit: usize,
}

impl BoundedOutput {
    fn new(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            limit,
        }
    }

    fn into_inner(self) -> Vec<u8> {
        self.bytes
    }
}

impl io::Write for BoundedOutput {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        let remaining = self
            .limit
            .checked_sub(self.bytes.len())
            .ok_or_else(|| io::Error::new(io::ErrorKind::WriteZero, "output limit exceeded"))?;
        if input.len() > remaining {
            return Err(io::Error::new(
                io::ErrorKind::WriteZero,
                "output limit exceeded",
            ));
        }
        self.bytes
            .try_reserve(input.len())
            .map_err(|_error| io::Error::other("output allocation failed"))?;
        self.bytes.extend_from_slice(input);
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FieldTransferClass {
    /// The typed field model explicitly guarantees inert metadata semantics.
    Passive,
    /// The instruction names an external document, data source, or update.
    External,
    /// The instruction can invoke active content or application behavior.
    Active,
    /// The instruction is not covered by a transfer policy.
    Unknown,
}

fn field_transfer_class(field: &crate::Field<'_>) -> FieldTransferClass {
    use crate::FieldType;
    let parsed_code = field.parsed_code();
    match field.field_type {
        // HYPERLINK is deliberately allowed only because HyperlinkField's
        // existing contract is inert: this crate retains the target and
        // cached result but never resolves, opens, fetches, or activates it.
        FieldType::Hyperlink => {
            if matches!(
                &parsed_code,
                crate::ParsedFieldCode::Hyperlink(code)
                    if code.external_target.is_none() && code.bookmark.is_some()
            ) {
                FieldTransferClass::Passive
            } else {
                FieldTransferClass::Unknown
            }
        },
        FieldType::Reference if matches!(&parsed_code, crate::ParsedFieldCode::Reference(_)) => {
            FieldTransferClass::Passive
        },
        FieldType::PageReference
            if matches!(&parsed_code, crate::ParsedFieldCode::PageReference(_)) =>
        {
            FieldTransferClass::Passive
        },
        FieldType::NoteReference
            if matches!(&parsed_code, crate::ParsedFieldCode::NoteReference(_)) =>
        {
            FieldTransferClass::Passive
        },
        FieldType::ReferencedDocument
        | FieldType::Dde
        | FieldType::DdeAuto
        | FieldType::Link
        | FieldType::Include
        | FieldType::Import
        | FieldType::IncludeText
        | FieldType::IncludePicture
        | FieldType::Database
        | FieldType::MergeField
        | FieldType::MailMergeData
        | FieldType::MergeRecord
        | FieldType::MergeSequence
        | FieldType::MailMergeNext
        | FieldType::MailMergeNextIf
        | FieldType::MailMergeSkipIf
        | FieldType::AddressBlock
        | FieldType::GreetingLine
        | FieldType::MergeBarcode => FieldTransferClass::External,
        FieldType::MacroButton
        | FieldType::GoToButton
        | FieldType::Print
        | FieldType::Embed
        | FieldType::AddIn
        | FieldType::Control
        | FieldType::HtmlControl
        | FieldType::Ask
        | FieldType::FillIn
        | FieldType::Shape
        | FieldType::FormText
        | FieldType::FormCheckbox
        | FieldType::FormDropdown => FieldTransferClass::Active,
        FieldType::Page
        | FieldType::Date
        | FieldType::Toc
        | FieldType::TocEntry
        | FieldType::TableOfAuthorities
        | FieldType::TableOfAuthoritiesEntry
        | FieldType::Bookmark
        | FieldType::Equation
        | FieldType::Barcode
        | FieldType::DisplayBarcode
        | FieldType::BidiOutline
        | FieldType::Private => {
            if matches!(&parsed_code, crate::ParsedFieldCode::Malformed(_)) {
                FieldTransferClass::Unknown
            } else {
                FieldTransferClass::Passive
            }
        },
        _ => FieldTransferClass::Unknown,
    }
}

fn validate_transferable_field_kind(field: &crate::Field<'_>) -> Result<(), Error> {
    match field_transfer_class(field) {
        FieldTransferClass::Passive => Ok(()),
        FieldTransferClass::External => Err(Error::UnsupportedSource(
            "field transfer refuses external, referenced-document, or mail-merge instructions",
        )),
        FieldTransferClass::Active => Err(Error::UnsupportedSource(
            "field transfer refuses active field instructions",
        )),
        FieldTransferClass::Unknown => Err(Error::UnsupportedSource(
            "field transfer requires a validated passive field policy",
        )),
    }
}

fn owned_field(
    field: &crate::Field<'_>,
    limits: crate::ParseLimits,
) -> Result<crate::Field<'static>, Error> {
    let instruction = super::clone_bounded_text(
        field.instruction.as_ref(),
        limits,
        "could not reserve transferred field instruction",
    )?;
    let result = super::clone_bounded_text(
        field.result.as_ref(),
        limits,
        "could not reserve transferred field result",
    )?;
    refuse_active_shape_links(&field.shapes, &field.shape_groups)?;
    let mut shapes = Vec::new();
    shapes
        .try_reserve(field.shapes.len())
        .map_err(|_error| Error::Write("could not reserve transferred field shapes".to_string()))?;
    for shape in &field.shapes {
        shapes.push(owned_shape(shape, limits)?);
    }
    let mut shape_groups = Vec::new();
    shape_groups
        .try_reserve(field.shape_groups.len())
        .map_err(|_error| {
            Error::Write("could not reserve transferred field shape groups".to_string())
        })?;
    for group in &field.shape_groups {
        shape_groups.push(owned_shape_group(group, limits)?);
    }
    let mut drawing_order = Vec::new();
    drawing_order
        .try_reserve(field.drawing_order.len())
        .map_err(|_error| {
            Error::Write("could not reserve transferred field drawing order".to_string())
        })?;
    drawing_order.extend_from_slice(&field.drawing_order);
    let mut result_events = Vec::new();
    result_events
        .try_reserve(field.result_events.len())
        .map_err(|_error| {
            Error::Write("could not reserve transferred field story events".to_string())
        })?;
    result_events.extend_from_slice(&field.result_events);
    Ok(crate::Field {
        field_type: field.field_type,
        instruction: Cow::Owned(instruction),
        result: Cow::Owned(result),
        status: field.status,
        shapes,
        shape_groups,
        drawing_order,
        result_events,
        owner: field.owner,
        position: field.position,
        range_end: field.range_end,
    })
}

const MAX_TRANSFER_SHAPE_NESTING_DEPTH: usize = 64;

fn owned_shape(
    shape: &crate::Shape<'_>,
    limits: crate::ParseLimits,
) -> Result<crate::Shape<'static>, Error> {
    owned_shape_at_depth(shape, limits, 0)
}

fn owned_shape_at_depth(
    shape: &crate::Shape<'_>,
    limits: crate::ParseLimits,
    depth: usize,
) -> Result<crate::Shape<'static>, Error> {
    if depth >= MAX_TRANSFER_SHAPE_NESTING_DEPTH {
        return Err(Error::UnsupportedSource(
            "shape transfer refuses excessive nested drawing depth",
        ));
    }
    if shape.result.is_some() {
        return Err(Error::UnsupportedSource(
            "shape transfer refuses legacy shape-result ownership",
        ));
    }
    let text = super::clone_bounded_text(
        shape.text.as_ref(),
        limits,
        "could not reserve transferred shape text",
    )?;
    let name = super::clone_bounded_text(
        shape.name.as_ref(),
        limits,
        "could not reserve transferred shape name",
    )?;

    let mut text_shapes = Vec::new();
    text_shapes
        .try_reserve(shape.text_shapes.len())
        .map_err(|_error| {
            Error::Write("could not reserve transferred shape text drawings".to_string())
        })?;
    for nested in &shape.text_shapes {
        text_shapes.push(owned_shape_at_depth(nested, limits, depth + 1)?);
    }

    let mut text_shape_groups = Vec::new();
    text_shape_groups
        .try_reserve(shape.text_shape_groups.len())
        .map_err(|_error| {
            Error::Write("could not reserve transferred shape text groups".to_string())
        })?;
    for group in &shape.text_shape_groups {
        text_shape_groups.push(owned_shape_group_at_depth(group, limits, depth + 1)?);
    }

    let mut text_drawing_order = Vec::new();
    text_drawing_order
        .try_reserve(shape.text_drawing_order.len())
        .map_err(|_error| {
            Error::Write("could not reserve transferred shape drawing order".to_string())
        })?;
    text_drawing_order.extend_from_slice(&shape.text_drawing_order);

    let mut text_story_events = Vec::new();
    text_story_events
        .try_reserve(shape.text_story_events.len())
        .map_err(|_error| {
            Error::Write("could not reserve transferred shape story events".to_string())
        })?;
    text_story_events.extend_from_slice(&shape.text_story_events);

    let mut properties = Vec::new();
    properties
        .try_reserve(shape.properties.len())
        .map_err(|_error| {
            Error::Write("could not reserve transferred shape properties".to_string())
        })?;
    for property in &shape.properties {
        properties.push(owned_shape_property(property, limits)?);
    }

    let mut info = Vec::new();
    info.try_reserve(shape.info.len())
        .map_err(|_error| Error::Write("could not reserve transferred shape info".to_string()))?;
    info.extend_from_slice(&shape.info);

    Ok(crate::Shape {
        position: shape.position,
        instruction_present: shape.instruction_present,
        shape_type: shape.shape_type,
        geometry: shape.geometry,
        fill: shape.fill,
        border: shape.border,
        line: shape.line,
        text: Cow::Owned(text),
        text_destination_present: shape.text_destination_present,
        text_formatting: shape.text_formatting,
        text_shapes,
        text_shape_groups,
        text_drawing_order,
        text_story_events,
        wrap_mode: shape.wrap_mode,
        behind_doc: shape.behind_doc,
        is_background: shape.is_background,
        locked: shape.locked,
        name: Cow::Owned(name),
        properties,
        result: None,
        info,
    })
}

fn owned_shape_group(
    group: &crate::ShapeGroup<'_>,
    limits: crate::ParseLimits,
) -> Result<crate::ShapeGroup<'static>, Error> {
    owned_shape_group_at_depth(group, limits, 0)
}

fn owned_shape_group_at_depth(
    group: &crate::ShapeGroup<'_>,
    limits: crate::ParseLimits,
    depth: usize,
) -> Result<crate::ShapeGroup<'static>, Error> {
    if depth >= MAX_TRANSFER_SHAPE_NESTING_DEPTH {
        return Err(Error::UnsupportedSource(
            "shape transfer refuses excessive nested drawing depth",
        ));
    }
    if group.result.is_some() {
        return Err(Error::UnsupportedSource(
            "shape transfer refuses legacy shape-result ownership",
        ));
    }
    let name = super::clone_bounded_text(
        group.name.as_ref(),
        limits,
        "could not reserve transferred shape-group name",
    )?;

    let mut shapes = Vec::new();
    shapes.try_reserve(group.shapes.len()).map_err(|_error| {
        Error::Write("could not reserve transferred shape-group shapes".to_string())
    })?;
    for shape in &group.shapes {
        shapes.push(owned_shape_at_depth(shape, limits, depth + 1)?);
    }

    let mut groups = Vec::new();
    groups
        .try_reserve(group.groups.len())
        .map_err(|_error| Error::Write("could not reserve transferred shape groups".to_string()))?;
    for nested in &group.groups {
        groups.push(owned_shape_group_at_depth(nested, limits, depth + 1)?);
    }

    let mut child_order = Vec::new();
    child_order
        .try_reserve(group.child_order.len())
        .map_err(|_error| {
            Error::Write("could not reserve transferred shape-group child order".to_string())
        })?;
    child_order.extend_from_slice(&group.child_order);

    let mut info = Vec::new();
    info.try_reserve(group.info.len()).map_err(|_error| {
        Error::Write("could not reserve transferred shape-group info".to_string())
    })?;
    info.extend_from_slice(&group.info);

    let mut properties = Vec::new();
    properties
        .try_reserve(group.properties.len())
        .map_err(|_error| {
            Error::Write("could not reserve transferred shape-group properties".to_string())
        })?;
    for property in &group.properties {
        properties.push(owned_shape_property(property, limits)?);
    }

    Ok(crate::ShapeGroup {
        position: group.position,
        name: Cow::Owned(name),
        shapes,
        groups,
        child_order,
        info,
        geometry: group.geometry,
        properties,
        result: None,
    })
}

fn owned_shape_property(
    property: &crate::ShapeProperty<'_>,
    limits: crate::ParseLimits,
) -> Result<crate::ShapeProperty<'static>, Error> {
    let name = super::clone_bounded_text(
        property.name.as_ref(),
        limits,
        "could not reserve transferred shape-property name",
    )?;
    let value = super::clone_bounded_text(
        property.value.as_ref(),
        limits,
        "could not reserve transferred shape-property value",
    )?;
    let binary_value = property
        .binary_value
        .as_ref()
        .map(|bytes| {
            super::clone_bounded_bytes(
                bytes.as_ref(),
                limits,
                "could not reserve transferred shape-property binary value",
            )
            .map(Cow::Owned)
        })
        .transpose()?;
    let hyperlink = property
        .hyperlink
        .as_ref()
        .map(|link| owned_shape_hyperlink(link, limits))
        .transpose()?;
    Ok(crate::ShapeProperty {
        name: Cow::Owned(name),
        value: Cow::Owned(value),
        binary_value,
        theme_value: property.theme_value,
        hyperlink,
    })
}

fn owned_shape_hyperlink(
    hyperlink: &crate::ShapeHyperlink<'_>,
    limits: crate::ParseLimits,
) -> Result<crate::ShapeHyperlink<'static>, Error> {
    Ok(crate::ShapeHyperlink {
        location: owned_optional_shape_text(
            hyperlink.location.as_ref(),
            limits,
            "could not reserve transferred shape-hyperlink location",
        )?,
        source: owned_optional_shape_text(
            hyperlink.source.as_ref(),
            limits,
            "could not reserve transferred shape-hyperlink source",
        )?,
        friendly_name: owned_optional_shape_text(
            hyperlink.friendly_name.as_ref(),
            limits,
            "could not reserve transferred shape-hyperlink friendly name",
        )?,
    })
}

fn owned_optional_shape_text(
    value: Option<&Cow<'_, str>>,
    limits: crate::ParseLimits,
    context: &'static str,
) -> Result<Option<Cow<'static, str>>, Error> {
    value
        .map(|value| super::clone_bounded_text(value.as_ref(), limits, context))
        .transpose()
        .map(|value| value.map(Cow::Owned))
}

fn owned_transfer_table(
    table: &crate::Table<'_>,
    limits: crate::ParseLimits,
    depth: usize,
) -> Result<crate::Table<'static>, Error> {
    if depth >= MAX_TRANSFER_SHAPE_NESTING_DEPTH {
        return Err(Error::UnsupportedSource(
            "nested-table transfer refuses excessive table depth",
        ));
    }
    let mut output = crate::Table::new();
    output.set_direction(table.direction());
    for row in table.rows() {
        let mut owned_row = crate::Row::new();
        owned_row.set_table_style(row.table_style());
        owned_row.set_table_rsid(row.table_rsid());
        owned_row.set_direction(row.direction());
        owned_row.set_layout(*row.layout());
        owned_row.set_borders(row.borders().clone());
        owned_row.set_shading(row.shading());
        owned_row.set_geometry(row.geometry());
        owned_row.set_autoformat_flags(row.autoformat_flags());
        owned_row.set_banding(row.banding());
        owned_row.set_revision(row.revision());
        for cell in row.cells() {
            let text = super::clone_bounded_text(
                cell.text(),
                limits,
                "could not reserve nested-table cell text",
            )?;
            let mut owned_cell = crate::Cell::with_distances(
                Cow::Owned(text),
                cell.padding().clone(),
                cell.spacing().clone(),
            );
            owned_cell.set_layout(*cell.layout());
            owned_cell.set_merge(cell.merge());
            owned_cell.set_right_boundary(cell.right_boundary());
            owned_cell.set_preferred_width(cell.preferred_width());
            owned_cell.set_revision(cell.revision());
            owned_cell.set_borders(cell.borders().clone());
            owned_cell.set_shading(cell.shading());
            for nested in cell.nested_tables() {
                owned_cell.add_nested_table(
                    nested.text_offset,
                    owned_transfer_table(&nested.table, limits, depth + 1)?,
                )?;
            }

            let mut shapes = Vec::new();
            shapes.try_reserve(cell.shapes().len()).map_err(|_error| {
                Error::Write("could not reserve nested-table shapes".to_string())
            })?;
            for shape in cell.shapes() {
                shapes.push(owned_shape(shape, limits)?);
            }
            let mut shape_groups = Vec::new();
            shape_groups
                .try_reserve(cell.shape_groups().len())
                .map_err(|_error| {
                    Error::Write("could not reserve nested-table shape groups".to_string())
                })?;
            for group in cell.shape_groups() {
                shape_groups.push(owned_shape_group(group, limits)?);
            }
            let mut drawing_order = Vec::new();
            drawing_order
                .try_reserve(cell.drawing_order().len())
                .map_err(|_error| {
                    Error::Write("could not reserve nested-table drawing order".to_string())
                })?;
            drawing_order.extend_from_slice(cell.drawing_order());
            let mut story_events = Vec::new();
            story_events
                .try_reserve(cell.story_events().len())
                .map_err(|_error| {
                    Error::Write("could not reserve nested-table story events".to_string())
                })?;
            story_events.extend_from_slice(cell.story_events());
            owned_cell.set_story_content(shapes, shape_groups, drawing_order, story_events)?;
            owned_row.try_add_cell(owned_cell)?;
        }
        owned_row.set_padding(row.padding().clone());
        owned_row.set_spacing(row.spacing().clone());
        owned_row.set_cell_defaults(row.cell_defaults().clone());
        owned_row.set_positioning(row.positioning().clone());
        output.try_add_row(owned_row)?;
    }
    Ok(output)
}

fn validate_table_transfer(
    source: &Snapshot,
    table: &crate::Table<'_>,
    depth: usize,
) -> Result<usize, Error> {
    let mut node_count = 1usize;
    validate_table_transfer_at_depth(source, table, depth, &mut node_count)
}

fn validate_table_transfer_at_depth(
    source: &Snapshot,
    table: &crate::Table<'_>,
    depth: usize,
    node_count: &mut usize,
) -> Result<usize, Error> {
    if depth >= MAX_TRANSFER_SHAPE_NESTING_DEPTH {
        return Err(Error::UnsupportedSource(
            "nested-table transfer refuses excessive table depth",
        ));
    }
    let mut drawing_count = 0usize;
    for row in table.rows() {
        *node_count = node_count.checked_add(1).ok_or(Error::UnsupportedSource(
            "nested-table transfer node count overflows",
        ))?;
        if *node_count > MAX_TRANSFER_TABLE_NODES {
            return Err(Error::UnsupportedSource(
                "nested-table transfer exceeds bounded table ownership",
            ));
        }
        for cell in row.cells() {
            *node_count = node_count.checked_add(1).ok_or(Error::UnsupportedSource(
                "nested-table transfer node count overflows",
            ))?;
            if *node_count > MAX_TRANSFER_TABLE_NODES {
                return Err(Error::UnsupportedSource(
                    "nested-table transfer exceeds bounded table ownership",
                ));
            }
            refuse_drawing_story_fields(cell.shapes(), cell.shape_groups())?;
            refuse_active_shape_links(cell.shapes(), cell.shape_groups())?;
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
                *node_count = node_count.checked_add(1).ok_or(Error::UnsupportedSource(
                    "nested-table transfer node count overflows",
                ))?;
                if *node_count > MAX_TRANSFER_TABLE_NODES {
                    return Err(Error::UnsupportedSource(
                        "nested-table transfer exceeds bounded table ownership",
                    ));
                }
                drawing_count = drawing_count.saturating_add(validate_table_transfer_at_depth(
                    source,
                    &nested.table,
                    depth + 1,
                    node_count,
                )?);
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
            let (shapes, shape_groups, drawing_order, mut events) = cell.take_story_content();
            for event in &mut events {
                if let crate::CellStoryEvent::Field(field) = event {
                    field.field_index = *indices.get(&field.field_index).ok_or_else(|| {
                        crate::RtfError::MalformedDocument(
                            "nested-table field remapping is incomplete".to_string(),
                        )
                    })?;
                }
            }
            cell.set_story_content(shapes, shape_groups, drawing_order, events)?;
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
    validate_transferable_field_kind(field)?;
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
    for shape in shapes {
        refuse_shape_story_fields(shape, 0)?;
    }
    for group in groups {
        refuse_shape_group_story_fields(group, 0)?;
    }
    Ok(())
}

fn refuse_active_shape_links(
    shapes: &[crate::Shape<'_>],
    groups: &[crate::ShapeGroup<'_>],
) -> Result<(), Error> {
    for shape in shapes {
        refuse_shape_active_link(shape, 0)?;
    }
    for group in groups {
        refuse_shape_group_active_link(group, 0)?;
    }
    Ok(())
}

fn refuse_shape_story_fields(shape: &crate::Shape<'_>, depth: usize) -> Result<(), Error> {
    if depth >= MAX_TRANSFER_SHAPE_NESTING_DEPTH {
        return Err(Error::UnsupportedSource(
            "drawing transfer refuses excessive nested drawing depth",
        ));
    }
    if shape
        .text_story_events
        .iter()
        .any(|event| matches!(event, crate::StoryEvent::Field(_)))
    {
        return Err(Error::UnsupportedSource(
            "drawing transfer refuses unresolved shape-text field handles",
        ));
    }
    for nested in &shape.text_shapes {
        refuse_shape_story_fields(nested, depth + 1)?;
    }
    for group in &shape.text_shape_groups {
        refuse_shape_group_story_fields(group, depth + 1)?;
    }
    Ok(())
}

fn refuse_shape_group_story_fields(
    group: &crate::ShapeGroup<'_>,
    depth: usize,
) -> Result<(), Error> {
    if depth >= MAX_TRANSFER_SHAPE_NESTING_DEPTH {
        return Err(Error::UnsupportedSource(
            "drawing transfer refuses excessive nested drawing depth",
        ));
    }
    for shape in &group.shapes {
        refuse_shape_story_fields(shape, depth + 1)?;
    }
    for nested in &group.groups {
        refuse_shape_group_story_fields(nested, depth + 1)?;
    }
    Ok(())
}

fn refuse_shape_active_link(shape: &crate::Shape<'_>, depth: usize) -> Result<(), Error> {
    if depth >= MAX_TRANSFER_SHAPE_NESTING_DEPTH {
        return Err(Error::UnsupportedSource(
            "shape transfer refuses excessive nested drawing depth",
        ));
    }
    if shape
        .properties
        .iter()
        .any(|property| property.hyperlink.is_some())
    {
        return Err(Error::UnsupportedSource(
            "shape transfer refuses active hyperlink metadata",
        ));
    }
    for nested in &shape.text_shapes {
        refuse_shape_active_link(nested, depth + 1)?;
    }
    for group in &shape.text_shape_groups {
        refuse_shape_group_active_link(group, depth + 1)?;
    }
    Ok(())
}

fn refuse_shape_group_active_link(
    group: &crate::ShapeGroup<'_>,
    depth: usize,
) -> Result<(), Error> {
    if depth >= MAX_TRANSFER_SHAPE_NESTING_DEPTH {
        return Err(Error::UnsupportedSource(
            "shape transfer refuses excessive nested drawing depth",
        ));
    }
    if group
        .properties
        .iter()
        .any(|property| property.hyperlink.is_some())
    {
        return Err(Error::UnsupportedSource(
            "shape transfer refuses active hyperlink metadata",
        ));
    }
    for shape in &group.shapes {
        refuse_shape_active_link(shape, depth + 1)?;
    }
    for nested in &group.groups {
        refuse_shape_group_active_link(nested, depth + 1)?;
    }
    Ok(())
}

fn shape_dependency_count(shape: &crate::Shape<'_>) -> Result<usize, Error> {
    shape_dependency_count_at_depth(shape, 0)
}

fn shape_dependency_count_at_depth(shape: &crate::Shape<'_>, depth: usize) -> Result<usize, Error> {
    if depth >= MAX_TRANSFER_SHAPE_NESTING_DEPTH {
        return Err(Error::UnsupportedSource(
            "shape transfer refuses excessive nested drawing depth",
        ));
    }
    let mut count = shape
        .text_shapes
        .len()
        .checked_add(shape.text_shape_groups.len())
        .ok_or(Error::UnsupportedSource(
            "shape transfer dependency count overflows",
        ))?;
    for nested in &shape.text_shapes {
        count = count
            .checked_add(shape_dependency_count_at_depth(nested, depth + 1)?)
            .ok_or(Error::UnsupportedSource(
                "shape transfer dependency count overflows",
            ))?;
    }
    for group in &shape.text_shape_groups {
        count = count
            .checked_add(shape_group_dependency_count_at_depth(group, depth + 1)?)
            .ok_or(Error::UnsupportedSource(
                "shape transfer dependency count overflows",
            ))?;
    }
    Ok(count)
}

fn shape_group_dependency_count_at_depth(
    group: &crate::ShapeGroup<'_>,
    depth: usize,
) -> Result<usize, Error> {
    if depth >= MAX_TRANSFER_SHAPE_NESTING_DEPTH {
        return Err(Error::UnsupportedSource(
            "shape transfer refuses excessive nested drawing depth",
        ));
    }
    let mut count =
        group
            .shapes
            .len()
            .checked_add(group.groups.len())
            .ok_or(Error::UnsupportedSource(
                "shape transfer dependency count overflows",
            ))?;
    for shape in &group.shapes {
        count = count
            .checked_add(shape_dependency_count_at_depth(shape, depth + 1)?)
            .ok_or(Error::UnsupportedSource(
                "shape transfer dependency count overflows",
            ))?;
    }
    for nested in &group.groups {
        count = count
            .checked_add(shape_group_dependency_count_at_depth(nested, depth + 1)?)
            .ok_or(Error::UnsupportedSource(
                "shape transfer dependency count overflows",
            ))?;
    }
    Ok(count)
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
