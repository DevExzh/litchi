//! Staged structural document verbs of [`Edit`] (record 0358).
//!
//! Each verb records one operation that the commit applies, in staging
//! order, to its working document; consecutive verbs share that document.

use super::{Edit, FamilyResult, OperationResult};
use litchi_core::Position;

impl Edit {
    /// Add a heading to the end of the document.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn add_heading(&mut self, text: impl Into<String>, level: u8) -> FamilyResult<&mut Self> {
        let text: String = text.into();
        self.document_operation(move |document| {
            document
                .add_heading(&text, level)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Append a paragraph containing one simple ODF hyperlink.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn add_hyperlink(
        &mut self,
        href: impl Into<String>,
        text: impl Into<String>,
    ) -> FamilyResult<&mut Self> {
        let href: String = href.into();
        let text: String = text.into();
        self.document_operation(move |document| {
            document
                .add_hyperlink(&href, &text)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Append a paragraph containing a fully configured ODF hyperlink.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn add_hyperlink_element(
        &mut self,
        hyperlink: crate::Hyperlink,
    ) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .add_hyperlink_element(hyperlink.clone())
                .map(|_| OperationResult::Unit)
        })
    }

    /// Add an existing list to the end of the document.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn add_list(&mut self, list: crate::elements::text::List) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .add_list(list.clone())
                .map(|_| OperationResult::Unit)
        })
    }

    /// Add an empty master page and its referenced page layout.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn add_master_page(
        &mut self,
        name: impl Into<String>,
        page_layout_name: impl Into<String>,
    ) -> FamilyResult<&mut Self> {
        let name: String = name.into();
        let page_layout_name: String = page_layout_name.into();
        self.document_operation_at(format!("/styles/master-page/{name}"), move |document| {
            document
                .add_master_page(&name, &page_layout_name)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Add a new plain paragraph to the end of the document.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn add_paragraph(&mut self, text: impl Into<String>) -> FamilyResult<&mut Self> {
        let text: String = text.into();
        self.document_operation(move |document| {
            document.add_paragraph(&text).map(|_| OperationResult::Unit)
        })
    }

    /// Append a complete typed section without rewriting existing body content.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn add_section(&mut self, section: &crate::Section) -> FamilyResult<&mut Self> {
        let section = section.clone();
        self.document_operation(move |document| {
            document
                .add_section(&section)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Add an existing table to the end of the document.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn add_table(&mut self, table: crate::elements::table::Table) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .add_table(table.clone())
                .map(|_| OperationResult::Unit)
        })
    }

    /// Atomically append a declaration in insertion order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn add_tracked_change(&mut self, change: crate::TrackChange) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .add_tracked_change(change.clone())
                .map(|_| OperationResult::Unit)
        })
    }

    /// Append a line break at a checked semantic paragraph position.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn append_line_break_at(&mut self, paragraph: Position) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .append_line_break_at(paragraph)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove the document-wide bibliography formatting policy.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn clear_bibliography_configuration(&mut self) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .clear_bibliography_configuration()
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove all projected top-level body content.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn clear_content(&mut self) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document.clear_content();
            Ok(OperationResult::Unit)
        })
    }

    /// Remove content-part font-face declarations and return the old value.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn clear_content_font_face_declarations(&mut self) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .clear_content_font_face_declarations()
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove one header/footer region from an existing master page.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn clear_header_footer(
        &mut self,
        master_page_name: impl Into<String>,
        kind: crate::header_footer::Kind,
    ) -> FamilyResult<&mut Self> {
        let master_page_name: String = master_page_name.into();
        self.document_operation_at(
            format!("/styles/master-page/{master_page_name}"),
            move |document| {
                document
                    .clear_header_footer(&master_page_name, kind)
                    .map(|_| OperationResult::Unit)
            },
        )
    }

    /// Remove document line-numbering configuration and return its old value.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn clear_line_numbering_configuration(&mut self) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .clear_line_numbering_configuration()
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove one stored note-class configuration and return its prior value.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn clear_notes_configuration(
        &mut self,
        note_class: crate::notes_configuration::Class,
    ) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .clear_notes_configuration(note_class)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove all top-level paragraphs while retaining other body elements.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn clear_paragraphs(&mut self) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document.clear_paragraphs();
            Ok(OperationResult::Unit)
        })
    }

    /// Remove every section wrapper/source while retaining mixed nested content.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn clear_sections(&mut self) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document.clear_sections().map(|_| OperationResult::Unit)
        })
    }

    /// Remove styles-part font-face declarations and return the old value.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn clear_styles_font_face_declarations(&mut self) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .clear_styles_font_face_declarations()
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove all top-level tables while retaining other body elements.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn clear_tables(&mut self) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document.clear_tables();
            Ok(OperationResult::Unit)
        })
    }

    /// Remove all declarations, policy, and correlated markers.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn clear_tracked_changes(&mut self) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .clear_tracked_changes()
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert a point bookmark at a paragraph end, or wrap the paragraph with a range bookmark.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_bookmark_target(
        &mut self,
        paragraph_index: usize,
        target: &crate::BookmarkTarget,
    ) -> FamilyResult<&mut Self> {
        let target = target.clone();
        self.document_operation(move |document| {
            document
                .insert_bookmark_target(paragraph_index, &target)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert a property into a form/control owner selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_form_property(
        &mut self,
        owner_index: usize,
        property: &crate::form::Property,
    ) -> FamilyResult<&mut Self> {
        let property = property.clone();
        self.document_operation(move |document| {
            document
                .insert_form_property(owner_index, &property)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert a fixed-text, hidden, or generic control into a form selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_generic_form_control(
        &mut self,
        form_index: usize,
        control: &crate::form::GenericFormControl,
    ) -> FamilyResult<&mut Self> {
        let control = control.clone();
        self.document_operation(move |document| {
            document
                .insert_generic_form_control(form_index, &control)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Stages `insert_grid_control`.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_grid_control(
        &mut self,
        form_index: usize,
        control: &crate::form::GridControl,
    ) -> FamilyResult<&mut Self> {
        let control = control.clone();
        self.document_operation(move |document| {
            document
                .insert_grid_control(form_index, &control)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert an image frame as a new paragraph at a specific index.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    /// [`Commit::results`](super::Commit::results) reports the generated name as
    /// [`OperationResult::Name`].
    pub fn insert_image(
        &mut self,
        index: usize,
        image: impl Into<Vec<u8>>,
        width: &crate::frame::Length,
        height: &crate::frame::Length,
        anchor: crate::frame::Anchor,
    ) -> FamilyResult<&mut Self> {
        let image: Vec<u8> = image.into();
        let width = width.clone();
        let height = height.clone();
        self.document_operation(move |document| {
            document
                .insert_image(index, &image, &width, &height, anchor)
                .map(OperationResult::Name)
        })
    }

    /// Insert an image-frame control into a form selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_image_frame_control(
        &mut self,
        form_index: usize,
        control: &crate::form::ImageFrameControl,
    ) -> FamilyResult<&mut Self> {
        let control = control.clone();
        self.document_operation(move |document| {
            document
                .insert_image_frame_control(form_index, &control)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert a button or checkbox into a form selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_interactive_control(
        &mut self,
        form_index: usize,
        control: &crate::form::InteractiveControl,
    ) -> FamilyResult<&mut Self> {
        let control = control.clone();
        self.document_operation(move |document| {
            document
                .insert_interactive_control(form_index, &control)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert a complete typed master page without rewriting unrelated styles.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_master_page(
        &mut self,
        page: &crate::header_footer::Master,
    ) -> FamilyResult<&mut Self> {
        let page = page.clone();
        self.document_operation_at(
            format!("/styles/master-page/{}", page.name),
            move |document| {
                document
                    .insert_master_page(&page)
                    .map(|_| OperationResult::Unit)
            },
        )
    }

    /// Insert a plain paragraph at a checked semantic paragraph position.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_paragraph_at(
        &mut self,
        position: Position,
        text: impl Into<String>,
    ) -> FamilyResult<&mut Self> {
        let text: String = text.into();
        self.document_operation(move |document| {
            document
                .insert_paragraph_at(position, &text)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert a password or file control into a form selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_password_file_control(
        &mut self,
        form_index: usize,
        control: &crate::form::PasswordFileControl,
    ) -> FamilyResult<&mut Self> {
        let control = control.clone();
        self.document_operation(move |document| {
            document
                .insert_password_file_control(form_index, &control)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert a point reference at a paragraph end, or wrap the paragraph with a range reference.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_reference_mark(
        &mut self,
        paragraph_index: usize,
        mark: &crate::ReferenceMark,
    ) -> FamilyResult<&mut Self> {
        let mark = mark.clone();
        self.document_operation(move |document| {
            document
                .insert_reference_mark(paragraph_index, &mark)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert a listbox or combobox into a form selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_selection_control(
        &mut self,
        form_index: usize,
        control: &crate::form::SelectionControl,
    ) -> FamilyResult<&mut Self> {
        let control = control.clone();
        self.document_operation(move |document| {
            document
                .insert_selection_control(form_index, &control)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert a plain-text text-box frame as a new paragraph at a specific index.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    /// [`Commit::results`](super::Commit::results) reports the generated name as
    /// [`OperationResult::Name`].
    pub fn insert_text_box(
        &mut self,
        index: usize,
        text: impl Into<String>,
        width: &crate::frame::Length,
        height: &crate::frame::Length,
        anchor: crate::frame::Anchor,
    ) -> FamilyResult<&mut Self> {
        let text: String = text.into();
        let width = width.clone();
        let height = height.clone();
        self.document_operation(move |document| {
            document
                .insert_text_box(index, &text, &width, &height, anchor)
                .map(OperationResult::Name)
        })
    }

    /// Insert a text or textarea control into a form selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_text_control(
        &mut self,
        form_index: usize,
        control: &crate::form::TextControl,
    ) -> FamilyResult<&mut Self> {
        let control = control.clone();
        self.document_operation(move |document| {
            document
                .insert_text_control(form_index, &control)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Append caller-authored index markup to `office:text` without refreshing its cache.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_text_index(&mut self, index: &crate::TextIndex) -> FamilyResult<&mut Self> {
        let index = index.clone();
        self.document_operation(move |document| {
            document
                .insert_text_index(&index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert a point mark at a paragraph end, or wrap the paragraph with a range mark.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_text_index_mark(
        &mut self,
        paragraph_index: usize,
        mark: &crate::TextIndexMark,
    ) -> FamilyResult<&mut Self> {
        let mark = mark.clone();
        self.document_operation(move |document| {
            document
                .insert_text_index_mark(paragraph_index, &mark)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert an inline `text:meta` into a paragraph.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_text_meta(
        &mut self,
        paragraph: Position,
        value: &crate::TextMeta,
    ) -> FamilyResult<&mut Self> {
        let value = value.clone();
        self.document_operation(move |document| {
            document
                .insert_text_meta(paragraph, &value)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert a typed value control into a form selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_typed_value_control(
        &mut self,
        form_index: usize,
        control: &crate::form::TypedValueControl,
    ) -> FamilyResult<&mut Self> {
        let control = control.clone();
        self.document_operation(move |document| {
            document
                .insert_typed_value_control(form_index, &control)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert a value-range control into a form selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_value_range_control(
        &mut self,
        form_index: usize,
        control: &crate::form::ValueRangeControl,
    ) -> FamilyResult<&mut Self> {
        let control = control.clone();
        self.document_operation(move |document| {
            document
                .insert_value_range_control(form_index, &control)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert a radio, frame, or image-button into a form selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_visual_control(
        &mut self,
        form_index: usize,
        control: &crate::form::VisualControl,
    ) -> FamilyResult<&mut Self> {
        let control = control.clone();
        self.document_operation(move |document| {
            document
                .insert_visual_control(form_index, &control)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert a model at the end of the existing `office:forms` container.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_xforms_model(&mut self, model: &crate::xforms::Model) -> FamilyResult<&mut Self> {
        let model = model.clone();
        self.document_operation(move |document| {
            document
                .insert_xforms_model(&model)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Mark a live insertion or format-change range using Unicode character offsets.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn mark_tracked_change_range(
        &mut self,
        change_id: impl Into<String>,
        start: crate::Position,
        end: crate::Position,
    ) -> FamilyResult<&mut Self> {
        let change_id: String = change_id.into();
        self.document_operation(move |document| {
            document
                .mark_tracked_change_range(&change_id, start.clone(), end.clone())
                .map(|_| OperationResult::Unit)
        })
    }

    /// Place a point deletion marker using a Unicode character offset.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn mark_tracked_deletion(
        &mut self,
        change_id: impl Into<String>,
        position: crate::Position,
    ) -> FamilyResult<&mut Self> {
        let change_id: String = change_id.into();
        self.document_operation(move |document| {
            document
                .mark_tracked_deletion(&change_id, position.clone())
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove bookmark markers while preserving enclosed text and markup.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_bookmark_target(&mut self, target_index: usize) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_bookmark_target(target_index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove a form property and remove its container when it becomes empty.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_form_property(&mut self, property_index: usize) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_form_property(property_index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove a fixed-text, hidden, or generic control selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_generic_form_control(&mut self, control_index: usize) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_generic_form_control(control_index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Stages `remove_grid_control`.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_grid_control(&mut self, control_index: usize) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_grid_control(control_index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove an image-frame control selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_image_frame_control(&mut self, control_index: usize) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_image_frame_control(control_index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove a button or checkbox selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_interactive_control(&mut self, control_index: usize) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_interactive_control(control_index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove one named master page without rewriting unrelated styles.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_master_page(&mut self, name: impl Into<String>) -> FamilyResult<&mut Self> {
        let name: String = name.into();
        self.document_operation_at(format!("/styles/master-page/{name}"), move |document| {
            document
                .remove_master_page(&name)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove one named outline numbering style and return its prior value.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_outline_style(&mut self, name: impl Into<String>) -> FamilyResult<&mut Self> {
        let name: String = name.into();
        self.document_operation(move |document| {
            document
                .remove_outline_style(&name)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove one paragraph selected by a checked semantic position.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_paragraph_at(&mut self, position: Position) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_paragraph_at(position)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove a password or file control selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_password_file_control(
        &mut self,
        control_index: usize,
    ) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_password_file_control(control_index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove marker elements while preserving enclosed text and markup.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_reference_mark(&mut self, mark_index: usize) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_reference_mark(mark_index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Delete a named section together with its enclosed content.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_section(&mut self, name: impl Into<String>) -> FamilyResult<&mut Self> {
        let name: String = name.into();
        self.document_operation(move |document| {
            document
                .remove_section(&name)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove a listbox or combobox selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_selection_control(&mut self, control_index: usize) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_selection_control(control_index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove one top-level table and return its typed value.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_table(&mut self, index: usize) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document.remove_table(index).map(|_| OperationResult::Unit)
        })
    }

    /// Remove a text or textarea control selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_text_control(&mut self, control_index: usize) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_text_control(control_index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove one named index and return its previous inert representation.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_text_index(&mut self, name: impl Into<String>) -> FamilyResult<&mut Self> {
        let name: String = name.into();
        self.document_operation(move |document| {
            document
                .remove_text_index(&name)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Stages `remove_text_index_mark`.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_text_index_mark(&mut self, mark_index: usize) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_text_index_mark(mark_index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove one inline `text:meta` occurrence.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_text_meta(&mut self, position: Position) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_text_meta(position)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove a typed value control selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_typed_value_control(&mut self, control_index: usize) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_typed_value_control(control_index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove a value-range control selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_value_range_control(&mut self, control_index: usize) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_value_range_control(control_index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Atomically remove one declaration container, returning its old value.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_variable_declaration_group(
        &mut self,
        part: crate::variable_declaration::Part,
        scope: &crate::variable_declaration::Scope,
        kind: crate::variable_declaration::Kind,
    ) -> FamilyResult<&mut Self> {
        let scope = scope.clone();
        self.document_operation(move |document| {
            document
                .remove_variable_declaration_group(part, &scope, kind)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove a radio, frame, or image-button selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_visual_control(&mut self, control_index: usize) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_visual_control(control_index)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove one model from `office:forms`.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_xforms_model(&mut self, position: Position) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .remove_xforms_model(position)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace a bookmark target selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_bookmark_target(
        &mut self,
        target_index: usize,
        replacement: &crate::BookmarkTarget,
    ) -> FamilyResult<&mut Self> {
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_bookmark_target(target_index, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace a form property selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_form_property(
        &mut self,
        property_index: usize,
        replacement: &crate::form::Property,
    ) -> FamilyResult<&mut Self> {
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_form_property(property_index, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace a fixed-text, hidden, or generic control selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_generic_form_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::GenericFormControl,
    ) -> FamilyResult<&mut Self> {
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_generic_form_control(control_index, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Stages `replace_grid_control`.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_grid_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::GridControl,
    ) -> FamilyResult<&mut Self> {
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_grid_control(control_index, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace an image-frame control selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_image_frame_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::ImageFrameControl,
    ) -> FamilyResult<&mut Self> {
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_image_frame_control(control_index, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace a button or checkbox selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_interactive_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::InteractiveControl,
    ) -> FamilyResult<&mut Self> {
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_interactive_control(control_index, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace one named master page without rewriting unrelated styles.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_master_page(
        &mut self,
        name: impl Into<String>,
        page: &crate::header_footer::Master,
    ) -> FamilyResult<&mut Self> {
        let name: String = name.into();
        let page = page.clone();
        self.document_operation_at(format!("/styles/master-page/{name}"), move |document| {
            document
                .replace_master_page(&name, &page)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace one paragraph's content with plain text at a checked position.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_paragraph_at(
        &mut self,
        position: Position,
        text: impl Into<String>,
    ) -> FamilyResult<&mut Self> {
        let text: String = text.into();
        self.document_operation(move |document| {
            document
                .replace_paragraph_at(position, &text)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace a password or file control selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_password_file_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::PasswordFileControl,
    ) -> FamilyResult<&mut Self> {
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_password_file_control(control_index, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace a reference target selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_reference_mark(
        &mut self,
        mark_index: usize,
        replacement: &crate::ReferenceMark,
    ) -> FamilyResult<&mut Self> {
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_reference_mark(mark_index, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace a listbox or combobox selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_selection_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::SelectionControl,
    ) -> FamilyResult<&mut Self> {
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_selection_control(control_index, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace a text or textarea control selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_text_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::TextControl,
    ) -> FamilyResult<&mut Self> {
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_text_control(control_index, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace one named index and return its previous inert representation.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_text_index(
        &mut self,
        name: impl Into<String>,
        replacement: &crate::TextIndex,
    ) -> FamilyResult<&mut Self> {
        let name: String = name.into();
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_text_index(&name, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Stages `replace_text_index_mark`.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_text_index_mark(
        &mut self,
        mark_index: usize,
        replacement: &crate::TextIndexMark,
    ) -> FamilyResult<&mut Self> {
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_text_index_mark(mark_index, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace one inline `text:meta` occurrence.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_text_meta(
        &mut self,
        position: Position,
        value: &crate::TextMeta,
    ) -> FamilyResult<&mut Self> {
        let value = value.clone();
        self.document_operation(move |document| {
            document
                .replace_text_meta(position, &value)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace a typed value control selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_typed_value_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::TypedValueControl,
    ) -> FamilyResult<&mut Self> {
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_typed_value_control(control_index, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace a value-range control selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_value_range_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::ValueRangeControl,
    ) -> FamilyResult<&mut Self> {
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_value_range_control(control_index, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace a radio, frame, or image-button selected in document order.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_visual_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::VisualControl,
    ) -> FamilyResult<&mut Self> {
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .replace_visual_control(control_index, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace one model in `office:forms`.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_xforms_model(
        &mut self,
        position: Position,
        model: &crate::xforms::Model,
    ) -> FamilyResult<&mut Self> {
        let model = model.clone();
        self.document_operation(move |document| {
            document
                .replace_xforms_model(position, &model)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert or replace the document-wide bibliography formatting policy.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_bibliography_configuration(
        &mut self,
        configuration: &crate::bibliography_configuration::Configuration,
    ) -> FamilyResult<&mut Self> {
        let configuration = configuration.clone();
        self.document_operation(move |document| {
            document
                .set_bibliography_configuration(&configuration)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Set or clear RDFa on a uniquely named bookmark start.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_bookmark_rdfa(
        &mut self,
        name: impl Into<String>,
        value: &crate::RdfaAttributes,
    ) -> FamilyResult<&mut Self> {
        let name: String = name.into();
        let value = value.clone();
        self.document_operation(move |document| {
            document
                .set_bookmark_rdfa(&name, &value)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace content-part font-face declarations and return the old value.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_content_font_face_declarations(
        &mut self,
        declarations: &crate::font_face::Declarations,
    ) -> FamilyResult<&mut Self> {
        let declarations = declarations.clone();
        self.document_operation(move |document| {
            document
                .set_content_font_face_declarations(&declarations)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Set plain text in one header/footer region of an existing master page.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_header_footer_text(
        &mut self,
        master_page_name: impl Into<String>,
        kind: crate::header_footer::Kind,
        text: impl Into<String>,
    ) -> FamilyResult<&mut Self> {
        let master_page_name: String = master_page_name.into();
        let text: String = text.into();
        self.document_operation_at(
            format!("/styles/master-page/{master_page_name}"),
            move |document| {
                document
                    .set_header_footer_text(&master_page_name, kind, &text)
                    .map(|_| OperationResult::Unit)
            },
        )
    }

    /// Replace one header/footer region with a complete XML fragment.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_header_footer_xml(
        &mut self,
        master_page_name: impl Into<String>,
        kind: crate::header_footer::Kind,
        xml: impl Into<String>,
    ) -> FamilyResult<&mut Self> {
        let master_page_name: String = master_page_name.into();
        let xml: String = xml.into();
        self.document_operation_at(
            format!("/styles/master-page/{master_page_name}"),
            move |document| {
                document
                    .set_header_footer_xml(&master_page_name, kind, &xml)
                    .map(|_| OperationResult::Unit)
            },
        )
    }

    /// Insert or replace document line-numbering configuration.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_line_numbering_configuration(
        &mut self,
        configuration: &crate::line_numbering::Configuration,
    ) -> FamilyResult<&mut Self> {
        let configuration = configuration.clone();
        self.document_operation(move |document| {
            document
                .set_line_numbering_configuration(&configuration)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Add an empty master page and its referenced page layout.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_list_level_label_alignment(
        &mut self,
        item: &crate::list_label_alignment::Style,
    ) -> FamilyResult<&mut Self> {
        let item = item.clone();
        self.document_operation(move |document| {
            document
                .set_list_level_label_alignment(&item)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Override the root MIME type written by `to_bytes`.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_mimetype(&mut self, mimetype: impl Into<String>) -> FamilyResult<&mut Self> {
        let mimetype: String = mimetype.into();
        self.document_operation(move |document| {
            document.set_mimetype(mimetype.clone());
            Ok(OperationResult::Unit)
        })
    }

    /// Insert or replace one stored footnote or endnote configuration.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_notes_configuration(
        &mut self,
        configuration: &crate::notes_configuration::Configuration,
    ) -> FamilyResult<&mut Self> {
        let configuration = configuration.clone();
        self.document_operation(move |document| {
            document
                .set_notes_configuration(&configuration)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace both stored note-class configurations and return the old values.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_notes_configurations(
        &mut self,
        configurations: &crate::notes_configuration::Configurations,
    ) -> FamilyResult<&mut Self> {
        let configurations = configurations.clone();
        self.document_operation(move |document| {
            document
                .set_notes_configurations(&configurations)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert or replace one named outline numbering style.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_outline_style(
        &mut self,
        style: &crate::outline_style::Style,
    ) -> FamilyResult<&mut Self> {
        let style = style.clone();
        self.document_operation(move |document| {
            document
                .set_outline_style(&style)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Create or replace typed columns in one existing page layout.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_page_layout_columns(
        &mut self,
        page_layout_name: impl Into<String>,
        columns: &crate::style::columns::Columns,
    ) -> FamilyResult<&mut Self> {
        let page_layout_name: String = page_layout_name.into();
        let columns = columns.clone();
        self.document_operation_at(
            format!("/styles/page-layout/{page_layout_name}"),
            move |document| {
                document
                    .set_page_layout_columns(&page_layout_name, &columns)
                    .map(|_| OperationResult::Unit)
            },
        )
    }

    /// Create or replace the typed footnote separator in one existing page layout.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_page_layout_footnote_separator(
        &mut self,
        page_layout_name: impl Into<String>,
        separator: &crate::footnote_separator::Separator,
    ) -> FamilyResult<&mut Self> {
        let page_layout_name: String = page_layout_name.into();
        let separator = separator.clone();
        self.document_operation_at(
            format!("/styles/page-layout/{page_layout_name}"),
            move |document| {
                document
                    .set_page_layout_footnote_separator(&page_layout_name, &separator)
                    .map(|_| OperationResult::Unit)
            },
        )
    }

    /// Create or replace typed header/footer properties in one page layout.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_page_layout_header_footer_properties(
        &mut self,
        page_layout_name: impl Into<String>,
        region: crate::header_footer::properties::Region,
        properties: &crate::header_footer::properties::StyleProperties,
    ) -> FamilyResult<&mut Self> {
        let page_layout_name: String = page_layout_name.into();
        let properties = properties.clone();
        self.document_operation_at(
            format!("/styles/page-layout/{page_layout_name}"),
            move |document| {
                document
                    .set_page_layout_header_footer_properties(
                        &page_layout_name,
                        region,
                        &properties,
                    )
                    .map(|_| OperationResult::Unit)
            },
        )
    }

    /// Replace one page layout with a complete XML fragment.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_page_layout_xml(
        &mut self,
        page_layout_name: impl Into<String>,
        page_layout_xml: impl Into<String>,
    ) -> FamilyResult<&mut Self> {
        let page_layout_name: String = page_layout_name.into();
        let page_layout_xml: String = page_layout_xml.into();
        self.document_operation_at(
            format!("/styles/page-layout/{page_layout_name}"),
            move |document| {
                document
                    .set_page_layout_xml(&page_layout_name, &page_layout_xml)
                    .map(|_| OperationResult::Unit)
            },
        )
    }

    /// Set, replace, or clear the document's `text:page-sequence`.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_page_sequence(
        &mut self,
        sequence: Option<&crate::page_sequence::Sequence>,
    ) -> FamilyResult<&mut Self> {
        let sequence = sequence.cloned();
        self.document_operation(move |document| {
            document
                .set_page_sequence(sequence.as_ref())
                .map(|_| OperationResult::Unit)
        })
    }

    /// Set or clear RDFa on one paragraph while retaining unrelated XML.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_paragraph_rdfa(
        &mut self,
        position: Position,
        value: &crate::RdfaAttributes,
    ) -> FamilyResult<&mut Self> {
        let value = value.clone();
        self.document_operation(move |document| {
            document
                .set_paragraph_rdfa(position, &value)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Add an empty master page and its referenced page layout.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_paragraph_style_drop_cap(
        &mut self,
        style: &crate::style::paragraph::drop_cap::Style,
    ) -> FamilyResult<&mut Self> {
        let style = style.clone();
        self.document_operation(move |document| {
            document
                .set_paragraph_style_drop_cap(&style)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace styles-part font-face declarations and return the old value.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_styles_font_face_declarations(
        &mut self,
        declarations: &crate::font_face::Declarations,
    ) -> FamilyResult<&mut Self> {
        let declarations = declarations.clone();
        self.document_operation(move |document| {
            document
                .set_styles_font_face_declarations(&declarations)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace, insert, or remove typed row properties on an existing table-row style.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_table_row_style_properties(
        &mut self,
        style: &crate::style::table::row::Style,
    ) -> FamilyResult<&mut Self> {
        let style = style.clone();
        self.document_operation(move |document| {
            document
                .set_table_row_style_properties(&style)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Replace, insert, or remove typed properties on an existing table style.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_table_style_properties(
        &mut self,
        style: &crate::style::table::table::Style,
    ) -> FamilyResult<&mut Self> {
        let style = style.clone();
        self.document_operation(move |document| {
            document
                .set_table_style_properties(&style)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Set or clear RDFa on one inline `text:meta` occurrence.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_text_meta_rdfa(
        &mut self,
        position: Position,
        value: &crate::RdfaAttributes,
    ) -> FamilyResult<&mut Self> {
        let value = value.clone();
        self.document_operation(move |document| {
            document
                .set_text_meta_rdfa(position, &value)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Atomically replace the declaration table and policy metadata.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_tracked_changes(
        &mut self,
        tracked: crate::TrackedChanges,
    ) -> FamilyResult<&mut Self> {
        self.document_operation(move |document| {
            document
                .set_tracked_changes(tracked.clone())
                .map(|_| OperationResult::Unit)
        })
    }

    /// Atomically insert or replace one declaration container.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn set_variable_declaration_group(
        &mut self,
        group: &crate::variable_declaration::Group,
    ) -> FamilyResult<&mut Self> {
        let group = group.clone();
        self.document_operation(move |document| {
            document
                .set_variable_declaration_group(&group)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove every marker for one declaration while retaining its live text.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn unmark_tracked_change(
        &mut self,
        change_id: impl Into<String>,
    ) -> FamilyResult<&mut Self> {
        let change_id: String = change_id.into();
        self.document_operation(move |document| {
            document
                .unmark_tracked_change(&change_id)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Remove one section wrapper/source while retaining all enclosed content bytes.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn unwrap_section(&mut self, name: impl Into<String>) -> FamilyResult<&mut Self> {
        let name: String = name.into();
        self.document_operation(move |document| {
            document
                .unwrap_section(&name)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Atomically update section metadata/source while preserving enclosed XML bytes.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn update_section(
        &mut self,
        name: impl Into<String>,
        replacement: &crate::Section,
    ) -> FamilyResult<&mut Self> {
        let name: String = name.into();
        let replacement = replacement.clone();
        self.document_operation(move |document| {
            document
                .update_section(&name, &replacement)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Atomically replace a declaration without changing marker identity.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn update_tracked_change(
        &mut self,
        id: impl Into<String>,
        replacement: crate::TrackChange,
    ) -> FamilyResult<&mut Self> {
        let id: String = id.into();
        self.document_operation(move |document| {
            document
                .update_tracked_change(&id, replacement.clone())
                .map(|_| OperationResult::Unit)
        })
    }

    /// Wrap one UTF-8 text range in a selected paragraph with ruby.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn wrap_ruby_annotation(
        &mut self,
        paragraph_index: usize,
        range: std::ops::Range<usize>,
        annotation: &crate::ruby_family::Annotation,
    ) -> FamilyResult<&mut Self> {
        let annotation = annotation.clone();
        self.document_operation(move |document| {
            document
                .wrap_ruby_annotation(paragraph_index, range.clone(), &annotation)
                .map(|_| OperationResult::Unit)
        })
    }

    /// Wrap an inclusive stable block range in a typed section.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn wrap_section(
        &mut self,
        section: &crate::Section,
        start: crate::Block,
        end: crate::Block,
    ) -> FamilyResult<&mut Self> {
        let section = section.clone();
        self.document_operation(move |document| {
            document
                .wrap_section(&section, start.clone(), end.clone())
                .map(|_| OperationResult::Unit)
        })
    }

    /// Insert one compact typed section-layout style into `styles.xml`.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_section_layout(
        &mut self,
        section: &crate::header_footer::SectionLayout,
    ) -> FamilyResult<&mut Self> {
        let section = section.clone();
        self.document_operation_at(
            format!("/styles/section/{}", section.name()),
            move |document| {
                document
                    .edit_layout_styles(|edit| edit.insert_section_layout(&section))
                    .map(|()| OperationResult::Unit)
            },
        )
    }

    /// Replace one named section-layout style with a typed compact owner.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn replace_section_layout(
        &mut self,
        section: &crate::header_footer::SectionLayout,
    ) -> FamilyResult<&mut Self> {
        let section = section.clone();
        self.document_operation_at(
            format!("/styles/section/{}", section.name()),
            move |document| {
                document
                    .edit_layout_styles(|edit| edit.replace_section_layout(&section))
                    .map(|()| OperationResult::Unit)
            },
        )
    }

    /// Remove one named section-layout style.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn remove_section_layout(&mut self, name: impl Into<String>) -> FamilyResult<&mut Self> {
        let name: String = name.into();
        self.document_operation_at(format!("/styles/section/{name}"), move |document| {
            document
                .edit_layout_styles(|edit| edit.remove_section_layout(&name))
                .map(|()| OperationResult::Unit)
        })
    }

    /// Insert a master page and its page layout prepared by
    /// [`header_footer::Snapshot::prepare_master_page_transfer`](crate::header_footer::Snapshot::prepare_master_page_transfer)
    /// on another document. Every linked dependency must have its `href` in
    /// `available` (resources the destination already holds); none is
    /// fetched.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_master_page_transfer(
        &mut self,
        transfer: &crate::header_footer::Transfer,
        available: &[&str],
    ) -> FamilyResult<&mut Self> {
        let transfer = transfer.clone();
        let available: Vec<String> = available.iter().map(|href| (*href).to_owned()).collect();
        self.document_operation_at(
            format!("/styles/master-page/{}", transfer.master_page_name()),
            move |document| {
                document
                    .edit_layout_styles(|edit| {
                        edit.insert_transfer_with(&transfer, |dependency| {
                            available.iter().any(|href| href == dependency.href())
                        })
                    })
                    .map(|()| OperationResult::Unit)
            },
        )
    }

    /// Insert a section layout prepared by
    /// [`header_footer::Snapshot::prepare_section_layout_transfer`](crate::header_footer::Snapshot::prepare_section_layout_transfer)
    /// on another document; `available` lists the dependency `href`s the
    /// destination already holds.
    ///
    /// Staged: the operation applies when [`Edit::commit`] runs, in staging
    /// order, and reports its errors there.
    pub fn insert_section_layout_transfer(
        &mut self,
        transfer: &crate::header_footer::SectionTransfer,
        available: &[&str],
    ) -> FamilyResult<&mut Self> {
        let transfer = transfer.clone();
        let available: Vec<String> = available.iter().map(|href| (*href).to_owned()).collect();
        self.document_operation_at(
            format!("/styles/section/{}", transfer.name()),
            move |document| {
                document
                    .edit_layout_styles(|edit| {
                        edit.insert_section_transfer_with(&transfer, |dependency| {
                            available.iter().any(|href| href == dependency.href())
                        })
                    })
                    .map(|()| OperationResult::Unit)
            },
        )
    }
}
