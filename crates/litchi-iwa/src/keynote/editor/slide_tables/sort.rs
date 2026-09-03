//! Source-built compatibility execution for Keynote slide tables.
//!
//! Exact packages are owned by `litchi-keynote`'s physical `Sort Now`
//! transaction. This module retains only the narrow writer needed for
//! `KeynoteDocumentBuilder` output, which the focused package intentionally
//! does not admit yet.

use super::*;
use litchi_keynote::slide::table::{
    lock::State as FocusedTableLockState,
    sort::{Order, RowRange},
};
use litchi_keynote::{SlideSelector, TableSelector};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct PhysicalTableSelection {
    slide_index: usize,
    table_index: usize,
    model_object_id: u64,
}

impl KeynoteEditor {
    /// Execute a source-built slide table's configured full-table sort order.
    ///
    /// This internal migration-host path exists only for
    /// `KeynoteDocumentBuilder` graphs carrying the explicit source-built
    /// template marker. Exact packages must use
    /// `litchi_keynote::Package::execute_slide_table_sort_order`; they never
    /// fall back to this writer. Returns `true` when one or more rows moved
    /// and `false` for an already stable body order.
    pub(super) fn apply_source_built_table_sort_order(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
    ) -> Result<bool> {
        require_source_built_compatibility(self)?;
        let selection = select_legacy_table(self, slide_index, model_object_id)?;
        execute_source_built_table_sort(self, selection, None)
    }

    /// Execute a source-built slide table's configured selected-row sort over
    /// one body-relative range.
    ///
    /// This internal migration-host path exists only for
    /// `KeynoteDocumentBuilder` graphs carrying the explicit source-built
    /// template marker. Exact packages must use
    /// `litchi_keynote::Package::execute_slide_table_sort_order_to_rows`; they
    /// never fall back to this writer. The range is half-open and body
    /// relative, excluding header and footer rows.
    pub(super) fn apply_source_built_table_sort_order_to_rows(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        rows: RowRange,
    ) -> Result<bool> {
        require_source_built_compatibility(self)?;
        let selection = select_legacy_table(self, slide_index, model_object_id)?;
        execute_source_built_table_sort(self, selection, Some(rows))
    }
}

fn select_legacy_table(
    editor: &KeynoteEditor,
    slide_index: usize,
    model_object_id: u64,
) -> Result<PhysicalTableSelection> {
    require_table_model(editor, slide_index, model_object_id)?;
    let tables = editor.slide_tables(slide_index)?;
    let mut table_index = None;
    for (index, table) in tables.iter().enumerate() {
        if table.model_object_id != model_object_id {
            continue;
        }
        if table_index.replace(index).is_some() {
            return Err(Error::ParseError(format!(
                "Keynote object {model_object_id} has ambiguous table ownership on slide {slide_index}"
            )));
        }
    }
    let table_index = table_index.ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote table model {model_object_id} is not owned by slide {slide_index}"
        ))
    })?;
    Ok(PhysicalTableSelection {
        slide_index,
        table_index,
        model_object_id,
    })
}

/// Recognize the legacy source-built graph without treating arbitrary exact
/// packages as eligible for fallback. The marker is emitted only by
/// `KeynoteDocumentBuilder` and survives its compatibility save/reopen path.
fn source_built_compatibility_package(editor: &KeynoteEditor) -> Result<bool> {
    const DOCUMENT_ENTRY: &str = "Index/Document.iwa";
    const DOCUMENT_IDENTIFIER: u64 = 1;
    const DOCUMENT_MESSAGE_TYPE: u32 = 1;
    const SOURCE_BUILT_TEMPLATE: &str = "Application/Litchi/Blank/Wide";

    let source_is_exact = editor.package().source_is_exact();
    let marker = (|| {
        let archive = editor.package().archive(DOCUMENT_ENTRY)?;
        let object = archive
            .object(DOCUMENT_IDENTIFIER)
            .ok_or_else(|| Error::InvalidFormat("Keynote document root is missing".to_owned()))?;
        let mut messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == DOCUMENT_MESSAGE_TYPE);
        let message = messages.next().ok_or_else(|| {
            Error::InvalidFormat("Keynote document payload is missing".to_owned())
        })?;
        if messages.next().is_some() {
            return Err(Error::InvalidFormat(
                "Keynote document payload is ambiguous".to_owned(),
            ));
        }
        let document = kn::DocumentArchive::decode(message.data.as_slice())?;
        Ok(document.super_.template_identifier.as_deref() == Some(SOURCE_BUILT_TEMPLATE))
    })();

    match marker {
        Ok(marked) => Ok(marked),
        // An exact/native package that does not expose the compatibility
        // marker belongs to the focused owner. Never route it through the
        // permissive Numbers writer merely because this narrow probe failed.
        Err(_) if source_is_exact => Ok(false),
        Err(error) => Err(error),
    }
}

fn require_source_built_compatibility(editor: &KeynoteEditor) -> Result<()> {
    if source_built_compatibility_package(editor)? {
        return Ok(());
    }
    Err(Error::InvalidFormat(
        "Keynote physical Sort Now is owned by litchi_keynote::Package for this package; source-built compatibility requires the KeynoteDocumentBuilder marker".to_owned(),
    ))
}

fn execute_source_built_table_sort(
    editor: &mut KeynoteEditor,
    selection: PhysicalTableSelection,
    rows: Option<RowRange>,
) -> Result<bool> {
    require_table_model(editor, selection.slide_index, selection.model_object_id)?;
    let (order, lock_state) = focused_table_sort_context(editor, selection)?;
    reject_locked_table(lock_state)?;
    let order = order.ok_or_else(|| {
        Error::ParseError(
            "Cannot execute a Keynote table sort without a configured table sort order".to_owned(),
        )
    })?;
    let mut staged = editor.package().clone();
    let changed = match rows {
        Some(rows) => crate::numbers::editor::apply_table_sort_order_to_rows_in_package(
            &mut staged,
            selection.model_object_id,
            &order,
            rows,
        )?,
        None => crate::numbers::editor::apply_table_sort_order_in_package(
            &mut staged,
            selection.model_object_id,
            &order,
        )?,
    };
    if !changed {
        return Ok(false);
    }
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    require_table_model(&verified, selection.slide_index, selection.model_object_id)?;
    if focused_table_sort_order(&verified, selection)?.as_ref() != Some(&order) {
        return Err(Error::InvalidFormat(
            "Keynote table sort execution did not preserve its sort order".to_owned(),
        ));
    }
    *editor = verified;
    Ok(true)
}

fn focused_table_sort_order(
    editor: &KeynoteEditor,
    selection: PhysicalTableSelection,
) -> Result<Option<Order>> {
    Ok(focused_table_sort_context(editor, selection)?.0)
}

fn focused_table_sort_context(
    editor: &KeynoteEditor,
    selection: PhysicalTableSelection,
) -> Result<(Option<Order>, FocusedTableLockState)> {
    let package = focused_table_package(editor)?;
    let order = package
        .slide_table_sort_order(
            SlideSelector::index(selection.slide_index),
            TableSelector::index(selection.table_index),
        )
        .map_err(|error| Error::InvalidFormat(format!("focused Keynote sort read: {error}")))?;
    let lock_state = package
        .slide_table_lock_state(
            SlideSelector::index(selection.slide_index),
            TableSelector::index(selection.table_index),
        )
        .map_err(|error| Error::InvalidFormat(format!("focused Keynote lock read: {error}")))?;
    Ok((order, lock_state))
}

fn reject_locked_table(lock_state: FocusedTableLockState) -> Result<()> {
    if lock_state == FocusedTableLockState::Locked {
        return Err(Error::InvalidFormat(
            "Cannot execute a Keynote table sort on a locked table".to_owned(),
        ));
    }
    Ok(())
}

fn focused_table_package(editor: &KeynoteEditor) -> Result<litchi_keynote::Package> {
    litchi_keynote::Package::from_bytes(&editor.to_bytes()?).map_err(|error| {
        Error::InvalidFormat(format!("focused Keynote sort source failed: {error}"))
    })
}
