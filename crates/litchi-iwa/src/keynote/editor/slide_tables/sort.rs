//! Typed sort-rule editing and execution for Keynote slide tables.

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
    /// Execute a slide table's configured full-table sort order.
    ///
    /// This selector-first compatibility-host operation is the programmatic
    /// equivalent of Keynote's physical **Sort Now** action. It moves only
    /// body rows, preserves the persisted rule, and exposes no native object
    /// identity. Returns `true` when rows moved and `false` when the selected
    /// table was already stable.
    pub fn execute_slide_table_sort_order<'slide>(
        &mut self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<bool> {
        let selection = select_semantic_table(self, slide.into(), table.into())?;
        execute_table_sort_selection(self, selection, None)
    }

    /// Execute a configured selected-row sort over one body-relative range.
    ///
    /// The half-open range excludes header and footer rows. Selection remains
    /// semantic and position-based; native model identities stay private to
    /// the compatibility host.
    pub fn execute_slide_table_sort_order_to_rows<'slide>(
        &mut self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        rows: RowRange,
    ) -> Result<bool> {
        let selection = select_semantic_table(self, slide.into(), table.into())?;
        execute_table_sort_selection(self, selection, Some(rows))
    }

    /// Execute a slide table's configured full-table sort order.
    ///
    /// This physically reorders only body rows and retains the native rule.
    /// The supported scalar subset and safety checks are owned by the focused
    /// [`litchi_keynote::slide::table::physical_sort`] adapter. Returns `true`
    /// when one or more rows moved and `false` for an already stable body
    /// order.
    #[deprecated(
        since = "0.0.1",
        note = "use execute_slide_table_sort_order with SlideSelector and TableSelector; native model IDs remain migration-host compatibility only"
    )]
    pub fn apply_slide_table_sort_order(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
    ) -> Result<bool> {
        let selection = select_legacy_table(self, slide_index, model_object_id)?;
        self.execute_slide_table_sort_order_selection(selection)
    }

    fn execute_slide_table_sort_order_selection(
        &mut self,
        selection: PhysicalTableSelection,
    ) -> Result<bool> {
        execute_table_sort_selection(self, selection, None)
    }

    /// Execute a slide table's configured selected-row sort over one range.
    ///
    /// The half-open range is body-relative, excluding header and footer rows.
    /// Returns `true` when one or more selected rows moved.
    #[deprecated(
        since = "0.0.1",
        note = "use execute_slide_table_sort_order_to_rows with SlideSelector and TableSelector; native model IDs remain migration-host compatibility only"
    )]
    pub fn apply_slide_table_sort_order_to_rows(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        rows: RowRange,
    ) -> Result<bool> {
        let selection = select_legacy_table(self, slide_index, model_object_id)?;
        self.execute_slide_table_sort_order_to_rows_selection(selection, rows)
    }

    fn execute_slide_table_sort_order_to_rows_selection(
        &mut self,
        selection: PhysicalTableSelection,
        rows: RowRange,
    ) -> Result<bool> {
        execute_table_sort_selection(self, selection, Some(rows))
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

fn select_semantic_table(
    editor: &KeynoteEditor,
    slide: SlideSelector<'_>,
    table: TableSelector,
) -> Result<PhysicalTableSelection> {
    let slides = editor.slides()?;
    let slide_index = match slide {
        SlideSelector::Position(position) => position.get(),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(Error::ParseError(
                    "Keynote slide name selector must not be empty".to_owned(),
                ));
            }
            let mut matches = slides
                .iter()
                .filter(|candidate| candidate.name.as_deref() == Some(name));
            let selected = matches.next().ok_or_else(|| {
                Error::ParseError("Keynote slide name selector matched no slide".to_owned())
            })?;
            if matches.next().is_some() {
                return Err(Error::ParseError(
                    "Keynote slide name selector is ambiguous".to_owned(),
                ));
            }
            selected.index
        },
        _ => {
            return Err(Error::ParseError(
                "Keynote slide selector is not supported by this compatibility host".to_owned(),
            ));
        },
    };
    if slide_index >= slides.len() {
        return Err(Error::ParseError(format!(
            "Keynote slide position {slide_index} is outside the {} slide presentation",
            slides.len()
        )));
    }
    let tables = editor.slide_tables(slide_index)?;
    let table_index = table.as_index();
    let selected = tables.get(table_index).ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote table position {table_index} is outside the selected slide's {} tables",
            tables.len()
        ))
    })?;
    Ok(PhysicalTableSelection {
        slide_index,
        table_index,
        model_object_id: selected.model_object_id,
    })
}

fn execute_table_sort_selection(
    editor: &mut KeynoteEditor,
    selection: PhysicalTableSelection,
    rows: Option<RowRange>,
) -> Result<bool> {
    if source_built_compatibility_package(editor)? {
        execute_source_built_table_sort(editor, selection, rows)
    } else {
        execute_focused_table_sort(
            editor,
            SlideSelector::index(selection.slide_index),
            TableSelector::index(selection.table_index),
            rows,
        )
    }
}

/// Recognize the legacy source-built graph without treating arbitrary exact
/// packages as eligible for fallback. The marker is emitted only by
/// `KeynoteDocumentBuilder` and survives its compatibility save/reopen path.
fn source_built_compatibility_package(editor: &KeynoteEditor) -> Result<bool> {
    const DOCUMENT_ENTRY: &str = "Index/Document.iwa";
    const DOCUMENT_IDENTIFIER: u64 = 1;
    const DOCUMENT_MESSAGE_TYPE: u32 = 1;
    const SOURCE_BUILT_TEMPLATE: &str = "Application/Litchi/Blank/Wide";

    let archive = editor.package().archive(DOCUMENT_ENTRY)?;
    let object = archive
        .object(DOCUMENT_IDENTIFIER)
        .ok_or_else(|| Error::InvalidFormat("Keynote document root is missing".to_owned()))?;
    let mut messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == DOCUMENT_MESSAGE_TYPE);
    let message = messages
        .next()
        .ok_or_else(|| Error::InvalidFormat("Keynote document payload is missing".to_owned()))?;
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(
            "Keynote document payload is ambiguous".to_owned(),
        ));
    }
    let document = kn::DocumentArchive::decode(message.data.as_slice())?;
    Ok(document.super_.template_identifier.as_deref() == Some(SOURCE_BUILT_TEMPLATE))
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

/// Execute the physical operation through the focused Keynote package owner.
///
/// The compatibility host only adopts a verified package commit. It does not
/// invoke the legacy Numbers writer, which keeps selector resolution, exact
/// source admission, row-affine topology checks, and transaction conflicts in
/// the focused format crate. A no-op commit is intentionally not serialized
/// back through the host so that its source bytes remain byte-for-byte stable.
fn execute_focused_table_sort(
    editor: &mut KeynoteEditor,
    slide: SlideSelector<'_>,
    table: TableSelector,
    rows: Option<RowRange>,
) -> Result<bool> {
    let package = focused_table_package(editor)?;
    let before = package
        .slide_table_sort_order(slide, table)
        .map_err(|error| {
            Error::InvalidFormat(format!("focused Keynote sort read failed: {error}"))
        })?;
    let commit = match rows {
        Some(rows) => package
            .execute_slide_table_sort_order_to_rows(slide, table, rows)
            .map_err(|error| {
                Error::InvalidFormat(format!("focused Keynote physical sort failed: {error}"))
            })?,
        None => package
            .execute_slide_table_sort_order(slide, table)
            .map_err(|error| {
                Error::InvalidFormat(format!("focused Keynote physical sort failed: {error}"))
            })?,
    };
    let changed = commit.diagnostics().changed();
    if !changed {
        if !commit.patch().is_noop() {
            return Err(Error::InvalidFormat(
                "focused Keynote physical sort reported an inconsistent no-op".to_owned(),
            ));
        }
        return Ok(false);
    }

    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote physical sort write failed: {error}"
        ))
    })?;
    let verified = KeynoteEditor::from_bytes(&bytes)?;
    let verified_package = focused_table_package(&verified)?;
    let after = verified_package
        .slide_table_sort_order(slide, table)
        .map_err(|error| {
            Error::InvalidFormat(format!("focused Keynote sort verify failed: {error}"))
        })?;
    if after != before {
        return Err(Error::InvalidFormat(
            "Keynote physical sort did not preserve its persisted sort order".to_owned(),
        ));
    }
    *editor = verified;
    Ok(true)
}

fn focused_table_package(editor: &KeynoteEditor) -> Result<litchi_keynote::Package> {
    litchi_keynote::Package::from_bytes(&editor.to_bytes()?).map_err(|error| {
        Error::InvalidFormat(format!("focused Keynote sort source failed: {error}"))
    })
}
