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
        self.execute_slide_table_sort_order_selection(selection)
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
        self.execute_slide_table_sort_order_to_rows_selection(selection, rows)
    }

    /// Execute a slide table's configured full-table sort order.
    ///
    /// This physically reorders only body rows and retains the native rule.
    /// The supported scalar subset and safety checks are the same as for
    /// [`crate::numbers::NumbersEditor::apply_table_sort_order`]. Returns
    /// `true` when one or more rows moved and `false` for an already stable
    /// body order.
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
        require_table_model(self, selection.slide_index, selection.model_object_id)?;
        let (order, lock_state) = focused_table_sort_context(self, selection)?;
        reject_locked_table(lock_state)?;
        let order = order.ok_or_else(|| {
            Error::ParseError(
                "Cannot execute a Keynote table sort without a configured table sort order"
                    .to_owned(),
            )
        })?;
        let mut staged = self.package().clone();
        if !crate::numbers::editor::apply_table_sort_order_in_package(
            &mut staged,
            selection.model_object_id,
            &order,
        )? {
            return Ok(false);
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, selection.slide_index, selection.model_object_id)?;
        if focused_table_sort_order(&verified, selection)?.as_ref() != Some(&order) {
            return Err(Error::InvalidFormat(
                "Keynote table sort execution did not preserve its sort order".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
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
        require_table_model(self, selection.slide_index, selection.model_object_id)?;
        let (order, lock_state) = focused_table_sort_context(self, selection)?;
        reject_locked_table(lock_state)?;
        let order = order.ok_or_else(|| {
            Error::ParseError(
                "Cannot execute a Keynote table sort without a configured table sort order"
                    .to_owned(),
            )
        })?;
        let mut staged = self.package().clone();
        if !crate::numbers::editor::apply_table_sort_order_to_rows_in_package(
            &mut staged,
            selection.model_object_id,
            &order,
            rows,
        )? {
            return Ok(false);
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, selection.slide_index, selection.model_object_id)?;
        if focused_table_sort_order(&verified, selection)?.as_ref() != Some(&order) {
            return Err(Error::InvalidFormat(
                "Keynote selected-row table sort did not preserve its sort order".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }
}

/// Read persisted sort configuration through the focused package owner while
/// retaining the historical model identifier at the physical executor edge.
fn focused_table_sort_order(
    editor: &KeynoteEditor,
    selection: PhysicalTableSelection,
) -> Result<Option<Order>> {
    Ok(focused_table_sort_context(editor, selection)?.0)
}

/// Read the persisted order and lock state before handing the model to the
/// legacy physical executor.  The focused package owns both persisted facts;
/// the executor must not move rows in a locked table.
fn focused_table_sort_context(
    editor: &KeynoteEditor,
    selection: PhysicalTableSelection,
) -> Result<(Option<Order>, FocusedTableLockState)> {
    let package = litchi_keynote::Package::from_bytes(&editor.to_bytes()?)
        .map_err(|error| Error::InvalidFormat(format!("focused Keynote sort source: {error}")))?;
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

fn reject_locked_table(lock_state: FocusedTableLockState) -> Result<()> {
    if lock_state == FocusedTableLockState::Locked {
        return Err(Error::InvalidFormat(
            "Cannot execute a Keynote table sort on a locked table".to_owned(),
        ));
    }
    Ok(())
}
