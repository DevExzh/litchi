use crate::{Error, Result};

/// Text displayed by one native Pop-Up Menu choice.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TableCellPopUpMenuItem(Box<str>);

impl TableCellPopUpMenuItem {
    /// Construct a choice from its displayed text.
    pub fn new(value: impl Into<Box<str>>) -> Self {
        Self(value.into())
    }

    /// Return the displayed text.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl<T: Into<Box<str>>> From<T> for TableCellPopUpMenuItem {
    fn from(value: T) -> Self {
        Self::new(value)
    }
}

impl AsRef<str> for TableCellPopUpMenuItem {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// Value assigned when Pop-Up Menu is applied to an empty cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TableCellPopUpMenuInitialSelection {
    /// Select the first menu item.
    #[default]
    FirstItem,
    /// Leave the cell blank until a user chooses an item.
    Blank,
}

/// Native interactive Pop-Up Menu format for one table cell.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TableCellPopUpMenuFormat {
    items: Box<[TableCellPopUpMenuItem]>,
    initial_selection: TableCellPopUpMenuInitialSelection,
}

impl TableCellPopUpMenuFormat {
    /// Construct a Pop-Up Menu with at least one choice.
    pub fn try_new(
        items: impl IntoIterator<Item = impl Into<TableCellPopUpMenuItem>>,
    ) -> Result<Self> {
        let items = items
            .into_iter()
            .map(Into::into)
            .collect::<Vec<_>>()
            .into_boxed_slice();
        if items.is_empty() {
            return Err(Error::InvalidFormat(
                "table-cell Pop-Up Menu must contain at least one item".to_owned(),
            ));
        }
        Ok(Self {
            items,
            initial_selection: TableCellPopUpMenuInitialSelection::FirstItem,
        })
    }

    /// Return the choices in native display order.
    pub fn items(&self) -> &[TableCellPopUpMenuItem] {
        &self.items
    }

    /// Return the empty-cell initialization policy.
    pub const fn initial_selection(&self) -> TableCellPopUpMenuInitialSelection {
        self.initial_selection
    }

    /// Replace the empty-cell initialization policy.
    pub const fn with_initial_selection(
        mut self,
        initial_selection: TableCellPopUpMenuInitialSelection,
    ) -> Self {
        self.initial_selection = initial_selection;
        self
    }
}

impl Default for TableCellPopUpMenuFormat {
    fn default() -> Self {
        Self::try_new(["Item 1", "Item 2", "Item 3"])
            .expect("the native Pop-Up Menu defaults are nonempty")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn menu_requires_an_item_and_preserves_order() {
        assert!(TableCellPopUpMenuFormat::try_new(Vec::<String>::new()).is_err());
        let format = TableCellPopUpMenuFormat::try_new(["Low", "High"])
            .unwrap()
            .with_initial_selection(TableCellPopUpMenuInitialSelection::Blank);
        assert_eq!(
            format
                .items()
                .iter()
                .map(TableCellPopUpMenuItem::as_str)
                .collect::<Vec<_>>(),
            ["Low", "High"]
        );
        assert_eq!(
            format.initial_selection(),
            TableCellPopUpMenuInitialSelection::Blank
        );
    }
}
