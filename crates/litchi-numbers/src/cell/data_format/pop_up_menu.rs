//! Pop-Up Menu cell-control values.

use std::fmt;

/// Maximum UTF-8 size of one menu item.
pub const MAX_ITEM_BYTES: usize = 4 * 1_024;

/// Errors returned by checked Pop-Up Menu construction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    /// A menu contains no choices.
    Empty,
    /// An item exceeds [`MAX_ITEM_BYTES`].
    ItemTooLong { length: usize, maximum: usize },
    /// An item contains a control character.
    ItemContainsControl { index: usize },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Empty => formatter.write_str("Pop-Up Menu must contain at least one item"),
            Self::ItemTooLong { length, maximum } => {
                write!(
                    formatter,
                    "Pop-Up Menu item is {length} bytes; maximum is {maximum}"
                )
            },
            Self::ItemContainsControl { index } => write!(
                formatter,
                "Pop-Up Menu item contains a control character at character index {index}"
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Result returned by checked Pop-Up Menu constructors.
pub type Result<T> = std::result::Result<T, Error>;

/// Text displayed by one Pop-Up Menu choice.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Item(Box<str>);

impl Item {
    /// Validates and stores a borrowed item without allocating on failure.
    ///
    /// # Errors
    ///
    /// Returns a typed error for oversized or control-containing text.
    pub fn new(value: &str) -> Result<Self> {
        validate(value)?;
        Ok(Self(value.into()))
    }

    /// Validates and stores an owned item without an extra copy.
    ///
    /// # Errors
    ///
    /// Returns a typed error for oversized or control-containing text.
    pub fn from_owned(value: String) -> Result<Self> {
        validate(&value)?;
        Ok(Self(value.into_boxed_str()))
    }

    /// Borrows the displayed item text.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for Item {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<&str> for Item {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<String> for Item {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        Self::from_owned(value)
    }
}

/// Value assigned when a menu is applied to an empty cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum InitialSelection {
    /// Select the first item.
    #[default]
    FirstItem,
    /// Leave the cell blank until a choice is made.
    Blank,
}

/// Pop-Up Menu cell-control format.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PopUpMenu {
    items: Box<[Item]>,
    initial_selection: InitialSelection,
}

impl PopUpMenu {
    /// Validates and constructs a menu in display order.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Empty`] when no items are supplied, or an item
    /// validation error when an item is oversized or contains controls.
    pub fn new<I, T>(items: I) -> Result<Self>
    where
        I: IntoIterator<Item = T>,
        T: AsRef<str>,
    {
        let collected_items = items
            .into_iter()
            .map(|item| Item::new(item.as_ref()))
            .collect::<Result<Vec<_>>>()?
            .into_boxed_slice();
        if collected_items.is_empty() {
            return Err(Error::Empty);
        }
        Ok(Self {
            items: collected_items,
            initial_selection: InitialSelection::FirstItem,
        })
    }

    /// Borrows choices in their display order.
    #[must_use]
    pub fn items(&self) -> &[Item] {
        &self.items
    }

    /// Returns the empty-cell initialization policy.
    #[must_use]
    pub const fn initial_selection(&self) -> InitialSelection {
        self.initial_selection
    }

    /// Replaces the empty-cell initialization policy.
    #[must_use]
    pub const fn with_initial_selection(mut self, value: InitialSelection) -> Self {
        self.initial_selection = value;
        self
    }
}

impl Default for PopUpMenu {
    fn default() -> Self {
        Self {
            items: Box::new([
                Item("Item 1".into()),
                Item("Item 2".into()),
                Item("Item 3".into()),
            ]),
            initial_selection: InitialSelection::FirstItem,
        }
    }
}

fn validate(value: &str) -> Result<()> {
    if value.len() > MAX_ITEM_BYTES {
        return Err(Error::ItemTooLong {
            length: value.len(),
            maximum: MAX_ITEM_BYTES,
        });
    }
    if let Some((index, _)) = value
        .chars()
        .enumerate()
        .find(|(_, character)| character.is_control())
    {
        return Err(Error::ItemContainsControl { index });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn menu_validation_rejects_empty_and_controlled_input() {
        assert_eq!(PopUpMenu::new(Vec::<&str>::new()), Err(Error::Empty));
        assert!(matches!(
            Item::new("High\nPriority"),
            Err(Error::ItemContainsControl { index: 4 })
        ));
    }

    #[test]
    fn menu_values_round_trip_order_and_selection() {
        let Ok(value) = PopUpMenu::new(["Low", "High"]) else {
            panic!("nonempty visible menu should construct");
        };
        let value = value.with_initial_selection(InitialSelection::Blank);
        assert_eq!(
            value.items().iter().map(Item::as_str).collect::<Vec<_>>(),
            ["Low", "High"]
        );
        assert_eq!(value.initial_selection(), InitialSelection::Blank);
    }
}
