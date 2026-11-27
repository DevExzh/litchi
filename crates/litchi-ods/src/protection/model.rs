//! Semantic ODS protection vocabulary.
//!
//! The wire-shaped protection records already live under `model`.  This
//! module gives the package facade names that carry their document context
//! and keeps the automatic cell-style projection separate from the sheet
//! protection records.

use crate::model::{
    protection as wire,
    style_protection::{self as style_wire, Protection},
};

pub(super) use style_wire::ConditionalStyle;
pub use wire::Key;

/// Protection metadata on the document's `office:spreadsheet` element.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Document {
    /// Whether the workbook structure is protected. `None` retains an
    /// omitted attribute, while `Some(false)` retains an explicit false.
    pub structure_protected: Option<bool>,
    /// Password-verifier metadata. This is never treated as an encryption
    /// key and is not verified by the facade.
    pub key: Key,
}

/// `LibreOffice`'s optional sheet edit permissions.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Permissions {
    pub select_protected_cells: Option<bool>,
    pub select_unprotected_cells: Option<bool>,
    pub insert_columns: Option<bool>,
    pub insert_rows: Option<bool>,
    pub delete_columns: Option<bool>,
    pub delete_rows: Option<bool>,
    pub use_auto_filter: Option<bool>,
    pub use_pivot: Option<bool>,
}

impl Permissions {
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.select_protected_cells.is_none()
            && self.select_unprotected_cells.is_none()
            && self.insert_columns.is_none()
            && self.insert_rows.is_none()
            && self.delete_columns.is_none()
            && self.delete_rows.is_none()
            && self.use_auto_filter.is_none()
            && self.use_pivot.is_none()
    }
}

/// Protection metadata for one worksheet, selected by its ODF table name.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Sheet {
    pub name: String,
    pub protected: Option<bool>,
    pub key: Key,
    pub permissions: Permissions,
}

impl Sheet {
    #[must_use]
    pub fn is_protected(&self) -> Option<bool> {
        self.protected
    }
}

/// One automatic table-cell style carrying a `style:cell-protect` value.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Style {
    pub name: String,
    pub parent_name: Option<String>,
    pub protection: Protection,
}

impl Style {
    pub fn new(name: impl Into<String>, protection: Protection) -> Self {
        Self {
            name: name.into(),
            parent_name: None,
            protection,
        }
    }

    pub fn with_parent_name(mut self, parent_name: impl Into<String>) -> Self {
        self.parent_name = Some(parent_name.into());
        self
    }
}

/// The automatic table-cell protection-style catalog in one ODS content
/// part. Conditional `style:map` rules are retained internally by the
/// transaction so editing this catalog does not discard them.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Styles {
    automatic: Vec<Style>,
    conditional: Vec<ConditionalStyle>,
}

impl Styles {
    #[must_use]
    pub fn automatic(&self) -> &[Style] {
        &self.automatic
    }

    #[must_use]
    pub fn get(&self, name: &str) -> Option<&Style> {
        self.automatic.iter().find(|style| style.name == name)
    }

    /// Borrow conditional `style:map` catalogs in source order.
    #[must_use]
    pub fn conditional(&self) -> &[ConditionalStyle] {
        &self.conditional
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.automatic.is_empty() && self.conditional.is_empty()
    }

    /// Replace the automatic protection-style catalog in a transaction draft.
    pub fn set_automatic(&mut self, styles: Vec<Style>) {
        self.automatic = styles;
    }

    /// Replace the inert conditional protection-style catalog in a transaction draft.
    pub fn set_conditional(&mut self, styles: Vec<ConditionalStyle>) {
        self.conditional = styles;
    }
}

pub(crate) fn document_from_wire(value: &wire::Protection) -> Document {
    Document {
        structure_protected: value.structure_protected,
        key: value.key.clone(),
    }
}

pub(crate) fn permissions_from_wire(value: &wire::Options) -> Permissions {
    Permissions {
        select_protected_cells: value.select_protected_cells,
        select_unprotected_cells: value.select_unprotected_cells,
        insert_columns: value.insert_columns,
        insert_rows: value.insert_rows,
        delete_columns: value.delete_columns,
        delete_rows: value.delete_rows,
        use_auto_filter: value.use_auto_filter,
        use_pivot: value.use_pivot,
    }
}

pub(crate) fn sheet_from_wire(name: String, value: &wire::Sheet) -> Sheet {
    Sheet {
        name,
        protected: value.protected,
        key: value.key.clone(),
        permissions: permissions_from_wire(&value.options),
    }
}

pub(crate) fn styles_from_wire(registry: &style_wire::CellStyleRegistry) -> Styles {
    let automatic = registry
        .automatic_protection_styles()
        .iter()
        .map(|style| Style {
            name: style.style_name.clone(),
            parent_name: style.parent_style_name.clone(),
            protection: style.protection,
        })
        .collect();
    Styles {
        automatic,
        conditional: registry.conditional_styles().to_vec(),
    }
}
