//! Semantic model for inert ODF DDE declarations and references.

use crate::variable_declaration::{Part, Scope};

/// A named DDE source declaration. It is retained but never contacted.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Declaration {
    pub part: Part,
    pub scope: Scope,
    pub name: String,
    pub application: String,
    pub topic: String,
    pub item: String,
    pub automatic_update: Option<bool>,
}

impl Declaration {
    /// ODF defaults `office:automatic-update` to `false`.
    pub fn effective_automatic_update(&self) -> bool {
        self.automatic_update.unwrap_or(false)
    }
}

/// One `text:dde-connection` occurrence referring to a declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Use {
    pub part: Part,
    pub scope: Scope,
    pub connection_name: String,
}

/// Package-wide DDE declarations and references in source order.
#[derive(Default)]
pub(crate) struct Connections {
    pub declarations: Vec<Declaration>,
    pub uses: Vec<Use>,
}
