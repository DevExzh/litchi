//! Workbook-owned semantic state used by formula resolution.

use std::sync::Arc;

use crate::external_link::Link;
use crate::package::error::Result;
use crate::package::formula::{Definition, ExternalSheet, Scope, View};

use super::relationships::SupportingLink;

/// Stable identity for the PivotTable scope attached to one formula.
///
/// Keeping this key typed prevents cache, worksheet, and view coordinates
/// from being mixed accidentally while keeping the key allocation-free after
/// the context is built.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct PivotScopeKey {
    cache_id: u32,
    sheet_index: usize,
    view_name: String,
}

impl PivotScopeKey {
    fn from_scope(scope: &Scope) -> Self {
        Self {
            cache_id: scope.cache_id,
            sheet_index: scope.sheet_index,
            view_name: scope.view_name.clone(),
        }
    }

    pub(super) fn cache_id(&self) -> u32 {
        self.cache_id
    }

    pub(super) fn sheet_index(&self) -> usize {
        self.sheet_index
    }

    pub(super) fn view_name(&self) -> &str {
        &self.view_name
    }
}

/// Cached metadata for one external workbook relationship.
#[derive(Debug, Clone, PartialEq)]
pub struct ExternalBook {
    pub(crate) metadata: Link,
}

impl ExternalBook {
    pub(crate) fn metadata(&self) -> Link {
        self.metadata.clone()
    }

    pub(crate) fn metadata_ref(&self) -> &Link {
        &self.metadata
    }
}

/// Workbook data required to render context-dependent XLSB formula tokens.
///
/// A workbook owns one instance and worksheet decoders borrow it while
/// rendering formulas. The immutable collections are reference counted so
/// deriving a sheet- or PivotTable-specific view does not copy metadata.
#[derive(Debug, Clone, Default)]
pub struct Context {
    pub(crate) worksheet_names: Arc<[String]>,
    pub(crate) supporting_links: Arc<[SupportingLink]>,
    pub(crate) external_sheets: Arc<[ExternalSheet]>,
    pub(crate) external_books: Arc<[ExternalBook]>,
    pub(crate) defined_names: Arc<[String]>,
    pub(crate) tables: Arc<[Definition]>,
    pub(crate) pivot_views: Arc<[View]>,
    pub(crate) pivot_name_scopes: Arc<[Scope]>,
    pub(crate) active_pivot_scope: Option<PivotScopeKey>,
    pub(crate) current_sheet: Option<usize>,
}

impl Context {
    /// Bind formula rendering to one consuming worksheet without copying
    /// workbook metadata.
    pub(crate) fn for_sheet(&self, sheet_index: usize) -> Self {
        let mut context = self.clone();
        context.current_sheet = Some(sheet_index);
        context
    }

    /// Bind formula-local `BrtBeginPName` metadata to an exact PivotTable
    /// view, validating the relationship before the context is returned.
    pub fn for_pivot_formula(&self, scope: Scope) -> Result<Self> {
        let mut context = self.clone();
        context.current_sheet = Some(scope.sheet_index);
        context.active_pivot_scope = Some(PivotScopeKey::from_scope(&scope));
        context.pivot_name_scopes = vec![scope].into();
        context.validate_active_pivot_scope()?;
        Ok(context)
    }
}
