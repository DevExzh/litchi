//! Typed slicer cache and view values.
//!
//! The model intentionally stops at inert metadata. It does not evaluate
//! PivotCaches, apply selections, refresh data, or calculate cross-filter
//! results.

use crate::package::error::{Error, Result};

pub(crate) const MAX_CACHES: usize = 4096;
pub(crate) const MAX_VIEWS: usize = 16_384;
pub(crate) const MAX_ITEMS: usize = 1_000_000;
pub(crate) const MAX_PIVOT_TABLES: usize = 65_536;
pub(crate) const MAX_NAME_UNITS: usize = 32_767;
pub(crate) const MAX_CACHE_NAME_UNITS: usize = 255;

/// Slicer item ordering from `fSortOrder`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SortOrder {
    /// Ascending order.
    Ascending,
    /// Descending order.
    Descending,
}

impl SortOrder {
    pub(crate) fn from_wire(value: u8, field: &'static str) -> Result<Self> {
        match value {
            1 => Ok(Self::Ascending),
            2 => Ok(Self::Descending),
            _ => Err(Error::Unrecognized {
                typ: field.to_string(),
                val: value.to_string(),
            }),
        }
    }

    pub(crate) const fn wire(self) -> u8 {
        match self {
            Self::Ascending => 1,
            Self::Descending => 2,
        }
    }
}

impl Default for SortOrder {
    fn default() -> Self {
        Self::Ascending
    }
}

/// Slicer cross-filter display mode from `fCrossFilter`/`iCrossFilter`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CrossFilter {
    /// Do not style or separately sort items without data.
    None,
    /// Style items without data and sort them at the bottom.
    ShowItemsWithoutData,
    /// Style items without data without separately sorting them.
    ShowItemsWithoutDataUnsorted,
}

impl CrossFilter {
    pub(crate) fn from_wire(value: u8, field: &'static str) -> Result<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::ShowItemsWithoutData),
            2 => Ok(Self::ShowItemsWithoutDataUnsorted),
            _ => Err(Error::Unrecognized {
                typ: field.to_string(),
                val: value.to_string(),
            }),
        }
    }

    pub(crate) const fn wire(self) -> u8 {
        match self {
            Self::None => 0,
            Self::ShowItemsWithoutData => 1,
            Self::ShowItemsWithoutDataUnsorted => 2,
        }
    }
}

impl Default for CrossFilter {
    fn default() -> Self {
        Self::None
    }
}

/// A non-OLAP cache item. The index addresses the associated PivotCache item.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Item {
    /// Zero-based PivotCache item index.
    pub cache_index: u32,
    /// Whether this item is selected by the inert snapshot.
    pub selected: bool,
    /// Whether the item has no data in the associated PivotCache.
    pub no_data: bool,
}

impl Item {
    /// Construct an unselected cache item.
    #[must_use]
    pub const fn new(cache_index: u32) -> Self {
        Self {
            cache_index,
            selected: false,
            no_data: false,
        }
    }
}

/// Native/non-OLAP slicer cache source.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Native {
    /// Associated non-OLAP PivotCache identifier.
    pub cache_id: u32,
    /// Item ordering in the view.
    pub sort_order: SortOrder,
    /// Cross-filter display policy.
    pub cross_filter: CrossFilter,
    /// Whether Excel custom lists participate in sorting.
    pub sort_using_custom_lists: bool,
    /// Whether unused PivotCache items are shown.
    pub show_all_items: bool,
    /// Cached item selection state.
    pub items: Vec<Item>,
}

impl Native {
    /// Construct an empty native cache source.
    #[must_use]
    pub fn new(cache_id: u32) -> Self {
        Self {
            cache_id,
            sort_order: SortOrder::Ascending,
            cross_filter: CrossFilter::None,
            sort_using_custom_lists: false,
            show_all_items: false,
            items: Vec::new(),
        }
    }
}

/// Table-backed slicer cache source.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Table {
    /// Associated table column identifier.
    pub column: u32,
    /// Associated table identifier.
    pub table_id: u32,
    /// Item ordering in the view.
    pub sort_order: SortOrder,
    /// Cross-filter display policy.
    pub cross_filter: CrossFilter,
    /// Whether Excel custom lists participate in sorting.
    pub sort_using_custom_lists: bool,
}

impl Table {
    /// Construct a table-backed source.
    #[must_use]
    pub const fn new(column: u32, table_id: u32) -> Self {
        Self {
            column,
            table_id,
            sort_order: SortOrder::Ascending,
            cross_filter: CrossFilter::None,
            sort_using_custom_lists: false,
        }
    }
}

/// Bounded marker for an OLAP cache that this slice does not author.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Olap {
    /// Associated PivotCache identifier.
    pub pivot_cache_id: u32,
}

/// The source kind owned by a cache definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Source {
    /// Non-OLAP PivotCache source.
    Native(Native),
    /// Structured table source.
    Table(Table),
    /// OLAP source metadata. OLAP nested item ranges remain read-only and
    /// are rejected by the bounded authoring codec.
    Olap(Olap),
}

/// A PivotTable view associated with a slicer cache.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotTable {
    /// Worksheet tab identifier.
    pub sheet_id: u32,
    /// PivotTable view name.
    pub name: String,
}

/// A BIFF12 slicer cache definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Cache {
    /// Unique cache name.
    pub name: String,
    /// Cache field or MDX hierarchy name.
    pub hierarchy: String,
    /// PivotTable views using this cache.
    pub pivot_tables: Vec<PivotTable>,
    /// Source-specific cache state.
    pub source: Source,
}

impl Cache {
    /// Construct a native cache definition.
    #[must_use]
    pub fn native(name: impl Into<String>, hierarchy: impl Into<String>, source: Native) -> Self {
        Self {
            name: name.into(),
            hierarchy: hierarchy.into(),
            pivot_tables: Vec::new(),
            source: Source::Native(source),
        }
    }

    /// Construct a table-backed cache definition.
    #[must_use]
    pub fn table(name: impl Into<String>, hierarchy: impl Into<String>, source: Table) -> Self {
        Self {
            name: name.into(),
            hierarchy: hierarchy.into(),
            pivot_tables: Vec::new(),
            source: Source::Table(source),
        }
    }

    /// Construct an inert OLAP cache descriptor for read-only inspection.
    #[must_use]
    pub fn olap(
        name: impl Into<String>,
        hierarchy: impl Into<String>,
        pivot_cache_id: u32,
    ) -> Self {
        Self {
            name: name.into(),
            hierarchy: hierarchy.into(),
            pivot_tables: Vec::new(),
            source: Source::Olap(Olap { pivot_cache_id }),
        }
    }
}

/// One worksheet slicer view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct View {
    /// Unique workbook-scoped view name.
    pub name: String,
    /// Referenced cache name.
    pub cache_name: String,
    /// Zero-based first visible item.
    pub start_item: u32,
    /// Number of display columns.
    pub column_count: u32,
    /// OLAP level; zero for native/table sources.
    pub level: u32,
    /// Row height in EMUs.
    pub row_height: u32,
    /// Whether the caption is displayed.
    pub caption_visible: bool,
    /// Optional explicit caption.
    pub caption: Option<String>,
    /// Optional slicer style name.
    pub style: Option<String>,
    /// Whether the view position is locked.
    pub locked_position: bool,
}

impl View {
    /// Construct a native/table view with default layout flags.
    #[must_use]
    pub fn new(name: impl Into<String>, cache_name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            cache_name: cache_name.into(),
            start_item: 0,
            column_count: 1,
            level: 0,
            row_height: 0,
            caption_visible: true,
            caption: None,
            style: None,
            locked_position: false,
        }
    }
}

/// The slicer views stored in one worksheet slicers part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Views {
    /// Views in stream order.
    pub items: Vec<View>,
}

impl Views {
    /// Construct an empty worksheet view collection.
    #[must_use]
    pub const fn new() -> Self {
        Self { items: Vec::new() }
    }
}

pub(crate) fn validate_name(value: &str, field: &'static str, max_units: usize) -> Result<()> {
    let units = value.encode_utf16().count();
    if units == 0 || units > max_units || value.contains('\0') {
        return Err(Error::Unrecognized {
            typ: field.to_string(),
            val: format!("invalid UTF-16 length or NUL ({} units)", units),
        });
    }
    Ok(())
}

pub(crate) fn validate_cache(value: &Cache) -> Result<()> {
    validate_name(&value.name, "slicer cache name", MAX_CACHE_NAME_UNITS)?;
    validate_name(&value.hierarchy, "slicer cache hierarchy", MAX_NAME_UNITS)?;
    if value.pivot_tables.len() > MAX_PIVOT_TABLES {
        return Err(Error::InvalidLength {
            expected: MAX_PIVOT_TABLES,
            found: value.pivot_tables.len(),
        });
    }
    let mut pivot_keys = std::collections::HashSet::with_capacity(value.pivot_tables.len());
    for pivot in &value.pivot_tables {
        validate_name(&pivot.name, "slicer PivotTable name", MAX_NAME_UNITS)?;
        if !pivot_keys.insert((pivot.sheet_id, pivot.name.to_ascii_lowercase())) {
            return Err(Error::Unrecognized {
                typ: "slicer PivotTable collection".to_string(),
                val: "duplicate PivotTable reference".to_string(),
            });
        }
    }
    match &value.source {
        Source::Native(native) => {
            if native.items.is_empty() || native.items.len() > MAX_ITEMS {
                return Err(Error::InvalidLength {
                    expected: MAX_ITEMS,
                    found: native.items.len(),
                });
            }
            if !native.items.iter().any(|item| item.selected) {
                return Err(Error::Unrecognized {
                    typ: "BrtSlicerCacheNativeItem".to_string(),
                    val: "at least one item must be selected".to_string(),
                });
            }
            let mut indices = std::collections::HashSet::with_capacity(native.items.len());
            for item in &native.items {
                if !indices.insert(item.cache_index) {
                    return Err(Error::Unrecognized {
                        typ: "BrtSlicerCacheNativeItem".to_string(),
                        val: format!("duplicate cache index {}", item.cache_index),
                    });
                }
                if native.cross_filter == CrossFilter::None && item.no_data {
                    return Err(Error::Unrecognized {
                        typ: "BrtSlicerCacheNativeItem".to_string(),
                        val: "no-data item requires cross filtering".to_string(),
                    });
                }
            }
        },
        Source::Table(_) | Source::Olap(_) => {},
    }
    Ok(())
}

pub(crate) fn validate_views(value: &Views) -> Result<()> {
    if value.items.len() > MAX_VIEWS {
        return Err(Error::InvalidLength {
            expected: MAX_VIEWS,
            found: value.items.len(),
        });
    }
    let mut names = std::collections::HashSet::with_capacity(value.items.len());
    for view in &value.items {
        validate_name(&view.name, "slicer view name", MAX_NAME_UNITS)?;
        validate_name(&view.cache_name, "slicer view cache name", MAX_NAME_UNITS)?;
        if view.column_count == 0 || view.column_count > 20_000 {
            return Err(Error::Unrecognized {
                typ: "BrtBeginSlicer column count".to_string(),
                val: view.column_count.to_string(),
            });
        }
        if !names.insert(view.name.to_ascii_lowercase()) {
            return Err(Error::Unrecognized {
                typ: "slicer view collection".to_string(),
                val: "duplicate view name".to_string(),
            });
        }
        if view.caption.is_some() && view.caption.as_deref().is_some_and(str::is_empty) {
            return Err(Error::Unrecognized {
                typ: "slicer view caption".to_string(),
                val: "caption cannot be empty".to_string(),
            });
        }
        if let Some(caption) = &view.caption {
            validate_name(caption, "slicer view caption", MAX_NAME_UNITS)?;
        }
        if let Some(style) = &view.style {
            validate_name(style, "slicer view style", MAX_NAME_UNITS)?;
        }
    }
    Ok(())
}
