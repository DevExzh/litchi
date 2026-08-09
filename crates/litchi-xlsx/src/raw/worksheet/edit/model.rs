//! Semantic selectors and orthogonal worksheet edit effects.

use std::collections::BTreeMap;

use litchi_sheet::{COLUMNS, Cell as Address, Column, ROWS, Rect, Row};

use crate::cell::{Content, Text};
use crate::column::Width;
use crate::error::{Result, invalid};
use crate::layout::{self, Descent};
use crate::outline::Outline;
use crate::row::Height;

#[derive(Debug, PartialEq, Eq)]
pub(crate) enum Payload {
    Set(Content),
    /// Reuse one exact workbook shared-string item, retaining rich runs.
    SharedString {
        index: usize,
        text: Text,
    },
    /// Ensure an explicit empty cell record exists.
    Clear,
    /// Clear only when another effect or the base snapshot retains the cell.
    ClearIfPresent,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum StyleEffect {
    Set(u32),
    Reset,
}

/// Orthogonal effects on one cell record.
///
/// `Remove` owns the whole record. An `Update` may independently change its
/// payload and local style, allowing proven-disjoint effects to be joined.
#[derive(Debug, PartialEq, Eq)]
pub(crate) enum Action {
    Update {
        payload: Option<Payload>,
        style: Option<StyleEffect>,
    },
    Remove,
}

impl Action {
    pub(crate) fn set(content: Content) -> Self {
        Self::Update {
            payload: Some(Payload::Set(content)),
            style: None,
        }
    }

    pub(crate) const fn clear(create: bool) -> Self {
        Self::Update {
            payload: Some(if create {
                Payload::Clear
            } else {
                Payload::ClearIfPresent
            }),
            style: None,
        }
    }

    pub(crate) const fn style(key: u32) -> Self {
        Self::Update {
            payload: None,
            style: Some(StyleEffect::Set(key)),
        }
    }

    pub(crate) const fn reset_style() -> Self {
        Self::Update {
            payload: None,
            style: Some(StyleEffect::Reset),
        }
    }

    pub(crate) const fn payload(&self) -> Option<&Payload> {
        match self {
            Self::Update { payload, .. } => payload.as_ref(),
            Self::Remove => None,
        }
    }

    pub(crate) const fn creates_missing(&self) -> bool {
        match self {
            Self::Update { payload, style } => {
                matches!(
                    payload,
                    Some(Payload::Set(_) | Payload::SharedString { .. } | Payload::Clear)
                ) || matches!(style, Some(StyleEffect::Set(_)))
            },
            Self::Remove => false,
        }
    }

    pub(crate) const fn overlaps(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::Remove, _) | (_, Self::Remove) => true,
            (
                Self::Update {
                    payload: left_payload,
                    style: left_style,
                },
                Self::Update {
                    payload: right_payload,
                    style: right_style,
                },
            ) => {
                (left_payload.is_some() && right_payload.is_some())
                    || (left_style.is_some() && right_style.is_some())
            },
        }
    }

    pub(crate) fn merge(&mut self, other: Self) {
        if let (
            Self::Update { payload, style },
            Self::Update {
                payload: other_payload,
                style: other_style,
            },
        ) = (self, other)
        {
            // `Edit::join` proves these facets disjoint before moving either
            // map. The conditional assignments keep this primitive total and
            // panic-free if an internal caller ever violates that contract.
            if payload.is_none() {
                *payload = other_payload;
            }
            if style.is_none() {
                *style = other_style;
            }
        }
    }

    pub(crate) fn set_payload(&mut self, effect: Payload) {
        *self = match std::mem::replace(self, Self::Remove) {
            Self::Update { style, .. } => Self::Update {
                payload: Some(effect),
                style,
            },
            Self::Remove => Self::Update {
                payload: Some(effect),
                style: None,
            },
        };
    }

    pub(crate) fn set_style(&mut self, effect: StyleEffect) {
        *self = match std::mem::replace(self, Self::Remove) {
            Self::Update { payload, .. } => Self::Update {
                payload,
                style: Some(effect),
            },
            Self::Remove => Self::Update {
                payload: None,
                style: Some(effect),
            },
        };
    }
}

/// One checked height mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum HeightEffect {
    Set(Height),
    Reset,
}

/// One checked typographic-descent mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum DescentEffect {
    Set(Descent),
    Reset,
}

/// Orthogonal effects on one stored row record.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct RowAction {
    pub(crate) hidden: Option<bool>,
    pub(crate) height: Option<HeightEffect>,
    pub(crate) descent: Option<DescentEffect>,
    pub(crate) style: Option<StyleEffect>,
    pub(crate) outline: Option<Outline>,
    pub(crate) collapsed: Option<bool>,
    pub(crate) thick_top: Option<bool>,
    pub(crate) thick_bottom: Option<bool>,
    pub(crate) phonetic: Option<bool>,
}

impl RowAction {
    #[cfg(test)]
    pub(crate) const fn hide() -> Self {
        Self {
            hidden: Some(true),
            height: None,
            descent: None,
            style: None,
            outline: None,
            collapsed: None,
            thick_top: None,
            thick_bottom: None,
            phonetic: None,
        }
    }

    #[cfg(test)]
    pub(crate) const fn show() -> Self {
        Self {
            hidden: Some(false),
            height: None,
            descent: None,
            style: None,
            outline: None,
            collapsed: None,
            thick_top: None,
            thick_bottom: None,
            phonetic: None,
        }
    }

    pub(crate) const fn materializes(self) -> bool {
        matches!(self.hidden, Some(true))
            || matches!(self.height, Some(HeightEffect::Set(_)))
            || matches!(self.descent, Some(DescentEffect::Set(_)))
            || matches!(self.style, Some(StyleEffect::Set(_)))
            || matches!(self.outline, Some(level) if level.get() != 0)
            || matches!(self.collapsed, Some(true))
            || matches!(self.thick_top, Some(true))
            || matches!(self.thick_bottom, Some(true))
            || matches!(self.phonetic, Some(true))
    }

    pub(crate) const fn overlaps(self, other: Self) -> bool {
        (self.hidden.is_some() && other.hidden.is_some())
            || (self.height.is_some() && other.height.is_some())
            || (self.descent.is_some() && other.descent.is_some())
            || (self.style.is_some() && other.style.is_some())
            || (self.outline.is_some() && other.outline.is_some())
            || (self.collapsed.is_some() && other.collapsed.is_some())
            || (self.thick_top.is_some() && other.thick_top.is_some())
            || (self.thick_bottom.is_some() && other.thick_bottom.is_some())
            || (self.phonetic.is_some() && other.phonetic.is_some())
    }

    pub(crate) fn merge(&mut self, other: Self) {
        if self.hidden.is_none() {
            self.hidden = other.hidden;
        }
        if self.height.is_none() {
            self.height = other.height;
        }
        if self.descent.is_none() {
            self.descent = other.descent;
        }
        if self.style.is_none() {
            self.style = other.style;
        }
        if self.outline.is_none() {
            self.outline = other.outline;
        }
        if self.collapsed.is_none() {
            self.collapsed = other.collapsed;
        }
        if self.thick_top.is_none() {
            self.thick_top = other.thick_top;
        }
        if self.thick_bottom.is_none() {
            self.thick_bottom = other.thick_bottom;
        }
        if self.phonetic.is_none() {
            self.phonetic = other.phonetic;
        }
    }
}

/// Set or remove one optional worksheet-default value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum OptionalEffect<T> {
    Set(T),
    Reset,
}

/// Orthogonal effects on the worksheet's `sheetFormatPr` record.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct DefaultsEffects {
    pub(crate) base_width: Option<OptionalEffect<u8>>,
    pub(crate) width: Option<OptionalEffect<layout::Width>>,
    pub(crate) height: Option<layout::Height>,
    pub(crate) hidden: Option<bool>,
    pub(crate) thick_top: Option<bool>,
    pub(crate) thick_bottom: Option<bool>,
    pub(crate) descent: Option<DescentEffect>,
}

impl DefaultsEffects {
    pub(crate) fn fields(self) -> layout::Fields {
        let mut fields = layout::Fields::empty();
        if self.base_width.is_some() {
            fields.insert(layout::Fields::BASE_WIDTH);
        }
        if self.width.is_some() {
            fields.insert(layout::Fields::WIDTH);
        }
        if self.height.is_some() {
            fields.insert(layout::Fields::HEIGHT);
        }
        if self.hidden.is_some() {
            fields.insert(layout::Fields::HIDDEN);
        }
        if self.thick_top.is_some() {
            fields.insert(layout::Fields::THICK_TOP);
        }
        if self.thick_bottom.is_some() {
            fields.insert(layout::Fields::THICK_BOTTOM);
        }
        if self.descent.is_some() {
            fields.insert(layout::Fields::DESCENT);
        }
        fields
    }

    pub(crate) const fn materializes(self) -> bool {
        matches!(self.base_width, Some(OptionalEffect::Set(_)))
            || matches!(self.width, Some(OptionalEffect::Set(_)))
            || self.height.is_some()
            || matches!(self.hidden, Some(true))
            || matches!(self.thick_top, Some(true))
            || matches!(self.thick_bottom, Some(true))
            || matches!(self.descent, Some(DescentEffect::Set(_)))
    }

    fn merge(&mut self, other: Self) {
        if self.base_width.is_none() {
            self.base_width = other.base_width;
        }
        if self.width.is_none() {
            self.width = other.width;
        }
        if self.height.is_none() {
            self.height = other.height;
        }
        if self.hidden.is_none() {
            self.hidden = other.hidden;
        }
        if self.thick_top.is_none() {
            self.thick_top = other.thick_top;
        }
        if self.thick_bottom.is_none() {
            self.thick_bottom = other.thick_bottom;
        }
        if self.descent.is_none() {
            self.descent = other.descent;
        }
    }
}

/// Whole-record deletion or facet-level worksheet-default updates.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct DefaultsAction {
    remove: bool,
    effects: DefaultsEffects,
}

impl DefaultsAction {
    pub(crate) fn update(&mut self) -> &mut DefaultsEffects {
        self.remove = false;
        &mut self.effects
    }

    pub(crate) const fn remove() -> Self {
        Self {
            remove: true,
            effects: DefaultsEffects {
                base_width: None,
                width: None,
                height: None,
                hidden: None,
                thick_top: None,
                thick_bottom: None,
                descent: None,
            },
        }
    }

    pub(crate) const fn is_remove(self) -> bool {
        self.remove
    }

    pub(crate) const fn effects(self) -> DefaultsEffects {
        self.effects
    }

    pub(crate) fn fields(self) -> layout::Fields {
        if self.remove {
            layout::Fields::all()
        } else {
            self.effects.fields()
        }
    }

    pub(crate) const fn materializes(self) -> bool {
        !self.remove && self.effects.materializes()
    }

    pub(crate) fn overlaps(self, other: Self) -> bool {
        self.remove || other.remove || self.fields().intersects(other.fields())
    }

    pub(crate) fn merge(&mut self, other: Self) {
        if !self.remove && !other.remove {
            self.effects.merge(other.effects);
        }
    }
}

/// One checked width mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum WidthEffect {
    Set(Width),
    Reset,
}

/// Orthogonal effects on one effective column-property record.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct ColumnAction {
    pub(crate) hidden: Option<bool>,
    pub(crate) width: Option<WidthEffect>,
    pub(crate) style: Option<StyleEffect>,
    pub(crate) best_fit: Option<bool>,
    pub(crate) outline: Option<Outline>,
    pub(crate) collapsed: Option<bool>,
    pub(crate) phonetic: Option<bool>,
}

impl ColumnAction {
    #[cfg(test)]
    pub(crate) const fn hide() -> Self {
        Self {
            hidden: Some(true),
            width: None,
            style: None,
            best_fit: None,
            outline: None,
            collapsed: None,
            phonetic: None,
        }
    }

    #[cfg(test)]
    pub(crate) const fn show() -> Self {
        Self {
            hidden: Some(false),
            width: None,
            style: None,
            best_fit: None,
            outline: None,
            collapsed: None,
            phonetic: None,
        }
    }

    pub(crate) const fn materializes(self) -> bool {
        matches!(self.hidden, Some(true))
            || matches!(self.width, Some(WidthEffect::Set(_)))
            || matches!(self.style, Some(StyleEffect::Set(_)))
            || matches!(self.best_fit, Some(true))
            || matches!(self.outline, Some(level) if level.get() != 0)
            || matches!(self.collapsed, Some(true))
            || matches!(self.phonetic, Some(true))
    }

    pub(crate) const fn overlaps(self, other: Self) -> bool {
        (self.hidden.is_some() && other.hidden.is_some())
            || (self.width.is_some() && other.width.is_some())
            || (self.style.is_some() && other.style.is_some())
            || (self.best_fit.is_some() && other.best_fit.is_some())
            || (self.outline.is_some() && other.outline.is_some())
            || (self.collapsed.is_some() && other.collapsed.is_some())
            || (self.phonetic.is_some() && other.phonetic.is_some())
    }

    pub(crate) fn merge(&mut self, other: Self) {
        if self.hidden.is_none() {
            self.hidden = other.hidden;
        }
        if self.width.is_none() {
            self.width = other.width;
        }
        if self.style.is_none() {
            self.style = other.style;
        }
        if self.best_fit.is_none() {
            self.best_fit = other.best_fit;
        }
        if self.outline.is_none() {
            self.outline = other.outline;
        }
        if self.collapsed.is_none() {
            self.collapsed = other.collapsed;
        }
        if self.phonetic.is_none() {
            self.phonetic = other.phonetic;
        }
    }
}

/// Move-only worksheet rewrite plan with orthogonal cell and row facets.
#[derive(Debug, Default)]
pub(crate) struct Plan {
    pub(crate) defaults: Option<DefaultsAction>,
    pub(crate) cells: BTreeMap<Address, Action>,
    pub(crate) rows: BTreeMap<Row, RowAction>,
    pub(crate) columns: BTreeMap<Column, ColumnAction>,
}

impl Plan {
    pub(crate) fn cells(cells: BTreeMap<Address, Action>) -> Self {
        Self {
            defaults: None,
            cells,
            rows: BTreeMap::new(),
            columns: BTreeMap::new(),
        }
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.defaults.is_none()
            && self.cells.is_empty()
            && self.rows.is_empty()
            && self.columns.is_empty()
    }
}

impl From<BTreeMap<Address, Action>> for Plan {
    fn from(value: BTreeMap<Address, Action>) -> Self {
        Self::cells(value)
    }
}

/// Final add/remove effects for one worksheet merge container.
#[derive(Debug, Default)]
pub(crate) struct MergePlan {
    pub(crate) add: Vec<Rect>,
    pub(crate) remove: Vec<Rect>,
}

impl MergePlan {
    pub(crate) fn is_empty(&self) -> bool {
        self.add.is_empty() && self.remove.is_empty()
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct SelectionRange {
    pub(super) first_row: u32,
    pub(super) first_column: u32,
    pub(super) last_row: u32,
    pub(super) last_column: u32,
}

impl SelectionRange {
    pub(super) fn from_rect(range: Rect) -> Self {
        let (end_row, end_column) = range.end();
        Self {
            first_row: range.start().row().get(),
            first_column: range.start().column().get(),
            last_row: end_row - 1,
            last_column: end_column - 1,
        }
    }

    pub(super) fn cell_or_area(value: &str) -> Result<Self> {
        let (first, last) = value.split_once(':').unwrap_or((value, value));
        if last.contains(':') {
            return Err(invalid(format!("invalid cell range '{value}'")));
        }
        let first = Address::from_a1(first)?;
        let last = Address::from_a1(last)?;
        if first.row() > last.row() || first.column() > last.column() {
            return Err(invalid(format!("reversed cell range '{value}'")));
        }
        Ok(Self {
            first_row: first.row().get(),
            first_column: first.column().get(),
            last_row: last.row().get(),
            last_column: last.column().get(),
        })
    }

    pub(super) fn selection(value: &str) -> Result<Self> {
        if let Ok(range) = Self::cell_or_area(value) {
            return Ok(range);
        }
        let (first, last) = value
            .split_once(':')
            .ok_or_else(|| invalid(format!("invalid selection range '{value}'")))?;
        if first
            .bytes()
            .all(|byte| byte == b'$' || byte.is_ascii_alphabetic())
            && last
                .bytes()
                .all(|byte| byte == b'$' || byte.is_ascii_alphabetic())
        {
            let first = column(first)?;
            let last = column(last)?;
            if first > last {
                return Err(invalid(format!("reversed column range '{value}'")));
            }
            return Ok(Self {
                first_row: 0,
                first_column: first,
                last_row: ROWS - 1,
                last_column: last,
            });
        }
        if first
            .bytes()
            .all(|byte| byte == b'$' || byte.is_ascii_digit())
            && last
                .bytes()
                .all(|byte| byte == b'$' || byte.is_ascii_digit())
        {
            let first = row(first)?;
            let last = row(last)?;
            if first > last {
                return Err(invalid(format!("reversed row range '{value}'")));
            }
            return Ok(Self {
                first_row: first,
                first_column: 0,
                last_row: last,
                last_column: COLUMNS - 1,
            });
        }
        Err(invalid(format!("invalid selection range '{value}'")))
    }

    pub(super) fn contains(self, address: Address) -> bool {
        (self.first_row..=self.last_row).contains(&address.row().get())
            && (self.first_column..=self.last_column).contains(&address.column().get())
    }

    pub(super) fn starts_at(self, address: Address) -> bool {
        self.first_row == address.row().get() && self.first_column == address.column().get()
    }

    pub(super) fn overlaps(self, range: Rect) -> bool {
        let (end_row, end_column) = range.end();
        self.first_row < end_row
            && range.start().row().get() <= self.last_row
            && self.first_column < end_column
            && range.start().column().get() <= self.last_column
    }
}

// Column and row selector parsing is kept with the selector model.
fn column(value: &str) -> Result<u32> {
    let value = value.trim_start_matches('$');
    let mut column = 0u32;
    for byte in value.bytes() {
        if !byte.is_ascii_alphabetic() {
            return Err(invalid(format!("invalid column reference '{value}'")));
        }
        column = column
            .checked_mul(26)
            .and_then(|column| column.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1)))
            .ok_or_else(|| invalid(format!("column reference '{value}' overflows")))?;
    }
    Column::new(
        column
            .checked_sub(1)
            .ok_or_else(|| invalid(format!("invalid column reference '{value}'")))?,
    )
    .map(Column::get)
    .map_err(Into::into)
}

fn row(value: &str) -> Result<u32> {
    let value = value.trim_start_matches('$');
    let row = value
        .parse::<u32>()
        .ok()
        .and_then(|row| row.checked_sub(1))
        .ok_or_else(|| invalid(format!("invalid row reference '{value}'")))?;
    Row::new(row).map(Row::get).map_err(Into::into)
}
