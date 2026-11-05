//! The workbook tab catalog: the kind and visibility of every sheet tab.
//!
//! A workbook lists its tabs in `BrtBundleSh` records. Each record names the
//! tab, carries its visibility (`hsState`) and points at the sheet part
//! through a relationship whose type decides the [`Kind`]. The eager
//! and source-backed owners keep this small semantic catalog for every tab;
//! worksheet-only operations continue to use their worksheet ordinal maps.

use std::borrow::Cow;
use std::fmt;

use litchi_core::Position;
use litchi_opc::constants::relationship_type;

use crate::package::error::{Error, Result};

const CHART_SHEET_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const STRICT_CHART_SHEET_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
const DIALOG_SHEET_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
const STRICT_DIALOG_SHEET_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet";
const MACRO_SHEET_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";
const INTL_MACRO_SHEET_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";

/// The kind of a workbook tab, decided by the relationship from the
/// workbook part to its sheet part.
///
/// Only [`Kind::Worksheet`] tabs have a cell grid that
/// [`crate::Worksheet`] models. The other recognized kinds remain visible in
/// the complete tab catalog and are refused by worksheet-only materializers.
/// A valid very hidden tab can also be [`Kind::Unbound`] when its relationship
/// is absent and it has no physical payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    /// A worksheet: a grid of cells.
    Worksheet,
    /// A chart sheet: one chart that fills the tab.
    Chart,
    /// An Excel 5.0 dialog sheet.
    Dialog,
    /// An Excel 4.0 macro sheet (`xlMacrosheet`).
    Macro,
    /// An Excel 4.0 international macro sheet (`xlIntlMacrosheet`).
    IntlMacro,
    /// A very hidden tab with no relationship to a physical sheet part.
    ///
    /// The semantic tab remains selectable in the complete catalog, while
    /// payload operations return a typed unsupported-feature error.
    Unbound,
}

impl Kind {
    /// Whether this tab is a worksheet, the only kind with a cell grid.
    #[must_use]
    pub const fn is_worksheet(self) -> bool {
        matches!(self, Self::Worksheet)
    }

    /// Whether this tab is an Excel 4.0 macro sheet of either flavour.
    #[must_use]
    pub const fn is_macro(self) -> bool {
        matches!(self, Self::Macro | Self::IntlMacro)
    }

    /// Classify a workbook-to-sheet relationship type, or `None` when the
    /// type names no sheet kind recognized by this crate.
    pub(crate) fn from_relationship_type(reltype: &str) -> Option<Self> {
        match reltype {
            relationship_type::WORKSHEET | relationship_type::STRICT_WORKSHEET => {
                Some(Self::Worksheet)
            },
            CHART_SHEET_RELATIONSHIP | STRICT_CHART_SHEET_RELATIONSHIP => Some(Self::Chart),
            DIALOG_SHEET_RELATIONSHIP | STRICT_DIALOG_SHEET_RELATIONSHIP => Some(Self::Dialog),
            MACRO_SHEET_RELATIONSHIP => Some(Self::Macro),
            INTL_MACRO_SHEET_RELATIONSHIP => Some(Self::IntlMacro),
            _ => None,
        }
    }
}

/// The visibility of a workbook tab, from `BrtBundleSh.hsState`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum State {
    /// The tab is visible.
    Visible,
    /// The tab is hidden but can be shown by the user interface.
    Hidden,
    /// The tab is hidden from the ordinary user interface.
    VeryHidden,
}

impl State {
    /// Whether Excel shows this tab.
    #[must_use]
    pub const fn is_visible(self) -> bool {
        matches!(self, Self::Visible)
    }

    /// Whether this tab is hidden by either hidden state.
    #[must_use]
    pub const fn is_hidden(self) -> bool {
        matches!(self, Self::Hidden | Self::VeryHidden)
    }

    /// Whether this tab is very hidden.
    #[must_use]
    pub const fn is_very_hidden(self) -> bool {
        matches!(self, Self::VeryHidden)
    }

    /// Decode a stored `hsState` value. The wire parser also rejects values
    /// above two; keeping this conversion checked protects callers that use
    /// it when validating a retained parsed record.
    pub(crate) fn from_hs_state(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Visible),
            1 => Ok(Self::Hidden),
            2 => Ok(Self::VeryHidden),
            _other => Err(Error::Unrecognized {
                typ: "BrtBundleSh hsState",
                val: "unexpected value",
            }),
        }
    }
}

/// A semantic selector for a workbook tab.
///
/// Names and checked zero-based positions are the ordinary selection forms.
/// Native relationship and tab identifiers intentionally have no variant in
/// this API; callers that need them must use the explicit raw layer.
#[derive(Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Selector<'a> {
    /// Select by the tab's authored name.
    Name(Cow<'a, str>),
    /// Select by checked zero-based catalog position.
    Position(Position),
}

impl fmt::Debug for Selector<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Name(_) => formatter.debug_tuple("Name").field(&"<redacted>").finish(),
            Self::Position(position) => formatter.debug_tuple("Position").field(position).finish(),
        }
    }
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(Cow::Borrowed(value))
    }
}

impl From<String> for Selector<'_> {
    fn from(value: String) -> Self {
        Self::Name(Cow::Owned(value))
    }
}

impl<'a> From<&'a String> for Selector<'a> {
    fn from(value: &'a String) -> Self {
        Self::Name(Cow::Borrowed(value.as_str()))
    }
}

impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Position(Position::new(value))
    }
}

impl From<Position> for Selector<'_> {
    fn from(value: Position) -> Self {
        Self::Position(value)
    }
}

/// The private semantic metadata retained for one workbook tab.
///
/// The owned name is used by eager and source-backed selectors. Native
/// relationship identifiers and physical part names remain outside this
/// value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct TabMetadata {
    name: String,
    kind: Kind,
    state: State,
}

impl TabMetadata {
    pub(crate) fn new(name: String, kind: Kind, state: State) -> Self {
        Self { name, kind, state }
    }

    pub(crate) fn name(&self) -> &str {
        &self.name
    }

    pub(crate) const fn kind(&self) -> Kind {
        self.kind
    }

    pub(crate) const fn state(&self) -> State {
        self.state
    }
}

/// One workbook tab of any [`Kind`], in workbook order.
///
/// [`Self::position`] is the checked zero-based position among all tabs. The
/// optional [`Self::worksheet_position`] bridges to the worksheet-only APIs and is `None`
/// for chart, dialog, macro, and unbound tabs.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct Tab<'a> {
    name: &'a str,
    position: Position,
    worksheet_position: Option<Position>,
    kind: Kind,
    state: State,
}

impl fmt::Debug for Tab<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        // Sheet names are authored content. Keep them out of Debug output;
        // callers can request the name explicitly through the semantic API.
        formatter
            .debug_struct("Tab")
            .field("position", &self.position)
            .field("worksheet_position", &self.worksheet_position)
            .field("kind", &self.kind)
            .field("state", &self.state)
            .finish()
    }
}

impl<'a> Tab<'a> {
    pub(crate) const fn new(
        name: &'a str,
        position: Position,
        worksheet_position: Option<Position>,
        kind: Kind,
        state: State,
    ) -> Self {
        Self {
            name,
            position,
            worksheet_position,
            kind,
            state,
        }
    }

    /// The tab name, exactly as the workbook stores it.
    #[must_use]
    pub const fn name(&self) -> &'a str {
        self.name
    }

    /// The zero-based position of this tab among all tabs.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// The checked zero-based position of this tab among worksheet tabs, or
    /// `None` for the other recognized kinds.
    #[must_use]
    pub const fn worksheet_position(&self) -> Option<Position> {
        self.worksheet_position
    }

    /// The kind of this tab.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// The visibility of this tab.
    #[must_use]
    pub const fn state(&self) -> State {
        self.state
    }
}

#[cfg(test)]
mod tests {
    use super::{Kind, Position, Selector, State};

    #[test]
    fn relationship_types_classify_every_recognized_kind() {
        for (reltype, kind) in [
            (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                Kind::Worksheet,
            ),
            (
                "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet",
                Kind::Worksheet,
            ),
            (super::CHART_SHEET_RELATIONSHIP, Kind::Chart),
            (super::STRICT_CHART_SHEET_RELATIONSHIP, Kind::Chart),
            (super::DIALOG_SHEET_RELATIONSHIP, Kind::Dialog),
            (super::STRICT_DIALOG_SHEET_RELATIONSHIP, Kind::Dialog),
            (super::MACRO_SHEET_RELATIONSHIP, Kind::Macro),
            (super::INTL_MACRO_SHEET_RELATIONSHIP, Kind::IntlMacro),
        ] {
            assert_eq!(Kind::from_relationship_type(reltype), Some(kind));
        }
        assert_eq!(
            Kind::from_relationship_type("https://attacker.invalid/relationships/chartsheet"),
            None
        );
        assert!(Kind::Worksheet.is_worksheet());
        assert!(!Kind::Chart.is_worksheet());
        assert!(Kind::Macro.is_macro() && Kind::IntlMacro.is_macro());
        assert!(!Kind::Dialog.is_macro());
    }

    #[test]
    fn hs_state_round_trips_and_refuses_unknown_values() {
        for (raw, state) in [
            (0, State::Visible),
            (1, State::Hidden),
            (2, State::VeryHidden),
        ] {
            assert_eq!(State::from_hs_state(raw).unwrap(), state);
        }
        assert!(State::from_hs_state(3).is_err());
        assert!(State::Visible.is_visible() && !State::Visible.is_hidden());
        assert!(State::Hidden.is_hidden() && !State::Hidden.is_very_hidden());
        assert!(State::VeryHidden.is_hidden() && State::VeryHidden.is_very_hidden());
    }

    #[test]
    fn debug_does_not_include_authored_tab_name() {
        let tab = super::Tab::new(
            "private-authored-name",
            Position::new(4),
            None,
            Kind::Chart,
            State::Hidden,
        );
        let debug = format!("{tab:?}");
        assert!(!debug.contains("private-authored-name"));
        assert!(debug.contains("position"));
        assert!(debug.contains("Chart"));
    }

    #[test]
    fn selector_debug_redacts_authored_name() {
        let selector = Selector::from("private-authored-name");
        let debug = format!("{selector:?}");
        assert!(!debug.contains("private-authored-name"));
        assert!(debug.contains("redacted"));
        assert_eq!(
            Selector::from(3_usize),
            Selector::Position(Position::new(3))
        );
    }
}
