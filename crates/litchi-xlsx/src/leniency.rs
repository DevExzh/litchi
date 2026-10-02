//! Repair reporting for non-structural defects in `SpreadsheetML` workbooks.
//!
//! Two kinds of repair can happen while a workbook is read, and both are
//! *recorded*, so a caller can enumerate exactly what changed through
//! [`ToleranceReport`]. Nothing is swallowed silently, and nothing outside
//! the closed enumerations changes: ZIP and XML well-formedness, OPC package
//! grammar, worksheet framing, bounded resource limits, and encryption remain
//! hard errors under every variant.
//!
//! **Excel-compatible repairs** are applied under every [`Leniency`] variant,
//! including the default [`Leniency::Strict`]. They cover markup ECMA-376
//! constrains but real producers emit anyway — markup Excel and LibreOffice
//! read past (or silently repair) without failing the workbook, so rejecting
//! it would cost the caller a whole sheet over a stale reference that cannot
//! change what the conforming remainder of the file means. Each substitutes
//! the value documented on the corresponding defect variant.
//!
//! **Opt-in repairs** trade strict conformance for reach and stay gated by
//! [`Leniency`] variants. None exist yet, so [`Leniency::Strict`] is the only
//! variant; the judgment-call class on the roadmap (audit finding 4.8,
//! sheet-name authoring rules enforced at load time) joins here as an
//! explicit variant when it lands.
//!
//! The defect enumerations are sized to the litchi-xlsx section of
//! `docs/LENIENCY_AUDIT.md` (findings 4.1–4.12). A variant whose recording
//! site has not landed yet is documented as a *roadmap contract*: it fixes
//! the repair the finding will apply, and no read produces it today.
//!
//! # Where recording sites plug in
//!
//! This crate opens lazily: [`crate::Workbook`] and
//! [`crate::SourceBackedWorkbook`] validate the OPC graph and parse the
//! workbook catalog when opened, while worksheet, shared-string, and style
//! payloads are parsed on first use behind `&self` handles and published
//! through write-once caches. The report is therefore assembled at three
//! levels, with explicit `&mut ToleranceReport` threading inside each parse
//! and no interior mutability:
//!
//! * **Open time (workbook level).** Each workbook handle owns a report built
//!   while the catalog was parsed. Catalog findings (4.7 `activeTab` /
//!   `localSheetId`, 4.8 sheet names) thread `&mut ToleranceReport` into the
//!   catalog parse and record [`CatalogDefect`]s located at
//!   [`Location::Workbook`] or [`Location::DefinedName`].
//! * **Per worksheet.** A worksheet's parsed cell store carries its own
//!   report, filled while that store was built: the streaming worksheet
//!   parser threads `&mut ToleranceReport` through cell materialization
//!   (findings 4.3, 4.4, 4.6, 4.9 record [`CellDefect`]s at
//!   [`Location::Cell`]; 4.1 and 4.10 record [`WorksheetDefect`]s), and the
//!   style cross-reference pass that runs before the store is published
//!   (finding 4.2) takes the store mutably and records at the cell, row, or
//!   column it repaired. Because the report
//!   lives *in* the store, every path that builds one — owned workbook,
//!   source-backed workbook, edit and readback planning — carries the repairs
//!   it made, and a cached store reports the same repairs on every access.
//!   `Worksheet::tolerance_report` / `SourceWorksheet::tolerance_report`
//!   force that sheet's parse and return its report; per-sheet locations
//!   leave [`ToleratedDefect::sheet`] unset.
//! * **Workbook aggregate.** `Workbook::tolerance_report` /
//!   `SourceBackedWorkbook::tolerance_report` merge the open-time report with
//!   the report of every worksheet parsed *so far*, in sheet order, stamping
//!   each entry with its sheet position in [`ToleratedDefect::sheet`]. Sheets
//!   that were never read have not been parsed, so they cannot have been
//!   repaired; reading every sheet first yields the complete report.
//!   Shared-string repairs (4.5, 4.12) are recorded when the table is parsed
//!   and join the aggregate the same way, at [`Location::SharedStringTable`]
//!   or [`Location::SharedStringItem`].
//!
//! The source-backed facade also answers cold cell reads through a bounded
//! selected-range scanner that never builds a store. That scanner does not
//! repair anything: whenever a dependency it needs is irregular — a
//! shared-string index the table does not hold, a style index past the
//! catalog — it falls back to the authoritative store parse, which repairs
//! and records. Selected reads therefore never observe a defect the report
//! does not list.
//!
//! # Carrier status
//!
//! The plumbing is in place: both workbook handles carry an open-time
//! report, every worksheet store carries the report of the parse that built
//! it, and the aggregates above expose them. Live recording sites:
//!
//! * finding 4.2 — `Store::repair_style_indexes`, run by both read facades
//!   before a worksheet store is published, records
//!   [`CellDefect::StyleIndexOutOfRange`] at the cell, row, or column;
//! * finding 4.4 — cell materialization (`raw/worksheet/semantic.rs`)
//!   records [`CellDefect::SharedStringIndexOutOfRange`] and
//!   [`CellDefect::MissingSharedStringTable`] at the cell.
//!
//! Write paths stay strict: the value-only cell editor parses with repairs
//! refused and keeps its historical errors, and commit-time readback
//! verification still refuses stale style indexes, so an edit never makes a
//! repair durable. Every other variant below is a roadmap contract.
//!
//! This mirrors the contract `crates/litchi-doc/src/leniency.rs` establishes
//! for legacy Word documents and `crates/litchi-xls/src/leniency.rs` for
//! legacy workbooks, so the readers behave the same way.

use litchi_sheet::{Cell as Address, Column as ColumnIndex, Row as RowIndex};

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous workbook must not be able to make the report
/// grow without limit. Defects beyond this bound are counted, not stored; see
/// [`ToleranceReport::unrecorded`].
const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural workbook defects.
///
/// Structural defects — malformed ZIP or XML framing, a broken OPC package
/// grammar, unsupported/absent decryption material, or an exceeded resource
/// limit — are hard errors under every variant. The Excel-compatible repairs
/// listed by [`Defect`] are applied and recorded under every variant,
/// including [`Leniency::Strict`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Leniency {
    /// Reject every defect whose repair requires opting in.
    ///
    /// This is the default and currently the only variant: the
    /// Excel-compatible repairs are applied and recorded here too, and no
    /// opt-in gate exists yet.
    #[default]
    Strict,
}

/// A per-cell defect repaired the way Excel repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// cell-level repairs a read applies.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CellDefect {
    /// A cell's `s` attribute (or a row's / column's style) names an index at
    /// or beyond the workbook's `cellXfs` count (audit finding 4.2).
    ///
    /// ECMA-376 §18.3.1.4 requires the index to reference `cellXfs` but gives
    /// no basis for rejecting the sheet over a stale one; Excel substitutes
    /// the default format and LibreOffice clamps. The reader drops the stale
    /// index, so the cell, row, or column reads with the mandatory default
    /// format (`cellXfs` entry 0). `MAX_CELL_STYLE` remains an allocation
    /// guard and is still a hard error.
    StyleIndexOutOfRange,
    /// A `t="s"` cell names a shared-string index at or beyond the table's
    /// length (audit finding 4.4).
    ///
    /// Excel repairs the cell; LibreOffice resolves the lookup to empty. The
    /// reader substitutes an empty string, matching the LibreOffice result.
    /// An index that is not an unsigned integer at all stays a hard error.
    SharedStringIndexOutOfRange,
    /// A `t="s"` cell appears in a workbook that has no shared-string part
    /// (audit finding 4.4).
    ///
    /// The reader substitutes an empty string for each such cell, exactly as
    /// for an out-of-range index into an empty table.
    MissingSharedStringTable,
    /// A number `<v>` does not parse as a finite `xsd:double` (audit finding
    /// 4.3). *Roadmap contract:* the cell is demoted to `Cell::Unknown`
    /// preserving the lexical, instead of failing the sheet.
    MalformedNumber,
    /// A boolean `<v>` is outside `{0, 1, true, false}` (audit finding 4.3).
    /// *Roadmap contract:* parsed case-insensitively, with any non-zero
    /// number read as true, as LibreOffice does.
    MalformedBoolean,
    /// A `t="d"` value is outside the accepted ISO 8601 layouts (audit
    /// finding 4.3). *Roadmap contract:* demoted to text.
    MalformedDate,
    /// A shared-formula group is structurally inconsistent — master without
    /// `ref` or expression, duplicate master, follower without master
    /// (audit finding 4.6). *Roadmap contract:* the master text is kept
    /// verbatim and orphaned followers read with an `Unknown` formula
    /// instead of a translated one; the location is the follower cell.
    BrokenSharedFormulaGroup,
    /// A scalar or array `<f>` starts with `=` (audit finding 4.9).
    /// *Roadmap contract:* one leading `=` is stripped.
    FormulaLeadingEquals,
    /// A scalar or array `<f>` is empty (audit finding 4.9). *Roadmap
    /// contract:* the cell reads without a formula.
    EmptyFormula,
}

impl CellDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::StyleIndexOutOfRange => {
                "dropped the stale style index; the default cell format (cellXfs 0) applies"
            },
            Self::SharedStringIndexOutOfRange => {
                "substituted an empty string for the out-of-range shared-string reference"
            },
            Self::MissingSharedStringTable => {
                "substituted an empty string; the workbook has no shared-string table"
            },
            Self::MalformedNumber => "demoted the unparseable number to an unknown cell",
            Self::MalformedBoolean => "read the nonstandard boolean lexical leniently",
            Self::MalformedDate => "demoted the unparseable date to text",
            Self::BrokenSharedFormulaGroup => {
                "kept the shared-formula master verbatim; the orphaned member reads as unknown"
            },
            Self::FormulaLeadingEquals => "stripped the leading '=' from the formula",
            Self::EmptyFormula => "dropped the empty formula",
        }
    }
}

/// A worksheet-level structural-advisory defect repaired the way Excel
/// repairs it.
///
/// Every variant is a *roadmap contract* (audit findings 4.1 and 4.10): no
/// read produces it yet.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum WorksheetDefect {
    /// A row `ht` lies outside 0..=409 points (finding 4.1). The height is
    /// clamped to the domain edge.
    RowHeightOutOfRange,
    /// A column `width` lies outside 0..=255 characters (finding 4.1). The
    /// width is clamped to the domain edge.
    ColumnWidthOutOfRange,
    /// A `sheetFormatPr` default is negative or out of range (finding 4.1).
    /// The default is dropped, so the application default applies.
    SheetFormatDefaultOutOfRange,
    /// `dimension`, `sheetFormatPr`, `cols`, or `mergeCells` appears out of
    /// schema order, or `dimension` is duplicated or unparseable (finding
    /// 4.10). The element is accepted position-independently; an unusable
    /// `dimension` is ignored.
    MisplacedElement,
    /// An advisory count disagrees with the records it describes, or a
    /// collection is empty (finding 4.10). The records are authoritative.
    CountMismatch,
    /// Rows appear out of ascending order or a row index repeats (finding
    /// 4.10). Rows are sorted; a duplicate's cells merge last-wins.
    RowOutOfOrder,
}

impl WorksheetDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::RowHeightOutOfRange => "clamped the row height to 0..=409 points",
            Self::ColumnWidthOutOfRange => "clamped the column width to 0..=255 characters",
            Self::SheetFormatDefaultOutOfRange => "dropped the out-of-range sheet format default",
            Self::MisplacedElement => "accepted the out-of-order worksheet element",
            Self::CountMismatch => "ignored the advisory count; the records are authoritative",
            Self::RowOutOfOrder => "sorted the out-of-order rows",
        }
    }
}

/// A workbook-catalog (`xl/workbook.xml`) defect.
///
/// Every variant is a *roadmap contract* (audit findings 4.7 and 4.8): no
/// read produces it yet.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CatalogDefect {
    /// `workbookView/@activeTab` names a sheet past the catalog (finding
    /// 4.7). The active tab is clamped to the last sheet.
    ActiveTabOutOfRange,
    /// A `definedName/@localSheetId` names a sheet past the catalog (finding
    /// 4.7). The name is demoted to workbook-global scope.
    LocalSheetIdOutOfRange,
    /// A sheet name violates Excel's authoring rules — over 31 characters,
    /// a reserved character, or a leading/trailing apostrophe (finding 4.8).
    /// The name is retained verbatim. This is the one judgment-call repair on
    /// the xlsx roadmap and will be gated behind an opt-in [`Leniency`]
    /// variant.
    SheetNameRules,
}

impl CatalogDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::ActiveTabOutOfRange => "clamped the active tab to the last sheet",
            Self::LocalSheetIdOutOfRange => {
                "demoted the sheet-scoped defined name to workbook scope"
            },
            Self::SheetNameRules => "retained the non-conforming sheet name verbatim",
        }
    }
}

/// A shared-string table (`xl/sharedStrings.xml`) defect.
///
/// Every variant is a *roadmap contract* (audit findings 4.5 and 4.12): no
/// read produces it yet.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SharedStringsDefect {
    /// `sst/@count` or `sst/@uniqueCount` is unparseable or out of range
    /// (finding 4.5). The hint is ignored, as the crate's tolerant twin
    /// parser already does.
    CountHint,
    /// A rich-text run lacks its `<t>`, or an item mixes simple and rich
    /// text (finding 4.12). The offending run is dropped; the item's
    /// remaining text is kept.
    RunStructure,
    /// A run property is duplicated, unknown, or malformed (finding 4.12).
    /// The property is dropped; an unknown underline reads as single.
    RunProperty,
    /// An `_xHHHH_` escape encodes an unpaired surrogate (finding 4.12). The
    /// escape decodes to U+FFFD.
    UnpairedEscape,
}

impl SharedStringsDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::CountHint => "ignored the malformed shared-string count hint",
            Self::RunStructure => "dropped the malformed rich-text run; the item text was kept",
            Self::RunProperty => "dropped the malformed run property",
            Self::UnpairedEscape => "decoded the unpaired surrogate escape as U+FFFD",
        }
    }
}

/// A theme color-scheme (`a:clrScheme`) defect.
///
/// Every variant is a *roadmap contract* (audit finding 4.11): no read
/// produces it yet. The theme parser is not on the workbook-open path.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ThemeDefect {
    /// Color-scheme slots appear out of order. Slots are matched by name.
    MisorderedSlot,
    /// A color-scheme slot is missing. The slot takes its default.
    MissingSlot,
    /// `a:sysClr` lacks the optional `lastClr`. The color resolves by name,
    /// else black.
    MissingLastColor,
    /// `a:srgbClr/@val` is not six hex digits. The slot takes its default.
    MalformedColor,
}

impl ThemeDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::MisorderedSlot => "matched the color-scheme slot by name",
            Self::MissingSlot => "defaulted the missing color-scheme slot",
            Self::MissingLastColor => "resolved the system color by name",
            Self::MalformedColor => "defaulted the slot with a malformed color",
        }
    }
}

/// One class of defect repaired while a workbook was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// a read may change about a workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A per-cell (or per-row/column style) defect.
    Cell(CellDefect),
    /// A worksheet structural-advisory defect.
    Worksheet(WorksheetDefect),
    /// A workbook-catalog defect.
    Catalog(CatalogDefect),
    /// A shared-string table defect.
    SharedStrings(SharedStringsDefect),
    /// A theme color-scheme defect.
    Theme(ThemeDefect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::Cell(defect) => defect.repair(),
            Self::Worksheet(defect) => defect.repair(),
            Self::Catalog(defect) => defect.repair(),
            Self::SharedStrings(defect) => defect.repair(),
            Self::Theme(defect) => defect.repair(),
        }
    }
}

impl From<CellDefect> for Defect {
    fn from(defect: CellDefect) -> Self {
        Self::Cell(defect)
    }
}

impl From<WorksheetDefect> for Defect {
    fn from(defect: WorksheetDefect) -> Self {
        Self::Worksheet(defect)
    }
}

impl From<CatalogDefect> for Defect {
    fn from(defect: CatalogDefect) -> Self {
        Self::Catalog(defect)
    }
}

impl From<SharedStringsDefect> for Defect {
    fn from(defect: SharedStringsDefect) -> Self {
        Self::SharedStrings(defect)
    }
}

impl From<ThemeDefect> for Defect {
    fn from(defect: ThemeDefect) -> Self {
        Self::Theme(defect)
    }
}

/// Where in its part a repaired defect was found.
///
/// Worksheet locations are relative to the worksheet named by
/// [`ToleratedDefect::sheet`]; coordinates are zero-based.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Location {
    /// The workbook catalog part as a whole (for example `workbookView`).
    Workbook,
    /// The zero-based position of a `definedName` in the catalog.
    DefinedName(u32),
    /// The shared-string table as a whole (for example its count hints).
    SharedStringTable,
    /// The zero-based index of one shared-string item (`si`).
    SharedStringItem(u32),
    /// The worksheet part as a whole (for example its `dimension`).
    Worksheet,
    /// One worksheet cell.
    Cell(Address),
    /// One worksheet row's properties.
    Row(RowIndex),
    /// One worksheet column-range record (`col/@min..=@max`).
    Columns {
        /// The first column the record covers.
        first: ColumnIndex,
        /// The last column the record covers.
        last: ColumnIndex,
    },
    /// The zero-based slot of a theme color-scheme entry.
    ColorSchemeSlot(u8),
}

/// One repaired defect, with enough context to locate it in the workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    /// Which defect class was repaired.
    pub defect: Defect,
    /// The zero-based workbook position of the worksheet the defect was
    /// found in, when it was found in a worksheet and reported through a
    /// workbook-level aggregate; `None` for workbook-scoped parts and in a
    /// single worksheet's own report.
    pub sheet: Option<u32>,
    /// Where in its part the defect was found.
    pub location: Location,
}

/// Everything a read repaired, in the order the defects were found.
///
/// An empty report means every parse so far conformed and no substitution
/// was made. See the [module documentation](self) for how per-worksheet
/// reports aggregate into a workbook report.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

/// The report of a parse that cannot repair anything, such as a chart sheet's
/// absent cell store.
pub(crate) static CLEAN: ToleranceReport = ToleranceReport {
    defects: Vec::new(),
    unrecorded: 0,
};

impl ToleranceReport {
    /// Whether the parse completed without any repair.
    #[inline]
    #[must_use]
    pub fn is_clean(&self) -> bool {
        self.defects.is_empty() && self.unrecorded == 0
    }

    /// The individually recorded defects, in discovery order.
    #[inline]
    #[must_use]
    pub fn defects(&self) -> &[ToleratedDefect] {
        &self.defects
    }

    /// Iterate the individually recorded defects, in discovery order.
    #[inline]
    pub fn iter(&self) -> std::slice::Iter<'_, ToleratedDefect> {
        self.defects.iter()
    }

    /// Defects that were repaired but not stored, once the bound was reached.
    #[inline]
    #[must_use]
    pub fn unrecorded(&self) -> u32 {
        self.unrecorded
    }

    /// Total repairs performed, recorded or not.
    #[inline]
    #[must_use]
    pub fn total(&self) -> u64 {
        self.defects.len() as u64 + u64::from(self.unrecorded)
    }

    /// Record one repair, saturating at [`MAX_RECORDED_DEFECTS`] stored entries.
    pub(crate) fn record(&mut self, defect: impl Into<Defect>, location: Location) {
        self.push(ToleratedDefect {
            defect: defect.into(),
            sheet: None,
            location,
        });
    }

    fn push(&mut self, tolerated: ToleratedDefect) {
        if self.defects.len() < MAX_RECORDED_DEFECTS {
            self.defects.push(tolerated);
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }

    /// Merge another parse's report into this one, keeping discovery order.
    ///
    /// The stored-entry bound applies to the merged report as a whole; entries
    /// that no longer fit join the unrecorded count.
    pub(crate) fn append(&mut self, other: &Self) {
        for tolerated in other.iter() {
            self.push(*tolerated);
        }
        self.unrecorded = self.unrecorded.saturating_add(other.unrecorded);
    }

    /// Merge one worksheet's report, stamping each entry with the worksheet's
    /// workbook position.
    pub(crate) fn append_sheet(&mut self, other: &Self, position: usize) {
        let sheet = u32::try_from(position).ok();
        for tolerated in other.iter() {
            self.push(ToleratedDefect {
                sheet,
                ..*tolerated
            });
        }
        self.unrecorded = self.unrecorded.saturating_add(other.unrecorded);
    }
}

impl<'a> IntoIterator for &'a ToleranceReport {
    type Item = &'a ToleratedDefect;
    type IntoIter = std::slice::Iter<'a, ToleratedDefect>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cell(row: u32, column: u32) -> Location {
        Location::Cell(Address::at(row, column).unwrap())
    }

    #[test]
    fn strict_is_the_default() {
        assert_eq!(Leniency::default(), Leniency::Strict);
    }

    #[test]
    fn a_fresh_report_is_clean() {
        let report = ToleranceReport::default();
        assert!(report.is_clean());
        assert_eq!(report.total(), 0);
        assert!(report.defects().is_empty());
        assert_eq!(report.iter().count(), 0);
    }

    #[test]
    fn recorded_defects_are_returned_in_discovery_order() {
        let mut report = ToleranceReport::default();
        report.record(CellDefect::StyleIndexOutOfRange, cell(0, 0));
        report.record(CellDefect::SharedStringIndexOutOfRange, cell(4, 2));
        report.record(CatalogDefect::ActiveTabOutOfRange, Location::Workbook);

        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert_eq!(
            (&report)
                .into_iter()
                .map(|d| (d.defect, d.sheet, d.location))
                .collect::<Vec<_>>(),
            vec![
                (
                    Defect::Cell(CellDefect::StyleIndexOutOfRange),
                    None,
                    cell(0, 0)
                ),
                (
                    Defect::Cell(CellDefect::SharedStringIndexOutOfRange),
                    None,
                    cell(4, 2)
                ),
                (
                    Defect::Catalog(CatalogDefect::ActiveTabOutOfRange),
                    None,
                    Location::Workbook
                ),
            ]
        );
    }

    /// A hostile part must not be able to grow the report without limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for _ in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(CellDefect::StyleIndexOutOfRange, Location::Worksheet);
        }
        for _ in 0..3 {
            report.record(CellDefect::StyleIndexOutOfRange, Location::Worksheet);
        }

        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 8);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 8);
    }

    #[test]
    fn appended_reports_merge_within_the_bound() {
        let mut base = ToleranceReport::default();
        base.record(CatalogDefect::ActiveTabOutOfRange, Location::Workbook);

        let mut parsed = ToleranceReport::default();
        for _ in 0..MAX_RECORDED_DEFECTS {
            parsed.record(CellDefect::StyleIndexOutOfRange, cell(1, 1));
        }
        assert_eq!(parsed.unrecorded(), 0);

        base.append(&parsed);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 1);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 1);
        assert_eq!(
            base.defects()[0],
            ToleratedDefect {
                defect: Defect::Catalog(CatalogDefect::ActiveTabOutOfRange),
                sheet: None,
                location: Location::Workbook,
            }
        );

        let mut clean = ToleranceReport::default();
        clean.append(&ToleranceReport::default());
        clean.append_sheet(&ToleranceReport::default(), 3);
        assert!(clean.is_clean());
    }

    #[test]
    fn sheet_reports_are_stamped_with_their_position() {
        let mut sheet = ToleranceReport::default();
        sheet.record(CellDefect::MissingSharedStringTable, cell(2, 3));
        sheet.unrecorded = 2;

        let mut workbook = ToleranceReport::default();
        workbook.append_sheet(&sheet, 1);
        assert_eq!(workbook.defects()[0].sheet, Some(1));
        assert_eq!(workbook.defects()[0].location, cell(2, 3));
        assert_eq!(workbook.unrecorded(), 2);
        assert_eq!(workbook.total(), 3);
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        let defects = [
            Defect::from(CellDefect::StyleIndexOutOfRange),
            Defect::from(CellDefect::SharedStringIndexOutOfRange),
            Defect::from(CellDefect::MissingSharedStringTable),
            Defect::from(CellDefect::MalformedNumber),
            Defect::from(CellDefect::MalformedBoolean),
            Defect::from(CellDefect::MalformedDate),
            Defect::from(CellDefect::BrokenSharedFormulaGroup),
            Defect::from(CellDefect::FormulaLeadingEquals),
            Defect::from(CellDefect::EmptyFormula),
            Defect::from(WorksheetDefect::RowHeightOutOfRange),
            Defect::from(WorksheetDefect::ColumnWidthOutOfRange),
            Defect::from(WorksheetDefect::SheetFormatDefaultOutOfRange),
            Defect::from(WorksheetDefect::MisplacedElement),
            Defect::from(WorksheetDefect::CountMismatch),
            Defect::from(WorksheetDefect::RowOutOfOrder),
            Defect::from(CatalogDefect::ActiveTabOutOfRange),
            Defect::from(CatalogDefect::LocalSheetIdOutOfRange),
            Defect::from(CatalogDefect::SheetNameRules),
            Defect::from(SharedStringsDefect::CountHint),
            Defect::from(SharedStringsDefect::RunStructure),
            Defect::from(SharedStringsDefect::RunProperty),
            Defect::from(SharedStringsDefect::UnpairedEscape),
            Defect::from(ThemeDefect::MisorderedSlot),
            Defect::from(ThemeDefect::MissingSlot),
            Defect::from(ThemeDefect::MissingLastColor),
            Defect::from(ThemeDefect::MalformedColor),
        ];
        for defect in defects {
            assert!(!defect.repair().is_empty());
        }
    }

    /// The carrier is wired through every open path: a conformant workbook
    /// whose sheets have all been parsed reports clean, per sheet and in the
    /// workbook aggregate.
    #[test]
    fn conformant_fixture_opens_with_a_clean_report() {
        let fixture = include_bytes!("../../../test-data/ooxml/xlsx/DateFormatTests.xlsx");

        let workbook = crate::Workbook::from_slice(fixture).unwrap();
        assert!(workbook.tolerance_report().is_clean());
        for sheet in workbook.sheets() {
            assert!(sheet.tolerance_report().unwrap().is_clean());
        }
        let report = workbook.tolerance_report();
        assert!(report.is_clean());
        assert_eq!(report.total(), 0);
        assert_eq!(report.unrecorded(), 0);

        let package = crate::Package::from_slice(fixture).unwrap();
        assert!(package.workbook().unwrap().tolerance_report().is_clean());

        let source =
            crate::SourceBackedWorkbook::from_reader(std::io::Cursor::new(fixture.as_slice()))
                .unwrap();
        for sheet in source.sheets() {
            assert!(sheet.tolerance_report().unwrap().is_clean());
        }
        assert!(source.tolerance_report().is_clean());
    }
}
