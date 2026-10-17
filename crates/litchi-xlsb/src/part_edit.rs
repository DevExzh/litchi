//! The staged part-level engine behind the XLSB worksheet feature edit
//! families (unified-ops record 0322).
//!
//! [`Staged`] holds one immutable [`Workbook`] (borrowed, or materialized
//! from a [`Package`]) and a staged copy of its OPC graph. A family stages
//! whole-part replacements, additions and removals, records every part it
//! touched, and commits: [`PartPatch`] captures the exact before/after state
//! (content type, payload and relationships) of each touched part, and an
//! unchanged staging returns the source snapshot itself, so an exact no-op
//! keeps every byte. The source snapshot never changes.
//!
//! Worksheet streams are edited at the record level: [`SheetLayout`] splits
//! the records after the cell table into top-level units (one record, or a
//! begin/end collection), so a family replaces, inserts or removes exactly
//! its own units and copies every other byte. A rewritten worksheet keeps its
//! binary index current ([`crate::worksheet_index::maintain`]).
//!
//! The families built on it ([`crate::page_breaks::edit`],
//! [`crate::auto_filter::edit`], [`crate::hyperlinks::edit`],
//! [`crate::data_validation::edit`], [`crate::conditional_formatting::edit`],
//! [`crate::comments::edit`] and [`crate::drawing::edit`]) expose the ADR
//! 0034 transaction shape through [`family!`].

use std::collections::BTreeSet;
use std::ops::Range;
use std::sync::Arc;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};

use crate::Workbook;
use crate::package::Package;
use crate::package::error::{Error, Result};
use crate::raw::{Kind, Records, kind};

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Relationship {
    id: String,
    kind: String,
    target: String,
    external: bool,
}

/// One part's exact state: its content type, payload and relationships.
#[derive(Debug, Clone, PartialEq, Eq)]
struct State {
    content_type: String,
    blob: Arc<Vec<u8>>,
    relationships: Vec<Relationship>,
}

impl State {
    fn read(part: &dyn Part) -> Self {
        let mut relationships: Vec<_> = part
            .rels()
            .iter()
            .map(|relationship| Relationship {
                id: relationship.r_id().to_owned(),
                kind: relationship.reltype().to_owned(),
                target: relationship.target_ref().to_owned(),
                external: relationship.is_external(),
            })
            .collect();
        relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
        Self {
            content_type: part.content_type().to_owned(),
            blob: part.blob_arc(),
            relationships,
        }
    }

    fn of(package: &OpcPackage, name: &PackURI) -> Option<Self> {
        package.get_part(name).ok().map(Self::read)
    }

    fn to_part(&self, name: PackURI) -> Result<BlobPart> {
        let mut part =
            BlobPart::new_shared(name, self.content_type.clone(), Arc::clone(&self.blob));
        for relationship in &self.relationships {
            part.rels_mut().try_add_relationship(
                relationship.kind.clone(),
                relationship.target.clone(),
                relationship.id.clone(),
                if relationship.external {
                    TargetMode::External
                } else {
                    TargetMode::Internal
                },
            )?;
        }
        Ok(part)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Delta {
    name: PackURI,
    before: Option<State>,
    after: Option<State>,
}

/// The exact, reversible part-level change of one staged edit.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct PartPatch {
    deltas: Vec<Delta>,
}

impl PartPatch {
    pub(crate) fn is_noop(&self) -> bool {
        self.deltas.is_empty()
    }

    pub(crate) fn parts(&self) -> usize {
        self.deltas.len()
    }

    /// The bytes the patch retains: every before and after payload.
    pub(crate) fn retained_bytes(&self) -> usize {
        self.deltas.iter().fold(0usize, |total, delta| {
            let size = |state: &Option<State>| state.as_ref().map_or(0, |state| state.blob.len());
            total
                .saturating_add(size(&delta.before))
                .saturating_add(size(&delta.after))
        })
    }

    pub(crate) fn inverse(&self) -> Self {
        Self {
            deltas: self
                .deltas
                .iter()
                .map(|delta| Delta {
                    name: delta.name.clone(),
                    before: delta.after.clone(),
                    after: delta.before.clone(),
                })
                .collect(),
        }
    }

    fn capture(before: &OpcPackage, after: &OpcPackage, names: &BTreeSet<String>) -> Result<Self> {
        let mut deltas = Vec::new();
        for name in names {
            let name = PackURI::new(name.as_str())?;
            let before = State::of(before, &name);
            let after = State::of(after, &name);
            if before != after {
                deltas.push(Delta {
                    name,
                    before,
                    after,
                });
            }
        }
        Ok(Self { deltas })
    }

    /// Replay onto `package`, refusing a package that is not the patch's
    /// before state.
    pub(crate) fn apply(&self, package: &Package) -> Result<Package> {
        if self.is_noop() {
            return Ok(package.clone());
        }
        crate::signature_policy::ensure_unsigned(package.opc_package())?;
        let mut target = package.opc_package().clone();
        for delta in &self.deltas {
            if State::of(&target, &delta.name) != delta.before {
                return Err(Error::PatchConflict {
                    part: delta.name.as_str().to_owned(),
                });
            }
        }
        for delta in &self.deltas {
            target.remove_part(&delta.name);
            if let Some(after) = &delta.after {
                target.add_part(Box::new(after.to_part(delta.name.clone())?));
            }
        }
        package.reparse_candidate_retaining_projection(target)
    }
}

/// The snapshot an edit stages over.
#[derive(Debug)]
enum Base<'a> {
    /// A package and the workbook materialized from it once.
    Package(&'a Package, Box<Workbook>),
    /// A workbook staged without first copying it into a package.
    Workbook(&'a Workbook),
}

/// A staged part-level edit of one workbook snapshot.
#[derive(Debug)]
pub(crate) struct Staged<'a> {
    base: Base<'a>,
    staged: OpcPackage,
    touched: BTreeSet<String>,
}

impl<'a> Staged<'a> {
    /// Stage over a package.
    pub(crate) fn of_package(package: &'a Package) -> Result<Self> {
        let workbook = Box::new(package.workbook()?);
        Ok(Self {
            staged: package.opc_package().clone(),
            base: Base::Package(package, workbook),
            touched: BTreeSet::new(),
        })
    }

    /// Stage over a workbook.
    pub(crate) fn of_workbook(workbook: &'a Workbook) -> Self {
        Self {
            staged: workbook.package.clone(),
            base: Base::Workbook(workbook),
            touched: BTreeSet::new(),
        }
    }

    /// The source workbook (its catalog, formula context and styles).
    pub(crate) fn workbook(&self) -> &Workbook {
        match &self.base {
            Base::Package(_, workbook) => workbook,
            Base::Workbook(workbook) => workbook,
        }
    }

    /// The staged OPC graph.
    pub(crate) const fn opc(&self) -> &OpcPackage {
        &self.staged
    }

    /// The staged OPC graph for a change: refuses a signed package first.
    /// The caller records every part it changes with [`Self::touch`].
    pub(crate) fn opc_mut(&mut self) -> Result<&mut OpcPackage> {
        crate::signature_policy::ensure_unsigned(&self.staged)?;
        Ok(&mut self.staged)
    }

    pub(crate) fn touch(&mut self, name: &PackURI) {
        self.touched.insert(name.as_str().to_owned());
    }

    /// The part name of the worksheet at complete-catalog position `sheet`;
    /// refuses a missing tab and a chart, dialog or macro sheet.
    pub(crate) fn worksheet(&self, sheet: usize) -> Result<PackURI> {
        worksheet_uri(self.workbook(), sheet)
    }

    /// The staged payload of `name`.
    pub(crate) fn blob(&self, name: &PackURI) -> Result<&[u8]> {
        Ok(self.staged.get_part(name)?.blob())
    }

    /// Replace the payload of an existing part; identical bytes are a no-op.
    pub(crate) fn replace_blob(&mut self, name: &PackURI, bytes: Vec<u8>) -> Result<bool> {
        if self.staged.get_part(name)?.blob() == bytes.as_slice() {
            return Ok(false);
        }
        self.opc_mut()?.get_part_mut(name)?.set_blob(bytes);
        self.touch(name);
        Ok(true)
    }

    /// Replace a worksheet stream and keep its binary index current.
    pub(crate) fn replace_worksheet(&mut self, name: &PackURI, bytes: Vec<u8>) -> Result<bool> {
        let original = self.staged.get_part(name)?.blob_arc();
        if original.as_slice() == bytes.as_slice() {
            return Ok(false);
        }
        let index = crate::worksheet_index::find_index_target(&self.staged, name)?;
        let package = self.opc_mut()?;
        crate::worksheet_index::maintain(package, name, &original, &bytes)?;
        package.get_part_mut(name)?.set_blob(bytes);
        self.touch(name);
        if let Some(index) = index {
            self.touch(&index);
        }
        Ok(true)
    }

    /// Add a new part.
    pub(crate) fn add_part(&mut self, part: BlobPart) -> Result<()> {
        let name = part.partname().clone();
        self.opc_mut()?.add_part(Box::new(part));
        self.touch(&name);
        Ok(())
    }

    /// Remove a part; `false` when it is absent.
    pub(crate) fn remove_part(&mut self, name: &PackURI) -> Result<bool> {
        let removed = self.opc_mut()?.remove_part(name);
        if removed {
            self.touch(name);
        }
        Ok(removed)
    }

    /// The first free part name `{prefix}{N}{suffix}`.
    pub(crate) fn free_name(&self, prefix: &str, suffix: &str) -> Result<PackURI> {
        for index in 1_u32..=1_000_000 {
            let candidate = PackURI::new(format!("{prefix}{index}{suffix}"))?;
            if self.staged.get_part(&candidate).is_err() {
                return Ok(candidate);
            }
        }
        Err(invalid("no free part name is left"))
    }

    /// Capture the patch and publish the staged package; an unchanged
    /// staging returns the source snapshot.
    pub(crate) fn commit(self) -> Result<(Package, PartPatch)> {
        let source = match &self.base {
            Base::Package(package, _) => package.opc_package(),
            Base::Workbook(workbook) => &workbook.package,
        };
        let patch = PartPatch::capture(source, &self.staged, &self.touched)?;
        let package = match self.base {
            Base::Package(package, _) if patch.is_noop() => package.clone(),
            Base::Workbook(workbook) if patch.is_noop() => workbook.to_package(),
            Base::Package(package, _) => {
                package.reparse_candidate_retaining_projection(self.staged)?
            },
            Base::Workbook(workbook) => Package::reparse_workbook_candidate(workbook, self.staged)?,
        };
        Ok((package, patch))
    }
}

/// The worksheet part of a complete-catalog position.
pub(crate) fn worksheet_uri(workbook: &Workbook, sheet: usize) -> Result<PackURI> {
    let worksheet = workbook
        .worksheet_for_catalog_position(sheet)
        .ok_or_else(|| {
            if sheet < workbook.sheets().len() {
                Error::UnsupportedFeature("worksheet features belong to worksheets".to_string())
            } else {
                Error::WorksheetNotFound(format!("catalog position {sheet}"))
            }
        })?;
    workbook.worksheet_uri(worksheet)
}

/// The semantic worksheet features one feature pass decodes: everything
/// the full worksheet reader decodes after the cell table, without decoding
/// the cells.
pub(crate) struct FeaturePass {
    pub(crate) validations: Vec<crate::package::data_validation::Validation>,
    pub(crate) formats: Vec<crate::conditional_formatting::Formatting>,
}

/// Run the full reader's feature decoding over a worksheet stream, with the
/// formula context of complete-catalog position `sheet`.
pub(crate) fn feature_pass(
    bytes: &[u8],
    context: &crate::package::formula::Context,
    sheet: usize,
    leniency: crate::leniency::Leniency,
) -> Result<FeaturePass> {
    let context = context.for_sheet(sheet);
    let stream = crate::package::records::Stream::new(std::io::Cursor::new(bytes));
    let mut reader = crate::package::cells_reader::CellsReader::new(
        stream,
        &[],
        &context,
        usize::MAX,
        leniency,
    )?;
    reader.skip_to_features()?;
    Ok(FeaturePass {
        validations: reader.data_validations,
        formats: reader.conditional_formattings,
    })
}

impl Staged<'_> {
    /// The feature pass of the staged worksheet `name` at catalog position
    /// `sheet`.
    pub(crate) fn feature_pass(&self, name: &PackURI, sheet: usize) -> Result<FeaturePass> {
        let workbook = self.workbook();
        feature_pass(
            self.blob(name)?,
            &workbook.formula_context,
            sheet,
            workbook.leniency(),
        )
    }
}

// ---------------------------------------------------------------------------
// Worksheet stream layout

/// The stream rank of a worksheet record after the cell table, in
/// [MS-XLSB] section 2.1.7.62 order. Authored units are inserted before the
/// first top-level unit of a higher rank.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(crate) enum Rank {
    SheetCalcProp,
    SheetProtectionIso,
    SheetProtection,
    RangeProtectionIso,
    RangeProtection,
    ScenarioManager,
    AutoFilter,
    SortState,
    DataConsolidation,
    UserSheetViews,
    MergeCells,
    PhoneticInfo,
    ConditionalFormatting,
    DataValidations,
    Hyperlink,
    PrintOptions,
    Margins,
    PageSetup,
    HeaderFooter,
    RowBreaks,
    ColumnBreaks,
    BigName,
    CellWatches,
    IgnoredErrors,
    SmartTags,
    Drawing,
    LegacyDrawing,
    LegacyDrawingHeaderFooter,
    BackgroundPicture,
    OleObjects,
    ActiveXControls,
    WebPublishItems,
    ListParts,
    FutureRecords,
}

/// The collections (begin and end kinds) and single records the layout
/// knows, with their stream rank; `None` ranks an alternate-content block.
const BLOCKS: &[(Kind, Kind, Option<Rank>)] = &[
    (
        kind::BEGIN_SCEN_MAN,
        kind::END_SCEN_MAN,
        Some(Rank::ScenarioManager),
    ),
    (
        kind::BEGIN_A_FILTER,
        kind::END_A_FILTER,
        Some(Rank::AutoFilter),
    ),
    (
        kind::BEGIN_SORT_STATE,
        kind::END_SORT_STATE,
        Some(Rank::SortState),
    ),
    (
        kind::BEGIN_D_CON,
        kind::END_D_CON,
        Some(Rank::DataConsolidation),
    ),
    (
        kind::BEGIN_USER_SH_VIEWS,
        kind::END_USER_SH_VIEWS,
        Some(Rank::UserSheetViews),
    ),
    (
        kind::BEGIN_MERGE_CELLS,
        kind::END_MERGE_CELLS,
        Some(Rank::MergeCells),
    ),
    (
        kind::BEGIN_COND_FORMATTING,
        kind::END_COND_FORMATTING,
        Some(Rank::ConditionalFormatting),
    ),
    (
        kind::BEGIN_D_VALS,
        kind::END_D_VALS,
        Some(Rank::DataValidations),
    ),
    (
        kind::BEGIN_HEADER_FOOTER,
        kind::END_HEADER_FOOTER,
        Some(Rank::HeaderFooter),
    ),
    (kind::BEGIN_RW_BRK, kind::END_RW_BRK, Some(Rank::RowBreaks)),
    (
        kind::BEGIN_COL_BRK,
        kind::END_COL_BRK,
        Some(Rank::ColumnBreaks),
    ),
    (
        kind::BEGIN_CELL_WATCHES,
        kind::END_CELL_WATCHES,
        Some(Rank::CellWatches),
    ),
    (
        kind::BEGIN_CELL_IGNORE_ECS,
        kind::END_CELL_IGNORE_ECS,
        Some(Rank::IgnoredErrors),
    ),
    (
        kind::BEGIN_SMART_TAGS,
        kind::END_SMART_TAGS,
        Some(Rank::SmartTags),
    ),
    (
        kind::BEGIN_OLE_OBJECTS,
        kind::END_OLE_OBJECTS,
        Some(Rank::OleObjects),
    ),
    (
        kind::BEGIN_ACTIVE_X_CONTROLS,
        kind::END_ACTIVE_X_CONTROLS,
        Some(Rank::ActiveXControls),
    ),
    (
        kind::BEGIN_WEB_PUB_ITEMS,
        kind::END_WEB_PUB_ITEMS,
        Some(Rank::WebPublishItems),
    ),
    (
        kind::BEGIN_LIST_PARTS,
        kind::END_LIST_PARTS,
        Some(Rank::ListParts),
    ),
    (kind::FRT_BEGIN, kind::FRT_END, Some(Rank::FutureRecords)),
    (kind::AC_BEGIN, kind::AC_END, None),
];

const SINGLES: &[(Kind, Rank)] = &[
    (kind::SHEET_CALC_PROP, Rank::SheetCalcProp),
    (kind::SHEET_PROTECTION_ISO, Rank::SheetProtectionIso),
    (kind::SHEET_PROTECTION, Rank::SheetProtection),
    (kind::RANGE_PROTECTION_ISO, Rank::RangeProtectionIso),
    (kind::RANGE_PROTECTION, Rank::RangeProtection),
    (kind::PHONETIC_INFO, Rank::PhoneticInfo),
    (kind::H_LINK, Rank::Hyperlink),
    (kind::PRINT_OPTIONS, Rank::PrintOptions),
    (kind::MARGINS, Rank::Margins),
    (kind::PAGE_SETUP, Rank::PageSetup),
    (kind::BIG_NAME, Rank::BigName),
    (kind::DRAWING, Rank::Drawing),
    (kind::LEGACY_DRAWING, Rank::LegacyDrawing),
    (kind::LEGACY_DRAWING_HF, Rank::LegacyDrawingHeaderFooter),
    (kind::BK_HIM, Rank::BackgroundPicture),
];

/// One top-level unit after the cell table: a single record, or a begin/end
/// collection with every record it nests.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Unit {
    /// The kind of the unit's first record.
    pub(crate) kind: Kind,
    /// The unit's stream rank, when the layout knows its kind.
    pub(crate) rank: Option<Rank>,
    /// The unit's byte span.
    pub(crate) span: Range<usize>,
}

/// The top-level record units of a worksheet stream after its cell table.
#[derive(Debug, Clone)]
pub(crate) struct SheetLayout {
    pub(crate) units: Vec<Unit>,
    /// The offset of `BrtEndSheet`.
    pub(crate) end_sheet: usize,
}

impl SheetLayout {
    /// Scan a worksheet stream. Unbalanced collections and a stream without
    /// a cell table or end record are refused.
    pub(crate) fn scan(data: &[u8]) -> Result<Self> {
        let mut records = Records::new(data);
        let mut after_cells = false;
        let mut units = Vec::new();
        let mut stack: Vec<Kind> = Vec::new();
        let mut open: Option<(Kind, Option<Rank>, usize)> = None;
        while let Some(record) = records.next() {
            let record = record?;
            let start = record.offset();
            let end = records.offset();
            let kind = record.kind();
            if !after_cells {
                if kind == kind::END_SHEET_DATA {
                    after_cells = true;
                }
                continue;
            }
            if let Some(top) = stack.last().copied() {
                if kind == top {
                    stack.pop();
                    if stack.is_empty()
                        && let Some((first, rank, begin)) = open.take()
                    {
                        units.push(Unit {
                            kind: first,
                            rank,
                            span: begin..end,
                        });
                    }
                } else if let Some((_, close, _)) =
                    BLOCKS.iter().find(|(begin, _, _)| *begin == kind)
                {
                    stack.push(*close);
                } else if BLOCKS.iter().any(|(_, close, _)| *close == kind) {
                    return Err(invalid(format!(
                        "worksheet record 0x{:04X} closes a collection it did not open",
                        kind.get()
                    )));
                }
                continue;
            }
            if kind == kind::END_SHEET {
                return Ok(Self {
                    units,
                    end_sheet: start,
                });
            }
            if let Some((_, close, rank)) = BLOCKS.iter().find(|(begin, _, _)| *begin == kind) {
                stack.push(*close);
                open = Some((kind, *rank, start));
                continue;
            }
            if BLOCKS.iter().any(|(_, close, _)| *close == kind) {
                return Err(invalid(format!(
                    "worksheet record 0x{:04X} closes a collection it did not open",
                    kind.get()
                )));
            }
            let rank = SINGLES
                .iter()
                .find(|(single, _)| *single == kind)
                .map(|(_, rank)| *rank);
            units.push(Unit {
                kind,
                rank,
                span: start..end,
            });
        }
        if !after_cells {
            return Err(invalid("worksheet stream has no cell table"));
        }
        Err(invalid("worksheet stream has no BrtEndSheet"))
    }

    /// The units whose first record is `kind`, in stream order.
    pub(crate) fn units_of(&self, kind: Kind) -> impl Iterator<Item = &Unit> + '_ {
        self.units.iter().filter(move |unit| unit.kind == kind)
    }

    /// The offset a new unit of `rank` is inserted at: before the first
    /// top-level unit of a higher known rank, else before `BrtEndSheet`.
    pub(crate) fn insertion(&self, rank: Rank) -> usize {
        self.units
            .iter()
            .find(|unit| unit.rank.is_some_and(|other| other > rank))
            .map_or(self.end_sheet, |unit| unit.span.start)
    }
}

/// Replace `span` of `data` with `replacement`.
pub(crate) fn splice(data: &[u8], span: Range<usize>, replacement: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(data.len() - span.len() + replacement.len());
    out.extend_from_slice(&data[..span.start]);
    out.extend_from_slice(replacement);
    out.extend_from_slice(&data[span.end..]);
    out
}

/// Encode one record.
pub(crate) fn record(kind: Kind, payload: &[u8]) -> Result<Vec<u8>> {
    let mut out = Vec::with_capacity(payload.len() + 6);
    crate::raw::Encoder::new(&mut out).write_record(kind, payload)?;
    Ok(out)
}

/// The `Patch`, `Commit` and `Diagnostics` types of one family, and its
/// `Error` re-export (ADR 0034 §3).
macro_rules! family {
    ($what:literal) => {
        pub use crate::package::PackageError as Error;

        #[doc = concat!("The exact, reversible part-level change of a ", $what, " edit.")]
        #[derive(Debug, Clone, Default, PartialEq, Eq)]
        pub struct Patch(pub(crate) crate::part_edit::PartPatch);

        impl Patch {
            /// Whether the patch changes nothing.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                self.0.is_noop()
            }

            /// The reverse patch.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self(self.0.inverse())
            }

            /// The payload bytes the patch retains (every before and after
            /// part image).
            #[must_use]
            pub fn retained_bytes(&self) -> usize {
                self.0.retained_bytes()
            }
        }

        #[doc = concat!("What a committed ", $what, " edit changed.")]
        #[derive(Debug, Clone, Copy, PartialEq, Eq)]
        pub struct Diagnostics {
            parts: usize,
        }

        impl Diagnostics {
            /// Whether the edit changed the package.
            #[must_use]
            pub const fn changed(&self) -> bool {
                self.parts > 0
            }

            /// The number of package parts the edit added, changed or
            /// removed.
            #[must_use]
            pub const fn parts(&self) -> usize {
                self.parts
            }
        }

        #[doc = concat!("The result of a ", $what, " edit's `commit`.")]
        #[derive(Debug)]
        pub struct Commit {
            package: crate::Package,
            patch: Patch,
            diagnostics: Diagnostics,
        }

        impl Commit {
            pub(crate) fn new(
                (package, patch): (crate::Package, crate::part_edit::PartPatch),
            ) -> Self {
                let diagnostics = Diagnostics {
                    parts: patch.parts(),
                };
                Self {
                    package,
                    patch: Patch(patch),
                    diagnostics,
                }
            }

            /// The package after the edit.
            #[must_use]
            pub const fn package(&self) -> &crate::Package {
                &self.package
            }

            /// Consume the commit into the package after the edit.
            #[must_use]
            pub fn into_package(self) -> crate::Package {
                self.package
            }

            /// The exact, reversible patch.
            #[must_use]
            pub const fn patch(&self) -> &Patch {
                &self.patch
            }

            /// What the edit changed.
            #[must_use]
            pub const fn diagnostics(&self) -> &Diagnostics {
                &self.diagnostics
            }

            /// Consume the commit into the package and the patch.
            #[must_use]
            pub fn into_parts(self) -> (crate::Package, Patch) {
                (self.package, self.patch)
            }
        }
    };
}

pub(crate) use family;

#[cfg(test)]
mod tests {
    use super::*;

    fn stream(records: &[(Kind, &[u8])]) -> Vec<u8> {
        let mut out = Vec::new();
        let mut encoder = crate::raw::Encoder::new(&mut out);
        for (kind, payload) in records {
            encoder.write_record(*kind, payload).unwrap();
        }
        out
    }

    #[test]
    fn layout_splits_units_and_ranks_insertions() {
        let data = stream(&[
            (kind::BEGIN_SHEET, &[]),
            (kind::BEGIN_SHEET_DATA, &[]),
            (kind::ROW_HDR, &[0; 17]),
            (kind::END_SHEET_DATA, &[]),
            (kind::SHEET_PROTECTION, &[0; 66]),
            (kind::BEGIN_A_FILTER, &[0; 16]),
            (kind::BEGIN_FILTER_COLUMN, &[0; 6]),
            (kind::END_FILTER_COLUMN, &[]),
            (kind::END_A_FILTER, &[]),
            (kind::MARGINS, &[0; 48]),
            (kind::FRT_BEGIN, &[0; 4]),
            (kind::FRT_END, &[]),
            (kind::END_SHEET, &[]),
        ]);
        let layout = SheetLayout::scan(&data).unwrap();
        let kinds: Vec<_> = layout.units.iter().map(|unit| unit.kind).collect();
        assert_eq!(
            kinds,
            [
                kind::SHEET_PROTECTION,
                kind::BEGIN_A_FILTER,
                kind::MARGINS,
                kind::FRT_BEGIN
            ]
        );
        let margins = &layout.units[2];
        assert_eq!(layout.insertion(Rank::Hyperlink), margins.span.start);
        assert_eq!(
            layout.insertion(Rank::RowBreaks),
            layout.units[3].span.start
        );
        assert_eq!(
            layout.insertion(Rank::AutoFilter),
            layout.units[2].span.start
        );
        assert_eq!(
            layout.insertion(Rank::SheetCalcProp),
            layout.units[0].span.start
        );
    }

    #[test]
    fn layout_refuses_unbalanced_collections() {
        let data = stream(&[
            (kind::BEGIN_SHEET, &[]),
            (kind::BEGIN_SHEET_DATA, &[]),
            (kind::END_SHEET_DATA, &[]),
            (kind::END_D_VALS, &[]),
            (kind::END_SHEET, &[]),
        ]);
        assert!(SheetLayout::scan(&data).is_err());
    }
}
