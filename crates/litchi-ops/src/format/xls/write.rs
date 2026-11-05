//! Unified scalar XLS writes over the canonical native package transaction.
//!
//! The BIFF8 editor owns all publication and dependency checks.  This module
//! only admits the unified target/property shape, expands bounded selectors,
//! coerces the common scalar union, and publishes one native `Package::edit`
//! for the complete operation.

use std::collections::HashSet;
use std::mem::size_of;

use litchi_core::{FileFormat, Position};
use litchi_xls::cell_values::{Reference, Selector, Value as NativeValue};

use super::node::{CellState, Item, Node};
use super::spec::{CELL_WRITE_TABLE, CellSetter, Kind, SHEET_WRITE_TABLE, SheetSetter};
use super::{Nav, Xls};
use crate::error::{Code, Error, Result};
use crate::props::{IntoProps, PropValue, Props};
use crate::registry::{Access, Operation};
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::Scalar;

/// The patch an XLS write step retains: a primary edit, a metadata edit or
/// a worksheet-feature edit (record 0327). Each exposes the exact package
/// images the generic history compares.
#[allow(
    clippy::large_enum_variant,
    reason = "the step history boxes the patch; the primary patch is the common one"
)]
pub(crate) enum StepPatch {
    /// A primary content edit.
    Primary(litchi_xls::Patch),
    /// A metadata edit.
    Metadata(litchi_xls::package::metadata::Patch),
    /// A worksheet-feature edit.
    Features(litchi_xls::sheet_features::edit::Patch),
}

impl StepPatch {
    pub(crate) fn before(&self) -> &[u8] {
        match self {
            Self::Primary(patch) => patch.before(),
            Self::Metadata(patch) => patch.before_bytes(),
            Self::Features(patch) => patch.before(),
        }
    }

    pub(crate) fn after(&self) -> &[u8] {
        match self {
            Self::Primary(patch) => patch.after(),
            Self::Metadata(patch) => patch.after_bytes(),
            Self::Features(patch) => patch.after(),
        }
    }
}

/// The native proof and immutable endpoints of one unified XLS write.
pub(crate) struct WriteCommit {
    pub(crate) before: Xls,
    pub(crate) after: Xls,
    pub(crate) patch: StepPatch,
    pub(crate) operations: usize,
    pub(crate) snapshot_bytes: usize,
}

#[derive(Debug, Clone)]
enum CellAssignment {
    Value(NativeValue),
    Clear,
    /// A formula expression without its leading `=`.
    Formula(String),
    /// Alignment, border-style and fill changes to the cell XF.
    Format(litchi_xls::cell_values::FormatChange),
    Noop,
}

#[derive(Debug, Clone)]
struct CellPlan {
    assignment: CellAssignment,
    staged_bytes: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct CellTarget {
    sheet: usize,
    reference: Reference,
    present: bool,
    formula: bool,
}

/// Stable cell identity for coalescing.  Presence is a property of the
/// authored state, not of the address: remove/reinsert must resolve to this
/// same key so a signed cancellation can be recognized.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct CellIdentity {
    sheet: usize,
    reference: Reference,
}

impl CellTarget {
    fn identity(self) -> CellIdentity {
        CellIdentity {
            sheet: self.sheet,
            reference: self.reference,
        }
    }
}

#[derive(Debug, Clone)]
enum SheetAssignment {
    Name(String),
}

const WRITE_ENTRY_OVERHEAD: usize = 64;

/// Admit one scalar cell or worksheet-name mutation and publish it through
/// the native immutable XLS package owner.
pub(crate) fn set_with_limits<'p, 'v>(
    file: &Xls,
    target: TargetArg<'p>,
    props: impl IntoProps<'v>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.admit_edit_for_write()?;
    let props = Props::from_props(props)?;

    // `/metadata` names the document properties unless the call carries a
    // sheet property: a worksheet named `metadata` stays renamable there.
    let result = if crate::format::ole_metadata::is_target(&target) && !is_sheet_property(&props) {
        file.set_metadata(&props, limits)
    } else if let Some(commit) = file.try_set_feature(&target, &props, limits)? {
        Ok(commit)
    } else if is_sheet_property(&props) {
        set_sheet_with_limits(file, target, &props, limits)
    } else if let Some(commit) = super::layout::set_layout(file, &target, &props, limits)? {
        Ok(commit)
    } else {
        set_cells_with_limits(file, target, &props, limits)
    };
    file.check_source_for_write()?;
    result
}

/// Remove existing stored cells selected by a path, range, or scoped query.
/// Missing coordinates are exact no-ops; all present coordinates are removed
/// by one native transaction so row and worksheet neighbors never shift.
pub(crate) fn remove_with_limits(
    file: &Xls,
    target: TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.admit_edit_for_write()?;
    if let Some(commit) = file.try_remove_feature(&target, limits)? {
        return Ok(commit);
    }
    if let Some(commit) = super::structure::remove_sheet(file, &target, limits)? {
        return Ok(commit);
    }
    if let Some(commit) = super::layout::remove_axis(file, &target, limits)? {
        return Ok(commit);
    }
    let targets = resolve_removal_targets(file, &target, limits)?;
    let staged = targets
        .len()
        .checked_mul(WRITE_ENTRY_OVERHEAD)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    if staged > limits.max_staged_bytes() {
        return Err(staged_limit(staged, limits.max_staged_bytes()));
    }
    if targets.is_empty() || !targets.iter().any(|target| target.present) {
        return no_op(file, limits);
    }
    if limits.max_operations() < 1 {
        return Err(operation_limit(1, limits.max_operations()));
    }
    if targets
        .iter()
        .any(|target| target.present && target.formula)
    {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "formula cells require a dependency-aware removal handler",
        )
        .with_format(FileFormat::Xls));
    }

    file.preflight_snapshot_bytes(limits.max_snapshot_bytes(), staged)?;
    let mut edit = file
        .editable_package_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?
        .edit()
        .map_err(native_edit_error)?;
    for target in targets {
        if target.present {
            edit.remove_cell(
                Selector::Position(Position::new(target.sheet)),
                target.reference,
            )
            .map_err(native_edit_error)?;
        }
    }
    let native = edit.commit().map_err(native_edit_error)?;
    finish_commit(file, native, 1, limits)
}

fn is_sheet_property(props: &Props<'_>) -> bool {
    props
        .iter()
        .any(|(key, _)| key.eq_ignore_ascii_case("name"))
}

fn set_sheet_with_limits(
    file: &Xls,
    target: TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    let (assignment, staged_bytes) = admit_sheet_property(props, limits)?;
    let positions = resolve_sheet_targets(file, &target, limits)?;
    let staged = positions
        .len()
        .checked_mul(staged_bytes)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    if staged > limits.max_staged_bytes() {
        return Err(staged_limit(staged, limits.max_staged_bytes()));
    }
    if positions.len() > 1 {
        return Err(Error::new(
            Code::InvalidArgument,
            "an XLS worksheet rename requires exactly one sheet target",
        )
        .with_format(FileFormat::Xls)
        .with_property("name"));
    }
    let Some(position) = positions.into_iter().next() else {
        return no_op(file, limits);
    };
    if limits.max_operations() < 1 {
        return Err(operation_limit(1, limits.max_operations()));
    }

    let SheetAssignment::Name(name) = assignment;
    super::structure::validate_sheet_name(file, Some(position), &name)?;
    if file
        .worksheet_names()?
        .get(position)
        .is_some_and(|current| current == &name)
    {
        return no_op(file, limits);
    }
    file.preflight_snapshot_bytes(limits.max_snapshot_bytes(), staged)?;
    let mut edit = file
        .editable_package_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?
        .edit()
        .map_err(native_edit_error)?;
    edit.rename_sheet(Selector::Position(Position::new(position)), &name)
        .map_err(native_edit_error)?;
    let native = edit.commit().map_err(native_edit_error)?;
    finish_commit(file, native, 1, limits)
}

fn set_cells_with_limits(
    file: &Xls,
    target: TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    let plan = admit_cell_property(props, limits)?;
    let targets = resolve_cell_targets(file, &target, limits)?;
    let staged = targets
        .len()
        .checked_mul(plan.staged_bytes)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    if staged > limits.max_staged_bytes() {
        return Err(staged_limit(staged, limits.max_staged_bytes()));
    }
    if targets.is_empty() {
        return no_op(file, limits);
    }
    if matches!(&plan.assignment, CellAssignment::Noop) {
        return no_op(file, limits);
    }
    if limits.max_operations() < 1 {
        return Err(operation_limit(1, limits.max_operations()));
    }
    if let CellAssignment::Format(change) = &plan.assignment {
        // Formatting rewrites only the cell's XF reference, so formula cells
        // and absent coordinates are admitted alike.
        file.preflight_snapshot_bytes(limits.max_snapshot_bytes(), staged)?;
        let mut edit = file
            .editable_package_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?
            .edit()
            .map_err(native_edit_error)?;
        for target in targets {
            edit.set_cell_format(
                Selector::Position(Position::new(target.sheet)),
                target.reference,
                change,
            )
            .map_err(format_error)?;
        }
        let native = edit.commit().map_err(native_edit_error)?;
        return finish_commit(file, native, 1, limits);
    }
    if targets
        .iter()
        .any(|target| target.present && target.formula)
    {
        let property = match &plan.assignment {
            CellAssignment::Clear => "clear",
            CellAssignment::Formula(_) => "formula",
            CellAssignment::Value(_) | CellAssignment::Format(_) | CellAssignment::Noop => "value",
        };
        return Err(Error::new(
            Code::UnsupportedCapability,
            "formula cells require formula-cache or formula-token editing",
        )
        .with_format(FileFormat::Xls)
        .with_property(property));
    }

    file.preflight_snapshot_bytes(limits.max_snapshot_bytes(), staged)?;
    let mut edit = file
        .editable_package_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?
        .edit()
        .map_err(native_edit_error)?;
    if let CellAssignment::Formula(expression) = &plan.assignment {
        // Each coordinate is written once: an absent cell gains a Formula,
        // and a scalar cell is replaced with one keeping its XF.
        for target in targets {
            edit.set_formula(
                Selector::Position(Position::new(target.sheet)),
                target.reference,
                expression,
            )
            .map_err(formula_error)?;
        }
        let native = edit.commit().map_err(native_edit_error)?;
        return finish_commit(file, native, 1, limits);
    }
    for target in targets {
        let selector = Selector::Position(Position::new(target.sheet));
        if matches!(&plan.assignment, CellAssignment::Clear) {
            // Keep an explicit BIFF8 blank record so the existing cell's
            // style survives and the read projection retains its `empty`
            // fact. A virtual coordinate is already clear and stays a no-op.
            if target.present {
                edit.set_value(selector, target.reference, NativeValue::Blank)
                    .map_err(native_edit_error)?;
            }
            continue;
        }
        let value = match &plan.assignment {
            CellAssignment::Value(value) => value.clone(),
            CellAssignment::Clear => unreachable!("clear handled above"),
            CellAssignment::Formula(_) => unreachable!("formula handled above"),
            CellAssignment::Format(_) => unreachable!("format handled above"),
            CellAssignment::Noop => unreachable!("no-op handled above"),
        };
        if target.present {
            edit.set_value(selector, target.reference, value)
                .map_err(native_edit_error)?;
        } else {
            edit.insert_cell(selector, target.reference, value)
                .map_err(native_edit_error)?;
        }
    }
    let native = edit.commit().map_err(native_edit_error)?;
    finish_commit(file, native, 1, limits)
}

fn admit_sheet_property(
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<(SheetAssignment, usize)> {
    let mut name = None;
    let mut staged_bytes = WRITE_ENTRY_OVERHEAD;
    for (key, value) in props.iter() {
        let (_, entry) = SHEET_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Xls)
                .with_property(key.as_str().to_owned())
        })?;
        let entry = *entry;
        entry
            .prepare(Operation::Set, value, |setter, value| match setter {
                SheetSetter::Name => {
                    if name.is_some() {
                        return Err(duplicate_sheet_property());
                    }
                    let Some(Scalar::Text(value)) = value else {
                        return Err(Error::new(
                            Code::InvalidValue,
                            "an XLS worksheet name requires text",
                        )
                        .with_format(FileFormat::Xls)
                        .with_property("name"));
                    };
                    staged_bytes = staged_bytes
                        .checked_add(value.len())
                        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
                    name = Some(value.to_owned());
                    Ok(())
                },
            })
            .map_err(|error| error.with_format(FileFormat::Xls))?;
    }
    let Some(name) = name else {
        return Err(Error::new(
            Code::InvalidArgument,
            "an XLS worksheet write requires the `name` property",
        )
        .with_format(FileFormat::Xls));
    };
    Ok((SheetAssignment::Name(name), staged_bytes))
}

fn admit_cell_property(props: &Props<'_>, limits: &SemanticLimits) -> Result<CellPlan> {
    let mut assignment = CellAssignment::Noop;
    let mut staged_bytes = WRITE_ENTRY_OVERHEAD;
    for (key, value) in props.iter() {
        let (_, entry) = CELL_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Xls)
                .with_property(key.as_str().to_owned())
        })?;
        let entry = *entry;
        // The shared Scalar registry preserves an exact integer only through
        // i64. BIFF8 also has an exact IEEE number domain, so retain the
        // unsigned values that round-trip through f64 (including 2^63) while
        // refusing values that would be rounded. This exception is scoped to
        // the XLS value setter; other formats keep their scalar policy.
        if matches!(entry.access, Access::Set(CellSetter::Value))
            && let PropValue::Unsigned(value) = value
        {
            let number = *value as f64;
            if !number.is_finite() || number as u128 != u128::from(*value) {
                return Err(invalid_scalar());
            }
            stage_cell_value(
                &mut assignment,
                &mut staged_bytes,
                NativeValue::Number(number),
                limits,
            )?;
            continue;
        }
        entry
            .prepare(Operation::Set, value, |setter, value| match setter {
                CellSetter::Value => {
                    let native = match value {
                        Some(value) => native_scalar(value)?,
                        None => NativeValue::Blank,
                    };
                    stage_cell_value(&mut assignment, &mut staged_bytes, native, limits)
                },
                CellSetter::Clear => {
                    if !matches!(assignment, CellAssignment::Noop) {
                        return Err(Error::new(
                            Code::InvalidArgument,
                            "cell properties value and clear cannot be supplied together",
                        )
                        .with_format(FileFormat::Xls)
                        .with_property("clear"));
                    }
                    let Some(Scalar::Bool(true)) = value else {
                        return Err(Error::new(Code::InvalidValue, "clear must be true")
                            .with_format(FileFormat::Xls)
                            .with_property("clear"));
                    };
                    staged_bytes = staged_bytes
                        .checked_add(1)
                        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
                    assignment = CellAssignment::Clear;
                    Ok(())
                },
                CellSetter::Horizontal
                | CellSetter::Vertical
                | CellSetter::WrapText
                | CellSetter::Indent
                | CellSetter::ShrinkToFit
                | CellSetter::BorderLeft
                | CellSetter::BorderRight
                | CellSetter::BorderTop
                | CellSetter::BorderBottom
                | CellSetter::FillPattern
                | CellSetter::Fill
                | CellSetter::FillBg => {
                    let change = match &mut assignment {
                        CellAssignment::Noop => {
                            assignment = CellAssignment::Format(Default::default());
                            let CellAssignment::Format(change) = &mut assignment else {
                                unreachable!("the format assignment was just installed")
                            };
                            change
                        },
                        CellAssignment::Format(change) => change,
                        CellAssignment::Value(_)
                        | CellAssignment::Clear
                        | CellAssignment::Formula(_) => {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "XLS cell formatting is set separately from value, clear and formula",
                            )
                            .with_format(FileFormat::Xls)
                            .with_suggestion("apply the formatting in its own set call"));
                        },
                    };
                    stage_format(change, setter, value)?;
                    staged_bytes = staged_bytes
                        .checked_add(8)
                        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
                    Ok(())
                },
                CellSetter::Formula => {
                    if !matches!(assignment, CellAssignment::Noop) {
                        return Err(Error::new(
                            Code::InvalidArgument,
                            "cell properties value, clear and formula are mutually exclusive",
                        )
                        .with_format(FileFormat::Xls)
                        .with_property("formula"));
                    }
                    let Some(Scalar::Text(expression)) = value else {
                        return Err(invalid_formula());
                    };
                    // OfficeCLI and the XLSX adapter spell formulas without
                    // the leading `=`; the readback omits it as well.
                    if expression.is_empty()
                        || expression.starts_with('=')
                        || expression.encode_utf16().count() > MAX_FORMULA_UNITS
                    {
                        return Err(invalid_formula());
                    }
                    staged_bytes = staged_bytes
                        .checked_add(expression.len())
                        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
                    assignment = CellAssignment::Formula(expression.to_owned());
                    Ok(())
                },
            })
            .map_err(|error| error.with_format(FileFormat::Xls))?;
    }
    Ok(CellPlan {
        assignment,
        staged_bytes,
    })
}

/// Record one formatting assignment; each field may be named once.
fn stage_format(
    change: &mut litchi_xls::cell_values::FormatChange,
    setter: CellSetter,
    value: Option<Scalar<'_>>,
) -> Result<()> {
    use super::style;
    let invalid = |property: &'static str| {
        Error::new(
            Code::InvalidValue,
            "the XLS cell format value is not supported",
        )
        .with_format(FileFormat::Xls)
        .with_property(property)
    };
    let duplicate = |property: &'static str| {
        Error::new(
            Code::InvalidArgument,
            "the same XLS cell property was supplied more than once",
        )
        .with_format(FileFormat::Xls)
        .with_property(property)
    };
    let token = |value: &Option<Scalar<'_>>| match value {
        Some(Scalar::Token(token)) => Some(*token),
        _ => None,
    };
    macro_rules! assign {
        ($field:ident, $property:literal, $value:expr) => {{
            let value = $value.ok_or_else(|| invalid($property))?;
            if change.$field.replace(value).is_some() {
                return Err(duplicate($property));
            }
        }};
    }
    match setter {
        CellSetter::Horizontal => assign!(
            horizontal,
            "alignment.horizontal",
            token(&value).and_then(style::horizontal)
        ),
        CellSetter::Vertical => assign!(
            vertical,
            "alignment.vertical",
            token(&value).and_then(style::vertical)
        ),
        CellSetter::WrapText => assign!(
            wrap_text,
            "alignment.wrapText",
            match value {
                Some(Scalar::Bool(value)) => Some(value),
                _ => None,
            }
        ),
        CellSetter::Indent => assign!(
            indent,
            "alignment.indent",
            match value {
                Some(Scalar::Int(value)) => u8::try_from(value).ok().filter(|value| *value <= 15),
                _ => None,
            }
        ),
        CellSetter::ShrinkToFit => assign!(
            shrink_to_fit,
            "alignment.shrinkToFit",
            match value {
                Some(Scalar::Bool(value)) => Some(value),
                _ => None,
            }
        ),
        CellSetter::BorderLeft => {
            assign!(border_left, "border.left", style::border(token(&value)))
        },
        CellSetter::BorderRight => {
            assign!(border_right, "border.right", style::border(token(&value)))
        },
        CellSetter::BorderTop => assign!(border_top, "border.top", style::border(token(&value))),
        CellSetter::BorderBottom => {
            assign!(border_bottom, "border.bottom", style::border(token(&value)))
        },
        CellSetter::FillPattern => assign!(
            fill_pattern,
            "fillPattern",
            token(&value).and_then(style::fill_pattern)
        ),
        CellSetter::Fill => {
            assign!(fill_foreground, "fill", style::rgb(value));
            // A bare fill color is a solid fill, as in the XLSX adapter.
            if change.fill_pattern.is_none() {
                change.fill_pattern = Some(litchi_xls::FillPattern::Solid);
            }
        },
        CellSetter::FillBg => assign!(fill_background, "fillBg", style::rgb(value)),
        CellSetter::Value | CellSetter::Clear | CellSetter::Formula => {},
    }
    Ok(())
}

fn native_scalar(value: Scalar<'_>) -> Result<NativeValue> {
    match value {
        Scalar::Bool(value) => Ok(NativeValue::Boolean(value)),
        Scalar::Text(value) => {
            validate_cell_text(value)?;
            Ok(NativeValue::Text(value.to_string()))
        },
        Scalar::Int(value) => {
            let number = value as f64;
            if number.is_finite() && number as i128 == i128::from(value) {
                Ok(NativeValue::Number(number))
            } else {
                Err(invalid_scalar())
            }
        },
        Scalar::Number(value) if valid_number(value) => Ok(NativeValue::Number(value)),
        Scalar::Number(_) => Err(invalid_scalar()),
        Scalar::Token(_) | Scalar::Length(_) | Scalar::Color(_) | Scalar::List(_) => {
            Err(invalid_scalar())
        },
    }
}

fn stage_cell_value(
    assignment: &mut CellAssignment,
    staged_bytes: &mut usize,
    value: NativeValue,
    limits: &SemanticLimits,
) -> Result<()> {
    if !matches!(assignment, CellAssignment::Noop) {
        return Err(duplicate_cell_property("value"));
    }
    *staged_bytes = staged_bytes
        .checked_add(native_value_bytes(&value))
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    *assignment = CellAssignment::Value(value);
    Ok(())
}

fn valid_number(value: f64) -> bool {
    value.is_normal() || value.to_bits() == 0
}

fn validate_cell_text(value: &str) -> Result<()> {
    if value.encode_utf16().count() > 65_535
        || value
            .chars()
            .any(|character| !is_xml_10_character(character))
    {
        return Err(invalid_scalar());
    }
    Ok(())
}

const fn is_xml_10_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
        || (character >= '\u{20}' && character <= '\u{d7ff}')
        || (character >= '\u{e000}' && character <= '\u{fffd}')
        || (character >= '\u{10000}' && character <= '\u{10ffff}')
}

fn native_value_bytes(value: &NativeValue) -> usize {
    match value {
        NativeValue::Text(value) => value.len(),
        _ => 8,
    }
}

fn resolve_sheet_targets(
    file: &Xls,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<usize>> {
    let mut positions = Vec::new();
    let mut seen = HashSet::new();
    if let Some(path) = target.path() {
        let node = crate::path::resolve::get(&Nav::new(file), path)?;
        append_sheet(&mut positions, &mut seen, node, limits)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(file), query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(invalid_query());
        }
        for node in walk {
            append_sheet(&mut positions, &mut seen, node?, limits)?;
        }
    } else {
        return Err(empty_target());
    }
    file.check_source()?;
    Ok(positions)
}

fn append_sheet(
    positions: &mut Vec<usize>,
    seen: &mut HashSet<usize>,
    node: Node<'_>,
    limits: &SemanticLimits,
) -> Result<()> {
    if node.element() != Kind::Sheet {
        return Err(Error::new(
            Code::UnsupportedElement,
            "an XLS worksheet write target must be a sheet",
        )
        .with_format(FileFormat::Xls));
    }
    let position = node.sheet_position();
    if seen.contains(&position) {
        return Ok(());
    }
    let count = positions
        .len()
        .checked_add(1)
        .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
    if count > limits.max_expanded_targets() {
        return Err(expanded_limit(count, limits.max_expanded_targets()));
    }
    positions
        .try_reserve(1)
        .map_err(|_| allocation("XLS worksheet targets"))?;
    seen.try_reserve(1)
        .map_err(|_| allocation("XLS worksheet target identities"))?;
    seen.insert(position);
    positions.push(position);
    Ok(())
}

fn resolve_cell_targets(
    file: &Xls,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<CellTarget>> {
    let mut targets = Vec::new();
    let mut seen = HashSet::new();
    let mut non_cell = false;
    if let Some(path) = target.path() {
        let node = crate::path::resolve::get(&Nav::new(file), path)?;
        append_cell_node(&mut targets, &mut seen, &mut non_cell, node, limits)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(file), query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(invalid_query());
        }
        for node in walk {
            append_cell_node(&mut targets, &mut seen, &mut non_cell, node?, limits)?;
        }
    } else {
        return Err(empty_target());
    }
    file.check_source()?;
    if non_cell {
        return Err(Error::new(
            Code::UnsupportedElement,
            "an XLS cell write target must be a cell or range",
        )
        .with_format(FileFormat::Xls));
    }
    Ok(targets)
}

fn resolve_removal_targets(
    file: &Xls,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<CellTarget>> {
    resolve_cell_targets(file, target, limits).map_err(|error| {
        if error.code() == Code::UnsupportedElement {
            Error::new(
                Code::UnsupportedCapability,
                "XLS removal supports worksheet tabs, empty rows and columns, stored cells and ranges",
            )
            .with_format(FileFormat::Xls)
        } else {
            error
        }
    })
}

fn append_cell_node(
    targets: &mut Vec<CellTarget>,
    seen: &mut HashSet<(usize, Reference)>,
    non_cell: &mut bool,
    node: Node<'_>,
    limits: &SemanticLimits,
) -> Result<()> {
    match node.element() {
        Kind::Cell => {
            let Some((column, row)) = node.cell() else {
                return Err(
                    Error::new(Code::InvalidPath, "the XLS cell reference is invalid")
                        .with_format(FileFormat::Xls),
                );
            };
            let reference = Reference::new(row, column).map_err(native_value_error)?;
            let (present, formula) = cell_state_facts(node.cell_state()?);
            append_cell_target(
                targets,
                seen,
                node.sheet_position(),
                reference,
                present,
                formula,
                limits,
            )
        },
        Kind::Range => {
            let Item::Range(range) = node.item() else {
                return Err(Error::new(Code::InvalidPath, "the XLS range is invalid")
                    .with_format(FileFormat::Xls));
            };
            let (first_row, first_column, last_row, last_column) = range.rect;
            let rows = last_row.saturating_sub(first_row).saturating_add(1) as usize;
            let columns = last_column.saturating_sub(first_column).saturating_add(1) as usize;
            let area = rows
                .checked_mul(columns)
                .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
            let total = targets.len().checked_add(area).unwrap_or(usize::MAX);
            if area > limits.max_expanded_targets() || total > limits.max_expanded_targets() {
                return Err(expanded_limit(total, limits.max_expanded_targets()));
            }
            targets
                .try_reserve(area)
                .map_err(|_| allocation("XLS range cell targets"))?;
            seen.try_reserve(area)
                .map_err(|_| allocation("XLS range cell identities"))?;
            for row in first_row..=last_row {
                for column in first_column..=last_column {
                    let reference = Reference::new(row, column).map_err(native_value_error)?;
                    let (present, formula) = range
                        .grid
                        .record((row, column))
                        .map_or((false, false), |cell| (true, cell.formula));
                    append_cell_target(
                        targets,
                        seen,
                        node.sheet_position(),
                        reference,
                        present,
                        formula,
                        limits,
                    )?;
                }
            }
            Ok(())
        },
        _ => {
            *non_cell = true;
            Ok(())
        },
    }
}

fn append_cell_target(
    targets: &mut Vec<CellTarget>,
    seen: &mut HashSet<(usize, Reference)>,
    sheet: usize,
    reference: Reference,
    present: bool,
    formula: bool,
    limits: &SemanticLimits,
) -> Result<()> {
    if seen.contains(&(sheet, reference)) {
        return Ok(());
    }
    let count = targets
        .len()
        .checked_add(1)
        .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
    if count > limits.max_expanded_targets() {
        return Err(expanded_limit(count, limits.max_expanded_targets()));
    }
    targets
        .try_reserve(1)
        .map_err(|_| allocation("XLS cell targets"))?;
    seen.try_reserve(1)
        .map_err(|_| allocation("XLS cell target identities"))?;
    seen.insert((sheet, reference));
    targets.push(CellTarget {
        sheet,
        reference,
        present,
        formula,
    });
    Ok(())
}

fn cell_state_facts(state: CellState<'_>) -> (bool, bool) {
    match state {
        CellState::Stored(cell) => (true, cell.formula),
        CellState::Missing | CellState::Stub => (false, false),
    }
}

pub(super) fn finish_commit(
    file: &Xls,
    native: litchi_xls::Commit,
    fallback_operations: usize,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.check_source_for_write()?;
    let changed = native.diagnostics().changed();
    // Native diagnostics count physical edits and therefore become zero for
    // an exact signed identity patch. Keep the authored leaf count as the
    // fallback even when the final package is unchanged.
    let operations = native
        .diagnostics()
        .operation_count()
        .max(fallback_operations);
    let native_snapshot_bytes = native.patch().retained_bytes();
    if native_snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(
            native_snapshot_bytes,
            limits.max_snapshot_bytes(),
        ));
    }
    let (package, patch) = native.into_parts();
    if !changed && file.signature() == crate::format::signature::Signature::Signed {
        // The write owner is an unsigned private copy. Rebuild the identity
        // patch over the retained signed package so native replay and the
        // facade compare the same physical base bytes.
        let original = file.original_package_for_noop()?;
        let minimum = file
            .retained_source_len()
            .saturating_mul(2)
            .saturating_add(size_of::<litchi_xls::Patch>())
            .saturating_add(size_of::<litchi_xls::package::edit::Edit>());
        if minimum > limits.max_snapshot_bytes() {
            return Err(snapshot_limit(minimum, limits.max_snapshot_bytes()));
        }
        let (package, patch) = original
            .edit()
            .map_err(native_edit_error)?
            .commit()
            .map_err(native_edit_error)?
            .into_parts();
        return finish_parts(
            file,
            package,
            patch,
            operations,
            fallback_operations,
            native_snapshot_bytes,
            limits,
        );
    }
    finish_parts(
        file,
        package,
        patch,
        operations,
        fallback_operations,
        native_snapshot_bytes,
        limits,
    )
}

fn finish_parts(
    file: &Xls,
    package: litchi_xls::Package,
    patch: litchi_xls::Patch,
    operations: usize,
    fallback_operations: usize,
    prior_snapshot_bytes: usize,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    let changed = !patch.is_noop();
    // The public operation bound counts mutation calls, not the native
    // resource/cell operations needed to implement each call.
    if fallback_operations > limits.max_operations() {
        return Err(operation_limit(
            fallback_operations,
            limits.max_operations(),
        ));
    }
    let snapshot_bytes = prior_snapshot_bytes.max(patch.retained_bytes());
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let after = if changed {
        file.with_package(package)?
    } else {
        file.clone()
    };
    file.check_source_for_write()?;
    Ok(WriteCommit {
        before: file.clone(),
        after,
        patch: StepPatch::Primary(patch),
        operations,
        snapshot_bytes,
    })
}

fn no_op(file: &Xls, limits: &SemanticLimits) -> Result<WriteCommit> {
    let native = file
        .editable_package_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?
        .edit()
        .map_err(native_edit_error)?
        .commit()
        .map_err(native_edit_error)?;
    finish_commit(file, native, 0, limits)
}

fn duplicate_cell_property(property: &'static str) -> Error {
    Error::new(
        Code::InvalidArgument,
        "the same XLS cell property was supplied more than once",
    )
    .with_format(FileFormat::Xls)
    .with_property(property)
}

fn duplicate_sheet_property() -> Error {
    Error::new(
        Code::InvalidArgument,
        "the same XLS sheet property was supplied more than once",
    )
    .with_format(FileFormat::Xls)
    .with_property("name")
}

/// Excel's formula-text bound (8,192 characters); the BIFF8 token bound is
/// checked by the native tokenizer.
const MAX_FORMULA_UNITS: usize = 8_192;

fn invalid_formula() -> Error {
    Error::new(
        Code::InvalidValue,
        "an XLS formula requires nonempty expression text without a leading '='",
    )
    .with_format(FileFormat::Xls)
    .with_property("formula")
}

/// Map a native format refusal: a color outside the workbook palette is an
/// invalid value for this workbook.
fn format_error(error: litchi_xls::package::edit::Error) -> Error {
    use litchi_xls::package::edit::Error as Native;
    match &error {
        Native::CellValues(litchi_xls::Error::UnsupportedFeature(_)) => Error::new(
            Code::InvalidValue,
            "the XLS fill color is not an entry of the workbook palette",
        )
        .with_format(FileFormat::Xls)
        .with_property("fill")
        .with_suggestion("use one of the workbook's 56 palette colors"),
        _ => native_edit_error(error),
    }
}

/// Map a native formula refusal: an expression the BIFF8 tokenizer cannot
/// encode is an invalid value, not a corrupt workbook.
fn formula_error(error: litchi_xls::package::edit::Error) -> Error {
    use litchi_xls::package::edit::Error as Native;
    match &error {
        Native::CellValues(
            litchi_xls::Error::InvalidFormula(_)
            | litchi_xls::Error::InvalidCellReference(_)
            | litchi_xls::Error::InvalidData(_),
        ) => Error::new(
            Code::InvalidValue,
            "the XLS formula expression cannot be encoded as BIFF8 tokens",
        )
        .with_format(FileFormat::Xls)
        .with_property("formula")
        .with_suggestion(
            "use same-sheet references, numbers, text, operators and built-in functions",
        ),
        _ => native_edit_error(error),
    }
}

fn invalid_scalar() -> Error {
    Error::new(
        Code::InvalidValue,
        "XLS cell values support text, bool, and finite BIFF8 numbers",
    )
    .with_format(FileFormat::Xls)
    .with_property("value")
}

fn empty_target() -> Error {
    Error::new(Code::InvalidArgument, "the mutation target is empty").with_format(FileFormat::Xls)
}

fn allocation(resource: &'static str) -> Error {
    Error::new(
        Code::InternalError,
        "could not allocate XLS mutation targets",
    )
    .with_format(FileFormat::Xls)
    .with_suggestion(resource)
}

fn expanded_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::ExpandedTargets,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xls)
}

fn staged_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::StagedBytes,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xls)
}

fn snapshot_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::SnapshotBytes,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xls)
}

fn operation_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::Operations,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xls)
}

fn native_error(error: litchi_xls::package::edit::Error) -> Error {
    Error::from(error).with_format(FileFormat::Xls)
}

pub(super) fn native_edit_error(error: litchi_xls::package::edit::Error) -> Error {
    native_error(error)
}

fn native_value_error(error: litchi_xls::Error) -> Error {
    Error::from(error).with_format(FileFormat::Xls)
}

fn invalid_query() -> Error {
    Error::new(
        Code::InvalidSelector,
        "the mutation selector contains an unsupported predicate",
    )
    .with_format(FileFormat::Xls)
}

/// Native-package coalescing for direct XLS cell mutations.
///
/// The unified target resolver remains the admission authority.  Once a
/// mutation has been proven to be one direct cell path, the run owns one
/// `Package::Edit` and appends disjoint cell operations to it.  A repeated
/// coordinate is a barrier: the following command is re-resolved only after
/// the preceding native edit has been published, which preserves insert,
/// remove, and missing-cell semantics.
/// CSV/TSV import through one native edit (record 0266).
pub(crate) mod find;
pub(crate) mod import;
#[cfg(feature = "eval")]
pub(crate) mod refresh;

pub(crate) mod coalesce {
    use super::{
        CellAssignment, CellIdentity, CellTarget, FileFormat, Kind, NativeValue, Result, TargetArg,
        WRITE_ENTRY_OVERHEAD, Xls, admit_cell_property, cell_state_facts, finish_commit,
        native_edit_error, native_value_error, snapshot_limit, staged_limit,
    };
    use crate::error::{Code, Error};
    use crate::props::Props;
    use crate::transaction::{Error as TransactionError, SemanticLimits};
    use litchi_core::Position;
    use litchi_core::sheet::CellValue;
    use litchi_xls::Package;
    use litchi_xls::cell_values::{Reference, Selector};
    use std::collections::HashSet;
    use std::mem::size_of;

    #[derive(Debug, Clone)]
    pub(crate) enum Action {
        Set(NativeValue),
        Clear,
        Remove,
        Noop,
    }

    #[derive(Debug, Clone)]
    pub(crate) struct Plan {
        target: CellTarget,
        action: Action,
        staged_bytes: usize,
        changed: bool,
        /// The authored scalar before this run.  Retaining it lets a signed
        /// run prove a cancellation without comparing reserialized BIFF.
        original: Option<NativeValue>,
        /// Whether the authored cell existed in the original package.  This
        /// stays separate from the staged target's current presence so a
        /// remove/reinsert pair cannot turn an absent coordinate into a
        /// false cancellation.
        original_present: bool,
        /// Stored errors and formulas do not have a comparable scalar value.
        /// They remain barriers for signed cancellation even when a later
        /// operation happens to produce `None`.
        original_known: bool,
    }

    impl Plan {
        fn key(&self) -> CellIdentity {
            self.target.identity()
        }

        fn charge(&self) -> usize {
            self.staged_bytes
        }
    }

    struct Item {
        plan: Plan,
        index: usize,
    }

    /// One native primary XLS package edit for a run of distinct direct cells.
    pub(crate) struct Run {
        base: Xls,
        source: Package,
        edit: litchi_xls::package::edit::Edit,
        items: Vec<Item>,
        seen: HashSet<CellIdentity>,
        staged_bytes: usize,
        max_staged_bytes: usize,
        limits: SemanticLimits,
        max_snapshot_bytes: usize,
        allow_repeated: bool,
    }

    impl Xls {
        /// Admit a direct scalar cell assignment.  Ranges, queries, sheet
        /// targets, formulas, and empty property sets remain on the ordinary
        /// format handler so their existing resolution and refusal policy is
        /// unchanged.
        pub(crate) fn cell_plan(
            &self,
            target: &TargetArg<'_>,
            props: &Props<'_>,
            limits: &SemanticLimits,
        ) -> Result<Option<Plan>> {
            self.cell_plan_with_source(target, props, limits, None)
        }

        fn cell_plan_with_source(
            &self,
            target: &TargetArg<'_>,
            props: &Props<'_>,
            limits: &SemanticLimits,
            source: Option<&litchi_xls::Edit>,
        ) -> Result<Option<Plan>> {
            self.check_outside_run(source)?;
            let result = (|| {
                let Some(path) = target.path() else {
                    return Ok(None);
                };
                let Some(step) = path.path().steps().last() else {
                    return Ok(None);
                };
                if step.cell_ref().is_none() {
                    return Ok(None);
                }
                // `name` is the worksheet setter's discriminator.  Leave
                // it to the ordinary dispatcher so a malformed sheet-cell
                // combination keeps the existing sheet-property error.
                if super::is_sheet_property(props) {
                    return Ok(None);
                }
                let admitted = admit_cell_property(props, limits)?;
                // Formula writes keep the ordinary handler: each one is a
                // removal and an authored Formula, not a scalar staging.
                if matches!(
                    admitted.assignment,
                    CellAssignment::Noop | CellAssignment::Formula(_) | CellAssignment::Format(_)
                ) {
                    return Ok(None);
                }
                let spelling = target;
                let Some((target, current)) = direct_target(self, spelling, limits, source)? else {
                    return Ok(None);
                };
                let signed = self.signature() == crate::format::signature::Signature::Signed;
                // A coalesced signed run resolves each command through its
                // staged native edit so later commands see earlier values.
                // The cancellation proof must retain the value from the
                // original package, however; otherwise a changed-then-restored
                // cell would compare against the changed intermediate value
                // and lose the exact signed no-op route.
                let (original_target, original) = if signed && source.is_some() {
                    let Some((original_target, original)) =
                        direct_target(self, spelling, limits, None)?
                    else {
                        return Ok(None);
                    };
                    (original_target, original)
                } else {
                    (target, signed.then(|| current.clone()).flatten())
                };
                // Formula cells are deliberately left to the ordinary setter,
                // which reports the dependency-aware unsupported capability.
                if target.formula {
                    return Ok(None);
                }
                if limits.max_operations() < 1 {
                    return Err(operation_limit(1, limits.max_operations()));
                }
                if admitted.staged_bytes > limits.max_staged_bytes() {
                    return Err(staged_limit(
                        admitted.staged_bytes,
                        limits.max_staged_bytes(),
                    ));
                }
                let changed = match &admitted.assignment {
                    CellAssignment::Value(value) => {
                        !target.present || current.as_ref() != Some(value)
                    },
                    CellAssignment::Clear => {
                        target.present && current.as_ref() != Some(&NativeValue::Blank)
                    },
                    CellAssignment::Formula(_)
                    | CellAssignment::Format(_)
                    | CellAssignment::Noop => false,
                };
                let action = match admitted.assignment {
                    CellAssignment::Value(value) => Action::Set(value),
                    CellAssignment::Clear => Action::Clear,
                    CellAssignment::Formula(_)
                    | CellAssignment::Format(_)
                    | CellAssignment::Noop => return Ok(None),
                };
                let original_known = !signed || !original_target.present || original.is_some();
                Ok(Some(Plan {
                    target,
                    action,
                    staged_bytes: admitted.staged_bytes,
                    changed,
                    original,
                    original_present: signed && original_target.present,
                    original_known,
                }))
            })();
            self.check_outside_run(source).and(result)
        }

        /// Admit a direct cell removal.  An absent coordinate is retained as
        /// an explicit no-op so it still consumes one unified mutation item;
        /// the native edit never receives an invalid absent-cell removal.
        pub(crate) fn remove_plan(
            &self,
            target: &TargetArg<'_>,
            limits: &SemanticLimits,
        ) -> Result<Option<Plan>> {
            self.remove_plan_with_source(target, limits, None)
        }

        /// The plans' own source checks. A coalesced run (`source` is the
        /// run's leaf edit) checks the source when it opens and around its
        /// commit, and each item's resolution and leaf `source_cell` read
        /// still observe the source, so a plan inside a run adds no check
        /// of its own (record 0337: about five `fstat`s an item on a
        /// file-backed owner).
        fn check_outside_run(&self, source: Option<&litchi_xls::Edit>) -> Result<()> {
            if source.is_some() {
                Ok(())
            } else {
                self.check_source()
            }
        }

        fn remove_plan_with_source(
            &self,
            target: &TargetArg<'_>,
            limits: &SemanticLimits,
            source: Option<&litchi_xls::Edit>,
        ) -> Result<Option<Plan>> {
            self.check_outside_run(source)?;
            let result = (|| {
                let Some(path) = target.path() else {
                    return Ok(None);
                };
                let Some(step) = path.path().steps().last() else {
                    return Ok(None);
                };
                if step.cell_ref().is_none() {
                    return Ok(None);
                }
                let spelling = target;
                let Some((target, current)) = direct_target(self, spelling, limits, source)? else {
                    return Ok(None);
                };
                let signed = self.signature() == crate::format::signature::Signature::Signed;
                let (original_target, original) = if signed && source.is_some() {
                    let Some((original_target, original)) =
                        direct_target(self, spelling, limits, None)?
                    else {
                        return Ok(None);
                    };
                    (original_target, original)
                } else {
                    (target, signed.then(|| current.clone()).flatten())
                };
                if target.formula {
                    return Ok(None);
                }
                if limits.max_operations() < 1 {
                    return Err(operation_limit(1, limits.max_operations()));
                }
                if WRITE_ENTRY_OVERHEAD > limits.max_staged_bytes() {
                    return Err(staged_limit(
                        WRITE_ENTRY_OVERHEAD,
                        limits.max_staged_bytes(),
                    ));
                }
                let original_known = !signed || !original_target.present || original.is_some();
                Ok(Some(Plan {
                    action: if target.present {
                        Action::Remove
                    } else {
                        Action::Noop
                    },
                    target,
                    staged_bytes: WRITE_ENTRY_OVERHEAD,
                    changed: target.present,
                    original,
                    original_present: signed && original_target.present,
                    original_known,
                }))
            })();
            self.check_outside_run(source).and(result)
        }
    }

    impl Run {
        pub(crate) fn new(
            base: &Xls,
            limits: &SemanticLimits,
            remaining_snapshot: usize,
        ) -> Result<Self> {
            base.check_source_for_write()?;
            let source_bytes = base.retained_source_len();
            let minimum = source_bytes
                .saturating_mul(2)
                .saturating_add(size_of::<litchi_xls::Patch>())
                .saturating_add(size_of::<litchi_xls::package::edit::Edit>());
            if minimum > remaining_snapshot {
                return Err(snapshot_limit(minimum, remaining_snapshot));
            }
            let source = base.editable_package_for_write_with_snapshot_limit(remaining_snapshot)?;
            let edit = source.edit().map_err(native_edit_error)?;
            base.check_source_for_write()?;
            Ok(Self {
                base: base.clone(),
                source,
                edit,
                items: Vec::new(),
                seen: HashSet::new(),
                staged_bytes: 0,
                max_staged_bytes: limits.max_staged_bytes(),
                limits: *limits,
                max_snapshot_bytes: remaining_snapshot,
                allow_repeated: base.signature() == crate::format::signature::Signature::Signed,
            })
        }

        pub(crate) fn cell_plan(
            &self,
            target: &TargetArg<'_>,
            props: &Props<'_>,
            limits: &SemanticLimits,
        ) -> Result<Option<Plan>> {
            self.base
                .cell_plan_with_source(target, props, limits, Some(&self.edit))
        }

        pub(crate) fn remove_plan(
            &self,
            target: &TargetArg<'_>,
            limits: &SemanticLimits,
        ) -> Result<Option<Plan>> {
            self.base
                .remove_plan_with_source(target, limits, Some(&self.edit))
        }

        pub(crate) fn first_index(&self) -> Option<usize> {
            self.items.first().map(|item| item.index)
        }

        pub(crate) fn requires_barrier(&self, plan: &Plan, limits: &SemanticLimits) -> bool {
            // A signed cancellation may repeat an originally stored cell
            // even after an intermediate remove made the staged target
            // absent. Original absence stays a barrier so insertion is never
            // mistaken for restoring a missing coordinate.
            (self.seen.contains(&plan.key()) && !(self.allow_repeated && plan.original_present))
                || self
                    .staged_bytes
                    .checked_add(plan.charge())
                    .is_none_or(|bytes| bytes > limits.max_staged_bytes())
        }

        pub(crate) fn stage(&mut self, plan: Plan, index: usize) -> Result<bool> {
            if (self.seen.contains(&plan.key()) && !(self.allow_repeated && plan.original_present))
                || self
                    .staged_bytes
                    .checked_add(plan.charge())
                    .is_none_or(|bytes| bytes > self.max_staged_bytes)
            {
                return Err(Error::new(
                    Code::InternalError,
                    "XLS cell coalescing barrier was not flushed",
                )
                .with_format(FileFormat::Xls));
            }
            let staged = self
                .staged_bytes
                .saturating_add(plan.charge())
                .saturating_add(size_of::<NativeValue>());
            let observed = self
                .base
                .retained_source_len()
                .saturating_mul(2)
                .saturating_add(size_of::<litchi_xls::Patch>())
                .saturating_add(size_of::<litchi_xls::package::edit::Edit>())
                .saturating_add(staged)
                .saturating_add(size_of::<Item>());
            if observed > self.max_snapshot_bytes {
                return Err(snapshot_limit(observed, self.max_snapshot_bytes));
            }
            // Staging reads no source: the run checks it when it opens and
            // around its commit (record 0337).
            self.items
                .try_reserve(1)
                .map_err(|source| TransactionError::Allocation { source })?;
            self.seen
                .try_reserve(1)
                .map_err(|source| TransactionError::Allocation { source })?;
            match &plan.action {
                Action::Set(value) => {
                    let selector = Selector::Position(Position::new(plan.target.sheet));
                    if plan.target.present {
                        self.edit
                            .set_value(selector, plan.target.reference, value.clone())
                            .map_err(native_edit_error)?;
                    } else {
                        self.edit
                            .insert_cell(selector, plan.target.reference, value.clone())
                            .map_err(native_edit_error)?;
                    }
                },
                Action::Clear => {
                    if plan.target.present {
                        self.edit
                            .set_value(
                                Selector::Position(Position::new(plan.target.sheet)),
                                plan.target.reference,
                                NativeValue::Blank,
                            )
                            .map_err(native_edit_error)?;
                    }
                },
                Action::Remove => {
                    self.edit
                        .remove_cell(
                            Selector::Position(Position::new(plan.target.sheet)),
                            plan.target.reference,
                        )
                        .map_err(native_edit_error)?;
                },
                Action::Noop => {},
            }
            self.seen.insert(plan.key());
            self.staged_bytes = self.staged_bytes.saturating_add(plan.charge());
            let changed = plan.changed;
            self.items.push(Item { plan, index });
            Ok(changed)
        }

        pub(crate) fn commit(self) -> Result<(Xls, super::WriteCommit)> {
            let Self {
                base,
                source,
                edit,
                items,
                limits,
                max_snapshot_bytes,
                allow_repeated,
                ..
            } = self;
            let snapshot_limits = limits.with_snapshot_budget(max_snapshot_bytes);
            let logical_noop = if allow_repeated {
                logical_noop(&items)
            } else {
                false
            };
            let original = if logical_noop {
                Some(base.original_package_for_noop()?)
            } else {
                None
            };
            base.check_source_for_write()?;
            let result = edit
                .commit()
                .map_err(native_edit_error)
                .and_then(|native| finish_commit(&base, native, items.len(), &snapshot_limits))
                .and_then(|commit| {
                    if logical_noop {
                        let staged_snapshot_bytes = commit.snapshot_bytes;
                        let original = original
                            .as_ref()
                            .ok_or_else(|| Error::signed(FileFormat::Xls))?;
                        let empty = original
                            .edit()
                            .map_err(native_edit_error)?
                            .commit()
                            .map_err(native_edit_error)?;
                        let mut commit =
                            finish_commit(&base, empty, items.len(), &snapshot_limits)?;
                        commit.operations = items.len();
                        commit.snapshot_bytes = commit.snapshot_bytes.max(staged_snapshot_bytes);
                        Ok(commit)
                    } else {
                        Ok(commit)
                    }
                });
            let freshness = base.check_source_for_write();
            match freshness.and(result) {
                Ok(commit) => Ok((base, commit)),
                Err(error) => Err(locate_failure(&base, &source, &items, error)),
            }
        }
    }

    fn logical_noop(items: &[Item]) -> bool {
        type State = (bool, bool, Option<NativeValue>);
        let mut final_values = std::collections::HashMap::<CellIdentity, State>::new();
        let mut originals = std::collections::HashMap::<CellIdentity, State>::new();
        for item in items {
            let plan = &item.plan;
            let key = plan.key();
            let original = originals.entry(key).or_insert_with(|| {
                (
                    plan.original_present,
                    plan.original_known,
                    plan.original.clone(),
                )
            });
            let state = final_values.entry(key).or_insert_with(|| original.clone());
            match &plan.action {
                Action::Set(value) => *state = (true, true, Some(value.clone())),
                Action::Clear if plan.target.present => {
                    *state = (true, true, Some(NativeValue::Blank));
                },
                Action::Remove if plan.target.present => *state = (false, true, None),
                Action::Clear | Action::Remove | Action::Noop => {},
            }
        }
        !final_values.is_empty()
            && final_values
                .into_iter()
                .all(|(key, value)| value.1 && originals.get(&key) == Some(&value))
    }

    fn apply_item(edit: &mut litchi_xls::package::edit::Edit, item: &Item) -> Result<()> {
        match &item.plan.action {
            Action::Set(value) => {
                let selector = Selector::Position(Position::new(item.plan.target.sheet));
                if item.plan.target.present {
                    edit.set_value(selector, item.plan.target.reference, value.clone())
                        .map_err(native_edit_error)?;
                } else {
                    edit.insert_cell(selector, item.plan.target.reference, value.clone())
                        .map_err(native_edit_error)?;
                }
            },
            Action::Clear if item.plan.target.present => edit
                .set_value(
                    Selector::Position(Position::new(item.plan.target.sheet)),
                    item.plan.target.reference,
                    NativeValue::Blank,
                )
                .map_err(native_edit_error)?,
            Action::Remove if item.plan.target.present => edit
                .remove_cell(
                    Selector::Position(Position::new(item.plan.target.sheet)),
                    item.plan.target.reference,
                )
                .map_err(native_edit_error)?,
            Action::Clear | Action::Remove | Action::Noop => {},
        }
        Ok(())
    }

    fn replay(base: &Xls, source: &Package, items: &[Item]) -> Result<Package> {
        base.check_source_for_write()?;
        let mut edit = source.edit().map_err(native_edit_error)?;
        for item in items {
            apply_item(&mut edit, item).map_err(|error| error.with_item(item.index))?;
        }
        let package = edit.commit().map_err(native_edit_error)?.into_package();
        base.check_source_for_write()?;
        Ok(package)
    }

    fn locate_failure(base: &Xls, source: &Package, items: &[Item], error: Error) -> Error {
        let Some(first) = items.first() else {
            return error;
        };
        if items.len() == 1
            || matches!(
                error.code(),
                Code::IoError
                    | Code::Cancelled
                    | Code::LimitExceeded
                    | Code::InternalError
                    | Code::DocumentProtected
            )
        {
            return error.with_item(first.index);
        }
        let middle = items.len() / 2;
        let prefix = match replay(base, source, &items[..middle]) {
            Ok(package) => package,
            Err(error) => return locate_failure(base, source, &items[..middle], error),
        };
        match replay(base, &prefix, &items[middle..]) {
            Err(error) => locate_failure(base, &prefix, &items[middle..], error),
            Ok(_) => error.with_item(first.index),
        }
    }

    fn operation_limit(observed: usize, maximum: usize) -> Error {
        Error::from(TransactionError::Limit {
            kind: crate::transaction::LimitKind::Operations,
            observed,
            maximum,
        })
        .with_format(FileFormat::Xls)
    }

    fn direct_target(
        file: &Xls,
        target: &TargetArg<'_>,
        limits: &SemanticLimits,
        source: Option<&litchi_xls::Edit>,
    ) -> Result<Option<(CellTarget, Option<NativeValue>)>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        let Some(step) = path.path().steps().last() else {
            return Ok(None);
        };
        if step.cell_ref().is_none() {
            return Ok(None);
        }
        if limits.max_expanded_targets() == 0 {
            return Err(super::expanded_limit(1, limits.max_expanded_targets()));
        }
        let nav = if source.is_some() {
            super::Nav::coordinates(file)
        } else {
            super::Nav::new(file)
        };
        let node = crate::path::resolve::get(&nav, path)?;
        if node.element() != Kind::Cell {
            return Ok(None);
        }
        let Some((column, row)) = node.cell() else {
            return Err(
                Error::new(Code::InvalidPath, "the XLS cell reference is invalid")
                    .with_format(FileFormat::Xls),
            );
        };
        let reference = Reference::new(row, column).map_err(native_value_error)?;
        let (present, formula, current) = if let Some(source) = source {
            let cell = source
                .source_cell(
                    Selector::Position(Position::new(node.sheet_position())),
                    reference,
                )
                .map_err(native_edit_error)?;
            (
                cell.is_some(),
                cell.is_some_and(|cell| {
                    cell.storage() == litchi_xls::cell_values::Storage::Formula
                }),
                cell.map(|cell| cell.value().clone()),
            )
        } else {
            let state = node.cell_state()?;
            let (present, formula) = cell_state_facts(state);
            let current = match state {
                super::CellState::Stored(cell) => native_value(&cell.value),
                super::CellState::Missing | super::CellState::Stub => None,
            };
            (present, formula, current)
        };
        // Inside a run the leaf `source_cell` read above checked the source.
        if source.is_none() {
            file.check_source()?;
        }
        Ok(Some((
            CellTarget {
                sheet: node.sheet_position(),
                reference,
                present,
                formula,
            },
            current,
        )))
    }

    fn native_value(value: &CellValue) -> Option<NativeValue> {
        match value {
            CellValue::Empty => Some(NativeValue::Blank),
            CellValue::Bool(value) => Some(NativeValue::Boolean(*value)),
            CellValue::Int(value) => Some(NativeValue::Number(*value as f64)),
            CellValue::Float(value) | CellValue::DateTime(value) => {
                Some(NativeValue::Number(*value))
            },
            CellValue::String(value) => Some(NativeValue::Text(value.clone())),
            CellValue::Error(_) | CellValue::Formula { .. } => None,
        }
    }
}
