//! XLSX rich-text runs (record 0338): `/<sheet>/<A1>/run[N]` (also
//! `r[N]`), OfficeCLI's `run` element.
//!
//! A run is one `r` of a cell's rich string: an inline rich string or a
//! shared-string item with runs, read with the leaf's
//! [`Worksheet::rich_text`](litchi_xlsx::Worksheet::rich_text). Its keys are
//! OfficeCLI's (`ExcelHandler.Helpers.Node.cs` `RunToNode`): `bold`,
//! `italic`, `strike`, `underline`, `superscript`, `subscript`, `size`,
//! `color` and `font`; `text` is the node text and write-only as a key.
//!
//! Writes go through [`litchi_xlsx::RichRun::apply_font`], which keeps every
//! run property a change does not name, and publish the cell's runs as its
//! own inline rich string ([`litchi_xlsx::workbook::edit::Sheet::set_rich_text`]),
//! as find/replace does (record 0295). `set` on a run changes it, `add`
//! under a cell inserts one (a plain text cell becomes rich, its text the
//! first run, as OfficeCLI's `AddRun` does), `remove` deletes one; as
//! OfficeCLI's removal, no run left is empty text and one unformatted run
//! left is plain text.
//!
//! A cell's `runs` property (`add`/`set` with `type=richtext`) is a JSON
//! array of run objects with the same keys (`ApplyRichTextToCell`).

use std::sync::Arc;

use litchi_core::FileFormat;
use litchi_xlsx::{self as xlsx, Address};

use super::node::{Item, Node, RunItem, SheetHandle};
use super::write::{Commit, edit_with_snapshot_limit};
use super::{Nav, Xlsx};
use crate::error::{Code, Error, Result};
use crate::node::Value;
use crate::node::spec::PropSpec;
use crate::path::navigate::{Lookup, Select};
use crate::props::{PropValue, Props};
use crate::registry::{Access, Alias, Constraint, Entry, Operation, Table};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;
use crate::value::{ColorContext, Length, LengthSpec, Scalar, ValueType};

/// The run keys a write admits.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum RunSetter {
    Text,
    Bold,
    Italic,
    Strike,
    Underline,
    Superscript,
    Subscript,
    Size,
    Color,
    Font,
}

const fn both(name: &'static str, ty: ValueType, setter: RunSetter) -> Entry<RunSetter> {
    Entry {
        spec: PropSpec::new(name, ty),
        access: Access::SetAndAdd(setter),
        constraint: Constraint::None,
        clear: true,
    }
}

/// The run registry: read keys in OfficeCLI's `RunToNode` order, `text`
/// write-only.
pub(crate) const RUN_WRITE_TABLE: Table<RunSetter, 10, 9> = Table::new(
    [
        Entry {
            spec: {
                let mut spec = PropSpec::new("text", ValueType::Text);
                spec.access.get = false;
                spec
            },
            access: Access::SetAndAdd(RunSetter::Text),
            constraint: Constraint::None,
            clear: false,
        },
        both("bold", ValueType::Bool, RunSetter::Bold),
        both("italic", ValueType::Bool, RunSetter::Italic),
        both("strike", ValueType::Bool, RunSetter::Strike),
        both(
            "underline",
            ValueType::Token(&super::spec::XLSX_UNDERLINE),
            RunSetter::Underline,
        ),
        both("superscript", ValueType::Bool, RunSetter::Superscript),
        both("subscript", ValueType::Bool, RunSetter::Subscript),
        Entry {
            spec: PropSpec::new("size", ValueType::Length(LengthSpec::FONT_SIZE)),
            access: Access::SetAndAdd(RunSetter::Size),
            constraint: Constraint::Length {
                min: Length::ZERO,
                max: Length::from_points(409),
            },
            clear: true,
        },
        both(
            "color",
            ValueType::Color(ColorContext::Spreadsheet),
            RunSetter::Color,
        ),
        Entry {
            spec: PropSpec::new("font", ValueType::Text),
            access: Access::SetAndAdd(RunSetter::Font),
            constraint: Constraint::NotEmpty,
            clear: true,
        },
    ],
    [
        Alias {
            name: "font.bold",
            entry: 1,
        },
        Alias {
            name: "font.italic",
            entry: 2,
        },
        Alias {
            name: "font.strike",
            entry: 3,
        },
        Alias {
            name: "font.underline",
            entry: 4,
        },
        Alias {
            name: "font.size",
            entry: 7,
        },
        Alias {
            name: "fontsize",
            entry: 7,
        },
        Alias {
            name: "font.color",
            entry: 8,
        },
        Alias {
            name: "font.name",
            entry: 9,
        },
        Alias {
            name: "fontname",
            entry: 9,
        },
    ],
);

/// The run element's properties (read and capability table).
pub(crate) const RUN_PROPS: &[PropSpec] = &RUN_WRITE_TABLE.readable();

/// The run read keys, by their index in [`RUN_PROPS`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RunKey {
    Text,
    Bold,
    Italic,
    Strike,
    Underline,
    Superscript,
    Subscript,
    Size,
    Color,
    Font,
}

impl RunKey {
    const ALL: [Self; 10] = [
        Self::Text,
        Self::Bold,
        Self::Italic,
        Self::Strike,
        Self::Underline,
        Self::Superscript,
        Self::Subscript,
        Self::Size,
        Self::Color,
        Self::Font,
    ];
}

/// The value of the run key at `index`.
pub(super) fn value(font: &xlsx::style::stylesheet::font::Font, index: usize) -> Option<Value<'_>> {
    use xlsx::style::{Script, Underline};
    Some(match RunKey::ALL.get(index)? {
        RunKey::Text => return None,
        RunKey::Bold => return font.bold.then_some(Value::Bool(true)),
        RunKey::Italic => return font.italic.then_some(Value::Bool(true)),
        RunKey::Strike => return font.strike.then_some(Value::Bool(true)),
        RunKey::Underline => Value::Token(match font.underline? {
            Underline::None => return None,
            Underline::Double => "double",
            Underline::SingleAccounting => "singleAccounting",
            Underline::DoubleAccounting => "doubleAccounting",
            Underline::Single => "single",
        }),
        RunKey::Superscript => {
            return matches!(font.script?, Script::Superscript).then_some(Value::Bool(true));
        },
        RunKey::Subscript => {
            return matches!(font.script?, Script::Subscript).then_some(Value::Bool(true));
        },
        RunKey::Size => Value::Length(super::node::points(font.size?)),
        RunKey::Color => super::node::font_color_value(font.color.as_ref()?)?,
        RunKey::Font => Value::Text(std::borrow::Cow::Borrowed(font.name.as_deref()?)),
    })
}

/// The rich-text runs of the cell at `at` of the sheet `handle`, or `None`
/// for a cell without runs.
pub(super) fn rich_text(handle: &SheetHandle, at: Address) -> Result<Option<xlsx::RichText>> {
    Ok(match handle {
        SheetHandle::Eager { sheet, .. } => sheet.rich_text(at)?,
        SheetHandle::Source { sheet } => sheet.rich_text(at)?,
    })
}

/// The run child of a cell node a path step selects.
pub(super) fn lookup<'f>(
    file: &'f Xlsx,
    cell: &Node<'f>,
    select: Select<'_>,
) -> Result<Lookup<Node<'f>, crate::node::spec::GapSpec>> {
    let Some((sheet, at)) = cell.write_cell_target() else {
        return Ok(Lookup::NotApplicable);
    };
    let Some(handle) = SheetHandle::at(file, sheet)? else {
        return Ok(Lookup::Missing);
    };
    let Some(rich) = rich_text(&handle, at)? else {
        return Ok(Lookup::Missing);
    };
    let count = rich.runs().len();
    let index = match select {
        Select::Ordinal(position) => Some(position.get()),
        Select::None => Some(0),
        Select::Last => count.checked_sub(1),
        _ => return Ok(Lookup::NotApplicable),
    };
    Ok(match index.filter(|index| *index < count) {
        Some(index) => Lookup::Found(Node::new_run(file, sheet, at, index, Arc::new(rich))),
        None => Lookup::Missing,
    })
}

/// The runs of a cell node, for a query walk that wants runs: `None` for a
/// cell without runs. Only a text cell is read for runs.
pub(super) fn cell_runs(cell: &Node<'_>) -> Result<Option<(usize, Address, Arc<xlsx::RichText>)>> {
    let Some((sheet, at)) = cell.write_cell_target() else {
        return Ok(None);
    };
    if !matches!(
        cell.cell_state()?,
        Some(super::node::CellState::Stored(xlsx::Cell::Value(
            xlsx::Value::Text(_)
        )))
    ) {
        return Ok(None);
    }
    let Some(handle) = SheetHandle::at(cell.file(), sheet)? else {
        return Ok(None);
    };
    Ok(rich_text(&handle, at)?.map(|rich| (sheet, at, Arc::new(rich))))
}

/// Admitted run assignments: the text and a font change.
#[derive(Debug, Default)]
pub(super) struct RunAssignments {
    text: Option<String>,
    font: xlsx::FontChange,
}

impl RunAssignments {
    /// Apply the assignments to `run`.
    fn apply(&self, run: &mut xlsx::RichRun) -> Result<()> {
        if let Some(text) = &self.text {
            run.set_text(text.clone());
        }
        run.apply_font(&self.font)
            .map_err(super::write::native_error)
    }

    /// A new run with the assignments.
    pub(super) fn run(&self) -> Result<xlsx::RichRun> {
        let mut run = xlsx::RichRun::new(self.text.clone().unwrap_or_default());
        run.apply_font(&self.font)
            .map_err(super::write::native_error)?;
        Ok(run)
    }
}

/// Admit run properties for `operation`.
pub(super) fn admit(props: &Props<'_>, operation: Operation) -> Result<RunAssignments> {
    let mut out = RunAssignments::default();
    let mut seen = [false; 10];
    let mut script: Option<RunSetter> = None;
    for (key, value) in props.iter() {
        let (index, entry) = RUN_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        let name = entry.spec.name;
        if std::mem::replace(&mut seen[index], true) {
            return Err(Error::new(
                Code::InvalidArgument,
                "the same XLSX run property was supplied more than once",
            )
            .with_property(name));
        }
        entry.prepare(operation, value, |setter, scalar| {
            let flag = |scalar: Option<Scalar<'_>>| match scalar {
                None => Ok(None),
                Some(Scalar::Bool(value)) => Ok(Some(value)),
                _ => Err(invalid(name, "run flags must be boolean or null")),
            };
            let font = std::mem::take(&mut out.font);
            out.font = match setter {
                RunSetter::Text => {
                    let Some(Scalar::Text(text)) = scalar else {
                        return Err(invalid(name, "run text must be text"));
                    };
                    super::write::validate_text(text, name)?;
                    out.text = Some(text.to_owned());
                    font
                },
                RunSetter::Bold => font.bold(flag(scalar)?),
                RunSetter::Italic => font.italic(flag(scalar)?),
                RunSetter::Strike => font.strike(flag(scalar)?),
                RunSetter::Underline => font.underline(match scalar {
                    None => None,
                    Some(Scalar::Token(token)) => Some(underline(token, name)?),
                    _ => return Err(invalid(name, "underline must be a token or null")),
                }),
                RunSetter::Superscript | RunSetter::Subscript => {
                    if script.replace(setter).is_some() {
                        return Err(Error::new(
                            Code::InvalidArgument,
                            "superscript and subscript cannot be supplied together",
                        )
                        .with_property(name));
                    }
                    let shifted = if setter == RunSetter::Superscript {
                        xlsx::style::Script::Superscript
                    } else {
                        xlsx::style::Script::Subscript
                    };
                    font.script(flag(scalar)?.map(|on| {
                        if on {
                            shifted
                        } else {
                            xlsx::style::Script::Baseline
                        }
                    }))
                },
                RunSetter::Size => font
                    .size(match scalar {
                        None => None,
                        Some(Scalar::Length(length)) => Some(length.points()),
                        _ => return Err(invalid(name, "run size must be a length or null")),
                    })
                    .map_err(|error| super::write::native_property_error(error, name))?,
                RunSetter::Color => font.color(match scalar {
                    None => None,
                    Some(Scalar::Color(color)) => {
                        Some(super::write::native_font_color(color, name)?)
                    },
                    _ => return Err(invalid(name, "run color must be a color or null")),
                }),
                RunSetter::Font => font
                    .name(match scalar {
                        None => None,
                        Some(Scalar::Text(text)) => Some(text),
                        _ => return Err(invalid(name, "run font must be text or null")),
                    })
                    .map_err(|error| super::write::native_property_error(error, name))?,
            };
            Ok(())
        })?;
    }
    Ok(out)
}

fn underline(token: &str, name: &'static str) -> Result<xlsx::style::Underline> {
    use xlsx::style::Underline;
    Ok(match token {
        "single" => Underline::Single,
        "double" => Underline::Double,
        "singleAccounting" => Underline::SingleAccounting,
        "doubleAccounting" => Underline::DoubleAccounting,
        "none" => Underline::None,
        _ => return Err(invalid(name, "underline has an unsupported token")),
    })
}

fn invalid(name: &'static str, message: &'static str) -> Error {
    Error::new(Code::InvalidValue, message)
        .with_property(name)
        .with_format(FileFormat::Xlsx)
}

/// Parse a cell's `runs` value: a JSON array of run objects, each with
/// `text` and run keys (OfficeCLI's `ApplyRichTextToCell`).
pub(super) fn parse_runs(value: &str, key: &str) -> Result<xlsx::RichText> {
    let refused = |message: String| {
        Error::new(Code::InvalidValue, message)
            .with_property(key.to_owned())
            .with_format(FileFormat::Xlsx)
    };
    let parsed: serde_json::Value = serde_json::from_str(value)
        .map_err(|error| refused(format!("`runs` is not JSON: {error}")))?;
    let serde_json::Value::Array(items) = parsed else {
        return Err(refused(
            "`runs` must be a JSON array of run objects".to_owned(),
        ));
    };
    if items.is_empty() {
        return Err(refused("`runs` needs at least one run".to_owned()));
    }
    let mut runs = Vec::with_capacity(items.len());
    for item in &items {
        let serde_json::Value::Object(object) = item else {
            return Err(refused(
                "each run in `runs` must be a JSON object".to_owned(),
            ));
        };
        let mut props = Props::new();
        for (name, value) in object {
            let value = match value {
                serde_json::Value::Null => PropValue::Null,
                serde_json::Value::Bool(value) => PropValue::Bool(*value),
                serde_json::Value::Number(number) => match number.as_i64() {
                    Some(value) => PropValue::Int(value),
                    None => PropValue::Number(number.as_f64().unwrap_or(f64::NAN)),
                },
                serde_json::Value::String(text) => PropValue::from(text.as_str()),
                _ => {
                    return Err(refused(format!(
                        "run key `{name}` in `runs` must be a scalar"
                    )));
                },
            };
            props.push(name.as_str(), value);
        }
        let assignments = admit(&props, Operation::Add).map_err(|error| {
            let inner = error.property().unwrap_or_default().to_owned();
            refused(format!("run key `{inner}` in `runs`: {}", error.message()))
        })?;
        runs.push(assignments.run()?);
    }
    Ok(xlsx::RichText::new(runs))
}

/// The cell's runs to edit: its rich text, or its plain text as one
/// unformatted run (OfficeCLI's `AddRun` keeps a plain cell's text as the
/// first run), or no run for an empty cell.
fn editable_runs(workbook: &xlsx::Package, sheet: usize, at: Address) -> Result<xlsx::RichText> {
    let workbook = workbook.workbook();
    let worksheet = workbook
        .sheet(sheet)
        .map_err(super::write::native_error)?
        .ok_or_else(missing_sheet)?;
    if let Some(rich) = worksheet
        .rich_text(at)
        .map_err(super::write::native_error)?
    {
        return Ok(rich);
    }
    Ok(
        match worksheet.cell(at).map_err(super::write::native_error)? {
            xlsx::cell::View::Stored(xlsx::Cell::Value(value)) => {
                let text = match value {
                    xlsx::Value::Text(text) => text.as_str().to_owned(),
                    xlsx::Value::Number(number) => number.as_str().to_owned(),
                    xlsx::Value::Bool(value) => if *value { "TRUE" } else { "FALSE" }.to_owned(),
                    xlsx::Value::Date(date) => date.as_str().to_owned(),
                    _ => return Err(not_rich("an error cell has no runs")),
                };
                if text.is_empty() {
                    xlsx::RichText::new(Vec::new())
                } else {
                    xlsx::RichText::plain(text)
                }
            },
            xlsx::cell::View::Stored(xlsx::Cell::Formula(_)) => {
                return Err(not_rich("a formula cell has no runs"));
            },
            xlsx::cell::View::Covered(_) => {
                return Err(not_rich("a merged follower cell has no runs"));
            },
            _ => xlsx::RichText::new(Vec::new()),
        },
    )
}

fn missing_sheet() -> Error {
    Error::new(
        Code::OutOfRange,
        "the XLSX worksheet position is out of range",
    )
    .with_format(FileFormat::Xlsx)
}

fn not_rich(message: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message).with_format(FileFormat::Xlsx)
}

impl Xlsx {
    /// `set` on a run path; `None` for other targets.
    pub(super) fn try_set_run(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if !path.path().steps().last().is_some_and(|step| {
            step.name().eq_ignore_ascii_case("run") || step.name().eq_ignore_ascii_case("r")
        }) {
            return Ok(None);
        }
        // Promote/admit before resolving the run node: resolving it may clone
        // the complete rich-text payload from a deferred source worksheet.
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let resolved = self.adopt_workbook(source.clone());
        let node = crate::path::resolve::get(&Nav::new(&resolved), path)?;
        let Item::Run(run) = node.item() else {
            return Ok(None);
        };
        let assignments = admit(props, Operation::Set)?;
        let (sheet, at, index) = (node.sheet_position(), run.at, run.index as usize);
        // Resolve the rich payload only after the native source owner has
        // admitted its snapshot budget. The path node supplies identity and
        // index; `editable_runs` below reads the bounded source once.
        self.edit_runs_with_source(source, sheet, at, false, None, limits, |rich| {
            let run = rich.run_mut(index).ok_or_else(|| {
                Error::new(Code::PathNotFound, "the run is no longer in the cell")
            })?;
            assignments.apply(run)
        })
        .map(Some)
    }

    /// `add` of a run under the cell `parent`: inserted at `index` (appended
    /// by default). Returns the commit and the new run's position.
    pub(super) fn add_run(
        &self,
        parent: &Node<'_>,
        props: &Props<'_>,
        index: Option<usize>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, usize)> {
        let Some((sheet, at)) = parent.write_cell_target() else {
            return Err(
                Error::new(Code::UnsupportedType, "an XLSX run is added under one cell")
                    .with_format(FileFormat::Xlsx),
            );
        };
        let assignments = admit(props, Operation::Add)?;
        let run = assignments.run()?;
        let mut placed = 0;
        let commit = self.edit_runs(sheet, at, limits, |rich| {
            let position = index.unwrap_or(rich.runs().len());
            rich.insert_run(position, run)
                .map_err(|error| super::write::native_property_error(error, "index"))?;
            placed = position;
            Ok(())
        })?;
        Ok((commit, placed))
    }

    /// `remove` of a run path; `None` for other targets.
    pub(super) fn try_remove_run(
        &self,
        target: &TargetArg<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if !path.path().steps().last().is_some_and(|step| {
            step.name().eq_ignore_ascii_case("run") || step.name().eq_ignore_ascii_case("r")
        }) {
            return Ok(None);
        }
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let resolved = self.adopt_workbook(source.clone());
        let node = crate::path::resolve::get(&Nav::new(&resolved), path)?;
        let Item::Run(run) = node.item() else {
            return Ok(None);
        };
        let (sheet, at, index) = (node.sheet_position(), run.at, run.index as usize);
        self.edit_runs_with_source(source, sheet, at, true, None, limits, |rich| {
            rich.remove_run(index)
                .map(drop)
                .map_err(super::write::native_error)
        })
        .map(Some)
    }

    /// Read the cell's runs from the editor, change them, and publish them
    /// as the cell's inline rich string in one native edit.
    fn edit_runs(
        &self,
        sheet: usize,
        at: Address,
        limits: &SemanticLimits,
        change: impl FnOnce(&mut xlsx::RichText) -> Result<()>,
    ) -> Result<Commit> {
        self.edit_runs_with(sheet, at, false, None, limits, change)
    }

    /// [`Self::edit_runs`]; after a removal (`plain`), the runs left
    /// collapse as OfficeCLI's run removal does: none is empty text, one
    /// unformatted run is plain text.
    /// `read` is the cell's runs as the resolved run node read them, so the
    /// editor does not read them again.
    fn edit_runs_with(
        &self,
        sheet: usize,
        at: Address,
        plain: bool,
        read: Option<xlsx::RichText>,
        limits: &SemanticLimits,
        change: impl FnOnce(&mut xlsx::RichText) -> Result<()>,
    ) -> Result<Commit> {
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        self.edit_runs_with_source(source, sheet, at, plain, read, limits, change)
    }

    fn edit_runs_with_source(
        &self,
        source: xlsx::Package,
        sheet: usize,
        at: Address,
        plain: bool,
        read: Option<xlsx::RichText>,
        limits: &SemanticLimits,
        change: impl FnOnce(&mut xlsx::RichText) -> Result<()>,
    ) -> Result<Commit> {
        // Start the bounded native edit before reading the source rich runs;
        // the read can decode a deferred worksheet payload.
        let mut edit = edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        let mut rich = match read {
            Some(rich) => rich,
            None => editable_runs(&source, sheet, at)?,
        };
        change(&mut rich)?;
        self.check_source_for_replay()?;
        let mut worksheet = edit
            .sheet(sheet)
            .map_err(super::write::native_error)?
            .ok_or_else(missing_sheet)?;
        let collapsed = match rich.runs() {
            [] => Some(String::new()),
            [run] if !run.is_formatted() => Some(run.text().to_owned()),
            _ => None,
        };
        match collapsed.filter(|_| plain) {
            Some(text) => {
                worksheet
                    .set(at, xlsx::cell::Content::Value(xlsx::Value::text(text)))
                    .map_err(super::write::native_error)?;
            },
            None => {
                worksheet
                    .set_rich_text(at, rich)
                    .map_err(super::write::native_error)?;
            },
        }
        let native = edit.commit().map_err(super::write::native_error)?;
        Commit::from_native_with_limit(self, source, native, limits.max_snapshot_bytes())
    }
}

impl RunItem {
    /// The run this node reads.
    pub(super) fn run(&self) -> Option<&xlsx::RichRun> {
        self.rich.runs().get(self.index as usize)
    }
}
