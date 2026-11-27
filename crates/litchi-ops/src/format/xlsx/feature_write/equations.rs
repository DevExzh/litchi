//! Add, set and remove of worksheet equations: Office Math equations and
//! OLE equation objects (MathType and Equation Editor 3.0).
//!
//! `add(/Sheet1, equation, {latex | omml, anchor, size})` (or under a cell,
//! which names the anchor) adds a text box anchored at that cell, the OMML
//! with a linear-text fallback, through
//! `drawing::edit::Edit::add_equation`; `set` replaces an equation's `latex`
//! or `omml` (`set_equation`), and `remove` deletes it with the pictures
//! and relationships nothing else uses (`remove_equation`). LaTeX that
//! loses formatting in the conversion is refused, so the caller can write
//! OMML instead.
//!
//! An `add` whose properties name `mtef` or `progId` (only
//! `Equation.DSMT4`) writes an OLE equation object instead
//! (`equation_objects::Edit::add`): MathType for `latex` and `omml`, the
//! data's own kind for `mtef`, with a PNG preview. `set` of an object's
//! `latex`, `omml` or `mtef` keeps its kind (`equation_objects::Edit::set`);
//! `remove` deletes it with the parts only it used. `mtef` is refused on an
//! Office Math equation, and a conversion to MTEF that loses formatting is
//! refused, so the caller can write MTEF instead. Writing needs the
//! `equation` feature.

use litchi_core::selector::Path;

use super::{Arg, Commit, Data, Feature, Parent, Props, Result, SemanticLimits, Xlsx, admit};
use super::{Code, Error, error, invalid};
#[cfg(feature = "equation")]
use super::{feature_path, part_edit_options};
use crate::format::xlsx::features::{self, EquationKey};
use crate::registry::Operation;

/// The only ProgID an `add` writes.
const MATH_TYPE: &str = "Equation.DSMT4";

/// What an `add` or `set` of an equation assigns.
#[derive(Default)]
struct Input {
    /// `latex`, `omml` or `mtef`, and its text.
    content: Option<(EquationKey, String)>,
    anchor: Option<String>,
    size: Option<f64>,
    prog_id: Option<String>,
}

impl Input {
    /// Whether the input writes an OLE equation object.
    fn is_object(&self) -> bool {
        self.prog_id.is_some() || matches!(self.content, Some((EquationKey::Mtef, _)))
    }
}

fn input(props: &Props<'_>, operation: Operation, limits: &SemanticLimits) -> Result<Input> {
    let mut input = Input::default();
    for (key, name, arg) in admit(&features::EQUATION_TABLE, operation, props)? {
        match key {
            EquationKey::Latex | EquationKey::Omml | EquationKey::Mtef => {
                let text = arg.text().unwrap_or_default();
                if text.len() > limits.max_staged_bytes() {
                    return Err(Error::new(
                        Code::LimitExceeded,
                        "the equation is larger than the staged-byte limit",
                    )
                    .with_format(litchi_core::FileFormat::Xlsx)
                    .with_property(name));
                }
                if input.content.replace((key, text)).is_some() {
                    return Err(error(
                        Code::InvalidArgument,
                        "an equation takes one of latex, omml and mtef",
                    ));
                }
            },
            EquationKey::Anchor => input.anchor = arg.text(),
            EquationKey::Size => match arg {
                Arg::Length(length) => input.size = Some(length.points()),
                _ => return Err(invalid(name, "size is a font size")),
            },
            EquationKey::ProgId => input.prog_id = arg.text(),
            EquationKey::Name | EquationKey::Error => {},
        }
    }
    Ok(input)
}

/// The cell an equation is anchored at: the `anchor` property, else the
/// cell or the first cell of the range the parent path named, else `A1`.
#[cfg_attr(
    not(feature = "equation"),
    allow(dead_code, reason = "used by the equation writer")
)]
fn anchor_cell(parent: &Parent, anchor: Option<&str>) -> Result<(u32, u32)> {
    let Some(reference) = anchor.or(parent.reference.as_deref()) else {
        return Ok((0, 0));
    };
    let first = reference
        .trim()
        .split(':')
        .next()
        .unwrap_or_default()
        .replace('$', "");
    let cell = litchi_core::selector::CellRef::parse(&first)
        .ok_or_else(|| invalid("anchor", "the anchor is not an A1 cell"))?;
    Ok((cell.column(), cell.row()))
}

#[cfg(not(feature = "equation"))]
fn unavailable() -> Error {
    error(
        Code::UnsupportedCapability,
        "writing equations needs the equation feature",
    )
    .with_suggestion("enable the litchi equation feature")
}

/// The leaf source of an admitted `latex` or `omml`.
#[cfg(feature = "equation")]
fn source(content: &(EquationKey, String)) -> litchi_xlsx::drawing::math::EquationSource<'_> {
    use litchi_xlsx::drawing::math::EquationSource;
    match content {
        (EquationKey::Omml, omml) => EquationSource::Omml(omml),
        (_, latex) => EquationSource::Latex(latex),
    }
}

#[cfg(feature = "equation")]
const fn content_key(content: &(EquationKey, String)) -> &'static str {
    match content.0 {
        EquationKey::Omml => "omml",
        _ => "latex",
    }
}

/// Refuse a LaTeX conversion that lost formatting: the caller decides
/// whether to write OMML instead.
#[cfg(feature = "equation")]
fn refuse_losses(losses: &[String]) -> Result<()> {
    match losses.first() {
        None => Ok(()),
        Some(loss) => Err(Error::new(
            Code::InvalidValue,
            format!("the LaTeX equation does not convert to OMML without loss: {loss}"),
        )
        .with_format(litchi_core::FileFormat::Xlsx)
        .with_property("latex")
        .with_suggestion("write the equation as omml")),
    }
}

/// A leaf equation error, with the property whose value caused it.
#[cfg(feature = "equation")]
fn leaf(failure: litchi_xlsx::Error, key: &'static str) -> Error {
    super::authored(failure).with_property(key)
}

impl Xlsx {
    /// `add(<sheet or cell>, equation, {latex | omml, anchor, size})`.
    pub(super) fn add_equation(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let input = input(props, Operation::Add, limits)?;
        if input.is_object() {
            return self.add_equation_object(parent, input, limits);
        }
        let Some(content) = input.content.as_ref() else {
            return Err(
                error(Code::MissingProperty, "an equation needs latex or omml")
                    .with_property("latex"),
            );
        };
        #[cfg(not(feature = "equation"))]
        {
            let _ = (content, parent, input.anchor, input.size);
            Err(unavailable())
        }
        #[cfg(feature = "equation")]
        {
            let (column, row) = anchor_cell(parent, input.anchor.as_deref())?;
            let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
            let mut edit = source.edit_drawings_with_options(part_edit_options(
                &source,
                limits.max_snapshot_bytes(),
            ))?;
            let written = edit
                .add_equation(
                    parent.sheet,
                    self::source(content),
                    litchi_xlsx::drawing::edit::Marker::cell(column, row),
                    input.size,
                )
                .map_err(|failure| leaf(failure, content_key(content)))?;
            refuse_losses(&written.losses)?;
            let commit = edit.commit()?;
            let changed = commit.diagnostics().changed();
            let native_patch_bytes = commit.patch().retained_bytes();
            let commit = self.publish_with_limit(
                source,
                commit.into_package(),
                changed,
                native_patch_bytes,
                limits.max_snapshot_bytes(),
            )?;
            Ok((
                commit,
                feature_path(parent.sheet, "equation", Some(written.index)),
            ))
        }
    }

    /// A `set` of an equation's `latex` or `omml`.
    pub(super) fn set_equation(
        &self,
        sheet: usize,
        feature: &Feature,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let input = input(props, Operation::Set, limits)?;
        let Some(content) = input.content.as_ref() else {
            return Err(error(
                Code::InvalidArgument,
                "no equation property was assigned",
            ));
        };
        match &feature.data {
            Data::Equation(_) => {},
            #[cfg(feature = "equation")]
            Data::OleEquation(object) => {
                return self.set_equation_object(sheet, object.shape_id(), content, limits);
            },
            _ => {
                return Err(error(
                    Code::InternalError,
                    "an equation node holds no equation",
                ));
            },
        }
        if content.0 == EquationKey::Mtef {
            return Err(invalid(
                "mtef",
                "mtef is written to OLE equation objects, not Office Math equations",
            )
            .with_suggestion("set latex or omml"));
        }
        #[cfg(not(feature = "equation"))]
        {
            let _ = (content, sheet);
            Err(unavailable())
        }
        #[cfg(feature = "equation")]
        {
            let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
            let mut edit = source.edit_drawings_with_options(part_edit_options(
                &source,
                limits.max_snapshot_bytes(),
            ))?;
            let written = edit
                .set_equation(sheet, feature.index as usize, self::source(content))
                .map_err(|failure| leaf(failure, content_key(content)))?;
            if let Some(written) = &written {
                refuse_losses(&written.losses)?;
            }
            let commit = edit.commit()?;
            let changed = commit.diagnostics().changed();
            let native_patch_bytes = commit.patch().retained_bytes();
            self.publish_with_limit(
                source,
                commit.into_package(),
                changed,
                native_patch_bytes,
                limits.max_snapshot_bytes(),
            )
        }
    }

    /// A `remove` of an equation, with its text box and picture.
    pub(super) fn remove_equation(
        &self,
        sheet: usize,
        feature: &Feature,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        #[cfg(not(feature = "equation"))]
        {
            let _ = (sheet, feature, limits);
            Err(unavailable())
        }
        #[cfg(feature = "equation")]
        {
            if let Data::OleEquation(object) = &feature.data {
                let shape = object.shape_id();
                return self.publish_equation_objects(limits, |edit| {
                    edit.remove(sheet, shape)?;
                    Ok(())
                });
            }
            let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
            let mut edit = source.edit_drawings_with_options(part_edit_options(
                &source,
                limits.max_snapshot_bytes(),
            ))?;
            edit.remove_equation(sheet, feature.index as usize)?;
            let commit = edit.commit()?;
            let changed = commit.diagnostics().changed();
            let native_patch_bytes = commit.patch().retained_bytes();
            self.publish_with_limit(
                source,
                commit.into_package(),
                changed,
                native_patch_bytes,
                limits.max_snapshot_bytes(),
            )
        }
    }
}

// ---------------------------------------------------------------------------
// OLE equation objects

#[cfg(feature = "equation")]
fn decode_mtef(text: &str) -> Result<Vec<u8>> {
    use base64::Engine as _;
    base64::engine::general_purpose::STANDARD
        .decode(text.trim())
        .map_err(|_| invalid("mtef", "mtef is the base64 of an Equation Native stream"))
}

/// Refuse a conversion to MTEF that lost formatting.
#[cfg(feature = "equation")]
fn refuse_mtef_losses(losses: &[String], key: &'static str) -> Result<()> {
    match losses.first() {
        None => Ok(()),
        Some(loss) => Err(Error::new(
            Code::InvalidValue,
            format!("the equation does not convert to MTEF without loss: {loss}"),
        )
        .with_format(litchi_core::FileFormat::Xlsx)
        .with_property(key)
        .with_suggestion("write the equation as mtef")),
    }
}

/// The leaf content of an admitted `latex`, `omml` or `mtef`, owned.
#[cfg(feature = "equation")]
enum ObjectContent {
    Latex(String),
    Omml(String),
    Mtef(Vec<u8>),
}

#[cfg(feature = "equation")]
impl ObjectContent {
    fn of(content: &(EquationKey, String)) -> Result<Self> {
        Ok(match content {
            (EquationKey::Mtef, text) => Self::Mtef(decode_mtef(text)?),
            (EquationKey::Omml, text) => Self::Omml(text.clone()),
            (_, text) => Self::Latex(text.clone()),
        })
    }

    const fn key(&self) -> &'static str {
        match self {
            Self::Latex(_) => "latex",
            Self::Omml(_) => "omml",
            Self::Mtef(_) => "mtef",
        }
    }

    fn source(&self) -> litchi_xlsx::equation_objects::Source<'_> {
        use litchi_xlsx::equation_objects::Source;
        match self {
            Self::Latex(latex) => Source::Latex(latex),
            Self::Omml(omml) => Source::Omml(omml),
            Self::Mtef(mtef) => Source::Mtef(mtef),
        }
    }
}

impl Xlsx {
    /// Stage `stage` in one native equation-object edit and publish it.
    #[cfg(feature = "equation")]
    fn publish_equation_objects(
        &self,
        limits: &SemanticLimits,
        stage: impl FnOnce(&mut litchi_xlsx::equation_objects::Edit) -> Result<()>,
    ) -> Result<Commit> {
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = source.edit_equation_objects_with_options(part_edit_options(
            &source,
            limits.max_snapshot_bytes(),
        ))?;
        stage(&mut edit)?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let native_patch_bytes = commit.patch().retained_bytes();
        self.publish_with_limit(
            source,
            commit.into_package(),
            changed,
            native_patch_bytes,
            limits.max_snapshot_bytes(),
        )
    }

    /// `add(<sheet or cell>, equation, {progId, latex | omml | mtef,
    /// anchor})`: an OLE equation object, numbered after the sheet's Office
    /// Math equations.
    fn add_equation_object(
        &self,
        parent: &Parent,
        input: Input,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        if let Some(prog_id) = &input.prog_id
            && !prog_id.eq_ignore_ascii_case(MATH_TYPE)
        {
            return Err(invalid(
                "progId",
                "an equation object is added as Equation.DSMT4 (MathType)",
            ));
        }
        if input.size.is_some() {
            return Err(invalid(
                "size",
                "size sizes an Office Math equation; an equation object is sized to its preview",
            ));
        }
        let Some(content) = input.content.as_ref() else {
            return Err(error(
                Code::MissingProperty,
                "an equation needs latex, omml or mtef",
            )
            .with_property("latex"));
        };
        #[cfg(not(feature = "equation"))]
        {
            let _ = (content, parent, limits);
            Err(unavailable())
        }
        #[cfg(feature = "equation")]
        {
            let (column, row) = anchor_cell(parent, input.anchor.as_deref())?;
            let content = ObjectContent::of(content)?;
            let key = content.key();
            let named = input.prog_id.is_some();
            let mut shape = 0;
            let commit = self.publish_equation_objects(limits, |edit| {
                let written = edit
                    .add(
                        parent.sheet,
                        row,
                        column,
                        content.source(),
                        &litchi_equation::Limits::DEFAULT,
                    )
                    .map_err(|failure| leaf(failure, key))?;
                refuse_mtef_losses(written.losses(), key)?;
                // `progId` names MathType; MTEF 3 data is Equation Editor 3.0.
                if named
                    && edit.equations(parent.sheet)?.iter().any(|object| {
                        object.shape_id() == written.shape_id()
                            && !object.prog_id().eq_ignore_ascii_case(MATH_TYPE)
                    })
                {
                    return Err(invalid("mtef", "Equation.DSMT4 needs MTEF 5 data"));
                }
                shape = written.shape_id();
                Ok(())
            })?;
            let index = features::list(
                commit.file(),
                &super::super::node::Sheet::new(
                    commit.file(),
                    parent.sheet,
                    super::super::node::SheetHandle::at(commit.file(), parent.sheet)?.ok_or_else(
                        || error(Code::InternalError, "the worksheet vanished"),
                    )?,
                ),
                super::Kind::Equation,
            )?
            .iter()
            .position(|feature| {
                matches!(&feature.data, Data::OleEquation(object) if object.shape_id() == shape)
            })
            .ok_or_else(|| {
                error(
                    Code::InternalError,
                    "the added equation object did not read back",
                )
            })?;
            Ok((commit, feature_path(parent.sheet, "equation", Some(index))))
        }
    }

    /// A `set` of an OLE equation object's `latex`, `omml` or `mtef`,
    /// keeping its kind.
    #[cfg(feature = "equation")]
    fn set_equation_object(
        &self,
        sheet: usize,
        shape: u32,
        content: &(EquationKey, String),
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let content = ObjectContent::of(content)?;
        let key = content.key();
        self.publish_equation_objects(limits, |edit| {
            let written = edit
                .set(
                    sheet,
                    shape,
                    content.source(),
                    &litchi_equation::Limits::DEFAULT,
                )
                .map_err(|failure| leaf(failure, key))?;
            refuse_mtef_losses(written.losses(), key)
        })
    }
}
