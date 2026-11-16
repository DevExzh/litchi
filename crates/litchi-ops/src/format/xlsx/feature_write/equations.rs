//! Add, set and remove of worksheet equations.
//!
//! `add(/Sheet1, equation, {latex | omml, anchor, size})` (or under a cell,
//! which names the anchor) adds a text box anchored at that cell, the OMML
//! with a linear-text fallback, through
//! `drawing::edit::Edit::add_equation`; `set` replaces an equation's `latex`
//! or `omml` (`set_equation`), and `remove` deletes it with the pictures
//! and relationships nothing else uses (`remove_equation`). Writing needs
//! the `equation` feature. LaTeX that loses formatting in the conversion is
//! refused, so the caller can write OMML instead.

use litchi_core::selector::Path;

use super::{Arg, Commit, Data, Feature, Parent, Props, Result, SemanticLimits, Xlsx, admit};
use super::{Code, Error, error, invalid};
#[cfg(feature = "equation")]
use super::{feature_path, part_edit_options};
use crate::format::xlsx::features::{self, EquationKey};
use crate::registry::Operation;

/// What an `add` or `set` of an equation assigns.
#[derive(Default)]
struct Input {
    /// `latex` or `omml`, and its text.
    content: Option<(EquationKey, String)>,
    anchor: Option<String>,
    size: Option<f64>,
}

fn input(props: &Props<'_>, operation: Operation, limits: &SemanticLimits) -> Result<Input> {
    let mut input = Input::default();
    for (key, name, arg) in admit(&features::EQUATION_TABLE, operation, props)? {
        match key {
            EquationKey::Latex | EquationKey::Omml => {
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
                        "an equation takes one of latex and omml",
                    ));
                }
            },
            EquationKey::Anchor => input.anchor = arg.text(),
            EquationKey::Size => match arg {
                Arg::Length(length) => input.size = Some(length.points()),
                _ => return Err(invalid(name, "size is a font size")),
            },
            EquationKey::Name => {},
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
        let Data::Equation(_) = &feature.data else {
            return Err(error(
                Code::InternalError,
                "an equation node holds no equation",
            ));
        };
        let input = input(props, Operation::Set, limits)?;
        let Some(content) = input.content.as_ref() else {
            return Err(error(
                Code::InvalidArgument,
                "no equation property was assigned",
            ));
        };
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
