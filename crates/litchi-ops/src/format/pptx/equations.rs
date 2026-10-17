//! Office Math equations (record 0321, L-PPTX-R11).
//!
//! `/slide[N]/shape[M]/equation[K]` is the K-th equation in the text of a
//! shape; `/slide[N]/equation[K]` the K-th of the slide. Each reads
//! `formula`, its OMML rendered LaTeX-like, and its text. Equations are
//! read from the slide's own XML (PowerPoint writes them in a
//! markup-compatibility choice whose fallback, a picture, is what the
//! shape scene reads). Writing them needs an OMML writer the leaf does not
//! have: `set`, `add` and `remove` are gaps.

use std::borrow::Cow;
use std::sync::Arc;

use litchi_pptx::shape::math::Equation;

use super::node::{Item, TreeItem};
use super::spec::EquationKey;
use super::{Pptx, PptxOwner};
use crate::Result;
use crate::node::Value;

/// One equation and where it is addressed.
#[derive(Clone)]
pub(super) struct EquationItem {
    pub(super) slide: u32,
    /// The shape whose text holds it, when it is addressed through it.
    pub(super) shape: Option<TreeItem>,
    /// Its position in its scope (the shape's or the slide's equations).
    pub(super) index: u32,
    pub(super) equation: Arc<Equation>,
}

/// The value of the equation property at `index`.
pub(super) fn value(item: &EquationItem, index: usize) -> Option<Value<'_>> {
    match EquationKey::at(index)? {
        EquationKey::Formula => Some(Value::Text(Cow::Borrowed(item.equation.formula()))),
    }
}

impl Pptx {
    /// The equations of the slide at `position`, in document order.
    pub(super) fn slide_equations(&self, position: usize) -> Result<Vec<Equation>> {
        self.finish(|| {
            if let Some(revision) = &self.0.revision {
                let presentation = revision.presentation()?;
                if position >= presentation.slide_count()? {
                    return Ok(Vec::new());
                }
                return Ok(presentation.slide(position)?.equations()?);
            }
            match &self.0.owner {
                PptxOwner::Source(presentation) => {
                    self.check_owner()?;
                    let Some(slide) = presentation.slide(position) else {
                        return Ok(Vec::new());
                    };
                    Ok(slide.equations()?)
                },
                PptxOwner::Eager(package) => {
                    let Some(slide) = super::eager_slide(self, package, position)? else {
                        return Ok(Vec::new());
                    };
                    Ok(slide.equations()?)
                },
            }
        })
    }
}

/// The equations of a slide (`shape: None`) or of one of its shapes.
pub(super) fn items(file: &Pptx, slide: u32, shape: Option<&TreeItem>) -> Result<Vec<Item>> {
    let id = shape.and_then(|tree| tree.handle.view().id());
    if shape.is_some() && id.is_none() {
        return Ok(Vec::new());
    }
    let mut out = Vec::new();
    for equation in file.slide_equations(slide as usize)? {
        if shape.is_some() && equation.shape_id() != id {
            continue;
        }
        out.push(Item::Equation(Box::new(EquationItem {
            slide,
            shape: shape.cloned(),
            index: u32::try_from(out.len()).unwrap_or(u32::MAX),
            equation: Arc::new(equation),
        })));
    }
    Ok(out)
}
