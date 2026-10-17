//! The children a query walk lists beyond each node's main axis (record
//! 0321): the presentation theme, a table's grid columns, a shape's or run's
//! hyperlink, a slide's media pictures and equations.
//!
//! The node view never lists these (`Children` with `Listing::View`), so a
//! view of an existing deck keeps its shape; paths and queries reach them.

use std::sync::Arc;

use super::Pptx;
use super::node::{Item, TreeContainer};
use super::spec::Kind;
use crate::Result;
use crate::path::navigate::KindSet;

/// The extra children of `parent` whose kinds are in `kinds`, in document
/// order after the main axis.
pub(super) fn walk_children(file: &Pptx, parent: &Item, kinds: KindSet<Kind>) -> Result<Vec<Item>> {
    let mut out = Vec::new();
    // A walk that wants every kind (`*`) lists the main axes only, as it
    // did before these elements were readable.
    if Kind::ALL.iter().all(|kind| kinds.contains(*kind)) {
        return Ok(out);
    }
    if let Item::Presentation = parent
        && kinds.contains(Kind::Theme)
        && let Some(theme) = file.presentation_theme()?
    {
        out.push(Item::Theme(super::theme::ThemeItem {
            theme: Arc::new(theme),
        }));
    }
    // `media` names both kinds; a walk that wants it lists each picture once.
    let media: &[Kind] = if kinds.contains(Kind::Media) {
        &[Kind::Media]
    } else {
        &[Kind::Video, Kind::Audio]
    };
    for &kind in media {
        if !kinds.contains(kind) {
            continue;
        }
        let container = match parent {
            Item::Slide(slide) => Some((
                TreeContainer::Slide(slide.position),
                slide.scene(file)?.roots(),
            )),
            Item::Tree(tree) if tree.kind == Kind::Group => {
                Some((tree.container, tree.handle.children()))
            },
            _ => None,
        };
        if let Some((container, handles)) = container {
            out.extend(super::media::items(container, handles, kind)?);
        }
    }
    if let Item::Tree(tree) = parent {
        if tree.kind == Kind::Table && kinds.contains(Kind::Col) {
            out.extend(super::columns::children(tree)?);
        }
        if tree.kind == Kind::Shape && kinds.contains(Kind::Hyperlink) {
            out.extend(super::links::children(file, parent)?);
        }
        if tree.kind == Kind::Shape
            && kinds.contains(Kind::Equation)
            && let Some(slide) = tree.container.slide()
        {
            out.extend(super::equations::items(file, slide, Some(tree))?);
        }
    }
    if let Item::Run(_) = parent
        && kinds.contains(Kind::Hyperlink)
    {
        out.extend(super::links::children(file, parent)?);
    }
    // Equations outside a shape (none in PowerPoint's output) are listed
    // under their slide.
    if let Item::Slide(slide) = parent
        && kinds.contains(Kind::Equation)
    {
        out.extend(
            super::equations::items(file, slide.position, None)?
                .into_iter()
                .filter(|item| {
                    matches!(item, Item::Equation(equation) if equation.equation.shape_id().is_none())
                }),
        );
    }
    Ok(out)
}
