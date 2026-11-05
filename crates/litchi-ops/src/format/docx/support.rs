//! Test-only access to the DOCX adapter for tests and examples outside the
//! crate: the counting-allocator test and the timing probe. `File::get`,
//! `find`, `query` and `view` (M1 step R5) are the public surface; this
//! module is not public API.

use litchi_core::FileFormat;

use super::{Docx, Nav};
use crate::Result;
use crate::node::Node;
use crate::path::{IntoPath, resolve};
use crate::query::walk::{PlanRef, Walk};
use crate::query::{Plan, Query};
use crate::view::mode::{Mode, Options};

/// Resolves `path` with `get`.
///
/// # Errors
///
/// As `File::get`.
pub fn get<'f, 'p>(file: &'f Docx, path: impl IntoPath<'p>) -> Result<Node<'f>> {
    let path = path.into_path()?;
    resolve::get(&Nav::new(file), &path).map(Node::Docx)
}

/// Resolves `path` with `find`.
///
/// # Errors
///
/// As `File::find`.
pub fn find<'f, 'p>(file: &'f Docx, path: impl IntoPath<'p>) -> Result<Option<Node<'f>>> {
    let path = path.into_path()?;
    Ok(resolve::find(&Nav::new(file), &path)?.map(Node::Docx))
}

/// Compiles `query` for DOCX documents.
///
/// # Errors
///
/// As [`Plan::compile`].
pub fn compile<'q>(query: &Query<'q>) -> Result<Plan<'q>> {
    Plan::compile(query, FileFormat::Docx)
}

/// Runs `plan` over `file`, calling `visit` with each match until `limit`
/// matches; returns the number of matches.
///
/// # Errors
///
/// As `File::query`.
pub fn run<'f>(
    file: &'f Docx,
    plan: &Plan<'_>,
    limit: usize,
    mut visit: impl FnMut(&Node<'f>),
) -> Result<usize> {
    let nav = Nav::new(file);
    let walk = Walk::new(&nav, PlanRef::Borrowed(plan))?.limit(limit);
    let mut matches = 0;
    for node in walk {
        visit(&Node::Docx(node?));
        matches += 1;
    }
    Ok(matches)
}

/// Writes the `mode` view (`text`, `annotated`, `outline`, `stats`) of
/// `file` into `out`.
///
/// # Errors
///
/// `invalid_argument` for an unknown mode name, and as `File::view_to`.
pub fn view(file: &Docx, mode: &str, out: &mut String) -> Result<()> {
    let mode = Mode::parse(mode)
        .ok_or_else(|| crate::Error::new(crate::Code::InvalidArgument, "unknown view mode"))?;
    super::views::view_to(file, mode, &Options::default(), out)
}

/// The leaf's own read of the body's block index, through the same owner:
/// the baseline the adapter's reads are measured against.
///
/// # Errors
///
/// The leaf's error for a main part that does not read.
pub fn leaf_body(file: &Docx) -> Result<litchi_docx::Blocks> {
    super::nav::load_body(file)
}
