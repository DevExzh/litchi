//! Locate one master of a package for the package-level master families
//! (record 0357).
//!
//! The live `DocumentContainer` names the main and title masters in its
//! master list and the notes and handout masters in its `DocumentAtom`. The
//! locator reads those records from the native snapshot (so an incrementally
//! saved package resolves its live records) and returns the selected
//! master's persist owner, context and exact source record bytes; the
//! family then validates that one record in its context.

use litchi_core::Position;

use super::{Kind, validation};
use crate::consts::RecordType;
use crate::master_layout::Context;
use crate::records::Record;
use crate::scoped::{self, Base, Failure};
use crate::slide_order::{self, records::Owner};

/// One located master record.
pub(crate) struct Located {
    /// The persisted owner holding the master record.
    pub(crate) owner: Owner,
    /// The master's contextual layout kind.
    pub(crate) context: Context,
    /// The exact persisted master record bytes.
    pub(crate) record: Vec<u8>,
}

/// The layout context of a master kind.
pub(crate) const fn context(kind: Kind) -> Context {
    match kind {
        Kind::Main => Context::Main,
        Kind::Title => Context::Title,
        Kind::Notes => Context::Notes,
        Kind::Handout => Context::Handout,
    }
}

/// Locate the master of `kind` at `position` among the masters of that kind,
/// in [`crate::Presentation::masters`] order. Notes and handout masters
/// accept only position 0.
pub(crate) fn locate(
    base: &Base,
    native: &slide_order::Snapshot,
    kind: Kind,
    position: Position,
    operation: &'static str,
) -> Result<Located, Failure> {
    let limits = base.open_options().source_limits_value();
    let document_bytes = base.owner_record(native, Owner::Document)?;
    let (document, _) =
        Record::parse(&document_bytes, 0, &limits).map_err(scoped::publication(operation))?;
    let missing = || scoped::not_found(operation, "no master of this kind at the position");

    let found = match kind {
        Kind::Main | Kind::Title => {
            let expected = if kind == Kind::Main {
                RecordType::MainMaster
            } else {
                RecordType::Slide
            };
            let mut found = None;
            let mut seen = 0usize;
            for list in document.children.iter().filter(|child| {
                child.record_type == RecordType::SlideListWithText && child.instance == 1
            }) {
                for reference in &list.children {
                    if reference.record_type != RecordType::SlidePersistAtom {
                        continue;
                    }
                    let (id, _) = validation::master_persist(reference)
                        .map_err(scoped::publication(operation))?;
                    let bytes = base.owner_record(native, Owner::Persist(id))?;
                    let (record, _) = Record::parse(&bytes, 0, &limits)
                        .map_err(scoped::publication(operation))?;
                    if record.record_type != expected {
                        continue;
                    }
                    if seen == position.get() {
                        found = Some((id, bytes));
                        break;
                    }
                    seen = seen.saturating_add(1);
                }
            }
            found
        },
        Kind::Notes | Kind::Handout => {
            let mut found = None;
            if position.get() == 0 {
                for atom in document
                    .children
                    .iter()
                    .filter(|child| child.record_type == RecordType::DocumentAtom)
                {
                    let (notes, handout) =
                        validation::document_atom(atom).map_err(scoped::publication(operation))?;
                    let id = if kind == Kind::Notes { notes } else { handout };
                    if id != 0 {
                        found = Some((id, base.owner_record(native, Owner::Persist(id))?));
                    }
                }
            }
            found
        },
    };
    let (id, record) = found.ok_or_else(missing)?;
    Ok(Located {
        owner: Owner::Persist(id),
        context: context(kind),
        record,
    })
}

/// Map a commit-time validation failure of a staged master: a candidate
/// that is not a valid master of its context (malformed or missing its
/// required atoms) is a content-free
/// [`litchi_core::edit::RefusalKind::InvalidValue`] refusal; limits and I/O
/// keep their package error.
pub(crate) fn candidate(operation: &'static str) -> impl FnOnce(crate::package::Error) -> Failure {
    move |error| match error {
        crate::package::Error::InvalidFormat(_) | crate::package::Error::Corrupted(_) => {
            Failure::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::InvalidValue,
                operation,
                "the staged record tree is not a valid master of its context",
            ))
        },
        error => Failure::Package(error),
    }
}
