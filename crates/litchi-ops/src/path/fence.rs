//! Source-freshness fence scopes (record 0262).
//!
//! A source-backed handle brackets each public read with a check that its
//! source has not changed (`Navigate::check_source`), so a changed source
//! wins over what the read produced. The resolver and the query walker
//! already bracket every `get`, `find` and query step that way; a [`Scope`]
//! tells the adapter's own nested reads that such a bracket is active for
//! its handle, so they skip their per-call checks instead of repeating them
//! for every child and property a step visits. On a file-backed source each
//! check reads the file's metadata, so per-visit checks dominated full
//! queries.
//!
//! The scope is per thread and per handle (keyed by the handle's address),
//! is restored on drop (also on unwind), and never outlives the bracket that
//! opened it: the bracket's closing check still runs after the scoped work,
//! so the observable result is the same as checking around each nested
//! read.

use std::cell::Cell;

thread_local! {
    /// The key of the handle whose reads an enclosing bracket fences; 0 for
    /// none.
    static ACTIVE: Cell<usize> = const { Cell::new(0) };
}

/// An active fence scope; dropping it restores the enclosing one.
pub(crate) struct Scope {
    previous: usize,
}

impl Scope {
    /// Marks the reads of the handle `key` as fenced by the caller until the
    /// scope drops. A `key` of 0 marks nothing.
    pub(crate) fn enter(key: usize) -> Self {
        let previous = ACTIVE.with(|active| active.replace(key));
        Self { previous }
    }
}

impl Drop for Scope {
    fn drop(&mut self) {
        ACTIVE.with(|active| active.set(self.previous));
    }
}

/// Whether an enclosing bracket fences the reads of the handle `key`.
pub(crate) fn is_active(key: usize) -> bool {
    key != 0 && ACTIVE.with(Cell::get) == key
}

#[cfg(test)]
mod tests {
    use super::{Scope, is_active};

    #[test]
    fn scopes_nest_and_restore() {
        assert!(!is_active(7));
        {
            let _outer = Scope::enter(7);
            assert!(is_active(7));
            assert!(!is_active(8));
            {
                let _inner = Scope::enter(8);
                assert!(is_active(8));
                assert!(!is_active(7));
            }
            assert!(is_active(7));
        }
        assert!(!is_active(7));
        let _none = Scope::enter(0);
        assert!(!is_active(0));
    }
}
