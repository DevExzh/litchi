//! A bounded per-thread cache of compiled plans (record 0291).
//!
//! A selector passed to `query`, `set` or `remove` as text or as a parsed
//! [`Query`] is compiled for the document's format on every call: about
//! 0.5 µs for a short RTF selector, a tenth of a single-target write.
//! Callers that repeat a selector can compile it once into a [`Plan`]
//! themselves; this cache gives the same saving to callers that do not.
//!
//! - **Key.** The parsed query (compared in full, after a hash), the format
//!   adapter it is compiled for (its format and its type), and whether the
//!   query may use native identities (`raw`). Compilation reads nothing
//!   else, so a cached plan is the plan a fresh compile would produce. Only
//!   successful compiles are kept; a failing selector compiles, and fails,
//!   every time.
//! - **Admission.** A query is cached the second time it is seen among the
//!   last [`SEEN`] distinct queries of its thread, so a selector run once
//!   pays only for its hash, not for an owned copy of itself.
//! - **Bounds.** At most [`CAPACITY`] plans per thread, replaced oldest
//!   first. Each holds an owned copy of its query, so a cached entry is
//!   bounded by the selector grammar's input limits.
//! - **Sharing.** Plans are shared by reference count (`Arc`), so a walk
//!   holding one stays `Send` and the cache never hands out a plan another
//!   walk can change. Nothing is shared between threads.

use std::any::type_name;
use std::cell::RefCell;
use std::hash::{Hash, Hasher};
use std::sync::Arc;

use litchi_core::FileFormat;
use litchi_core::selector::Query;

use super::compile::compile_with_raw;
use super::inspect::Inspect;
use super::plan::Plan;
use crate::Result;

/// Plans kept per thread.
const CAPACITY: usize = 32;

/// Distinct recent query hashes remembered per thread for admission.
const SEEN: usize = 64;

/// The compile target of an entry: the adapter's format and type.
#[derive(Clone, Copy, PartialEq, Eq)]
struct Target {
    format: FileFormat,
    adapter: &'static str,
    raw: bool,
}

struct Entry {
    hash: u64,
    target: Target,
    query: Query<'static>,
    plan: Arc<Plan<'static>>,
}

struct Cache {
    entries: Vec<Entry>,
    /// The slot the next admitted entry replaces once the cache is full.
    next: usize,
    /// Hashes of recently compiled, not yet admitted queries, in a ring.
    seen: [u64; SEEN],
    next_seen: usize,
}

thread_local! {
    static CACHE: RefCell<Cache> = const {
        RefCell::new(Cache {
            entries: Vec::new(),
            next: 0,
            seen: [0; SEEN],
            next_seen: 0,
        })
    };
}

/// A compiled plan: the shared cached one, or one compiled for this call.
pub(super) enum Compiled<'q> {
    Shared(Arc<Plan<'static>>),
    Fresh(Box<Plan<'q>>),
}

/// The plan of `query` for the adapter `N`, from the cache when this thread
/// compiled the same query for the same adapter recently.
///
/// # Errors
///
/// As [`compile_with_raw`]; failures are never cached.
pub(super) fn compile<'q, 'f, N: Inspect<'f>>(
    query: &Query<'q>,
    raw: bool,
) -> Result<Compiled<'q>> {
    let target = Target {
        format: N::FORMAT,
        adapter: type_name::<N>(),
        raw,
    };
    let hash = hash(query, target);
    let admit = CACHE.with(|cache| {
        let Ok(mut cache) = cache.try_borrow_mut() else {
            return Lookup::Skip;
        };
        if let Some(entry) = cache
            .entries
            .iter()
            .find(|entry| entry.hash == hash && entry.target == target && entry.query == *query)
        {
            return Lookup::Hit(Arc::clone(&entry.plan));
        }
        if cache.seen.contains(&hash) {
            return Lookup::Admit;
        }
        let slot = cache.next_seen;
        cache.seen[slot] = hash;
        cache.next_seen = (slot + 1) % SEEN;
        Lookup::Skip
    });
    match admit {
        Lookup::Hit(plan) => Ok(Compiled::Shared(plan)),
        Lookup::Skip => {
            compile_with_raw::<N>(query, raw).map(|plan| Compiled::Fresh(Box::new(plan)))
        },
        Lookup::Admit => {
            let owned = query.clone().into_owned();
            let plan = Arc::new(compile_with_raw::<N>(&owned, raw)?);
            CACHE.with(|cache| {
                if let Ok(mut cache) = cache.try_borrow_mut() {
                    cache.admit(Entry {
                        hash,
                        target,
                        query: owned,
                        plan: Arc::clone(&plan),
                    });
                }
            });
            Ok(Compiled::Shared(plan))
        },
    }
}

enum Lookup {
    Hit(Arc<Plan<'static>>),
    Admit,
    Skip,
}

impl Cache {
    fn admit(&mut self, entry: Entry) {
        let hash = entry.hash;
        if self.entries.len() < CAPACITY {
            self.entries.push(entry);
        } else {
            self.entries[self.next] = entry;
            self.next = (self.next + 1) % CAPACITY;
        }
        // An admitted hash leaves the admission ring, so a later collision
        // with another query does not admit that one on first sight.
        for seen in &mut self.seen {
            if *seen == hash {
                *seen = 0;
            }
        }
    }
}

/// The entry hash: the query and its compile target, with a small
/// multiplicative hasher (the full query is compared on a hash match).
fn hash(query: &Query<'_>, target: Target) -> u64 {
    let mut hasher = Fx(0);
    query.hash(&mut hasher);
    target.adapter.hash(&mut hasher);
    target.raw.hash(&mut hasher);
    // `FileFormat` does not implement `Hash`; the adapter type names it.
    hasher.finish() | 1
}

/// The `FxHasher` mixing step (rustc's): fast for the short keys a query
/// holds, and good enough with the full comparison behind it.
struct Fx(u64);

impl Fx {
    const SEED: u64 = 0x51_7c_c1_b7_27_22_0a_95;

    fn add(&mut self, word: u64) {
        self.0 = (self.0.rotate_left(5) ^ word).wrapping_mul(Self::SEED);
    }
}

impl Hasher for Fx {
    fn write(&mut self, bytes: &[u8]) {
        let mut chunks = bytes.chunks_exact(8);
        for chunk in &mut chunks {
            let mut word = [0; 8];
            word.copy_from_slice(chunk);
            self.add(u64::from_le_bytes(word));
        }
        let rest = chunks.remainder();
        if !rest.is_empty() {
            let mut word = [0; 8];
            word[..rest.len()].copy_from_slice(rest);
            self.add(u64::from_le_bytes(word));
        }
    }

    fn write_u8(&mut self, value: u8) {
        self.add(u64::from(value));
    }

    fn write_u32(&mut self, value: u32) {
        self.add(u64::from(value));
    }

    fn write_u64(&mut self, value: u64) {
        self.add(value);
    }

    fn write_usize(&mut self, value: usize) {
        self.add(value as u64);
    }

    fn finish(&self) -> u64 {
        self.0
    }
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use litchi_core::selector::Query;

    use super::super::compile::compile_with_raw;
    use super::super::fake::{Nav, SheetKind, WordKind};
    use super::{CAPACITY, Compiled, compile};

    type Word<'a> = &'a Nav<'a, WordKind>;

    fn shared(compiled: &Compiled<'_>) -> Option<usize> {
        match compiled {
            Compiled::Shared(plan) => Some(Arc::as_ptr(plan).addr()),
            Compiled::Fresh(_) => None,
        }
    }

    fn debug(compiled: &Compiled<'_>) -> String {
        match compiled {
            Compiled::Shared(plan) => format!("{plan:?}"),
            Compiled::Fresh(plan) => format!("{plan:?}"),
        }
    }

    #[test]
    fn a_repeated_query_is_compiled_once_and_shared() {
        // Each test runs on its own thread, so it starts with an empty cache.
        let text = String::from("/body/p[style=Heading1] > r[bold=true]");
        let query = Query::parse(&text).expect("query");
        let fresh = compile_with_raw::<Word<'_>>(&query, false).expect("plan");
        let first = compile::<Word<'_>>(&query, false).expect("first");
        assert_eq!(shared(&first), None, "a query seen once is not cached");
        let second = compile::<Word<'_>>(&query, false).expect("second");
        let third =
            compile::<Word<'_>>(&Query::parse(&text).expect("again"), false).expect("third");
        assert!(shared(&second).is_some());
        assert_eq!(shared(&second), shared(&third));
        for compiled in [&first, &second, &third] {
            assert_eq!(debug(compiled), format!("{fresh:?}"));
        }

        // Another adapter, or native identities, is another plan.
        let sheet = compile::<&Nav<'_, SheetKind>>(&Query::parse("row").expect("q"), false);
        assert!(sheet.is_ok());
        let raw = compile::<Word<'_>>(&query, true).expect("raw");
        assert_eq!(shared(&raw), None);
        let raw = compile::<Word<'_>>(&query, true).expect("raw again");
        assert!(shared(&raw).is_some() && shared(&raw) != shared(&second));
        assert!(matches!(&raw, Compiled::Shared(plan) if plan.raw));
    }

    #[test]
    fn failures_are_not_cached_and_the_cache_is_bounded() {
        let bad = Query::parse("slide").expect("query");
        for _ in 0..3 {
            assert!(compile::<Word<'_>>(&bad, false).is_err());
        }
        let texts: Vec<String> = (1..=CAPACITY * 2)
            .map(|index| format!("/body/p[{index}]"))
            .collect();
        let queries: Vec<Query<'_>> = texts
            .iter()
            .map(|text| Query::parse(text).expect("q"))
            .collect();
        for query in queries.iter().chain(&queries) {
            compile::<Word<'_>>(query, false).expect("plan");
        }
        super::CACHE.with(|cache| {
            let cache = cache.borrow();
            assert_eq!(cache.entries.len(), CAPACITY);
            // The newest half stays.
            for query in &queries[CAPACITY..] {
                assert!(cache.entries.iter().any(|entry| entry.query == *query));
            }
        });
    }
}
