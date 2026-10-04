//! Small, covariant, allocation-avoiding sequences for the AST.
//!
//! `smallvec::SmallVec` stores its element type behind the `Array::Item`
//! projection, which makes `SmallVec<[Step<'a>; N]>` invariant in `'a`; a
//! `Path<'static>` could then not be used where a `Path<'a>` is expected. These
//! two containers keep the zero-allocation fast path with plain arrays,
//! `Option`, and `Vec`, all of which are covariant, and without `unsafe`.

use std::fmt;
use std::hash::{Hash, Hasher};

/// Zero or one element inline; two or more on the heap.
#[derive(Clone)]
pub(super) struct OneOrMany<T> {
    // Invariant: when `one` is `Some`, `many` is empty.
    one: Option<T>,
    many: Vec<T>,
}

impl<T> OneOrMany<T> {
    pub(super) const fn new() -> Self {
        Self {
            one: None,
            many: Vec::new(),
        }
    }

    pub(super) fn as_slice(&self) -> &[T] {
        match &self.one {
            Some(one) => std::slice::from_ref(one),
            None => &self.many,
        }
    }

    pub(super) fn len(&self) -> usize {
        self.as_slice().len()
    }

    pub(super) fn is_empty(&self) -> bool {
        self.one.is_none() && self.many.is_empty()
    }

    pub(super) fn push(&mut self, value: T) {
        if !self.many.is_empty() {
            self.many.push(value);
        } else if let Some(first) = self.one.take() {
            self.many.reserve_exact(4);
            self.many.push(first);
            self.many.push(value);
        } else {
            self.one = Some(value);
        }
    }

    pub(super) fn into_vec(self) -> Vec<T> {
        match self.one {
            Some(one) => vec![one],
            None => self.many,
        }
    }

    pub(super) fn map<U>(self, mut map: impl FnMut(T) -> U) -> OneOrMany<U> {
        match self.one {
            Some(one) => OneOrMany {
                one: Some(map(one)),
                many: Vec::new(),
            },
            None => OneOrMany {
                one: None,
                many: self.many.into_iter().map(map).collect(),
            },
        }
    }
}

impl<T> Default for OneOrMany<T> {
    fn default() -> Self {
        Self::new()
    }
}

impl<T> FromIterator<T> for OneOrMany<T> {
    fn from_iter<I: IntoIterator<Item = T>>(iter: I) -> Self {
        let mut out = Self::new();
        for value in iter {
            out.push(value);
        }
        out
    }
}

impl<T: PartialEq> PartialEq for OneOrMany<T> {
    fn eq(&self, other: &Self) -> bool {
        self.as_slice() == other.as_slice()
    }
}

impl<T: Eq> Eq for OneOrMany<T> {}

impl<T: Hash> Hash for OneOrMany<T> {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.as_slice().hash(state);
    }
}

impl<T: fmt::Debug> fmt::Debug for OneOrMany<T> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_list().entries(self.as_slice()).finish()
    }
}

/// Up to `N` elements in a fixed array of placeholders; more spill wholesale
/// to the heap so the elements stay contiguous.
#[derive(Clone)]
pub(super) struct Inline<T, const N: usize> {
    len: usize,
    // Invariant: slots at and beyond `len` (or all slots once spilled) hold
    // the placeholder value, so derived structural state never leaks stale
    // elements.
    inline: [T; N],
    heap: Vec<T>,
}

/// A cheap, allocation-free placeholder for unused inline slots.
pub(super) trait Placeholder {
    fn placeholder() -> Self;
}

impl<T: Placeholder, const N: usize> Inline<T, N> {
    pub(super) fn new() -> Self {
        Self {
            len: 0,
            inline: std::array::from_fn(|_| T::placeholder()),
            heap: Vec::new(),
        }
    }

    pub(super) fn as_slice(&self) -> &[T] {
        if self.len <= N {
            &self.inline[..self.len]
        } else {
            &self.heap
        }
    }

    pub(super) fn len(&self) -> usize {
        self.len
    }

    pub(super) fn push(&mut self, value: T) {
        if self.len < N {
            self.inline[self.len] = value;
        } else {
            if self.len == N {
                self.heap.reserve_exact(N * 2);
                for slot in &mut self.inline {
                    self.heap.push(std::mem::replace(slot, T::placeholder()));
                }
            }
            self.heap.push(value);
        }
        self.len += 1;
    }

    pub(super) fn map<U: Placeholder>(self, mut map: impl FnMut(T) -> U) -> Inline<U, N> {
        let mut out = Inline::new();
        if self.len <= N {
            for value in self.inline.into_iter().take(self.len) {
                out.push(map(value));
            }
        } else {
            for value in self.heap {
                out.push(map(value));
            }
        }
        out
    }
}

impl<T: Placeholder, const N: usize> FromIterator<T> for Inline<T, N> {
    fn from_iter<I: IntoIterator<Item = T>>(iter: I) -> Self {
        let mut out = Self::new();
        for value in iter {
            out.push(value);
        }
        out
    }
}

impl<T: PartialEq + Placeholder, const N: usize> PartialEq for Inline<T, N> {
    fn eq(&self, other: &Self) -> bool {
        self.as_slice() == other.as_slice()
    }
}

impl<T: Eq + Placeholder, const N: usize> Eq for Inline<T, N> {}

impl<T: Hash + Placeholder, const N: usize> Hash for Inline<T, N> {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.as_slice().hash(state);
    }
}

impl<T: fmt::Debug + Placeholder, const N: usize> fmt::Debug for Inline<T, N> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_list().entries(self.as_slice()).finish()
    }
}
