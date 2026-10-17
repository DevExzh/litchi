//! A list that holds one element without a heap allocation.
//!
//! A DOCX `set` builds three lists per call: the resolved nodes, the
//! admitted targets and the admitted assignments. The common call has one
//! of each, and allocating each list cost about as much as resolving the
//! path (record 0302). [`Few`] keeps that element inline and spills to a
//! `Vec` from the second one on.

use std::ops::{Deref, DerefMut};

/// Zero elements, one inline, or two or more on the heap.
#[derive(Debug, Clone)]
pub(crate) enum Few<T> {
    Zero,
    One([T; 1]),
    // Invariant: two or more elements.
    Many(Vec<T>),
}

impl<T> Few<T> {
    /// An empty list.
    pub(crate) const fn new() -> Self {
        Self::Zero
    }

    /// An empty list that holds `capacity` elements without growing.
    pub(crate) fn with_capacity(capacity: usize) -> Self {
        if capacity > 1 {
            Self::Many(Vec::with_capacity(capacity))
        } else {
            Self::Zero
        }
    }

    /// Append `value`.
    pub(crate) fn push(&mut self, value: T) {
        match self {
            Self::Many(many) => many.push(value),
            Self::Zero => *self = Self::One([value]),
            Self::One(_) => {
                let Self::One([first]) = std::mem::replace(self, Self::Zero) else {
                    return;
                };
                *self = Self::Many(vec![first, value]);
            },
        }
    }

    /// Remove and return the last element.
    pub(crate) fn pop(&mut self) -> Option<T> {
        match self {
            Self::Zero => None,
            Self::Many(many) => many.pop(),
            Self::One(_) => match std::mem::replace(self, Self::Zero) {
                Self::One([only]) => Some(only),
                _ => None,
            },
        }
    }
}

impl<T> Default for Few<T> {
    fn default() -> Self {
        Self::new()
    }
}

impl<T> Deref for Few<T> {
    type Target = [T];

    fn deref(&self) -> &[T] {
        match self {
            Self::Zero => &[],
            Self::One(one) => one,
            Self::Many(many) => many,
        }
    }
}

impl<T> DerefMut for Few<T> {
    fn deref_mut(&mut self) -> &mut [T] {
        match self {
            Self::Zero => &mut [],
            Self::One(one) => one,
            Self::Many(many) => many,
        }
    }
}

impl<T> FromIterator<T> for Few<T> {
    fn from_iter<I: IntoIterator<Item = T>>(iter: I) -> Self {
        let mut few = Self::new();
        for value in iter {
            few.push(value);
        }
        few
    }
}

impl<'a, T> IntoIterator for &'a Few<T> {
    type Item = &'a T;
    type IntoIter = std::slice::Iter<'a, T>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

impl<T> IntoIterator for Few<T> {
    type Item = T;
    type IntoIter = std::iter::Chain<std::option::IntoIter<T>, std::vec::IntoIter<T>>;

    fn into_iter(self) -> Self::IntoIter {
        match self {
            Self::Zero => None.into_iter().chain(Vec::new()),
            Self::One([one]) => Some(one).into_iter().chain(Vec::new()),
            Self::Many(many) => None.into_iter().chain(many),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::Few;

    #[test]
    fn one_element_stays_inline_and_order_is_kept() {
        let mut few = Few::new();
        assert!(few.is_empty());
        few.push(1);
        assert!(matches!(few, Few::One(_)));
        assert_eq!(&*few, &[1]);
        few.push(2);
        few.push(3);
        assert_eq!(&*few, &[1, 2, 3]);
        assert_eq!(few.pop(), Some(3));
        assert_eq!(few.clone().into_iter().collect::<Vec<_>>(), [1, 2]);
        assert_eq!((&few).into_iter().copied().sum::<i32>(), 3);
        let mut one: Few<i32> = std::iter::once(7).collect();
        assert_eq!(one.clone().into_iter().collect::<Vec<_>>(), [7]);
        assert_eq!(one.pop(), Some(7));
        assert_eq!(one.pop(), None);
        assert!(Few::<i32>::with_capacity(1).is_empty());
        assert_eq!(Few::<i32>::new().into_iter().count(), 0);
    }
}
