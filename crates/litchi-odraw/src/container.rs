//! Checked, lazy traversal of `OfficeArt` container records.

use crate::{Error, Limit, Record, RecordKind, Result};

/// Resource ceilings for recursive `OfficeArt` traversal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum nested-container depth.
    pub max_depth: u16,
    /// Maximum records visited during one traversal.
    pub max_records: u32,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_depth: 64,
            max_records: 1_000_000,
        }
    }
}

/// A zero-copy iterator over the direct children of a container.
#[derive(Debug, Clone)]
pub struct Children<'data> {
    data: &'data [u8],
    offset: usize,
    done: bool,
}

impl<'data> Children<'data> {
    /// Creates an iterator over a validated container body or record sequence.
    #[must_use]
    pub const fn new(data: &'data [u8]) -> Self {
        Self {
            data,
            offset: 0,
            done: false,
        }
    }
}

impl<'data> Iterator for Children<'data> {
    type Item = Result<Record<'data>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done || self.offset == self.data.len() {
            return None;
        }

        if self.data.len().saturating_sub(self.offset) < 8 {
            self.done = true;
            return Some(Err(Error::TruncatedHeader {
                offset: self.offset,
                available: self.data.len().saturating_sub(self.offset),
            }));
        }

        match Record::parse(self.data, self.offset) {
            Ok((record, consumed)) => {
                if let Some(next) = self.offset.checked_add(consumed) {
                    self.offset = next;
                    Some(Ok(record))
                } else {
                    self.done = true;
                    Some(Err(Error::ArithmeticOverflow {
                        context: "child-record cursor",
                    }))
                }
            },
            Err(error) => {
                self.done = true;
                Some(Err(error))
            },
        }
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        if self.done {
            return (0, Some(0));
        }
        let remaining = self.data.len().saturating_sub(self.offset);
        (0, Some(remaining / 8))
    }
}

impl std::iter::FusedIterator for Children<'_> {}

/// A record proven to be an `OfficeArt` container.
#[derive(Debug, Clone)]
pub struct Container<'data> {
    record: Record<'data>,
}

impl<'data> Container<'data> {
    /// Validates and wraps a container record.
    ///
    /// # Errors
    ///
    /// Returns `Error::NotContainer` if `record` does not carry record version
    /// 15 and therefore cannot contain children.
    pub fn try_new(record: Record<'data>) -> Result<Self> {
        if !record.is_container() {
            return Err(Error::NotContainer {
                kind: record.kind(),
                raw_kind: record.raw_kind(),
            });
        }
        Ok(Self { record })
    }

    /// Returns the proven container record.
    #[must_use]
    pub const fn record(&self) -> &Record<'data> {
        &self.record
    }

    /// Lazily visits direct children without copying their payloads.
    #[must_use]
    pub const fn children(&self) -> Children<'data> {
        Children::new(self.record.data())
    }

    /// Finds the first direct child of `kind`.
    ///
    /// Absence is distinct from malformed input: any malformed child encountered
    /// before a match is returned to the caller.
    ///
    /// # Errors
    ///
    /// Returns the `Error` produced by the first malformed child encountered
    /// before a match, such as `Error::TruncatedHeader` or
    /// `Error::TruncatedPayload`.
    pub fn find(&self, kind: RecordKind) -> Result<Option<Record<'data>>> {
        for child_result in self.children() {
            let child = child_result?;
            if child.kind() == kind {
                return Ok(Some(child));
            }
        }
        Ok(None)
    }

    /// Collects every direct child of `kind`.
    ///
    /// # Errors
    ///
    /// Returns the `Error` produced by the first malformed child encountered,
    /// such as `Error::TruncatedHeader` or `Error::TruncatedPayload`.
    pub fn find_all(&self, kind: RecordKind) -> Result<Vec<Record<'data>>> {
        let mut matches = Vec::new();
        for child_result in self.children() {
            let child = child_result?;
            if child.kind() == kind {
                matches.push(child);
            }
        }
        Ok(matches)
    }

    /// Collects matching descendants using an iterative depth-first traversal.
    ///
    /// # Errors
    ///
    /// Returns an error from `Container::find_recursive_with` if a descendant
    /// record is malformed or the default `Limits` are exceeded.
    pub fn find_recursive(&self, kind: RecordKind) -> Result<Vec<Record<'data>>> {
        self.find_recursive_with(kind, Limits::default())
    }

    /// Collects matching descendants within explicit resource ceilings.
    ///
    /// # Errors
    ///
    /// Returns the `Error` produced by a malformed descendant record, or
    /// `Error::LimitExceeded` when `limits.max_records` or `limits.max_depth`
    /// is exceeded during traversal.
    pub fn find_recursive_with(
        &self,
        kind: RecordKind,
        limits: Limits,
    ) -> Result<Vec<Record<'data>>> {
        let mut matches = Vec::new();
        let mut stack = vec![self.children()];
        let mut visited = 0_u32;

        while let Some(children) = stack.last_mut() {
            match children.next() {
                Some(Ok(child)) => {
                    visited = visited.checked_add(1).ok_or(Error::LimitExceeded {
                        limit: Limit::Records,
                        maximum: limits.max_records,
                    })?;
                    if visited > limits.max_records {
                        return Err(Error::LimitExceeded {
                            limit: Limit::Records,
                            maximum: limits.max_records,
                        });
                    }
                    if child.kind() == kind {
                        matches.push(child.clone());
                    }
                    if child.is_container() {
                        let depth = u16::try_from(stack.len()).unwrap_or(u16::MAX);
                        if depth >= limits.max_depth {
                            return Err(Error::LimitExceeded {
                                limit: Limit::Depth,
                                maximum: u32::from(limits.max_depth),
                            });
                        }
                        stack.push(Children::new(child.data()));
                    }
                },
                Some(Err(error)) => return Err(error),
                None => {
                    stack.pop();
                },
            }
        }

        Ok(matches)
    }
}

impl<'data> TryFrom<Record<'data>> for Container<'data> {
    type Error = Error;

    fn try_from(record: Record<'data>) -> Result<Self> {
        Self::try_new(record)
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::*;

    fn atom(kind: u16, body: &[u8]) -> Vec<u8> {
        let mut bytes = vec![0x00, 0x00];
        bytes.extend_from_slice(&kind.to_le_bytes());
        let length = u32::try_from(body.len()).expect("body length fits in u32");
        bytes.extend_from_slice(&length.to_le_bytes());
        bytes.extend_from_slice(body);
        bytes
    }

    #[test]
    fn constructor_is_checked() {
        let bytes = atom(RecordKind::Sp.raw(), &[]);
        let (record, _) = Record::parse(&bytes, 0).expect("valid atom");

        assert!(matches!(
            Container::try_new(record),
            Err(Error::NotContainer {
                kind: RecordKind::Sp,
                ..
            })
        ));
    }

    #[test]
    fn iterates_children_without_copying() {
        let mut children = atom(RecordKind::Sp.raw(), &[1, 2, 3, 4]);
        children.extend_from_slice(&atom(RecordKind::Opt.raw(), &[5, 6]));
        let records: Result<Vec<_>> = Children::new(&children).collect();
        let parsed = records.expect("valid children");

        assert_eq!(parsed.len(), 2);
        assert_eq!(parsed[0].kind(), RecordKind::Sp);
        assert_eq!(parsed[1].kind(), RecordKind::Opt);
        assert_eq!(parsed[0].data(), &children[8..12]);
    }

    #[test]
    fn malformed_child_terminates_iteration() {
        let mut bytes = atom(RecordKind::Sp.raw(), &[]);
        bytes.extend_from_slice(&[0x00, 0x00, 0x0B, 0xF0, 32, 0, 0, 0, 1]);
        bytes.extend_from_slice(&atom(RecordKind::Dg.raw(), &[]));

        let mut children = Children::new(&bytes);
        assert!(matches!(children.next(), Some(Ok(_))));
        assert!(matches!(
            children.next(),
            Some(Err(Error::TruncatedPayload { .. }))
        ));
        assert!(children.next().is_none());
        assert!(children.next().is_none());
    }

    #[test]
    fn find_distinguishes_absence_from_malformed_input() {
        let bytes = [0x00, 0x00, 0x0A, 0xF0, 1, 0, 0, 0];
        let container_record = Record::parse(
            &[
                0x0F, 0x00, 0x04, 0xF0, 8, 0, 0, 0, 0, 0, 0x0A, 0xF0, 1, 0, 0, 0,
            ],
            0,
        )
        .expect("outer record is valid")
        .0;
        let container = Container::try_new(container_record).expect("container");

        assert!(matches!(
            container.find(RecordKind::Dg),
            Err(Error::TruncatedPayload { .. })
        ));
        assert_eq!(bytes.len(), 8);
    }

    #[test]
    fn recursive_traversal_enforces_record_and_depth_limits() {
        let atom = atom(RecordKind::Sp.raw(), &[]);
        let mut nested = vec![0x0F, 0x00, 0x04, 0xF0];
        let atom_length = u32::try_from(atom.len()).expect("atom length fits in u32");
        nested.extend_from_slice(&atom_length.to_le_bytes());
        nested.extend_from_slice(&atom);
        let mut root = vec![0x0F, 0x00, 0x02, 0xF0];
        let nested_length = u32::try_from(nested.len()).expect("nested length fits in u32");
        root.extend_from_slice(&nested_length.to_le_bytes());
        root.extend_from_slice(&nested);

        let record = Record::parse(&root, 0).expect("valid root").0;
        let container = Container::try_new(record).expect("container");
        assert!(matches!(
            container.find_recursive_with(
                RecordKind::Sp,
                Limits {
                    max_depth: 1,
                    max_records: 10,
                },
            ),
            Err(Error::LimitExceeded {
                limit: Limit::Depth,
                ..
            })
        ));
        assert!(matches!(
            container.find_recursive_with(
                RecordKind::Sp,
                Limits {
                    max_depth: 10,
                    max_records: 0,
                },
            ),
            Err(Error::LimitExceeded {
                limit: Limit::Records,
                ..
            })
        ));
    }
}
