//! Checked, lazy traversal of OfficeArt container records.

use crate::{Error, Limit, Record, RecordKind, Result};

/// Resource ceilings for recursive OfficeArt traversal.
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
            Ok((record, consumed)) => match self.offset.checked_add(consumed) {
                Some(next) => {
                    self.offset = next;
                    Some(Ok(record))
                },
                None => {
                    self.done = true;
                    Some(Err(Error::ArithmeticOverflow {
                        context: "child-record cursor",
                    }))
                },
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

/// A record proven to be an OfficeArt container.
#[derive(Debug, Clone)]
pub struct Container<'data> {
    record: Record<'data>,
}

impl<'data> Container<'data> {
    /// Validates and wraps a container record.
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
    pub const fn record(&self) -> &Record<'data> {
        &self.record
    }

    /// Lazily visits direct children without copying their payloads.
    pub const fn children(&self) -> Children<'data> {
        Children::new(self.record.data())
    }

    /// Finds the first direct child of `kind`.
    ///
    /// Absence is distinct from malformed input: any malformed child encountered
    /// before a match is returned to the caller.
    pub fn find(&self, kind: RecordKind) -> Result<Option<Record<'data>>> {
        for child in self.children() {
            let child = child?;
            if child.kind() == kind {
                return Ok(Some(child));
            }
        }
        Ok(None)
    }

    /// Collects every direct child of `kind`.
    pub fn find_all(&self, kind: RecordKind) -> Result<Vec<Record<'data>>> {
        let mut matches = Vec::new();
        for child in self.children() {
            let child = child?;
            if child.kind() == kind {
                matches.push(child);
            }
        }
        Ok(matches)
    }

    /// Collects matching descendants using an iterative depth-first traversal.
    pub fn find_recursive(&self, kind: RecordKind) -> Result<Vec<Record<'data>>> {
        self.find_recursive_with(kind, Limits::default())
    }

    /// Collects matching descendants within explicit resource ceilings.
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
    use super::*;

    fn atom(kind: u16, body: &[u8]) -> Vec<u8> {
        let mut bytes = vec![0x00, 0x00];
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&(body.len() as u32).to_le_bytes());
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
        let records = records.expect("valid children");

        assert_eq!(records.len(), 2);
        assert_eq!(records[0].kind(), RecordKind::Sp);
        assert_eq!(records[1].kind(), RecordKind::Opt);
        assert_eq!(records[0].data(), &children[8..12]);
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
        nested.extend_from_slice(&(atom.len() as u32).to_le_bytes());
        nested.extend_from_slice(&atom);
        let mut root = vec![0x0F, 0x00, 0x02, 0xF0];
        root.extend_from_slice(&(nested.len() as u32).to_le_bytes());
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
