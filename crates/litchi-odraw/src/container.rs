//! Checked, lazy traversal of `OfficeArt` container records.

use crate::{Error, LimitKind, Record, RecordKind, Result};

/// Resource ceilings for recursive `OfficeArt` traversal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Largest record body accepted by the raw framing parser.
    pub max_record_bytes: u32,
    /// Maximum nested-container depth.
    pub max_depth: u16,
    /// Maximum records visited during one traversal.
    pub max_records: u32,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_record_bytes: 256 * 1024 * 1024,
            max_depth: 64,
            max_records: 1_000_000,
        }
    }
}

impl Limits {
    /// Intersects two raw framing and traversal policies before a borrowed
    /// record is handed to a narrower consumer.
    #[must_use]
    pub const fn constrained_by(self, other: Self) -> Self {
        Self {
            max_record_bytes: if self.max_record_bytes < other.max_record_bytes {
                self.max_record_bytes
            } else {
                other.max_record_bytes
            },
            max_depth: if self.max_depth < other.max_depth {
                self.max_depth
            } else {
                other.max_depth
            },
            max_records: if self.max_records < other.max_records {
                self.max_records
            } else {
                other.max_records
            },
        }
    }
}

/// A zero-copy iterator over the direct children of a container.
#[derive(Debug, Clone)]
pub struct Children<'data> {
    data: &'data [u8],
    limits: Limits,
    offset: usize,
    seen: u32,
    done: bool,
}

impl<'data> Children<'data> {
    /// Creates an iterator over a validated container body or record sequence.
    #[must_use]
    pub const fn new(data: &'data [u8], limits: Limits) -> Self {
        Self {
            data,
            limits,
            offset: 0,
            seen: 0,
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

        if self.seen >= self.limits.max_records {
            self.done = true;
            return Some(Err(Error::LimitExceeded {
                limit: LimitKind::Records,
                observed: u64::from(self.seen) + 1,
                maximum: self.limits.max_records,
            }));
        }

        if self.data.len().saturating_sub(self.offset) < 8 {
            self.done = true;
            return Some(Err(Error::TruncatedHeader {
                offset: self.offset,
                available: self.data.len().saturating_sub(self.offset),
            }));
        }

        match Record::parse(self.data, self.offset, self.limits) {
            Ok((record, consumed)) => {
                self.seen += 1;
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
        if self.done || self.offset == self.data.len() {
            return (0, Some(0));
        }
        let remaining = self.data.len().saturating_sub(self.offset);
        let admitted = usize::try_from(self.limits.max_records - self.seen).unwrap_or(usize::MAX);
        // Every nonempty remainder can yield one pending framing or admission
        // error, including a zero-record budget and a one-to-seven-byte tail.
        // Count complete minimum-sized records, round up for that pending
        // error, and cap by the admitted records plus that one error.
        let wire_items = remaining.div_ceil(8);
        (0, Some(wire_items.min(admitted.saturating_add(1))))
    }
}

impl std::iter::FusedIterator for Children<'_> {}

/// A compact traversal cursor used by recursive searches.  Unlike the public
/// [`Children`] iterator it keeps the effective policy in the traversal
/// argument, so each stack frame only retains its source range and cursor.
#[derive(Clone, Copy)]
pub(crate) struct Frame<'data> {
    data: &'data [u8],
    offset: usize,
    done: bool,
}

impl<'data> Frame<'data> {
    pub(crate) const fn new(data: &'data [u8]) -> Self {
        Self {
            data,
            offset: 0,
            done: false,
        }
    }

    pub(crate) fn next(&mut self, limits: Limits) -> Option<Result<Record<'data>>> {
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

        match Record::parse(self.data, self.offset, limits) {
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
}

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
        Children::new(self.record.data(), self.record.limits())
    }

    /// Lazily visits direct children under an explicit caller policy.
    ///
    /// This is crate-private because public projections use the limits retained
    /// by the record at admission.  Typed decoders use it when a caller has
    /// supplied a narrower traversal policy for an already-admitted record.
    pub(crate) const fn children_under(&self, limits: Limits) -> Children<'data> {
        Children::new(
            self.record.data(),
            self.record.limits().constrained_by(limits),
        )
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

    /// Collects matching descendants within explicit resource ceilings.
    ///
    /// # Errors
    ///
    /// Returns the `Error` produced by a malformed descendant record, or
    /// `Error::LimitExceeded` when `limits.max_records` or `limits.max_depth`
    /// is exceeded during traversal.
    pub fn find_recursive(&self, kind: RecordKind, limits: Limits) -> Result<Vec<Record<'data>>> {
        let limits = self.record.limits().constrained_by(limits);
        let mut matches = Vec::new();
        let mut frame = Frame::new(self.record.data());
        let mut stack = Vec::new();
        let mut visited = 0_u32;

        loop {
            match frame.next(limits) {
                Some(Ok(child)) => {
                    visited = visited.checked_add(1).ok_or(Error::LimitExceeded {
                        limit: LimitKind::Records,
                        observed: u64::from(visited) + 1,
                        maximum: limits.max_records,
                    })?;
                    if visited > limits.max_records {
                        return Err(Error::LimitExceeded {
                            limit: LimitKind::Records,
                            observed: u64::from(visited),
                            maximum: limits.max_records,
                        });
                    }
                    if child.kind() == kind {
                        matches.push(child.clone());
                    }
                    if child.is_container() {
                        let depth =
                            u16::try_from(stack.len().saturating_add(1)).unwrap_or(u16::MAX);
                        if depth >= limits.max_depth {
                            return Err(Error::LimitExceeded {
                                limit: LimitKind::Depth,
                                observed: u64::from(depth),
                                maximum: u32::from(limits.max_depth),
                            });
                        }
                        stack.push(frame);
                        frame = Frame::new(child.data());
                    }
                },
                Some(Err(error)) => return Err(error),
                None => match stack.pop() {
                    Some(parent) => frame = parent,
                    None => break,
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
        let (record, _) = Record::parse(&bytes, 0, Limits::default()).expect("valid atom");

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
        let records: Result<Vec<_>> = Children::new(&children, Limits::default()).collect();
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

        let mut children = Children::new(&bytes, Limits::default());
        assert!(matches!(children.next(), Some(Ok(_))));
        assert!(matches!(
            children.next(),
            Some(Err(Error::TruncatedPayload { .. }))
        ));
        assert!(children.next().is_none());
        assert!(children.next().is_none());
    }

    #[test]
    fn child_iteration_enforces_the_record_admission_limit() {
        let mut bytes = atom(RecordKind::Sp.raw(), &[]);
        bytes.extend_from_slice(&atom(RecordKind::Opt.raw(), &[]));
        let mut children = Children::new(
            &bytes,
            Limits {
                max_records: 1,
                ..Limits::default()
            },
        );
        assert!(matches!(children.next(), Some(Ok(_))));
        assert!(matches!(
            children.next(),
            Some(Err(Error::LimitExceeded {
                limit: LimitKind::Records,
                observed: 2,
                maximum: 1,
            }))
        ));
        assert!(children.next().is_none());

        let mut none = Children::new(
            &bytes,
            Limits {
                max_records: 0,
                ..Limits::default()
            },
        );
        assert!(matches!(
            none.next(),
            Some(Err(Error::LimitExceeded {
                limit: LimitKind::Records,
                observed: 1,
                maximum: 0,
            }))
        ));
        assert!(none.next().is_none());
    }

    #[test]
    fn child_size_hint_allows_a_pending_error() {
        let bytes = atom(RecordKind::Sp.raw(), &[]);
        let mut zero = Children::new(
            &bytes,
            Limits {
                max_records: 0,
                ..Limits::default()
            },
        );
        assert_eq!(zero.size_hint(), (0, Some(1)));
        assert!(zero.next().is_some_and(|result| result.is_err()));
        assert_eq!(zero.size_hint(), (0, Some(0)));

        for length in 1..8 {
            let tail = vec![0_u8; length];
            let mut children = Children::new(&tail, Limits::default());
            assert_eq!(children.size_hint(), (0, Some(1)));
            assert!(children.next().is_some_and(|result| result.is_err()));
            assert_eq!(children.size_hint(), (0, Some(0)));
        }
    }

    #[test]
    fn find_distinguishes_absence_from_malformed_input() {
        let bytes = [0x00, 0x00, 0x0A, 0xF0, 1, 0, 0, 0];
        let container_record = Record::parse(
            &[
                0x0F, 0x00, 0x04, 0xF0, 8, 0, 0, 0, 0, 0, 0x0A, 0xF0, 1, 0, 0, 0,
            ],
            0,
            Limits::default(),
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
    fn recursive_traversal_resumes_each_parent_in_wire_order() {
        fn wrap(body: &[u8]) -> Vec<u8> {
            let mut bytes = vec![0x0f, 0, 0x04, 0xf0];
            bytes.extend_from_slice(&u32::try_from(body.len()).unwrap().to_le_bytes());
            bytes.extend_from_slice(body);
            bytes
        }
        let deepest = wrap(&atom(RecordKind::Sp.raw(), &[3]));
        let mut middle = atom(RecordKind::Sp.raw(), &[2]);
        middle.extend_from_slice(&deepest);
        middle.extend_from_slice(&atom(RecordKind::Sp.raw(), &[4]));
        let mut body = atom(RecordKind::Sp.raw(), &[1]);
        body.extend_from_slice(&wrap(&middle));
        body.extend_from_slice(&atom(RecordKind::Sp.raw(), &[5]));
        let bytes = wrap(&body);
        let root =
            Container::try_new(Record::parse(&bytes, 0, Limits::default()).unwrap().0).unwrap();
        let limits = Limits {
            max_depth: 3,
            max_records: 7,
            ..Limits::default()
        };
        let records = root.find_recursive(RecordKind::Sp, limits).unwrap();
        assert_eq!(
            records.iter().map(|r| r.data()[0]).collect::<Vec<_>>(),
            [1, 2, 3, 4, 5]
        );
        assert!(matches!(
            root.find_recursive(
                RecordKind::Sp,
                Limits {
                    max_records: 6,
                    ..limits
                }
            ),
            Err(Error::LimitExceeded {
                limit: LimitKind::Records,
                maximum: 6,
                ..
            })
        ));
        assert!(matches!(
            root.find_recursive(
                RecordKind::Sp,
                Limits {
                    max_depth: 2,
                    ..limits
                }
            ),
            Err(Error::LimitExceeded {
                limit: LimitKind::Depth,
                maximum: 2,
                ..
            })
        ));
        let flat = wrap(&atom(RecordKind::Sp.raw(), &[9]));
        let root =
            Container::try_new(Record::parse(&flat, 0, Limits::default()).unwrap().0).unwrap();
        assert_eq!(
            root.find_recursive(
                RecordKind::Sp,
                Limits {
                    max_depth: 0,
                    max_records: 1,
                    ..Limits::default()
                }
            )
            .unwrap()[0]
                .data(),
            &[9]
        );
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

        let record = Record::parse(&root, 0, Limits::default())
            .expect("valid root")
            .0;
        let container = Container::try_new(record).expect("container");
        assert!(matches!(
            container.find_recursive(
                RecordKind::Sp,
                Limits {
                    max_depth: 1,
                    max_records: 10,
                    ..Limits::default()
                },
            ),
            Err(Error::LimitExceeded {
                limit: LimitKind::Depth,
                ..
            })
        ));
        assert!(matches!(
            container.find_recursive(
                RecordKind::Sp,
                Limits {
                    max_depth: 10,
                    max_records: 0,
                    ..Limits::default()
                },
            ),
            Err(Error::LimitExceeded {
                limit: LimitKind::Records,
                ..
            })
        ));
    }

    #[test]
    fn recursive_traversal_intersects_the_retained_root_policy() {
        let mut body = atom(RecordKind::Sp.raw(), &[]);
        body.extend_from_slice(&atom(RecordKind::Sp.raw(), &[]));
        let mut root = vec![0x0F, 0x00, 0x04, 0xF0];
        root.extend_from_slice(&u32::try_from(body.len()).unwrap().to_le_bytes());
        root.extend_from_slice(&body);
        let origin = Limits {
            max_records: 1,
            ..Limits::default()
        };
        let record = Record::parse(&root, 0, origin).unwrap().0;
        let container = Container::try_new(record).unwrap();
        assert!(matches!(
            container.find_recursive(RecordKind::Sp, Limits::default()),
            Err(Error::LimitExceeded {
                limit: LimitKind::Records,
                maximum: 1,
                ..
            })
        ));

        let mut nested = vec![0x0F, 0x00, 0x04, 0xF0];
        nested.extend_from_slice(&8_u32.to_le_bytes());
        nested.extend_from_slice(&atom(RecordKind::Sp.raw(), &[]));
        let mut root = vec![0x0F, 0x00, 0x04, 0xF0];
        root.extend_from_slice(&u32::try_from(nested.len()).unwrap().to_le_bytes());
        root.extend_from_slice(&nested);
        let origin = Limits {
            max_depth: 1,
            ..Limits::default()
        };
        let record = Record::parse(&root, 0, origin).unwrap().0;
        let container = Container::try_new(record).unwrap();
        assert!(matches!(
            container.find_recursive(RecordKind::Sp, Limits::default()),
            Err(Error::LimitExceeded {
                limit: LimitKind::Depth,
                maximum: 1,
                ..
            })
        ));
    }
}
