//! Source-aware relationship-ID allocation for one SVG drawing transaction.
//!
//! The lifecycle planner may detach an SVG relationship and later attach a
//! replacement in the same drawing.  Re-scanning the current relationship map
//! for every attachment both reuses freed IDs and turns a dense batch into an
//! `O(attachments * relationships)` scan.  This allocator scans the source
//! relationship set once, then advances a checked monotonic candidate cursor
//! for the rest of the drawing plan.

use litchi_opc::Relationships;

use crate::error::{Result, allocation, invalid};

const BASE_ID: &str = "rIdSvg";
const MAX_CANDIDATE_ATTEMPTS: usize = 100_000;

/// Monotonic allocator for the generated SVG relationship-ID family.
///
/// The type and methods are visible only to the parent lifecycle planner.  A
/// clone is intentionally not provided: two independent cursors for one
/// drawing could issue the same transaction-local ID.
#[derive(Debug)]
pub(super) struct RelationshipIdAllocator {
    next_index: u64,
}

impl RelationshipIdAllocator {
    /// Build an allocator from the source relationship collection.
    ///
    /// Only the exact, case-sensitive `rIdSvg`/`rIdSvgN` family participates
    /// in the cursor.  Other XML IDs, including case variants and leading-zero
    /// spellings, remain independent IDs under XML's case-sensitive rules.
    pub(super) fn from_source(source: &Relationships) -> Result<Self> {
        let mut next_index = 0u64;
        for relationship in source.iter() {
            let Some(index) = generated_index(relationship.r_id()) else {
                continue;
            };
            let after = index.checked_add(1).ok_or_else(|| {
                invalid("SVG relationship ID suffix exceeds the checked allocator domain")
            })?;
            if after > next_index {
                next_index = after;
            }
        }
        Ok(Self { next_index })
    }

    /// Allocate the next source-aware ID that is absent from the current map.
    ///
    /// Advancing the cursor before returning makes the allocation durable for
    /// this transaction even when the caller subsequently removes the
    /// relationship or has not inserted the returned edge yet.
    pub(super) fn allocate(&mut self, current: &Relationships) -> Result<String> {
        for _ in 0..MAX_CANDIDATE_ATTEMPTS {
            let index = self.next_index;
            let next = index.checked_add(1).ok_or_else(|| {
                invalid("SVG relationship ID allocator exhausted its checked domain")
            })?;
            let candidate = candidate(index)?;
            self.next_index = next;
            if current.get(&candidate).is_none() {
                return Ok(candidate);
            }
        }
        Err(invalid(format!(
            "SVG relationship ID candidates exceed {MAX_CANDIDATE_ATTEMPTS}"
        )))
    }
}

fn generated_index(id: &str) -> Option<u64> {
    if id == BASE_ID {
        return Some(0);
    }
    let suffix = id.strip_prefix(BASE_ID)?;
    // The allocator emits no leading-zero suffix and never emits `rIdSvg0`.
    // Treating those spellings as independent preserves exact XML ID
    // semantics while still allowing a canonical candidate to be allocated.
    if suffix.is_empty() || suffix.as_bytes().first().is_none_or(|byte| *byte == b'0') {
        return None;
    }
    if !suffix.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    let mut value = 0u64;
    for byte in suffix.bytes() {
        value = value.checked_mul(10)?.checked_add(u64::from(byte - b'0'))?;
    }
    Some(value)
}

fn candidate(index: u64) -> Result<String> {
    if index == 0 {
        return Ok(BASE_ID.to_owned());
    }
    let mut output = String::new();
    output
        .try_reserve_exact(BASE_ID.len() + decimal_len(index))
        .map_err(|source| allocation("SVG relationship ID candidate", source))?;
    output.push_str(BASE_ID);
    use std::fmt::Write as _;
    write!(&mut output, "{index}")
        .map_err(|_| invalid("SVG relationship ID candidate formatting failed"))?;
    Ok(output)
}

fn decimal_len(mut value: u64) -> usize {
    let mut length = 1;
    while value >= 10 {
        value /= 10;
        length += 1;
    }
    length
}

#[cfg(test)]
mod tests {
    use litchi_opc::{Relationships, TargetMode};

    use super::RelationshipIdAllocator;

    fn relationships(ids: &[&str]) -> Relationships {
        let mut relationships = Relationships::new("/xl/drawings/drawing1.xml".to_owned());
        for (index, id) in ids.iter().enumerate() {
            relationships
                .try_add_relationship(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                        .to_owned(),
                    format!("../media/image{index}.svg"),
                    (*id).to_owned(),
                    TargetMode::Internal,
                )
                .expect("fixture relationship IDs are unique");
        }
        relationships
    }

    fn add(relationships: &mut Relationships, id: &str) {
        relationships
            .try_add_relationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                    .to_owned(),
                "../media/generated.svg".to_owned(),
                id.to_owned(),
                TargetMode::Internal,
            )
            .expect("generated fixture relationship ID is free");
    }

    #[test]
    fn dense_source_ids_advance_once_without_rechecking_the_prefix() {
        let source = relationships(&["rIdSvg", "rIdSvg1", "rIdSvg2"]);
        let mut allocator = RelationshipIdAllocator::from_source(&source).unwrap();
        assert_eq!(allocator.allocate(&source).unwrap(), "rIdSvg3");
        // The caller has not inserted the first candidate yet; it is still
        // consumed by this transaction's monotonic cursor.
        assert_eq!(allocator.allocate(&source).unwrap(), "rIdSvg4");
    }

    #[test]
    fn holes_and_case_variants_do_not_change_exact_id_semantics() {
        let source = relationships(&["rIdSvg", "rIdSvg2", "rIdSvg4", "ridsvg", "RIDSVG"]);
        let mut allocator = RelationshipIdAllocator::from_source(&source).unwrap();
        assert_eq!(allocator.allocate(&source).unwrap(), "rIdSvg5");

        let case_only = relationships(&["ridsvg", "RIDSVG"]);
        let mut case_allocator = RelationshipIdAllocator::from_source(&case_only).unwrap();
        assert_eq!(case_allocator.allocate(&case_only).unwrap(), "rIdSvg");
    }

    #[test]
    fn detached_and_transaction_issued_ids_are_never_reused() {
        let mut current = relationships(&["rIdSvg"]);
        let mut allocator = RelationshipIdAllocator::from_source(&current).unwrap();

        let first = allocator.allocate(&current).unwrap();
        add(&mut current, &first);
        current.remove(&first);
        assert_eq!(allocator.allocate(&current).unwrap(), "rIdSvg2");

        // A candidate that was issued but never inserted is also consumed.
        assert_eq!(allocator.allocate(&current).unwrap(), "rIdSvg3");
    }

    #[test]
    fn checked_suffix_growth_refuses_u64_overflow() {
        let source = relationships(&["rIdSvg18446744073709551614"]);
        let mut allocator = RelationshipIdAllocator::from_source(&source).unwrap();
        assert!(allocator.allocate(&source).is_err());

        let unrepresentable = relationships(&["rIdSvg18446744073709551615"]);
        assert!(RelationshipIdAllocator::from_source(&unrepresentable).is_err());
    }
}
