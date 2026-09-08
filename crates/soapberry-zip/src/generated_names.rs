//! A bounded, checked language for deterministic generated ZIP member names.
//!
//! A generated-name plan describes a finite sequence without materializing the
//! sequence.  Literals account for one member and indexed patterns account for
//! one or more decimal values.  The builder proves that the descriptors are
//! disjoint under exact and ASCII-case-equivalent comparison and that no
//! descriptor can produce a whole-component ancestor of another descriptor.
//! Consequently, a writer can retain only a scalar cursor while it publishes
//! the names in order.

use crate::{Error, ErrorKind};

const MAX_DECIMAL_DIGITS: usize = 20;

/// Budgets used while constructing a [`GeneratedNamePlan`].
///
/// `max_patterns` counts literal descriptors and indexed pattern pairs.  The
/// count is a bound on the plan representation, not on the number of members
/// emitted by an indexed range.  `max_pattern_bytes` counts the bytes in the
/// retained literal and prefix/suffix strings.  Decimal expansions are
/// bounded by the size of a `u64` and are accounted for by
/// [`GeneratedNamePlan::max_name_bytes`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct GeneratedNamePlanLimits {
    /// Maximum number of retained literal and indexed pattern descriptors.
    pub max_patterns: usize,
    /// Maximum bytes retained by all literal, prefix, and suffix strings.
    pub max_pattern_bytes: usize,
    /// Maximum number of generated members in the complete sequence.
    pub max_entries: u64,
}

impl GeneratedNamePlanLimits {
    /// Creates a generated-name plan budget.
    pub const fn new(max_patterns: usize, max_pattern_bytes: usize, max_entries: u64) -> Self {
        Self {
            max_patterns,
            max_pattern_bytes,
            max_entries,
        }
    }
}

/// A builder for a checked sequence of deterministic member names.
#[derive(Debug)]
pub struct GeneratedNamePlanBuilder {
    limits: GeneratedNamePlanLimits,
    descriptors: Vec<Descriptor>,
    pattern_count: usize,
    pattern_bytes: usize,
    entry_count: u64,
    max_name_bytes: usize,
}

/// A checked generated-name sequence.
///
/// The plan retains only the fixed descriptor set and scalar cursor state.
/// Indexed ranges are never expanded during construction or validation.
#[derive(Debug)]
pub struct GeneratedNamePlan {
    descriptors: Vec<Descriptor>,
    entry_count: u64,
    max_name_bytes: usize,
    cursor: Cursor,
}

#[derive(Debug)]
struct Cursor {
    descriptor: usize,
    pattern: usize,
    index: u64,
}

#[derive(Debug)]
enum Descriptor {
    Literal { name: String },
    Indexed(IndexedDescriptor),
}

#[derive(Debug)]
struct IndexedDescriptor {
    first: u64,
    last: u64,
    patterns: Vec<IndexedPattern>,
}

#[derive(Debug)]
struct IndexedPattern {
    prefix: String,
    suffix: String,
}

impl GeneratedNamePlanBuilder {
    /// Creates an empty builder under `limits`.
    pub fn new(limits: GeneratedNamePlanLimits) -> Result<Self, Error> {
        Ok(Self {
            limits,
            descriptors: Vec::new(),
            pattern_count: 0,
            pattern_bytes: 0,
            entry_count: 0,
            max_name_bytes: 0,
        })
    }

    /// Appends one canonical literal member name.
    pub fn push_literal(&mut self, name: &str) -> Result<(), Error> {
        validate_literal_name(name)?;
        self.ensure_pattern_budget(1, name.len())?;
        let next_entries = self
            .entry_count
            .checked_add(1)
            .ok_or_else(|| invalid("generated name plan entry count overflows u64"))?;
        if next_entries > self.limits.max_entries {
            return Err(invalid("generated name plan exceeds its entry budget"));
        }

        let candidate = Descriptor::Literal {
            name: clone_string(name, "generated name plan literal")?,
        };
        self.ensure_candidate_is_disjoint(&candidate)?;
        self.descriptors
            .try_reserve(1)
            .map_err(|source| allocation("generated name plan descriptors", source))?;
        self.descriptors.push(candidate);
        self.pattern_count += 1;
        self.pattern_bytes += name.len();
        self.entry_count = next_entries;
        self.max_name_bytes = self.max_name_bytes.max(name.len());
        Ok(())
    }

    /// Appends indexed patterns in supplied order for every checked decimal
    /// value in `first..first + count`.
    ///
    /// A pattern's prefix and suffix surround one canonical decimal run in
    /// the final path component.  The static strings are retained, while the
    /// range itself is represented by its endpoints.
    pub fn push_indexed(
        &mut self,
        first: u64,
        count: u64,
        patterns: &[(&str, &str)],
    ) -> Result<(), Error> {
        if count == 0 {
            return Err(invalid("generated name plan indexed range is empty"));
        }
        if patterns.is_empty() {
            return Err(invalid("generated name plan indexed range has no patterns"));
        }
        let last = first
            .checked_add(count - 1)
            .ok_or_else(|| invalid("generated name plan indexed range overflows u64"))?;
        let pattern_count = patterns.len();
        let static_bytes = patterns
            .iter()
            .try_fold(0usize, |total, (prefix, suffix)| {
                let pair_bytes = prefix
                    .len()
                    .checked_add(suffix.len())
                    .ok_or_else(|| invalid("generated name plan pattern bytes overflow usize"))?;
                total
                    .checked_add(pair_bytes)
                    .ok_or_else(|| invalid("generated name plan pattern bytes overflow usize"))
            })?;
        self.ensure_pattern_budget(pattern_count, static_bytes)?;
        for &(prefix, suffix) in patterns {
            validate_indexed_pattern(prefix, suffix, first, last)?;
        }
        let pattern_count_u64 = u64::try_from(pattern_count)
            .map_err(|_| invalid("generated name plan pattern count exceeds u64"))?;
        let added_entries = count
            .checked_mul(pattern_count_u64)
            .ok_or_else(|| invalid("generated name plan entry count overflows u64"))?;
        let next_entries = self
            .entry_count
            .checked_add(added_entries)
            .ok_or_else(|| invalid("generated name plan entry count overflows u64"))?;
        if next_entries > self.limits.max_entries {
            return Err(invalid("generated name plan exceeds its entry budget"));
        }

        let mut owned_patterns = Vec::new();
        owned_patterns
            .try_reserve_exact(pattern_count)
            .map_err(|source| allocation("generated name plan indexed patterns", source))?;
        for &(prefix, suffix) in patterns {
            owned_patterns.push(IndexedPattern {
                prefix: clone_string(prefix, "generated name plan prefix")?,
                suffix: clone_string(suffix, "generated name plan suffix")?,
            });
        }
        let candidate = Descriptor::Indexed(IndexedDescriptor {
            first,
            last,
            patterns: owned_patterns,
        });
        self.ensure_candidate_is_disjoint(&candidate)?;
        self.descriptors
            .try_reserve(1)
            .map_err(|source| allocation("generated name plan descriptors", source))?;
        self.descriptors.push(candidate);
        self.pattern_count += pattern_count;
        self.pattern_bytes += static_bytes;
        self.entry_count = next_entries;
        let digits = decimal_digits(last);
        for &(prefix, suffix) in patterns {
            let name_bytes = prefix
                .len()
                .checked_add(suffix.len())
                .and_then(|bytes| bytes.checked_add(digits))
                .ok_or_else(|| invalid("generated name plan name bytes overflow usize"))?;
            self.max_name_bytes = self.max_name_bytes.max(name_bytes);
        }
        Ok(())
    }

    /// Finishes the checked plan.
    pub fn finish(self) -> Result<GeneratedNamePlan, Error> {
        // Every mutation checks the new descriptor against all previous
        // descriptors.  Repeat the bounded descriptor proof at publication so
        // the invariant is local to the resulting capability even if the
        // builder implementation gains another insertion path later.
        validate_descriptor_set(&self.descriptors)?;
        let index = match self.descriptors.first() {
            Some(Descriptor::Literal { .. }) => 0,
            Some(Descriptor::Indexed(indexed)) => indexed.first,
            None => 0,
        };
        Ok(GeneratedNamePlan {
            descriptors: self.descriptors,
            entry_count: self.entry_count,
            max_name_bytes: self.max_name_bytes,
            cursor: Cursor {
                descriptor: 0,
                pattern: 0,
                index,
            },
        })
    }

    fn ensure_pattern_budget(&self, additional: usize, bytes: usize) -> Result<(), Error> {
        let next_patterns = self
            .pattern_count
            .checked_add(additional)
            .ok_or_else(|| invalid("generated name plan pattern count overflows usize"))?;
        if next_patterns > self.limits.max_patterns {
            return Err(invalid("generated name plan exceeds its pattern budget"));
        }
        let next_bytes = self
            .pattern_bytes
            .checked_add(bytes)
            .ok_or_else(|| invalid("generated name plan pattern bytes overflow usize"))?;
        if next_bytes > self.limits.max_pattern_bytes {
            return Err(invalid(
                "generated name plan exceeds its pattern-byte budget",
            ));
        }
        Ok(())
    }

    fn ensure_candidate_is_disjoint(&self, candidate: &Descriptor) -> Result<(), Error> {
        for existing in &self.descriptors {
            validate_descriptor_pair(existing, candidate)?;
        }
        if let Descriptor::Indexed(indexed) = candidate {
            for (left, first) in indexed.patterns.iter().enumerate() {
                for second in indexed.patterns.iter().skip(left + 1) {
                    check_family_pair(
                        indexed.first,
                        indexed.last,
                        first,
                        indexed.first,
                        indexed.last,
                        second,
                    )?;
                }
            }
        }
        Ok(())
    }
}

impl GeneratedNamePlan {
    /// Number of members represented by this plan.
    #[must_use]
    pub const fn entry_count(&self) -> u64 {
        self.entry_count
    }

    /// Maximum bytes in one generated member name.
    #[must_use]
    pub const fn max_name_bytes(&self) -> usize {
        self.max_name_bytes
    }

    /// Checks the exact next canonical member name without advancing.
    ///
    /// This is crate-private so only the generated writer adapters can hold
    /// the sequencing capability.  Public callers cannot manufacture or skip
    /// a checked cursor.
    pub(crate) fn check_next(&self, name: &str) -> Result<(), Error> {
        validate_literal_name(name)?;
        let descriptor = self
            .descriptors
            .get(self.cursor.descriptor)
            .ok_or_else(|| invalid("generated name plan is already complete"))?;
        let matches = match descriptor {
            Descriptor::Literal { name: expected } => expected == name,
            Descriptor::Indexed(indexed) => {
                let pattern = indexed
                    .patterns
                    .get(self.cursor.pattern)
                    .ok_or_else(|| invalid("generated name plan cursor pattern is invalid"))?;
                matches_exact_numbered_name(name, pattern, self.cursor.index)
            },
        };
        if matches {
            Ok(())
        } else {
            Err(invalid(
                "generated member name does not match the next plan item",
            ))
        }
    }

    /// Advances the private cursor after the matching entry has finalized.
    pub(crate) fn advance(&mut self) -> Result<(), Error> {
        let descriptor = self
            .descriptors
            .get(self.cursor.descriptor)
            .ok_or_else(|| invalid("generated name plan is already complete"))?;
        let indexed = match descriptor {
            Descriptor::Literal { .. } => None,
            Descriptor::Indexed(indexed) => Some((indexed.patterns.len(), indexed.last)),
        };
        match indexed {
            None => {
                self.cursor.descriptor += 1;
                self.reset_cursor_at_descriptor();
            },
            Some((pattern_count, last)) => {
                if self.cursor.pattern < pattern_count.saturating_sub(1) {
                    self.cursor.pattern += 1;
                } else if self.cursor.index < last {
                    self.cursor.pattern = 0;
                    self.cursor.index += 1;
                } else {
                    self.cursor.descriptor += 1;
                    self.reset_cursor_at_descriptor();
                }
            },
        }
        Ok(())
    }

    fn reset_cursor_at_descriptor(&mut self) {
        let index = match self.descriptors.get(self.cursor.descriptor) {
            Some(Descriptor::Indexed(indexed)) => indexed.first,
            Some(Descriptor::Literal { .. }) | None => 0,
        };
        self.cursor.pattern = 0;
        self.cursor.index = index;
    }

    /// Returns whether every plan item has been finalized.
    #[must_use]
    pub(crate) fn is_complete(&self) -> bool {
        self.cursor.descriptor >= self.descriptors.len()
    }
}

fn validate_descriptor_set(descriptors: &[Descriptor]) -> Result<(), Error> {
    for (left, first) in descriptors.iter().enumerate() {
        for second in descriptors.iter().skip(left + 1) {
            validate_descriptor_pair(first, second)?;
        }
        if let Descriptor::Indexed(indexed) = first {
            for (pattern_index, pattern) in indexed.patterns.iter().enumerate() {
                for second in indexed.patterns.iter().skip(pattern_index + 1) {
                    check_family_pair(
                        indexed.first,
                        indexed.last,
                        pattern,
                        indexed.first,
                        indexed.last,
                        second,
                    )?;
                }
            }
        }
    }
    Ok(())
}

fn validate_descriptor_pair(first: &Descriptor, second: &Descriptor) -> Result<(), Error> {
    match (first, second) {
        (Descriptor::Literal { name: left }, Descriptor::Literal { name: right }) => {
            check_literal_pair(left, right)
        },
        (Descriptor::Literal { name }, Descriptor::Indexed(indexed)) => {
            for pattern in &indexed.patterns {
                check_literal_family(name, indexed.first, indexed.last, pattern)?;
            }
            Ok(())
        },
        (Descriptor::Indexed(indexed), Descriptor::Literal { name }) => {
            for pattern in &indexed.patterns {
                check_literal_family(name, indexed.first, indexed.last, pattern)?;
            }
            Ok(())
        },
        (Descriptor::Indexed(left), Descriptor::Indexed(right)) => {
            for first in &left.patterns {
                for second in &right.patterns {
                    check_family_pair(
                        left.first,
                        left.last,
                        first,
                        right.first,
                        right.last,
                        second,
                    )?;
                }
            }
            Ok(())
        },
    }
}

fn check_literal_pair(left: &str, right: &str) -> Result<(), Error> {
    if folded_eq(left, right) {
        return Err(duplicate_conflict());
    }
    if is_component_prefix(left, right) || is_component_prefix(right, left) {
        return Err(ancestor_conflict());
    }
    Ok(())
}

fn check_literal_family(
    literal: &str,
    first: u64,
    last: u64,
    pattern: &IndexedPattern,
) -> Result<(), Error> {
    if literal_matches_family(literal, first, last, pattern) {
        return Err(duplicate_conflict());
    }
    let literal_depth = component_count(literal);
    let family = FamilyParts::new(pattern);
    if literal_depth < family.depth {
        if !component_prefix_matches(literal, family.parents, literal_depth - 1) {
            return Ok(());
        }
        let Some(parent) = component_at(family.parents, literal_depth - 1) else {
            return Ok(());
        };
        if folded_eq(
            component_at(literal, literal_depth - 1).unwrap_or_default(),
            parent,
        ) {
            return Err(ancestor_conflict());
        }
    } else if literal_depth > family.depth {
        if !component_prefix_matches(family.parents, literal, family.depth - 1) {
            return Ok(());
        }
        let Some(literal_component) = component_at(literal, family.depth - 1) else {
            return Ok(());
        };
        if component_matches_pattern(
            literal_component,
            first,
            last,
            family.final_prefix,
            &pattern.suffix,
        ) {
            return Err(ancestor_conflict());
        }
    }
    Ok(())
}

fn check_family_pair(
    first_start: u64,
    first_last: u64,
    first: &IndexedPattern,
    second_start: u64,
    second_last: u64,
    second: &IndexedPattern,
) -> Result<(), Error> {
    if folded_eq(&first.prefix, &second.prefix) && folded_eq(&first.suffix, &second.suffix) {
        if ranges_intersect(first_start, first_last, second_start, second_last) {
            return Err(duplicate_conflict());
        }
    } else if ambiguous_prefix_boundary(&first.prefix, &second.prefix) {
        return Err(unprovable_conflict());
    }

    let first_parts = FamilyParts::new(first);
    let second_parts = FamilyParts::new(second);
    if first_parts.depth < second_parts.depth
        && family_is_ancestor_of_family(first_start, first_last, first, &first_parts, &second_parts)
    {
        return Err(ancestor_conflict());
    }
    if second_parts.depth < first_parts.depth
        && family_is_ancestor_of_family(
            second_start,
            second_last,
            second,
            &second_parts,
            &first_parts,
        )
    {
        return Err(ancestor_conflict());
    }
    Ok(())
}

fn family_is_ancestor_of_family(
    ancestor_start: u64,
    ancestor_last: u64,
    ancestor: &IndexedPattern,
    ancestor_parts: &FamilyParts<'_>,
    descendant_parts: &FamilyParts<'_>,
) -> bool {
    debug_assert!(ancestor_parts.depth < descendant_parts.depth);
    if !component_prefix_matches(
        ancestor_parts.parents,
        descendant_parts.parents,
        ancestor_parts.depth - 1,
    ) {
        return false;
    }
    let Some(descendant_parent) = component_at(descendant_parts.parents, ancestor_parts.depth - 1)
    else {
        return false;
    };
    component_matches_pattern(
        descendant_parent,
        ancestor_start,
        ancestor_last,
        ancestor_parts.final_prefix,
        &ancestor.suffix,
    )
}

fn literal_matches_family(literal: &str, first: u64, last: u64, pattern: &IndexedPattern) -> bool {
    matches_numbered_name_in_range(literal, first, last, pattern)
}

fn matches_exact_numbered_name(name: &str, pattern: &IndexedPattern, value: u64) -> bool {
    matches_numbered_name_range(name, value, value, pattern, false)
}

fn matches_numbered_name_in_range(
    name: &str,
    first: u64,
    last: u64,
    pattern: &IndexedPattern,
) -> bool {
    matches_numbered_name_range(name, first, last, pattern, true)
}

fn matches_numbered_name_range(
    name: &str,
    first: u64,
    last: u64,
    pattern: &IndexedPattern,
    folded: bool,
) -> bool {
    let Some(static_len) = pattern.prefix.len().checked_add(pattern.suffix.len()) else {
        return false;
    };
    if name.len() < static_len {
        return false;
    }
    let suffix_start = name.len() - pattern.suffix.len();
    let middle_end = suffix_start;
    let middle_start = pattern.prefix.len();
    let prefix_matches = if folded {
        folded_bytes_eq(&name.as_bytes()[..middle_start], pattern.prefix.as_bytes())
    } else {
        name.as_bytes()[..middle_start] == pattern.prefix.as_bytes()[..]
    };
    let suffix_matches = if folded {
        folded_bytes_eq(&name.as_bytes()[suffix_start..], pattern.suffix.as_bytes())
    } else {
        name.as_bytes()[suffix_start..] == pattern.suffix.as_bytes()[..]
    };
    if middle_start > middle_end || !prefix_matches || !suffix_matches {
        return false;
    }
    let Some(value) = parse_decimal(&name.as_bytes()[middle_start..middle_end]) else {
        return false;
    };
    (first..=last).contains(&value)
}

fn component_matches_pattern(
    component: &str,
    first: u64,
    last: u64,
    prefix: &str,
    suffix: &str,
) -> bool {
    let Some(static_len) = prefix.len().checked_add(suffix.len()) else {
        return false;
    };
    if component.len() < static_len {
        return false;
    }
    let suffix_start = component.len() - suffix.len();
    let middle_start = prefix.len();
    if middle_start > suffix_start
        || !folded_bytes_eq(&component.as_bytes()[..middle_start], prefix.as_bytes())
        || !folded_bytes_eq(&component.as_bytes()[suffix_start..], suffix.as_bytes())
    {
        return false;
    }
    let Some(value) = parse_decimal(&component.as_bytes()[middle_start..suffix_start]) else {
        return false;
    };
    (first..=last).contains(&value)
}

struct FamilyParts<'a> {
    parents: &'a str,
    final_prefix: &'a str,
    depth: usize,
}

impl<'a> FamilyParts<'a> {
    fn new(pattern: &'a IndexedPattern) -> Self {
        let (parents, final_prefix) = match pattern.prefix.rfind('/') {
            Some(index) => (&pattern.prefix[..index], &pattern.prefix[index + 1..]),
            None => ("", pattern.prefix.as_str()),
        };
        Self {
            parents,
            final_prefix,
            depth: component_count(parents) + 1,
        }
    }
}

fn validate_indexed_pattern(
    prefix: &str,
    suffix: &str,
    first: u64,
    last: u64,
) -> Result<(), Error> {
    if !prefix.is_ascii() || !suffix.is_ascii() {
        return Err(invalid("generated name plan patterns must be ASCII"));
    }
    if prefix.as_bytes().last().is_some_and(u8::is_ascii_digit)
        || suffix.as_bytes().first().is_some_and(u8::is_ascii_digit)
    {
        return Err(invalid(
            "generated name plan numeric boundaries are ambiguous",
        ));
    }
    if suffix.contains('/') || suffix.contains('\\') || suffix.contains(':') {
        return Err(invalid(
            "generated name plan indexed suffix must remain in the final component",
        ));
    }
    let first_name = compose_name(prefix, suffix, first)?;
    validate_literal_name(&first_name)?;
    if first != last {
        let last_name = compose_name(prefix, suffix, last)?;
        validate_literal_name(&last_name)?;
    }
    Ok(())
}

fn validate_literal_name(name: &str) -> Result<(), Error> {
    if name.is_empty() || !name.is_ascii() || name.ends_with('/') {
        return Err(invalid(
            "generated name plan member name is not canonical ASCII",
        ));
    }
    if name.starts_with('/') || name.contains('\\') || name.contains(':') {
        return Err(invalid("generated name plan member name is not canonical"));
    }
    for component in name.split('/') {
        if component.is_empty() || component == "." || component == ".." {
            return Err(invalid("generated name plan member name is not canonical"));
        }
    }
    Ok(())
}

fn compose_name(prefix: &str, suffix: &str, value: u64) -> Result<String, Error> {
    let capacity = prefix
        .len()
        .checked_add(suffix.len())
        .and_then(|bytes| bytes.checked_add(decimal_digits(value)))
        .ok_or_else(|| invalid("generated name plan name bytes overflow usize"))?;
    let mut name = String::new();
    name.try_reserve_exact(capacity)
        .map_err(|source| allocation("generated name plan name", source))?;
    name.push_str(prefix);
    append_decimal(&mut name, value);
    name.push_str(suffix);
    Ok(name)
}

fn append_decimal(output: &mut String, mut value: u64) {
    let mut digits = [0u8; MAX_DECIMAL_DIGITS];
    let mut length = 0usize;
    loop {
        digits[length] = (value % 10) as u8;
        length += 1;
        value /= 10;
        if value == 0 {
            break;
        }
    }
    for digit in digits[..length].iter().rev() {
        output.push(char::from(b'0' + *digit));
    }
}

fn decimal_digits(mut value: u64) -> usize {
    let mut digits = 1;
    while value >= 10 {
        value /= 10;
        digits += 1;
    }
    digits
}

fn parse_decimal(bytes: &[u8]) -> Option<u64> {
    if bytes.is_empty() || (bytes.len() > 1 && bytes[0] == b'0') {
        return None;
    }
    let mut value = 0u64;
    for byte in bytes {
        if !byte.is_ascii_digit() {
            return None;
        }
        value = value
            .checked_mul(10)?
            .checked_add(u64::from(*byte - b'0'))?;
    }
    Some(value)
}

fn component_count(path: &str) -> usize {
    if path.is_empty() {
        0
    } else {
        path.split('/').count()
    }
}

fn component_at(path: &str, index: usize) -> Option<&str> {
    path.split('/').nth(index)
}

fn component_prefix_matches(shorter: &str, longer: &str, count: usize) -> bool {
    shorter
        .split('/')
        .take(count)
        .zip(longer.split('/'))
        .all(|(left, right)| folded_eq(left, right))
        && shorter.split('/').take(count).count() == count
        && longer.split('/').take(count).count() == count
}

fn is_component_prefix(shorter: &str, longer: &str) -> bool {
    let shorter_count = component_count(shorter);
    let longer_count = component_count(longer);
    shorter_count < longer_count && component_prefix_matches(shorter, longer, shorter_count)
}

fn ranges_intersect(
    first_start: u64,
    first_last: u64,
    second_start: u64,
    second_last: u64,
) -> bool {
    first_start <= second_last && second_start <= first_last
}

fn ambiguous_prefix_boundary(first: &str, second: &str) -> bool {
    let first_folded = first.as_bytes();
    let second_folded = second.as_bytes();
    let common = first_folded.len().min(second_folded.len());
    if !folded_bytes_eq(&first_folded[..common], &second_folded[..common]) {
        return false;
    }
    if first_folded.len() == second_folded.len() {
        return false;
    }
    let longer = if first_folded.len() > common {
        &first_folded[common..]
    } else {
        &second_folded[common..]
    };
    longer.first().is_some_and(u8::is_ascii_digit)
}

fn folded_eq(left: &str, right: &str) -> bool {
    folded_bytes_eq(left.as_bytes(), right.as_bytes())
}

fn folded_bytes_eq(left: &[u8], right: &[u8]) -> bool {
    left.len() == right.len()
        && left
            .iter()
            .zip(right)
            .all(|(left, right)| left.eq_ignore_ascii_case(right))
}

fn clone_string(value: &str, resource: &'static str) -> Result<String, Error> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| allocation(resource, source))?;
    owned.push_str(value);
    Ok(owned)
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    ErrorKind::Allocation { resource, source }.into()
}

fn invalid(message: &'static str) -> Error {
    ErrorKind::InvalidInput {
        msg: message.to_string(),
    }
    .into()
}

fn duplicate_conflict() -> Error {
    invalid("generated name plan contains duplicate or ASCII-case-equivalent names")
}

fn ancestor_conflict() -> Error {
    invalid("generated name plan contains a whole-component ancestor conflict")
}

fn unprovable_conflict() -> Error {
    invalid("generated name plan contains an ambiguous indexed-name boundary")
}

#[cfg(test)]
mod tests {
    use super::{GeneratedNamePlanBuilder, GeneratedNamePlanLimits};
    use crate::ErrorKind;

    fn limits(patterns: usize, bytes: usize, entries: u64) -> GeneratedNamePlanLimits {
        GeneratedNamePlanLimits::new(patterns, bytes, entries)
    }

    #[test]
    fn literal_and_indexed_cursor_are_ordered_without_expansion() {
        let mut builder = GeneratedNamePlanBuilder::new(limits(4, 128, 7)).unwrap();
        builder.push_literal("[Content_Types].xml").unwrap();
        builder
            .push_indexed(
                1,
                3,
                &[
                    ("ppt/slides/slide", ".xml"),
                    ("ppt/slides/slide", ".xml.rels"),
                ],
            )
            .unwrap();
        let mut plan = builder.finish().unwrap();
        assert_eq!(plan.entry_count(), 7);
        assert_eq!(plan.max_name_bytes(), 26);
        assert!(!plan.is_complete());
        for name in [
            "[Content_Types].xml",
            "ppt/slides/slide1.xml",
            "ppt/slides/slide1.xml.rels",
            "ppt/slides/slide2.xml",
            "ppt/slides/slide2.xml.rels",
            "ppt/slides/slide3.xml",
            "ppt/slides/slide3.xml.rels",
        ] {
            plan.check_next(name).unwrap();
            plan.advance().unwrap();
        }
        assert!(plan.is_complete());
        assert!(plan.check_next("anything").is_err());
    }

    #[test]
    fn exact_and_ascii_folded_literals_are_rejected() {
        let mut builder = GeneratedNamePlanBuilder::new(limits(4, 32, 4)).unwrap();
        builder.push_literal("ppt/slides/a.xml").unwrap();
        let error = builder.push_literal("PPT/slides/A.XML").unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
    }

    #[test]
    fn runtime_cursor_requires_exact_indexed_prefix_and_suffix_bytes() {
        let mut builder = GeneratedNamePlanBuilder::new(limits(2, 64, 2)).unwrap();
        builder.push_indexed(1, 1, &[("a/slide", ".XML")]).unwrap();
        let mut plan = builder.finish().unwrap();
        assert!(plan.check_next("a/SLIDE1.XML").is_err());
        assert!(plan.check_next("a/slide1.xml").is_err());
        plan.check_next("a/slide1.XML").unwrap();
        plan.advance().unwrap();
        assert!(plan.is_complete());
    }

    #[test]
    fn overlapping_and_ambiguous_families_are_rejected() {
        let mut builder = GeneratedNamePlanBuilder::new(limits(8, 128, 100)).unwrap();
        builder
            .push_indexed(1, 3, &[("ppt/slides/slide", ".xml")])
            .unwrap();
        assert!(
            builder
                .push_indexed(3, 2, &[("ppt/slides/SLIDE", ".XML")])
                .is_err()
        );

        let mut ambiguous = GeneratedNamePlanBuilder::new(limits(8, 128, 100)).unwrap();
        ambiguous.push_indexed(1, 2, &[("a", "x2")]).unwrap();
        assert!(ambiguous.push_indexed(1, 2, &[("a1x", "")]).is_err());
    }

    #[test]
    fn literal_family_and_family_family_ancestor_conflicts_are_proved() {
        let mut literal_ancestor = GeneratedNamePlanBuilder::new(limits(8, 128, 20)).unwrap();
        literal_ancestor.push_literal("ppt/slides").unwrap();
        assert!(
            literal_ancestor
                .push_indexed(1, 2, &[("ppt/slides/slide", ".xml")])
                .is_err()
        );

        let mut family_ancestor = GeneratedNamePlanBuilder::new(limits(8, 128, 20)).unwrap();
        family_ancestor
            .push_indexed(1, 2, &[("ppt/slides/slide", "")])
            .unwrap();
        assert!(
            family_ancestor
                .push_indexed(1, 2, &[("ppt/slides/slide1/", ".xml")])
                .is_err()
        );

        let mut fixed_parent = GeneratedNamePlanBuilder::new(limits(8, 128, 20)).unwrap();
        fixed_parent.push_indexed(1, 2, &[("a/", "")]).unwrap();
        assert!(
            fixed_parent
                .push_indexed(1, 2, &[("a/1/", ".xml")])
                .is_err()
        );
    }

    #[test]
    fn unequal_depth_literal_family_is_checked_both_directions() {
        let mut family_ancestor = GeneratedNamePlanBuilder::new(limits(8, 128, 20)).unwrap();
        family_ancestor
            .push_indexed(1, 2, &[("a/slide", "")])
            .unwrap();
        assert!(family_ancestor.push_literal("a/slide1/x").is_err());

        let mut literal_ancestor = GeneratedNamePlanBuilder::new(limits(8, 128, 20)).unwrap();
        literal_ancestor.push_literal("a/fixed").unwrap();
        assert!(
            literal_ancestor
                .push_indexed(1, 2, &[("a/fixed/", ".xml")])
                .is_err()
        );
    }

    #[test]
    fn invalid_syntax_boundaries_overflow_and_budgets_are_refused() {
        let mut invalid = GeneratedNamePlanBuilder::new(limits(8, 128, 20)).unwrap();
        assert!(invalid.push_literal("a//b").is_err());
        assert!(invalid.push_literal("a/../b").is_err());
        assert!(invalid.push_indexed(1, 1, &[("a1", ".xml")]).is_err());
        assert!(invalid.push_indexed(1, 1, &[("a", "2.xml")]).is_err());
        assert!(invalid.push_indexed(u64::MAX, 2, &[("a", ".xml")]).is_err());

        let mut patterns = GeneratedNamePlanBuilder::new(limits(1, 128, 20)).unwrap();
        patterns.push_literal("a").unwrap();
        assert!(patterns.push_literal("b").is_err());

        let mut bytes = GeneratedNamePlanBuilder::new(limits(4, 2, 20)).unwrap();
        assert!(bytes.push_literal("abc").is_err());

        let mut entries = GeneratedNamePlanBuilder::new(limits(4, 128, 2)).unwrap();
        assert!(
            entries
                .push_indexed(1, 2, &[("a", ".xml"), ("a", ".rels")])
                .is_err()
        );
    }

    #[test]
    fn ranges_are_checked_symbolically_at_u64_scale() {
        let mut builder = GeneratedNamePlanBuilder::new(limits(2, 32, u64::MAX)).unwrap();
        builder
            .push_indexed(u64::MAX - 1, 2, &[("x", ".bin")])
            .unwrap();
        let plan = builder.finish().unwrap();
        assert_eq!(plan.entry_count(), 2);
        assert_eq!(plan.max_name_bytes(), 25);
    }

    #[test]
    fn incomplete_and_repeated_cursor_operations_are_refused() {
        let mut builder = GeneratedNamePlanBuilder::new(limits(2, 32, 2)).unwrap();
        builder.push_literal("a").unwrap();
        let mut plan = builder.finish().unwrap();
        assert!(plan.check_next("b").is_err());
        plan.check_next("a").unwrap();
        plan.advance().unwrap();
        assert!(plan.is_complete());
        assert!(plan.advance().is_err());
    }

    #[test]
    fn cursor_initializes_each_following_indexed_descriptor() {
        let mut builder = GeneratedNamePlanBuilder::new(limits(3, 64, 5)).unwrap();
        builder.push_indexed(7, 2, &[("a/x", "")]).unwrap();
        builder.push_indexed(42, 2, &[("b/y", "")]).unwrap();
        builder.push_literal("c").unwrap();
        let mut plan = builder.finish().unwrap();
        for name in ["a/x7", "a/x8", "b/y42", "b/y43", "c"] {
            plan.check_next(name).unwrap();
            plan.advance().unwrap();
        }
        assert!(plan.is_complete());
    }

    #[test]
    fn root_level_families_use_one_component_depth() {
        let mut family_and_literal = GeneratedNamePlanBuilder::new(limits(4, 64, 8)).unwrap();
        family_and_literal.push_indexed(1, 2, &[("", "")]).unwrap();
        assert!(family_and_literal.push_literal("1/x").is_err());

        let mut family_and_family = GeneratedNamePlanBuilder::new(limits(4, 64, 8)).unwrap();
        family_and_family.push_indexed(1, 2, &[("", "")]).unwrap();
        assert!(family_and_family.push_indexed(1, 2, &[("1/", "")]).is_err());
    }
}
