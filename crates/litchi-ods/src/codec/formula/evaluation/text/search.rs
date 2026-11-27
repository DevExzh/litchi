//! Bounded Unicode case-folded SEARCH matching.
//!
//! SEARCH folds each source scalar with the caller-provided Unicode 17
//! default C+F table.  The matcher owns only the folded pattern, its KMP
//! failure table, and a ring of source positions for the current match.  The
//! haystack remains borrowed and is traversed once.  No normalization is
//! performed here.

use super::super::{EvaluationFailure, EvaluationResult, Evaluator};
use litchi_core::Resource;
use std::{collections::TryReserveError, mem::size_of};

use super::unicode;

/// The Unicode CaseFolding full mappings in the pinned data have at most
/// three scalars.  The Unicode owner writes one source scalar's mapping into
/// this fixed buffer and returns the number of initialized entries.
pub(super) const MAX_FOLD_SCALARS: usize = 3;

struct SearchPattern {
    folded: Vec<char>,
    failure: Vec<usize>,
}

#[derive(Clone, Copy, Default)]
struct SourceOrigin {
    source_position: usize,
    source_start: bool,
}

/// Return the first source-scalar position of a full Unicode C+F-folded match.
///
/// `start` and the returned position are one-based scalar positions.  A
/// folded match must begin and end at original source-scalar boundaries.  In
/// particular, a single `s` does not match the first half of `ß`, while `ss`
/// does match the complete expansion at that position.  Empty patterns match
/// at every valid `start` from 1 through `LEN(haystack) + 1`.
pub(super) fn find(
    evaluator: &mut Evaluator<'_, '_, '_>,
    needle: &str,
    haystack: &str,
    start: usize,
) -> EvaluationResult<Option<usize>> {
    find_with(evaluator, needle, haystack, start, |character, output| {
        let mapping = unicode::case_fold(character);
        let mut length = 0;
        for (index, folded) in mapping.iter().enumerate() {
            output[index] = folded;
            length = index + 1;
        }
        length
    })
}

fn find_with<F>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    needle: &str,
    haystack: &str,
    start: usize,
    mut fold: F,
) -> EvaluationResult<Option<usize>>
where
    F: FnMut(char, &mut [char; MAX_FOLD_SCALARS]) -> usize,
{
    let folded_len = count_folded_scalars(evaluator, needle, &mut fold)?;
    if folded_len == 0 {
        let mut charge = |amount| evaluator.charge_work(amount);
        return valid_empty_position(haystack, start, &mut charge);
    }

    let scratch_bytes = folded_len
        .checked_mul(
            size_of::<char>()
                .checked_add(size_of::<usize>())
                .and_then(|value| value.checked_add(size_of::<SourceOrigin>()))
                .ok_or_else(|| memory_limit(evaluator, usize::MAX))?,
        )
        .ok_or_else(|| memory_limit(evaluator, usize::MAX))?;
    let _scratch = evaluator.reserve_storage(scratch_bytes, "formula SEARCH matcher")?;
    let pattern = build_pattern(evaluator, needle, folded_len, &mut fold)?;
    let mut charge = |amount| evaluator.charge_work(amount);
    scan(&pattern, haystack, start, &mut fold, &mut charge)
}

fn count_folded_scalars<F>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    text: &str,
    fold: &mut F,
) -> EvaluationResult<usize>
where
    F: FnMut(char, &mut [char; MAX_FOLD_SCALARS]) -> usize,
{
    let mut output = ['\0'; MAX_FOLD_SCALARS];
    let mut count = 0usize;
    for character in text.chars() {
        evaluator.charge_work(1)?;
        let length = fold(character, &mut output);
        validate_fold_length(length)?;
        count = count
            .checked_add(length)
            .ok_or_else(|| memory_limit(evaluator, usize::MAX))?;
    }
    Ok(count)
}

fn build_pattern<F>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    needle: &str,
    folded_len: usize,
    fold: &mut F,
) -> EvaluationResult<SearchPattern>
where
    F: FnMut(char, &mut [char; MAX_FOLD_SCALARS]) -> usize,
{
    let mut folded = Vec::new();
    folded.try_reserve_exact(folded_len).map_err(allocation)?;
    let mut output = ['\0'; MAX_FOLD_SCALARS];
    for character in needle.chars() {
        evaluator.charge_work(1)?;
        let length = fold(character, &mut output);
        validate_fold_length(length)?;
        let next_length = folded
            .len()
            .checked_add(length)
            .ok_or_else(|| memory_limit(evaluator, usize::MAX))?;
        if next_length > folded_len {
            return Err(EvaluationFailure::InvalidExpression(
                "Unicode case-fold mapping changed during SEARCH pattern construction",
            ));
        }
        folded.extend_from_slice(&output[..length]);
    }
    if folded.len() != folded_len {
        return Err(EvaluationFailure::InvalidExpression(
            "Unicode case-fold mapping changed during SEARCH pattern construction",
        ));
    }

    let mut failure = Vec::new();
    failure.try_reserve_exact(folded_len).map_err(allocation)?;
    failure.resize(folded_len, 0);
    for index in 1..folded_len {
        let value = folded[index];
        let mut prefix = failure[index - 1];
        while prefix > 0 && folded[prefix] != value {
            evaluator.charge_work(1)?;
            prefix = failure[prefix - 1];
        }
        evaluator.charge_work(1)?;
        if folded[prefix] == value {
            prefix += 1;
        }
        failure[index] = prefix;
    }
    Ok(SearchPattern { folded, failure })
}

fn scan<F, C>(
    pattern: &SearchPattern,
    haystack: &str,
    start: usize,
    fold: &mut F,
    charge: &mut C,
) -> EvaluationResult<Option<usize>>
where
    F: FnMut(char, &mut [char; MAX_FOLD_SCALARS]) -> usize,
    C: FnMut(u64) -> EvaluationResult<()>,
{
    let mut origins = Vec::new();
    origins
        .try_reserve_exact(pattern.folded.len())
        .map_err(allocation)?;
    origins.resize(pattern.folded.len(), SourceOrigin::default());

    let mut iterator = haystack.chars();
    let start_index = start.saturating_sub(1);
    for _ in 0..start_index {
        if iterator.next().is_none() {
            return Ok(None);
        }
        charge(1)?;
    }

    let mut source_position = start;
    let mut stream_index = 0usize;
    let mut matched = 0usize;
    let mut output = ['\0'; MAX_FOLD_SCALARS];
    for character in iterator {
        let length = fold(character, &mut output);
        validate_fold_length(length)?;
        for (offset, &value) in output[..length].iter().enumerate() {
            charge(1)?;
            let next_stream_index =
                stream_index
                    .checked_add(1)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "SEARCH stream index overflow",
                    ))?;
            let origin_index = stream_index % pattern.folded.len();
            origins[origin_index] = SourceOrigin {
                source_position,
                source_start: offset == 0,
            };

            while matched > 0 && pattern.folded[matched] != value {
                charge(1)?;
                matched = pattern.failure[matched - 1];
            }
            if pattern.folded[matched] == value {
                matched += 1;
            }
            if matched == pattern.folded.len() {
                let first_index = next_stream_index - pattern.folded.len();
                let origin = origins[first_index % pattern.folded.len()];
                let ends_at_source_boundary = offset + 1 == length;
                matched = pattern.failure[matched - 1];
                if ends_at_source_boundary && origin.source_start {
                    return Ok(Some(origin.source_position));
                }
            }
            stream_index = next_stream_index;
        }
        source_position =
            source_position
                .checked_add(1)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "SEARCH source position overflow",
                ))?;
    }
    Ok(None)
}

fn valid_empty_position<C>(
    haystack: &str,
    start: usize,
    charge: &mut C,
) -> EvaluationResult<Option<usize>>
where
    C: FnMut(u64) -> EvaluationResult<()>,
{
    let start_index = start.saturating_sub(1);
    let mut iterator = haystack.chars();
    for _ in 0..start_index {
        if iterator.next().is_none() {
            return Ok(None);
        }
        charge(1)?;
    }
    Ok(Some(start))
}

fn validate_fold_length(length: usize) -> EvaluationResult<()> {
    if length <= MAX_FOLD_SCALARS {
        Ok(())
    } else {
        Err(EvaluationFailure::InvalidExpression(
            "Unicode case-fold mapping exceeded the fixed expansion bound",
        ))
    }
}

fn allocation(source: TryReserveError) -> EvaluationFailure {
    EvaluationFailure::Allocation {
        resource: "formula SEARCH matcher",
        source,
    }
}

fn memory_limit(evaluator: &Evaluator<'_, '_, '_>, observed: usize) -> EvaluationFailure {
    super::super::local_limit(
        Resource::Memory,
        u64::try_from(observed).unwrap_or(u64::MAX),
        u64::try_from(evaluator.limits.max_storage_bytes).unwrap_or(u64::MAX),
    )
}

#[cfg(test)]
mod tests {
    use super::{MAX_FOLD_SCALARS, SearchPattern, scan};
    use crate::codec::formula::evaluation::{EvaluationFailure, EvaluationResult};

    fn ascii_fold(character: char, output: &mut [char; MAX_FOLD_SCALARS]) -> usize {
        output[0] = character.to_ascii_lowercase();
        1
    }

    fn expansion_fold(character: char, output: &mut [char; MAX_FOLD_SCALARS]) -> usize {
        match character {
            'ß' => {
                output[..2].copy_from_slice(&['s', 's']);
                2
            },
            'İ' => {
                output[..2].copy_from_slice(&['i', '\u{307}']);
                2
            },
            '\u{fb03}' => {
                output[..3].copy_from_slice(&['f', 'f', 'i']);
                3
            },
            _ => ascii_fold(character, output),
        }
    }

    fn pattern(
        needle: &str,
        fold: impl Fn(char, &mut [char; MAX_FOLD_SCALARS]) -> usize,
    ) -> SearchPattern {
        let mut folded = Vec::new();
        let mut output = ['\0'; MAX_FOLD_SCALARS];
        for character in needle.chars() {
            let length = fold(character, &mut output);
            folded.extend_from_slice(&output[..length]);
        }
        let mut failure = vec![0; folded.len()];
        for index in 1..folded.len() {
            let mut prefix = failure[index - 1];
            while prefix > 0 && folded[prefix] != folded[index] {
                prefix = failure[prefix - 1];
            }
            if folded[prefix] == folded[index] {
                prefix += 1;
            }
            failure[index] = prefix;
        }
        SearchPattern { folded, failure }
    }

    fn run(
        pattern: &SearchPattern,
        haystack: &str,
        start: usize,
        mut fold: impl FnMut(char, &mut [char; MAX_FOLD_SCALARS]) -> usize,
    ) -> EvaluationResult<Option<usize>> {
        let mut charge = |_amount: u64| Ok::<(), EvaluationFailure>(());
        if pattern.folded.is_empty() {
            return super::valid_empty_position(haystack, start, &mut charge);
        }
        scan(pattern, haystack, start, &mut fold, &mut charge)
    }

    #[test]
    fn finds_folded_pattern_at_scalar_position() {
        let pattern = pattern("NEEDLE", ascii_fold);
        assert_eq!(
            run(&pattern, "find needle here", 1, ascii_fold).unwrap(),
            Some(6)
        );
        assert_eq!(
            run(&pattern, "find needle here", 7, ascii_fold).unwrap(),
            None
        );
        assert_eq!(
            run(&pattern, "find needle here", 17, ascii_fold).unwrap(),
            None
        );
    }

    #[test]
    fn expansion_matches_only_at_original_boundaries() {
        let single = pattern("s", expansion_fold);
        let double = pattern("ss", expansion_fold);
        let sharp_s = pattern("ß", expansion_fold);
        let ligature = pattern("ffi", expansion_fold);
        assert_eq!(run(&single, "ß", 1, expansion_fold).unwrap(), None);
        assert_eq!(run(&double, "ß", 1, expansion_fold).unwrap(), Some(1));
        assert_eq!(run(&single, "ßs", 1, expansion_fold).unwrap(), Some(2));
        assert_eq!(run(&double, "ss", 1, expansion_fold).unwrap(), Some(1));
        assert_eq!(run(&sharp_s, "ss", 1, expansion_fold).unwrap(), Some(1));
        assert_eq!(
            run(&ligature, "\u{fb03}", 1, expansion_fold).unwrap(),
            Some(1)
        );
    }

    fn corpus(alphabet: &[char], max_length: usize) -> Vec<String> {
        fn extend(
            prefix: &mut String,
            alphabet: &[char],
            remaining: usize,
            output: &mut Vec<String>,
        ) {
            output.push(prefix.clone());
            if remaining == 0 {
                return;
            }
            for &character in alphabet {
                prefix.push(character);
                extend(prefix, alphabet, remaining - 1, output);
                prefix.pop();
            }
        }

        let mut values = Vec::new();
        extend(&mut String::new(), alphabet, max_length, &mut values);
        values
    }

    fn reference_fold(text: &str) -> Vec<char> {
        let mut result = Vec::new();
        for character in text.chars() {
            match character {
                's' | 'S' | 'x' => result.push(character.to_ascii_lowercase()),
                'ß' => result.extend(['s', 's']),
                _ => unreachable!("test corpus character"),
            }
        }
        result
    }

    fn reference_search(needle: &str, haystack: &str, start: usize) -> Option<usize> {
        let haystack: Vec<char> = haystack.chars().collect();
        let needle = reference_fold(needle);
        let start_index = start.saturating_sub(1);
        if start_index > haystack.len() {
            return None;
        }
        if needle.is_empty() {
            return Some(start);
        }
        for source_start in start_index..haystack.len() {
            let mut folded = Vec::new();
            for &character in &haystack[source_start..] {
                folded.extend(reference_fold(&character.to_string()));
                if folded == needle {
                    return Some(source_start + 1);
                }
                if folded.len() >= needle.len() {
                    break;
                }
            }
        }
        None
    }

    #[test]
    fn exhaustive_boundary_matches_agree_with_reference() {
        let values = corpus(&['s', 'S', 'ß', 'x'], 3);
        for needle in &values {
            let pattern = pattern(needle, expansion_fold);
            for haystack in &values {
                let length = haystack.chars().count();
                for start in 1..=length + 2 {
                    let actual = run(&pattern, haystack, start, expansion_fold).unwrap();
                    let expected = reference_search(needle, haystack, start);
                    assert_eq!(
                        actual, expected,
                        "needle={needle:?} haystack={haystack:?} start={start}"
                    );
                }
            }
        }
    }

    #[test]
    fn case_fold_expansion_remains_unnormalized() {
        let first = pattern("i\u{307}", expansion_fold);
        assert_eq!(run(&first, "İ", 1, expansion_fold).unwrap(), Some(1));
        let canonically_different = pattern("i\u{301}", expansion_fold);
        assert_eq!(
            run(&canonically_different, "İ", 1, expansion_fold).unwrap(),
            None
        );
    }

    #[test]
    fn empty_pattern_accepts_start_through_end() {
        let empty = pattern("", ascii_fold);
        assert_eq!(run(&empty, "abc", 1, ascii_fold).unwrap(), Some(1));
        assert_eq!(run(&empty, "abc", 4, ascii_fold).unwrap(), Some(4));
        assert_eq!(run(&empty, "abc", 5, ascii_fold).unwrap(), None);
    }

    #[test]
    fn repeated_prefix_miss_stays_linear() {
        let pattern = pattern("aaaaab", ascii_fold);
        let haystack = "a".repeat(20_000);
        let mut charges = 0u64;
        let mut charge = |amount: u64| {
            charges += amount;
            Ok::<(), EvaluationFailure>(())
        };
        let result = scan(&pattern, &haystack, 1, &mut ascii_fold, &mut charge);
        assert_eq!(result.unwrap(), None);
        assert!(charges < 100_000, "KMP work unexpectedly grew: {charges}");
    }

    #[test]
    fn scan_propagates_work_failure() {
        let pattern = pattern("needle", ascii_fold);
        let mut charges = 0u64;
        let mut charge = |amount: u64| {
            charges += amount;
            if charges > 3 {
                Err(EvaluationFailure::InvalidExpression("test work limit"))
            } else {
                Ok(())
            }
        };
        let result = scan(&pattern, "a long haystack", 1, &mut ascii_fold, &mut charge);
        assert!(matches!(
            result,
            Err(EvaluationFailure::InvalidExpression("test work limit"))
        ));
    }
}
