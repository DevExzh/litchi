//! "Did you mean" suggestions. These run on the error path only, so they may
//! allocate.

/// Optimal-string-alignment distance (restricted Damerau-Levenshtein) over
/// lower-cased characters: insertion, deletion, substitution and a swap of
/// two adjacent characters each cost 1. This is OfficeCLI's
/// `EditDistance.Damerau`, applied case-insensitively as every OfficeCLI
/// suggester applies it.
pub(super) fn distance(a: &str, b: &str) -> usize {
    let a: Vec<char> = a.chars().flat_map(char::to_lowercase).collect();
    let b: Vec<char> = b.chars().flat_map(char::to_lowercase).collect();
    if a.is_empty() {
        return b.len();
    }
    if b.is_empty() {
        return a.len();
    }
    let width = b.len() + 1;
    let mut before: Vec<usize> = vec![0; width];
    let mut previous: Vec<usize> = (0..width).collect();
    let mut current: Vec<usize> = vec![0; width];
    for i in 1..=a.len() {
        current[0] = i;
        for j in 1..=b.len() {
            let cost = usize::from(a[i - 1] != b[j - 1]);
            let mut best = (previous[j] + 1)
                .min(current[j - 1] + 1)
                .min(previous[j - 1] + cost);
            if i > 1 && j > 1 && a[i - 1] == b[j - 2] && a[i - 2] == b[j - 1] {
                best = best.min(before[j - 2] + 1);
            }
            current[j] = best;
        }
        std::mem::swap(&mut before, &mut previous);
        std::mem::swap(&mut previous, &mut current);
    }
    previous[b.len()]
}

/// The candidate closest to `input` within OfficeCLI's threshold of
/// `max(2, len / 3)`, or `None` when nothing is that close, when the best
/// distance is shared by two candidates, or when `input` already equals a
/// candidate (OfficeCLI `AttributeFilter.NearestMatch`).
pub(crate) fn nearest<'c>(
    input: &str,
    candidates: impl IntoIterator<Item = &'c str>,
) -> Option<&'c str> {
    nearest_keyed(
        input,
        candidates.into_iter().map(|spelling| (spelling, spelling)),
    )
}

/// [`nearest`] over `(spelling, key)` pairs, returning the key. Spellings
/// that share a key (aliases of one value) are not a tie.
pub(crate) fn nearest_keyed<'c>(
    input: &str,
    candidates: impl IntoIterator<Item = (&'c str, &'c str)>,
) -> Option<&'c str> {
    let input = input.trim();
    if input.is_empty() {
        return None;
    }
    let length = input.chars().count();
    let threshold = (length / 3).max(2);
    let mut best: Option<(&'c str, usize)> = None;
    let mut tied = false;
    for (spelling, key) in candidates {
        // The distance is at least the difference in length.
        if spelling.chars().count().abs_diff(length) > threshold {
            continue;
        }
        let d = distance(input, spelling);
        if d == 0 || d > threshold {
            if d == 0 {
                // The input is a known spelling: nothing to suggest.
                return None;
            }
            continue;
        }
        match best {
            Some((_, best_distance)) if d > best_distance => {},
            Some((best_key, best_distance)) if d == best_distance => {
                if best_key != key {
                    tied = true;
                }
            },
            _ => {
                best = Some((key, d));
                tied = false;
            },
        }
    }
    if tied { None } else { best.map(|(key, _)| key) }
}
