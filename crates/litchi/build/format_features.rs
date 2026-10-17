//! The facade's format features, derived from its `Cargo.toml`.
//!
//! A format feature is a `[features]` entry that turns on a `litchi-ops`
//! format with a plain (not weak, `?`) `"litchi-ops/<format>"` item. Every
//! such feature makes `litchi::open` and `File` available, so the build
//! script emits `cfg(litchi_unified)` when any one is enabled, and no
//! hand-written `#[cfg(any(feature = ...))]` list can omit a format (record
//! 0334). The facade's unit tests include this file to pin the derived set.
//!
//! The parser reads only the `[features]` table: a key, `=`, and an array of
//! double-quoted strings, which may span lines. Comments and other tables are
//! skipped. It is deliberately small, because a build script cannot use the
//! workspace's TOML dependencies without adding them to every build.

/// The names of the format features declared in `manifest`, in file order.
pub fn format_features(manifest: &str) -> Vec<String> {
    let mut features = Vec::new();
    let mut in_features = false;
    let mut current: Option<(String, bool)> = None;
    for raw_line in manifest.lines() {
        let line = strip_comment(raw_line).trim();
        if current.is_none() && line.starts_with('[') {
            in_features = line == "[features]";
            continue;
        }
        if !in_features || line.is_empty() {
            continue;
        }
        let rest = match current.as_mut() {
            Some(_) => line,
            None => {
                let Some((key, value)) = line.split_once('=') else {
                    continue;
                };
                current = Some((key.trim().trim_matches('"').to_owned(), false));
                value.trim()
            },
        };
        let Some((_, enables_ops)) = current.as_mut() else {
            continue;
        };
        for item in quoted_items(rest) {
            if item
                .strip_prefix("litchi-ops/")
                .is_some_and(|f| !f.is_empty())
            {
                *enables_ops = true;
            }
        }
        if rest.contains(']')
            && let Some((name, true)) = current.take()
        {
            features.push(name);
        }
    }
    features
}

/// `line` up to an unquoted `#`.
fn strip_comment(line: &str) -> &str {
    let mut quoted = false;
    for (index, byte) in line.bytes().enumerate() {
        match byte {
            b'"' => quoted = !quoted,
            b'#' if !quoted => return &line[..index],
            _ => {},
        }
    }
    line
}

/// The double-quoted strings of `text`, without escapes (feature items have
/// none).
fn quoted_items(text: &str) -> impl Iterator<Item = &str> {
    text.split('"').skip(1).step_by(2)
}
