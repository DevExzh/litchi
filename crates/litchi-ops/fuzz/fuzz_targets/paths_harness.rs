//! Bounded Excel-notation and path normalization round trips.
use litchi_ops::path::{Path, PathArg};
pub fn exercise(data: &[u8]) -> bool {
    if data.len() > 32 * 1024 {
        return false;
    }
    let Ok(text) = std::str::from_utf8(data) else {
        return false;
    };
    let Ok(path) = PathArg::parse(text) else {
        return false;
    };
    let canonical = path.to_string();
    let reparsed = if path.is_absolute() {
        Path::parse(&canonical)
    } else {
        Path::parse_relative(&canonical)
    };
    assert_eq!(reparsed.expect("canonical path"), *path.path());
    true
}
