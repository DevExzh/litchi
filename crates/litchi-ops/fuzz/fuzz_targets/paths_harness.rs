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
    // The canonical spelling is a fixed point of `PathArg::parse` (record
    // 0338 F2), and the core grammar reads it as the same path.
    let canonical = path.to_string();
    let again = PathArg::parse(&canonical).expect("canonical path");
    assert_eq!(again.path(), path.path());
    assert_eq!(again.to_string(), canonical);
    let core = if path.is_absolute() {
        Path::parse(&canonical)
    } else {
        Path::parse_relative(&canonical)
    };
    assert_eq!(core.expect("canonical path"), *path.path());
    true
}
