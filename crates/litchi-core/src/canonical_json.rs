//! Canonical JSON bytes that do not depend on `serde_json` features.
//!
//! `serde_json::Map` iterates in insertion order when any crate in the build
//! enables the `preserve_order` feature, and in sorted key order otherwise.
//! Cargo unifies features across a workspace, so `serde_json::to_vec` of the
//! same [`serde_json::Value`] can produce different bytes depending on which
//! other crates are linked into the binary (record 0349). Durable envelopes,
//! digests and canonicality checks must instead use this writer, which always
//! emits:
//!
//! - object members sorted by the byte order of their keys;
//! - no insignificant whitespace;
//! - strings with `serde_json`'s escaping (which no feature changes);
//! - integers in plain decimal and other numbers in the shortest
//!   round-tripping `f64` form, whatever `arbitrary_precision` would keep.

use std::io;

use serde_json::{Number, Value};

/// Writes `value` as canonical JSON to `writer`.
///
/// # Errors
///
/// Returns the first error reported by `writer`.
pub fn to_writer<W: io::Write + ?Sized>(writer: &mut W, value: &Value) -> io::Result<()> {
    match value {
        Value::Null => writer.write_all(b"null"),
        Value::Bool(true) => writer.write_all(b"true"),
        Value::Bool(false) => writer.write_all(b"false"),
        Value::Number(number) => write_number(writer, number),
        Value::String(text) => write_string(writer, text),
        Value::Array(values) => {
            writer.write_all(b"[")?;
            for (index, child) in values.iter().enumerate() {
                if index != 0 {
                    writer.write_all(b",")?;
                }
                to_writer(writer, child)?;
            }
            writer.write_all(b"]")
        },
        Value::Object(members) => {
            let mut sorted = members.iter().collect::<Vec<_>>();
            if !sorted.is_sorted_by(|left, right| left.0 <= right.0) {
                sorted.sort_unstable_by(|left, right| left.0.cmp(right.0));
            }
            writer.write_all(b"{")?;
            for (index, (key, child)) in sorted.into_iter().enumerate() {
                if index != 0 {
                    writer.write_all(b",")?;
                }
                write_string(writer, key)?;
                writer.write_all(b":")?;
                to_writer(writer, child)?;
            }
            writer.write_all(b"}")
        },
    }
}

/// Returns `value` as canonical JSON bytes.
#[must_use]
pub fn to_vec(value: &Value) -> Vec<u8> {
    let mut bytes = Vec::new();
    // Writing to a `Vec` cannot fail.
    let _ = to_writer(&mut bytes, value);
    bytes
}

fn write_number<W: io::Write + ?Sized>(writer: &mut W, number: &Number) -> io::Result<()> {
    if let Some(value) = number.as_u64() {
        write!(writer, "{value}")
    } else if let Some(value) = number.as_i64() {
        write!(writer, "{value}")
    } else {
        // A `Number` always holds a finite value, so this writes the shortest
        // round-tripping decimal form rather than `null`.
        let value = number.as_f64().unwrap_or(f64::NAN);
        serde_json::to_writer(writer, &value).map_err(io::Error::from)
    }
}

fn write_string<W: io::Write + ?Sized>(writer: &mut W, text: &str) -> io::Result<()> {
    serde_json::to_writer(writer, text).map_err(io::Error::from)
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        reason = "test assertions unwrap expected success"
    )]

    use serde_json::{Map, json};

    use super::*;

    #[test]
    fn object_members_are_sorted_whatever_the_insertion_order() {
        // Under `preserve_order` this map iterates z, a, m; without it, a, m,
        // z. Either way the canonical bytes are identical.
        let mut map = Map::new();
        map.insert("z".to_owned(), json!(1));
        map.insert("a".to_owned(), json!({"y": [true, null], "b": "\u{1}\""}));
        map.insert("m".to_owned(), json!(-2));
        assert_eq!(
            to_vec(&Value::Object(map)),
            br#"{"a":{"b":"\u0001\"","y":[true,null]},"m":-2,"z":1}"#
        );
    }

    #[test]
    fn keys_sort_by_bytes_and_numbers_are_canonical() {
        let value = json!({"\u{e9}": 0, "B": 1.5, "a": u64::MAX, "_": i64::MIN, "f": 1e300});
        assert_eq!(
            String::from_utf8(to_vec(&value)).unwrap(),
            format!(
                "{{\"B\":1.5,\"_\":{},\"a\":{},\"f\":1e+300,\"\u{e9}\":0}}",
                i64::MIN,
                u64::MAX
            )
        );
    }

    #[test]
    fn parsed_noncanonical_input_reencodes_canonically() {
        let parsed: Value = serde_json::from_str(r#"{ "b" : 1.0, "a" : [ 2 ] }"#).unwrap();
        assert_eq!(to_vec(&parsed), br#"{"a":[2],"b":1.0}"#);
    }
}
