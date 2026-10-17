//! OfficeCLI batch-payload ingress (record 0334): bounded JSON decoding into
//! typed [`BatchItem`]s, then the typed lowering of their addressing fields.
//!
//! Every input either fails with a typed transport refusal or decodes to
//! items that
//! - carry a non-blank command (`BatchItem::validate`);
//! - serialize to the canonical wire form, which decodes back to equal items
//!   (a JSON number may drift a few ULPs; see `same_items`) and stays
//!   stable when serialized again;
//! - stay equal after `into_owned`;
//! - lower each path field through `PathArg::parse` (whose canonical spelling
//!   reparses to the same path) and the selector through `Query::parse`
//!   (a display fixed point), or end in a typed path/selector refusal.
//!
//! Decoding never executes a command; `batch_apply` covers execution.

use litchi_ops::batch::{self, SourceLimits};
use litchi_ops::path::{Path, PathArg};
use litchi_ops::query::Query;
use litchi_ops::{BatchItem, Code, PropValue};

/// Input bound: the default limits are 512 MiB and 100,000 items, far past
/// what a fuzzing iteration should decode.
pub const MAX_INPUT_BYTES: usize = 64 * 1024;
/// Item bound, small enough that the limit refusal is reachable.
pub const MAX_ITEMS: usize = 64;

pub fn limits() -> SourceLimits {
    SourceLimits::builder()
        .max_input_bytes(MAX_INPUT_BYTES)
        .and_then(|builder| builder.max_items(MAX_ITEMS))
        .and_then(|builder| builder.build())
        .expect("fuzz limits are in range")
}

/// Decode `data`; on success check every item and return the item count.
pub fn exercise(data: &[u8]) -> Option<usize> {
    let items = match batch::from_bytes_with_limits(data, &limits()) {
        Ok(items) => items,
        Err(error) => {
            assert!(
                matches!(error.code(), Code::CorruptBatch | Code::LimitExceeded),
                "unexpected decode refusal {:?}",
                error.code()
            );
            if data.len() > MAX_INPUT_BYTES {
                assert_eq!(error.code(), Code::LimitExceeded);
            }
            return None;
        },
    };
    assert!(items.len() <= MAX_ITEMS);
    for item in &items {
        check_item(item);
    }
    check_round_trip(&items);
    Some(items.len())
}

fn check_item(item: &BatchItem<'_>) {
    item.validate().expect("decoded items carry a command");
    assert!(!item.command.trim().is_empty());
    assert!(
        item.clone().into_owned() == *item,
        "into_owned changed an item"
    );

    for text in [
        &item.path,
        &item.path2,
        &item.parent,
        &item.from,
        &item.to,
        &item.after,
        &item.before,
    ]
    .into_iter()
    .flatten()
    {
        lower_path(text);
    }
    if let Some(selector) = &item.selector {
        match Query::parse(selector) {
            Ok(query) => {
                let canonical = query.to_string();
                let reparsed = Query::parse(&canonical).expect("canonical selector parses");
                assert_eq!(
                    reparsed.to_string(),
                    canonical,
                    "selector display is not a fixed point"
                );
            },
            Err(error) => assert!(error.offset() <= selector.len()),
        }
    }
    if let Some(props) = &item.props {
        assert_eq!(props.iter().count(), props.len());
        for (key, value) in props.iter() {
            assert!(
                props.get(key.as_str()).is_some(),
                "a decoded key is not found"
            );
            if let PropValue::Number(number) = value {
                assert!(number.is_finite(), "JSON numbers are finite");
            }
        }
    }
}

fn lower_path(text: &str) {
    match PathArg::parse(text) {
        Ok(path) => {
            // The canonical spelling reparses through the core grammar, as in
            // the `paths` target. It is not always a fixed point of
            // `PathArg::parse`: a relative name that needs no quotes in the
            // core grammar but contains `!` reads back as Excel sheet
            // notation (`'bod! my'` -> `bod! my` -> `/bod/ my`; record 0334
            // finding F2).
            let canonical = path.to_string();
            let again = if path.is_absolute() {
                Path::parse(&canonical)
            } else {
                Path::parse_relative(&canonical)
            };
            assert_eq!(again.expect("canonical path parses"), *path.path());
        },
        Err(error) => assert!(
            matches!(
                error.code(),
                Code::InvalidPath | Code::LimitExceeded | Code::MaxDepthExceeded
            ),
            "unexpected path refusal {:?}",
            error.code()
        ),
    }
}

fn check_round_trip(items: &[BatchItem<'_>]) {
    let wire = serde_json::to_vec(items).expect("decoded items serialize");
    let again = batch::from_bytes(&wire).expect("canonical wire form decodes");
    assert!(
        same_items(&again, items),
        "wire round trip changed the items"
    );
    let rewire = serde_json::to_vec(&again).expect("round-tripped items serialize");
    let third = batch::from_bytes(&rewire).expect("canonical wire form decodes again");
    assert!(
        same_items(&third, &again),
        "canonical wire form is not stable"
    );
}

/// How far a JSON number may drift per comparison, in units in the last place.
const MAX_ULPS: u64 = 4;

/// Item equality, except that a JSON number may come back a few ULPs away.
///
/// The workspace's `serde_json` is built without `float_roundtrip`, so its
/// decimal-to-`f64` conversion is not always correctly rounded: a long
/// literal (for example an integer past `u64::MAX`) and its shortest
/// re-serialization can decode to different doubles, and each further round
/// trip can move the value again (observed: 1 and 2 ULPs). Record 0334
/// reports this as finding F1; the harness tolerates that drift and nothing
/// else.
fn same_items(left: &[BatchItem<'_>], right: &[BatchItem<'_>]) -> bool {
    left.len() == right.len()
        && left.iter().zip(right).all(|(left, right)| {
            if left == right {
                return true;
            }
            let (Some(left_props), Some(right_props)) = (&left.props, &right.props) else {
                return false;
            };
            let mut left_rest = left.clone();
            let mut right_rest = right.clone();
            left_rest.props = None;
            right_rest.props = None;
            left_rest == right_rest
                && left_props.len() == right_props.len()
                && left_props
                    .iter()
                    .zip(right_props.iter())
                    .all(|((lk, lv), (rk, rv))| {
                        lk == rk
                            && (lv == rv
                                || matches!((lv, rv), (PropValue::Number(l), PropValue::Number(r))
                                if l.to_bits().abs_diff(r.to_bits()) <= MAX_ULPS))
                    })
        })
}
