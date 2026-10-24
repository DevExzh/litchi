//! OfficeCLI batch-payload ingress (record 0334): bounded JSON decoding into
//! typed [`BatchItem`]s, then the typed lowering of their addressing fields.
//!
//! Every input either fails with a typed transport refusal or decodes to
//! items that
//! - carry a non-blank command (`BatchItem::validate`);
//! - serialize to the canonical wire form, which decodes back to equal items
//!   (numbers included, record 0338 F1) and serializes to the same bytes
//!   again;
//! - stay equal after `into_owned`;
//! - lower each path field through `PathArg::parse` and the selector through
//!   `Query::parse`, each to a display fixed point, or end in a typed
//!   path/selector refusal.
//!
//! Decoding never executes a command; `batch_apply` covers execution.

use litchi_ops::batch::{self, SourceLimits};
use litchi_ops::path::PathArg;
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
    let items = match batch::from_bytes(data, &limits()) {
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
            // The canonical spelling is a fixed point of `PathArg::parse`
            // (record 0338 F2).
            let canonical = path.to_string();
            let again = PathArg::parse(&canonical).expect("canonical path parses");
            assert_eq!(again.path(), path.path(), "canonical path reparses");
            assert_eq!(
                again.to_string(),
                canonical,
                "path display is not a fixed point"
            );
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
    let again = batch::from_bytes(&wire, &batch::SourceLimits::default())
        .expect("canonical wire form decodes");
    // Numbers too: `serde_json` parses correctly rounded (record 0338 F1).
    assert!(again == items, "wire round trip changed the items");
    let rewire = serde_json::to_vec(&again).expect("round-tripped items serialize");
    let third = batch::from_bytes(&rewire, &batch::SourceLimits::default())
        .expect("canonical wire form decodes again");
    assert!(third == again, "canonical wire form is not stable");
    assert_eq!(rewire, wire, "canonical wire form is not a fixed point");
}
