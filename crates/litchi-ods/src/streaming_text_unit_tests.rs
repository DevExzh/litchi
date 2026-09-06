//! Scalar-reference checks for text encoding and exact refusal boundaries.

#![allow(clippy::expect_used, clippy::unwrap_used)]

use std::num::{NonZeroU64, NonZeroUsize};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits, Resource,
};

use super::{RowWriter, write_text};

#[derive(Debug)]
struct Snapshot {
    result_error: Option<String>,
    bytes: Vec<u8>,
    row_bytes: usize,
    local_limit: Option<(&'static str, u64, u64)>,
    execution_error: Option<ExecutionError>,
    parent_work: u64,
    child_work: u64,
}

fn execution_limits() -> ExecutionLimits {
    ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one task"),
        NonZeroU64::new(u64::MAX).expect("finite byte ceiling"),
        0,
    )
    .expect("valid execution policy")
}

fn context_with_work(parent_work: u64, child_work: u64) -> (Budget, Budget, ExecutionContext) {
    let parent = Budget::root(
        "streaming-text-parent",
        core_limits(Resource::Work, parent_work),
    );
    let child = parent.child(
        "streaming-text-child",
        core_limits(Resource::Work, child_work),
    );
    let (_source, token) = CancellationSource::pair();
    let context = ExecutionContext::new(child.clone(), token, execution_limits());
    (parent, child, context)
}

fn core_limits(resource: Resource, limit: u64) -> Limits {
    let mut values = [u64::MAX; 6];
    let index = match resource {
        Resource::Memory => 0,
        Resource::InputBytes => 1,
        Resource::OutputBytes => 2,
        Resource::Objects => 3,
        Resource::Depth => 4,
        Resource::Work => 5,
        _ => unreachable!("the six core resources are exhaustive"),
    };
    values[index] = limit;
    Limits::new(
        values[0], values[1], values[2], values[3], values[4], values[5],
    )
}

fn snapshot(value: &str, maximum: usize, parent_work: u64, child_work: u64) -> Snapshot {
    let (parent, child, context) = context_with_work(parent_work, child_work);
    let mut bytes = Vec::new();
    let mut writer = RowWriter::new(&mut bytes, &context, maximum);
    let result_error = write_text(&mut writer, value)
        .err()
        .map(|error| error.to_string());
    let row_bytes = writer.bytes;
    let local_limit = writer
        .limit_error
        .map(|limit| (limit.resource, limit.observed, limit.limit));
    let execution_error = writer.execution_error.clone();
    drop(writer);
    Snapshot {
        result_error,
        bytes,
        row_bytes,
        local_limit,
        execution_error,
        parent_work: parent.used(Resource::Work),
        child_work: child.used(Resource::Work),
    }
}

fn reference_pieces(value: &str) -> Vec<Vec<u8>> {
    value
        .chars()
        .map(|character| match character {
            '&' => b"&amp;".to_vec(),
            '<' => b"&lt;".to_vec(),
            '>' => b"&gt;".to_vec(),
            '"' => b"&quot;".to_vec(),
            '\'' => b"&apos;".to_vec(),
            '\r' => b"&#13;".to_vec(),
            '\n' => b"&#10;".to_vec(),
            '\t' => b"&#9;".to_vec(),
            character => character.to_string().into_bytes(),
        })
        .collect()
}

fn reference_output(value: &str) -> Vec<u8> {
    reference_pieces(value).into_iter().flatten().collect()
}

fn reference_prefix(value: &str, maximum: usize) -> (Vec<u8>, Option<u64>) {
    let mut prefix = Vec::new();
    for piece in reference_pieces(value) {
        let next = prefix.len().saturating_add(piece.len());
        if next > maximum {
            return (prefix, Some(next as u64));
        }
        prefix.extend_from_slice(&piece);
    }
    (prefix, None)
}

fn mixed_text() -> String {
    let mut value = "a".repeat(253);
    value.push_str("é🙂");
    value.push_str(&"b".repeat(513));
    for _ in 0..24 {
        value.push_str("spané🙂");
        value.push_str("&<>\"'\r\n\t");
    }
    value
}

fn assert_common_snapshot(actual: &Snapshot, expected_bytes: &[u8], expected_error: bool) {
    assert_eq!(actual.bytes, expected_bytes);
    assert_eq!(actual.row_bytes, expected_bytes.len());
    assert_eq!(actual.result_error.is_some(), expected_error);
    assert_eq!(actual.parent_work, expected_bytes.len() as u64);
    assert_eq!(actual.child_work, expected_bytes.len() as u64);
}

#[test]
fn mixed_unicode_and_escape_text_matches_independent_scalar_reference() {
    let value = mixed_text();
    let expected = reference_output(&value);
    let snapshot = snapshot(&value, expected.len(), u64::MAX, u64::MAX);
    assert_common_snapshot(&snapshot, &expected, false);
    assert!(snapshot.local_limit.is_none());
    assert!(snapshot.execution_error.is_none());
}

#[test]
fn every_row_window_offset_preserves_reference_buffer_prefix_and_local_limit() {
    let value = mixed_text();
    let expected = reference_output(&value);
    for maximum in 0..=expected.len() + 1 {
        let (expected_prefix, observed) = reference_prefix(&value, maximum);
        let snapshot = snapshot(&value, maximum, u64::MAX, u64::MAX);
        let failed = observed.is_some();
        assert_common_snapshot(&snapshot, &expected_prefix, failed);
        match observed {
            Some(observed) => {
                assert_eq!(
                    snapshot.local_limit,
                    Some(("row XML bytes", observed, maximum as u64))
                );
                assert!(snapshot.execution_error.is_none());
            },
            None => {
                assert!(snapshot.local_limit.is_none());
                assert!(snapshot.execution_error.is_none());
            },
        }
    }
}

#[test]
fn every_work_offset_preserves_reference_buffer_prefix_for_child_and_parent_budgets() {
    let value = mixed_text();
    let expected = reference_output(&value);
    for work_limit in 0..=expected.len() as u64 + 1 {
        let (expected_prefix, observed) = reference_prefix(&value, work_limit as usize);

        let child_snapshot = snapshot(&value, expected.len(), u64::MAX, work_limit);
        assert_common_snapshot(&child_snapshot, &expected_prefix, observed.is_some());
        if let Some(observed) = observed {
            if let Some(ExecutionError::ResourceLimit(limit)) = &child_snapshot.execution_error {
                assert_eq!(limit.scope.as_ref(), "streaming-text-child");
            }
            assert!(child_snapshot.local_limit.is_none());
            assert_eq!(
                child_snapshot
                    .execution_error
                    .as_ref()
                    .map(|error| match error {
                        ExecutionError::ResourceLimit(limit) => {
                            (limit.resource, limit.observed, limit.limit)
                        },
                        _ => panic!("expected a resource limit"),
                    }),
                Some((Resource::Work, observed, work_limit)),
            );
        } else {
            assert!(child_snapshot.execution_error.is_none());
        }

        let parent_snapshot = snapshot(&value, expected.len(), work_limit, u64::MAX);
        assert_common_snapshot(&parent_snapshot, &expected_prefix, observed.is_some());
        if let Some(observed) = observed {
            if let Some(ExecutionError::ResourceLimit(limit)) = &parent_snapshot.execution_error {
                assert_eq!(limit.scope.as_ref(), "streaming-text-parent");
            }
            assert!(parent_snapshot.local_limit.is_none());
            assert_eq!(
                parent_snapshot
                    .execution_error
                    .as_ref()
                    .map(|error| match error {
                        ExecutionError::ResourceLimit(limit) => {
                            (limit.resource, limit.observed, limit.limit)
                        },
                        _ => panic!("expected a resource limit"),
                    }),
                Some((Resource::Work, observed, work_limit)),
            );
        } else {
            assert!(parent_snapshot.execution_error.is_none());
        }
    }
}

#[test]
fn empty_text_is_a_zero_work_success_and_boundary_values_are_exact() {
    let empty = snapshot("", 0, 0, 0);
    assert_common_snapshot(&empty, &[], false);
    assert!(empty.local_limit.is_none());
    assert!(empty.execution_error.is_none());

    for value in ["é", "🙂", "&", "\r", "\n", "\t"] {
        let expected = reference_output(value);
        let snapshot = snapshot(
            value,
            expected.len(),
            expected.len() as u64,
            expected.len() as u64,
        );
        assert_common_snapshot(&snapshot, &expected, false);
    }
}

#[test]
fn competing_row_and_work_limits_preserve_scalar_refusal_precedence() {
    let value = mixed_text();
    let length = reference_output(&value).len();
    let boundaries = [0usize, 1, 254, 255, 256, 257, 258, 511, length];
    for maximum in boundaries {
        for work in boundaries {
            let (prefix, observed) = reference_prefix(&value, maximum.min(work));
            let actual = snapshot(&value, maximum, u64::MAX, work as u64);
            assert_common_snapshot(&actual, &prefix, observed.is_some());
            match observed {
                Some(next) if next > maximum as u64 => {
                    assert_eq!(
                        actual.local_limit,
                        Some(("row XML bytes", next, maximum as u64))
                    );
                    assert!(actual.execution_error.is_none());
                },
                Some(next) => {
                    assert!(actual.local_limit.is_none());
                    match actual.execution_error {
                        Some(ExecutionError::ResourceLimit(limit)) => {
                            assert_eq!(limit.resource, Resource::Work);
                            assert_eq!(limit.observed, next);
                            assert_eq!(limit.limit, work as u64);
                            assert_eq!(limit.scope.as_ref(), "streaming-text-child");
                        },
                        other => panic!("expected Work refusal, got {other:?}"),
                    }
                },
                None => {
                    assert!(actual.local_limit.is_none());
                    assert!(actual.execution_error.is_none());
                },
            }
        }
    }
}
