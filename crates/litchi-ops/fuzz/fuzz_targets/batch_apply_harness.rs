//! OfficeCLI batch execution over small created documents (record 0334).
//!
//! Inputs that decode (see `batch_payload_harness`) run as one batch against
//! a blank DOCX, XLSX and PPTX document, atomically and best effort with
//! OfficeCLI's skip-and-warn property policy. Execution must end in a commit
//! or a typed refusal, never a panic or an `internal_error`, and a failed
//! atomic batch leaves no trace: the caller's handle is immutable, so the
//! check is that a retry over the same handle behaves identically.

use std::sync::OnceLock;

use litchi_core::FileFormat;
use litchi_ops::batch::{self, Options, Unsupported};
use litchi_ops::{BatchMode, Code, File, ItemResult, create};

use crate::batch_payload_harness::limits;

/// Small enough that one iteration stays fast on a blank document.
const MAX_ITEMS: usize = 16;

fn documents() -> &'static [File; 3] {
    static DOCUMENTS: OnceLock<[File; 3]> = OnceLock::new();
    DOCUMENTS.get_or_init(|| {
        [FileFormat::Docx, FileFormat::Xlsx, FileFormat::Pptx]
            .map(|format| create(format).expect("blank document"))
    })
}

pub fn exercise(data: &[u8]) {
    let Ok(items) = batch::from_bytes(data, &limits()) else {
        return;
    };
    if items.len() > MAX_ITEMS {
        return;
    }
    for file in documents() {
        let atomic = Options::new().unsupported(Unsupported::Warn);
        let first = file.batch_with(&items, &atomic);
        match &first {
            Ok(commit) => assert_eq!(commit.value().len(), items.len()),
            Err(error) => typed(error.code()),
        }
        // The handle is a snapshot: the same batch behaves the same again.
        let second = file.batch_with(&items, &atomic);
        assert_eq!(
            first.is_ok(),
            second.is_ok(),
            "atomic batch is not repeatable"
        );
        if let (Err(first), Err(second)) = (&first, &second) {
            assert_eq!(first.code(), second.code());
        }

        let best_effort = atomic.mode(BatchMode::BestEffort);
        match file.batch_with(&items, &best_effort) {
            Ok(commit) => {
                assert_eq!(commit.value().len(), items.len());
                for result in commit.value() {
                    if let ItemResult::Failed(error) = result {
                        typed(error.code());
                    }
                }
            },
            Err(error) => typed(error.code()),
        }
    }
}

fn typed(code: Code) {
    assert_ne!(
        code,
        Code::InternalError,
        "batch execution reported an internal error"
    );
}
