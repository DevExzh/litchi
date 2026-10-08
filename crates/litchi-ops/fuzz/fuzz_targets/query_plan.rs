#![no_main]

//! Query plan compilation fuzz target (M1 step R3, record 0010 §5.2).
//!
//! Every UTF-8 input that parses as a selector compiles, for each of the
//! three fake format families, either to a plan or to a typed refusal
//! (`invalid_selector`, `unsupported_path`, `unsupported_capability`): never
//! an internal error or a panic. A compiled plan then runs over a small fake
//! document of its family for at most 64 matches, and ends in matches or in a
//! typed error the query contract allows. Grammar errors are the
//! `selector_grammar` target's business (`litchi-core/fuzz`), so inputs that
//! do not parse are skipped.

use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_ops::Code;
use litchi_ops::path::limits;
use litchi_ops::query::__fake::{Document, Family, compile_for, run};
use litchi_ops::query::Query;

/// One small document per family, built once.
fn documents() -> &'static [(Family, Document); 3] {
    static DOCUMENTS: OnceLock<[(Family, Document); 3]> = OnceLock::new();
    DOCUMENTS.get_or_init(|| {
        [
            (Family::Word, Document::word(12)),
            (Family::Presentation, Document::deck(2, 3)),
            (Family::Workbook, Document::workbook(6)),
        ]
    })
}

fuzz_target!(|data: &[u8]| {
    if data.len() > limits::MAX_INPUT_BYTES + 16 {
        return;
    }
    let Ok(input) = std::str::from_utf8(data) else {
        return;
    };
    let Ok(query) = Query::parse(input) else {
        return;
    };
    for (family, document) in documents() {
        let plan = match compile_for(*family, &query) {
            Ok(plan) => plan,
            Err(error) => {
                assert!(
                    matches!(
                        error.code(),
                        Code::InvalidSelector
                            | Code::UnsupportedPath
                            | Code::UnsupportedCapability
                            | Code::LimitExceeded
                            | Code::MaxDepthExceeded
                            | Code::InvalidPath
                    ),
                    "{family:?} {input:?}: {error}"
                );
                continue;
            },
        };
        assert_eq!(plan.alternatives(), query.alternatives().len());
        for diagnostic in plan.diagnostics() {
            assert!(!diagnostic.message().is_empty(), "{input:?}");
        }
        match run(document, &plan, 64) {
            Ok(outcome) => assert!(outcome.matches <= 64, "{input:?}"),
            Err(error) => assert!(
                matches!(
                    error.code(),
                    Code::LimitExceeded
                        | Code::MaxDepthExceeded
                        | Code::InvalidPath
                        | Code::InvalidSelector
                        | Code::PathNotFound
                        | Code::NotFound
                        | Code::UnsupportedPath
                        | Code::UnsupportedCapability
                ),
                "{family:?} {input:?}: {error}"
            ),
        }
    }
});
