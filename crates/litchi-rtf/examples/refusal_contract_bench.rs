//! Record 0359b: identical native edit probe before and after refusal migration.
#![allow(
    clippy::unwrap_used,
    reason = "a failed benchmark operation invalidates the run"
)]

use std::hint::black_box;
use std::time::Instant;

use litchi_rtf::edit::TextSpan;
use litchi_rtf::tail_append::TailSelector;
use litchi_rtf::{Package, UnderlineStyle};

fn measure(name: &str, mut run: impl FnMut(), iterations: u32, contract: &str) {
    let mut samples = Vec::new();
    for sample in 0..34 {
        let started = Instant::now();
        for _ in 0..iterations {
            run();
        }
        if sample >= 3 {
            samples.push(started.elapsed().as_secs_f64() * 1e6 / f64::from(iterations));
        }
    }
    samples.sort_by(f64::total_cmp);
    println!(
        "{{\"scenario\":\"{name}\",\"median_us\":{},\"workload\":\"rtf-refusal-contract-v1\",{contract}}}",
        samples[15]
    );
}

fn assert_refusal<T>(
    result: Result<T, litchi_rtf::edit::Error>,
    kind: &str,
    operation: &str,
    reason: &str,
) {
    match result {
        Err(litchi_rtf::edit::Error::Refused(refusal)) => {
            assert_eq!(refusal.kind().as_str(), kind);
            assert_eq!(refusal.operation(), operation);
            assert_eq!(refusal.reason(), reason);
        }
        Ok(_) => panic!("expected a typed refusal"),
        Err(other) => panic!("unexpected refusal error: {other:?}"),
    }
}

fn main() {
    let wire = format!("{{\\rtf1\\ansi {}end}}", "seed\\par ".repeat(1000));
    let source = Package::parse(&wire).unwrap();
    let span = TextSpan::new(0, 4).unwrap();
    let mut edit = source.edit();
    edit.replace_text(span, "tail").unwrap();
    let changed = edit.commit().unwrap();
    let protected = Package::parse(r"{\rtf1\ansi\readprot\enforceprot1 seed}").unwrap();

    // The native qualification output carries these stable contract fields;
    // keep the exact refusal classification/reason and unchanged-fidelity
    // assertions next to the timed routes so a bare `is_err()` cannot pass.
    let source_bytes = source.to_bytes().unwrap();
    let protected_bytes = protected.to_bytes().unwrap();
    assert_eq!(Package::parse(&source_bytes).unwrap().to_bytes().unwrap(), source_bytes);
    assert_eq!(changed.package().to_bytes().unwrap(), source.apply(changed.patch()).unwrap().to_bytes().unwrap());
    let mut unsupported_check = source.edit();
    assert_refusal(
        unsupported_check.set_text_underline(TextSpan::new(0, 0).unwrap(), UnderlineStyle::Single),
        "unsupported_content",
        "edit",
        "underline edits require non-empty text within one paragraph",
    );
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
    let mut protected_check = protected.edit();
    assert_refusal(
        protected_check.replace_text(span, "tail"),
        "protected",
        "edit",
        "document protection refuses changed publication",
    );
    assert_eq!(protected.to_bytes().unwrap(), protected_bytes);

    measure(
        "open_1k",
        || {
            black_box(Package::parse(black_box(&wire)).unwrap());
        },
        1,
        "\"outcome\":\"ok\",\"fidelity\":\"exact\"",
    );
    measure(
        "noop_1k",
        || {
            black_box(source.edit().commit().unwrap());
        },
        100,
        "\"outcome\":\"ok\",\"fidelity\":\"exact\"",
    );
    measure(
        "text_commit_1k",
        || {
            let mut edit = source.edit();
            edit.replace_text(span, "tail").unwrap();
            black_box(edit.commit().unwrap());
        },
        1,
        "\"outcome\":\"ok\",\"fidelity\":\"exact\"",
    );
    measure(
        "text_commit_save_1k",
        || {
            let mut edit = source.edit();
            edit.replace_text(span, "tail").unwrap();
            black_box(edit.commit().unwrap().package().to_bytes().unwrap());
        },
        1,
        "\"outcome\":\"ok\",\"fidelity\":\"exact\"",
    );
    measure(
        "replay_1k",
        || {
            black_box(source.apply(changed.patch()).unwrap());
        },
        100,
        "\"outcome\":\"ok\",\"fidelity\":\"exact\"",
    );
    measure(
        "tail_commit_1k",
        || {
            let mut edit = source.edit_tail_append(TailSelector::Body);
            edit.append_text_paragraphs(&["tail"]).unwrap();
            black_box(edit.commit().unwrap());
        },
        1,
        "\"outcome\":\"ok\",\"fidelity\":\"exact\"",
    );
    measure(
        "unsupported_content",
        || {
            let mut edit = source.edit();
            assert_refusal(
                black_box(edit.set_text_underline(TextSpan::new(0, 0).unwrap(), UnderlineStyle::Single)),
                "unsupported_content",
                "edit",
                "underline edits require non-empty text within one paragraph",
            );
        },
        1000,
        "\"outcome\":\"refused\",\"refusal_kind\":\"unsupported_content\",\"operation\":\"edit\",\"reason\":\"underline edits require non-empty text within one paragraph\",\"fidelity\":\"unchanged\"",
    );
    measure(
        "protected_commit",
        || {
            let mut edit = protected.edit();
            edit.replace_text(span, "tail").unwrap();
            assert_refusal(
                black_box(edit.commit()),
                "protected",
                "edit",
                "document protection refuses changed publication",
            );
        },
        100,
        "\"outcome\":\"refused\",\"refusal_kind\":\"protected\",\"operation\":\"edit\",\"reason\":\"document protection refuses changed publication\",\"fidelity\":\"unchanged\"",
    );
}
