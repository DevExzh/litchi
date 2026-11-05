//! Record 0359b: current native edit probe with checked refusal fidelity.
#![allow(
    clippy::unwrap_used,
    reason = "a failed benchmark operation invalidates the run"
)]

use std::hint::black_box;
use std::time::Instant;

use litchi_rtf::edit::TextSpan;
use litchi_rtf::tail_append::TailSelector;
use litchi_rtf::{Package, UnderlineStyle};

const UNSUPPORTED_REASON: &str = "underline edits require non-empty text within one paragraph";
const PROTECTED_REASON: &str = "document protection refuses changed publication";
const NATIVE_ERROR_API: &str = "litchi_rtf::edit::Error::Refused(Refusal)";
const APPLY_RESULT_API: &str = "litchi_rtf::Package::apply -> litchi_rtf::edit::Commit";

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

fn expected_source_text() -> String {
    format!("{}end", "seed\n".repeat(1000))
}

fn package_bytes(package: &Package) -> Vec<u8> {
    package.to_bytes().unwrap()
}

fn committed_bytes(commit: &litchi_rtf::edit::Commit) -> Vec<u8> {
    package_bytes(commit.package())
}

fn tail_committed_bytes(commit: &litchi_rtf::tail_append::Commit) -> Vec<u8> {
    package_bytes(commit.package())
}

fn replay_bytes(source: &Package, patch: &litchi_rtf::edit::Patch) -> Vec<u8> {
    committed_bytes(&replay_once(source, patch))
}

fn replay_once(source: &Package, patch: &litchi_rtf::edit::Patch) -> litchi_rtf::edit::Commit {
    source.apply(patch).unwrap()
}

fn assert_reopen_exact(bytes: &[u8]) {
    let text = std::str::from_utf8(bytes).unwrap();
    assert_eq!(package_bytes(&Package::parse(text).unwrap()), bytes);
}

fn assert_unsupported<T>(result: Result<T, litchi_rtf::edit::Error>) {
    match result {
        Err(litchi_rtf::edit::Error::Refused(refusal)) => {
            assert_eq!(refusal.kind().as_str(), "unsupported_content");
            assert_eq!(refusal.operation(), "edit");
            assert_eq!(refusal.reason(), UNSUPPORTED_REASON);
        },
        Ok(_) => panic!("expected a typed unsupported-content refusal"),
        Err(other) => panic!("unexpected unsupported-content error: {other:?}"),
    }
}

fn assert_protected<T>(result: Result<T, litchi_rtf::edit::Error>) {
    match result {
        Err(litchi_rtf::edit::Error::Refused(refusal)) => {
            assert_eq!(refusal.kind().as_str(), "protected");
            assert_eq!(refusal.operation(), "edit");
            assert_eq!(refusal.reason(), PROTECTED_REASON);
        },
        Ok(_) => panic!("expected a typed protected-document refusal"),
        Err(other) => panic!("unexpected protected-document error: {other:?}"),
    }
}

fn run_correctness_checks(
    source: &Package,
    changed: &litchi_rtf::edit::Commit,
    protected: &Package,
    span: TextSpan,
) {
    let source_bytes = package_bytes(source);
    let protected_bytes = package_bytes(protected);
    let expected_source_text = expected_source_text();
    let expected_changed_text = expected_source_text.replacen("seed", "tail", 1);

    // Exact reopen, commit publication, replay, inverse replay, and the
    // untouched source artifacts are correctness checks outside every clock.
    assert_eq!(source.text(), expected_source_text);
    assert_reopen_exact(&source_bytes);
    let changed_bytes = committed_bytes(changed);
    assert_eq!(changed.package().text(), expected_changed_text);
    assert_reopen_exact(&changed_bytes);
    let replayed_bytes = replay_bytes(source, changed.patch());
    assert_eq!(replayed_bytes, changed_bytes);
    assert_eq!(
        Package::parse(std::str::from_utf8(&replayed_bytes).unwrap())
            .unwrap()
            .text(),
        expected_changed_text
    );
    assert_eq!(
        replay_bytes(changed.package(), &changed.patch().inverse()),
        source_bytes
    );

    let noop = source.edit().commit().unwrap();
    assert_eq!(package_bytes(noop.package()), source_bytes);
    assert_eq!(noop.package().text(), expected_source_text);

    let mut unsupported_check = source.edit();
    assert_unsupported(
        unsupported_check.set_text_underline(TextSpan::new(0, 0).unwrap(), UnderlineStyle::Single),
    );
    assert_eq!(package_bytes(source), source_bytes);

    let mut protected_check = protected.edit();
    protected_check.replace_text(span, "tail").unwrap();
    assert_protected(protected_check.commit());
    assert_eq!(package_bytes(protected), protected_bytes);
}

fn main() {
    let wire = format!("{{\\rtf1\\ansi {}end}}", "seed\\par ".repeat(1000));
    let source = Package::parse(&wire).unwrap();
    let span = TextSpan::new(0, 4).unwrap();
    let mut edit = source.edit();
    edit.replace_text(span, "tail").unwrap();
    let changed = edit.commit().unwrap();
    let protected = Package::parse(r"{\rtf1\ansi\readprot\enforceprot1 seed}").unwrap();

    run_correctness_checks(&source, &changed, &protected, span);

    // The tail route has the same exact-reopen fidelity obligation as the
    // primary commit route, but its publication remains outside the timing
    // loops so route correctness cannot be inferred from elapsed output.
    let mut tail_edit = source.edit_tail_append(TailSelector::Body);
    tail_edit.append_text_paragraphs(&["tail"]).unwrap();
    let tail_commit = tail_edit.commit().unwrap();
    let tail_bytes = tail_committed_bytes(&tail_commit);
    let expected_source_text = expected_source_text();
    let expected_tail_text = format!("{expected_source_text}\ntail\n");
    assert_eq!(tail_commit.package().text(), expected_tail_text);
    assert_reopen_exact(&tail_bytes);
    assert_eq!(
        Package::parse(std::str::from_utf8(&tail_bytes).unwrap())
            .unwrap()
            .text(),
        expected_tail_text
    );

    if std::env::args().any(|arg| arg == "--check-only") {
        println!(
            "{{\"validation\":\"passed\",\"native_error_api\":\"{NATIVE_ERROR_API}\",\"apply_result_api\":\"{APPLY_RESULT_API}\",\"routes\":8}}"
        );
        return;
    }

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
            black_box(committed_bytes(&edit.commit().unwrap()));
        },
        1,
        "\"outcome\":\"ok\",\"fidelity\":\"exact\"",
    );
    measure(
        "replay_1k",
        || {
            black_box(replay_once(&source, changed.patch()));
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
            let result =
                edit.set_text_underline(TextSpan::new(0, 0).unwrap(), UnderlineStyle::Single);
            let _ = black_box(result);
        },
        1000,
        "\"outcome\":\"refused\",\"refusal_kind\":\"unsupported_content\",\"operation\":\"edit\",\"reason\":\"underline edits require non-empty text within one paragraph\",\"fidelity\":\"unchanged\"",
    );
    measure(
        "protected_commit",
        || {
            let mut edit = protected.edit();
            edit.replace_text(span, "tail").unwrap();
            let _ = black_box(edit.commit());
        },
        100,
        "\"outcome\":\"refused\",\"refusal_kind\":\"protected\",\"operation\":\"edit\",\"reason\":\"document protection refuses changed publication\",\"fidelity\":\"unchanged\"",
    );
}
