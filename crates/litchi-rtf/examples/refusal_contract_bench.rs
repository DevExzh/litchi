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

fn measure(name: &str, mut run: impl FnMut(), iterations: u32) {
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
    println!("{{\"scenario\":\"{name}\",\"median_us\":{}}}", samples[15]);
}

fn main() {
    let wire = format!("{{\\rtf1\\ansi {}end}}", "seed\\par ".repeat(1000));
    let source = Package::parse(&wire).unwrap();
    let span = TextSpan::new(0, 4).unwrap();
    let mut edit = source.edit();
    edit.replace_text(span, "tail").unwrap();
    let changed = edit.commit().unwrap();
    let protected = Package::parse(r"{\rtf1\ansi\readprot\enforceprot1 seed}").unwrap();

    measure(
        "open_1k",
        || {
            black_box(Package::parse(black_box(&wire)).unwrap());
        },
        1,
    );
    measure(
        "noop_1k",
        || {
            black_box(source.edit().commit().unwrap());
        },
        100,
    );
    measure(
        "text_commit_1k",
        || {
            let mut edit = source.edit();
            edit.replace_text(span, "tail").unwrap();
            black_box(edit.commit().unwrap());
        },
        1,
    );
    measure(
        "text_commit_save_1k",
        || {
            let mut edit = source.edit();
            edit.replace_text(span, "tail").unwrap();
            black_box(edit.commit().unwrap().package().to_bytes().unwrap());
        },
        1,
    );
    measure(
        "replay_1k",
        || {
            black_box(source.apply(changed.patch()).unwrap());
        },
        100,
    );
    measure(
        "tail_commit_1k",
        || {
            let mut edit = source.edit_tail_append(TailSelector::Body);
            edit.append_text_paragraphs(&["tail"]).unwrap();
            black_box(edit.commit().unwrap());
        },
        1,
    );
    measure(
        "unsupported_content",
        || {
            let mut edit = source.edit();
            let refused =
                edit.set_text_underline(TextSpan::new(0, 0).unwrap(), UnderlineStyle::Single);
            assert!(black_box(refused).is_err());
        },
        1000,
    );
    measure(
        "protected_commit",
        || {
            let mut edit = protected.edit();
            edit.replace_text(span, "tail").unwrap();
            assert!(black_box(edit.commit()).is_err());
        },
        100,
    );
}
