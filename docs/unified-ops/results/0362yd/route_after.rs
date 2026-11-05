use std::sync::Arc;

use litchi_core::{ReadAt, TextOutputOptions};

pub(super) const MODE: &str = "12385bc0453978b5406da808d430abae9ce77b5b";

fn owner(source: Arc<dyn ReadAt>) -> litchi_odp::SourceBackedPresentation {
    litchi_odp::SourceBackedPresentation::from_read_at(source).expect("open")
}

pub(super) fn open(source: Arc<dyn ReadAt>, _sink: &mut super::Digest) {
    let _ = owner(source);
}

pub(super) fn first_slide(source: Arc<dyn ReadAt>, _sink: &mut super::Digest) {
    let presentation = owner(source);
    std::hint::black_box(presentation.slide(0).expect("slide"));
}

pub(super) fn repeat_slide(source: Arc<dyn ReadAt>, _sink: &mut super::Digest) {
    let presentation = owner(source);
    std::hint::black_box(presentation.slide(0).expect("slide"));
    std::hint::black_box(presentation.slide(1).expect("slide"));
}

pub(super) fn first_text(source: Arc<dyn ReadAt>, _sink: &mut super::Digest) {
    let presentation = owner(source);
    std::hint::black_box(presentation.text().expect("text"));
}

pub(super) fn repeat_text(source: Arc<dyn ReadAt>, _sink: &mut super::Digest) {
    let presentation = owner(source);
    std::hint::black_box(presentation.text().expect("text"));
    std::hint::black_box(presentation.text().expect("text"));
}

pub(super) fn write_text(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let presentation = owner(source);
    presentation
        .write_text_to(sink, TextOutputOptions::default())
        .expect("write text");
}

pub(super) fn repeat_write_text(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let presentation = owner(source);
    presentation
        .write_text_to(sink, TextOutputOptions::default())
        .expect("write text");
    presentation
        .write_text_to(sink, TextOutputOptions::default())
        .expect("write text");
}

/// Keep the unchanged publication lifecycle common with f755903f45, which
/// predates the source-backed write owner method.
pub(super) fn unchanged_write(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let presentation = owner(source);
    let package = presentation.materialize().expect("materialize");
    package.write_to(sink).expect("write");
}

pub(super) fn append_commit_write(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let presentation = owner(source);
    let commit = presentation
        .edit_tail_append("Appended title", "Appended body")
        .expect("edit")
        .commit()
        .expect("commit");
    commit.package().write_to(sink).expect("write");
}

pub(super) fn replay(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let presentation = owner(source);
    let commit = presentation
        .edit_tail_append("Appended title", "Appended body")
        .expect("edit")
        .commit()
        .expect("commit");
    let replayed = presentation
        .apply_tail_append(commit.patch())
        .expect("replay");
    replayed.package().write_to(sink).expect("write");
}
