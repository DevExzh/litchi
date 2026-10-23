use std::sync::Arc;

use litchi_core::{ReadAt, TextOutputOptions};

pub(super) const MODE: &str = "UNBUILT-0362YE";

fn owner(source: Arc<dyn ReadAt>) -> litchi_odp::SourceBackedPresentation {
    litchi_odp::SourceBackedPresentation::from_read_at(source).expect("open")
}

pub(super) fn open(source: Arc<dyn ReadAt>, _sink: &mut super::Digest) {
    let _ = owner(source);
}

pub(super) fn first_slide(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let presentation = owner(source);
    let slide = presentation.slide(0).expect("slide").expect("first slide");
    sink.update_text(&slide.all_text());
}

pub(super) fn repeat_slide(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let presentation = owner(source);
    let first = presentation.slide(0).expect("slide").expect("first slide");
    sink.update_text(&first.all_text());
    let second = presentation.slide(1).expect("slide").expect("second slide");
    sink.update_text(&second.all_text());
}

pub(super) fn first_text(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let presentation = owner(source);
    sink.update_text(&presentation.text().expect("text"));
}

pub(super) fn repeat_text(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let presentation = owner(source);
    let first = presentation.text().expect("text");
    sink.update_text(&first);
    let second = presentation.text().expect("text");
    sink.update_text(&second);
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

pub(super) fn append_phases(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let open_baseline = super::phase_reset();
    let presentation = owner(source);
    super::phase_report("append.open", open_baseline);

    let stage_baseline = super::phase_reset();
    let edit = presentation
        .edit_tail_append("Appended title", "Appended body")
        .expect("edit");
    super::phase_report("append.stage", stage_baseline);

    let baseline = super::phase_reset();
    let commit = edit.commit().expect("commit");
    super::phase_report("append.commit", baseline);

    let baseline = super::phase_reset();
    commit.package().write_to(sink).expect("write");
    super::phase_report("append.write", baseline);
}

pub(super) fn replay_phases(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let open_baseline = super::phase_reset();
    let presentation = owner(source);
    let edit = presentation
        .edit_tail_append("Appended title", "Appended body")
        .expect("edit");
    let commit = edit.commit().expect("commit");
    super::phase_report("replay.prepare", open_baseline);

    let baseline = super::phase_reset();
    let replayed = presentation
        .apply_tail_append(commit.patch())
        .expect("replay");
    super::phase_report("replay.apply", baseline);

    let baseline = super::phase_reset();
    replayed.package().write_to(sink).expect("write");
    super::phase_report("replay.write", baseline);
}
