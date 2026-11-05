use std::sync::Arc;

use litchi_core::{ReadAt, TextOutputOptions};
use litchi_odf_common::core::SourceContentPublicationOptions;

pub(super) const MODE: &str = "0358a-pre-change-e4a0614f5d";

fn owner(source: Arc<dyn ReadAt>) -> litchi_odp::SourceBackedPresentation {
    litchi_odp::SourceBackedPresentation::from_read_at(source).expect("open")
}

pub(super) fn open(source: Arc<dyn ReadAt>, _sink: &mut super::Digest) {
    let _ = owner(source);
}

pub(super) fn first_slide(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let presentation = owner(source);
    let slide = presentation
        .slide(0)
        .expect("slide")
        .expect("first slide");
    sink.update_text(&slide.all_text());
}

pub(super) fn repeat_slide(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let presentation = owner(source);
    let first = presentation
        .slide(0)
        .expect("slide")
        .expect("first slide");
    sink.update_text(&first.all_text());
    let second = presentation
        .slide(1)
        .expect("slide")
        .expect("second slide");
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

/// The old source owner had no source-backed `write_to`; materialize and write
/// is the unchanged publication lifecycle available on both endpoints.
pub(super) fn unchanged_write(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let presentation = owner(source);
    let package = presentation.materialize().expect("materialize");
    package.write_to(sink).expect("write");
}

/// The 0358a baseline publication route: index the source and write the old
/// one-shot insertion plan. It has no Commit/Patch API to replay.
pub(super) fn append_commit_write(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let edit =
        litchi_odp::tail_append::Edit::from_read_at(source, "Appended title", "Appended body")
            .expect("edit");
    let options = SourceContentPublicationOptions::default();
    let plan = edit.plan(&options).expect("plan");
    plan.write_to(sink, options).expect("write");
}

/// Closest pre-Patch equivalent: repeat the old independent publication plan
/// once. The record labels this explicitly because e4a0614f5d predates replay.
pub(super) fn replay(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let edit =
        litchi_odp::tail_append::Edit::from_read_at(source, "Appended title", "Appended body")
            .expect("edit");
    let options = SourceContentPublicationOptions::default();
    let plan = edit.plan(&options).expect("plan");
    plan.write_to(sink, options).expect("write");
}

pub(super) fn append_phases(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let baseline = super::phase_reset();
    let edit =
        litchi_odp::tail_append::Edit::from_read_at(source, "Appended title", "Appended body")
            .expect("edit");
    super::phase_report("append.open_stage", baseline);

    let options = SourceContentPublicationOptions::default();
    let baseline = super::phase_reset();
    let plan = edit.plan(&options).expect("plan");
    super::phase_report("append.plan", baseline);

    let baseline = super::phase_reset();
    plan.write_to(sink, options).expect("write");
    super::phase_report("append.write", baseline);
}

pub(super) fn replay_phases(source: Arc<dyn ReadAt>, sink: &mut super::Digest) {
    let edit =
        litchi_odp::tail_append::Edit::from_read_at(source, "Appended title", "Appended body")
            .expect("edit");
    let options = SourceContentPublicationOptions::default();
    let baseline = super::phase_reset();
    let plan = edit.plan(&options).expect("plan");
    super::phase_report("replay.plan", baseline);

    let baseline = super::phase_reset();
    plan.write_to(sink, options).expect("write");
    super::phase_report("replay.write", baseline);
}
