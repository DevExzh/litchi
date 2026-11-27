//! What `finish` publishes for each kind of staged change, and when a signed
//! source is refused.

use super::super::Editor;
use crate::package::Error;

const SIGNED_REFUSAL: &str = "signed PPT is not eligible for a changed OLE publication";

fn fixture() -> Vec<u8> {
    let mut writer = crate::writer::Builder::new();
    let slide = writer.add_slide().unwrap();
    writer.add_textbox(slide, 10, 10, 240, 40, "text").unwrap();
    let mut output = std::io::Cursor::new(Vec::new());
    crate::raw::encode(&mut writer, &mut output).unwrap();
    output.into_inner()
}

fn open() -> Editor {
    Editor::open_records(fixture(), crate::SourceLimits::default()).unwrap()
}

/// A live record other than the document record.
fn other_persist_id(editor: &Editor) -> u32 {
    editor
        .persist_ids()
        .into_iter()
        .find(|id| *id != editor.document_persist_id)
        .unwrap()
}

fn restage_record(editor: &mut Editor, persist_id: u32) -> usize {
    let record = editor.persisted_record(persist_id).unwrap();
    let len = record.len();
    editor.replace_persisted_record(persist_id, record).unwrap();
    len
}

fn reopen(published: Vec<u8>) -> Editor {
    Editor::open_records(published, crate::SourceLimits::default()).unwrap()
}

fn assert_signed_refusal(editor: Editor) {
    match editor.finish() {
        Err(Error::Corrupted(message)) => assert_eq!(message, SIGNED_REFUSAL),
        other => panic!("expected the signed refusal, got {other:?}"),
    }
}

#[test]
fn an_unchanged_editor_publishes_the_source_bytes_signed_or_not() {
    let source = fixture();
    let editor = open();
    assert!(!editor.has_net_change(editor.objects()).unwrap());
    assert_eq!(editor.finish().unwrap(), source);

    let mut signed = open();
    signed.signed = true;
    assert_eq!(signed.finish().unwrap(), source);
}

#[test]
fn a_record_change_appends_the_record_and_keeps_the_document_record() {
    let source = open();
    let document_id = source.document_persist_id;
    let other = other_persist_id(&source);
    let source_len = u32::try_from(source.document.len()).unwrap();

    let mut editor = open();
    restage_record(&mut editor, other);
    let published = reopen(editor.finish().unwrap());
    assert_eq!(published.mappings[&other], source_len);
    assert_eq!(
        published.mappings[&document_id],
        source.mappings[&document_id]
    );
    assert_eq!(
        published.persisted_record(document_id).unwrap(),
        source.persisted_record(document_id).unwrap()
    );
}

#[test]
fn an_object_list_change_appends_a_rewritten_document_record() {
    let source = open();
    let document_id = source.document_persist_id;
    let other = other_persist_id(&source);
    let source_len = u32::try_from(source.document.len()).unwrap();

    let mut editor = open();
    editor.reorder(&[]).unwrap();
    let published = reopen(editor.finish().unwrap());
    assert_eq!(published.mappings[&document_id], source_len);
    assert_eq!(published.mappings[&other], source.mappings[&other]);
    let rewritten = published.persisted_record(document_id).unwrap();
    assert_ne!(rewritten, source.persisted_record(document_id).unwrap());

    // The document record is rewritten after the staged records, whichever
    // was staged first.
    let mut record_first = open();
    let record_len = u32::try_from(restage_record(&mut record_first, other)).unwrap();
    record_first.reorder(&[]).unwrap();
    let mut list_first = open();
    list_first.reorder(&[]).unwrap();
    restage_record(&mut list_first, other);
    let record_first = record_first.finish().unwrap();
    assert_eq!(record_first, list_first.finish().unwrap());
    let published = reopen(record_first);
    assert_eq!(published.mappings[&other], source_len);
    assert_eq!(published.mappings[&document_id], source_len + record_len);
    assert_eq!(published.persisted_record(document_id).unwrap(), rewritten);
}

#[test]
fn a_stream_replacement_is_published_without_a_net_object_change() {
    let source = fixture();
    let mut editor = open();
    editor
        .replace_stream_by_name("Extra", vec![1, 2, 3])
        .unwrap();
    assert!(!editor.has_net_change(editor.objects()).unwrap());
    let document_id = editor.document_persist_id;
    let document_offset = editor.mappings[&document_id];
    let published = editor.finish().unwrap();
    assert_ne!(published, source);
    let published = reopen(published);
    assert_eq!(published.mappings[&document_id], document_offset);
    assert!(
        published
            .streams
            .iter()
            .any(|(path, data)| path == &["Extra".to_string()] && data == &[1, 2, 3])
    );
}

#[test]
fn net_change_follows_the_staged_content_not_the_staging_calls() {
    let mut editor = open();
    let source_objects = editor.objects().clone();
    let other = other_persist_id(&editor);
    editor.reorder(&[]).unwrap();
    assert!(!editor.has_net_change(&source_objects).unwrap());
    let record = editor.persisted_record(other).unwrap();
    editor
        .replace_persisted_record(other, record.clone())
        .unwrap();
    assert!(!editor.has_net_change(&source_objects).unwrap());
    let mut changed = record;
    let last = changed.last_mut().unwrap();
    *last ^= 0xFF;
    editor.replace_persisted_record(other, changed).unwrap();
    assert!(editor.has_net_change(&source_objects).unwrap());
}

#[test]
fn a_signed_source_stages_every_change_and_is_refused_at_finish() {
    let mut records = open();
    records.signed = true;
    let other = other_persist_id(&records);
    restage_record(&mut records, other);
    assert!(records.is_record_staged(other));
    assert_signed_refusal(records);

    let mut list = open();
    list.signed = true;
    list.reorder(&[]).unwrap();
    assert_signed_refusal(list);

    let mut both = open();
    both.signed = true;
    restage_record(&mut both, other);
    both.reorder(&[]).unwrap();
    assert_signed_refusal(both);

    let mut stream = open();
    stream.signed = true;
    stream
        .replace_stream_by_name("Extra", vec![1, 2, 3])
        .unwrap();
    assert_signed_refusal(stream);
}

#[test]
fn unstaging_a_record_keeps_the_editor_changed() {
    let source = fixture();
    let mut editor = open();
    let other = other_persist_id(&editor);
    restage_record(&mut editor, other);
    editor.unstage_record(other);
    assert!(!editor.has_staged_records());
    let document_id = editor.document_persist_id;
    let mappings = editor.mappings.clone();
    let published = editor.finish().unwrap();
    assert_ne!(published, source);
    let published = reopen(published);
    assert_eq!(published.mappings, mappings);
    assert_eq!(
        published.persisted_record(document_id).unwrap(),
        open().persisted_record(document_id).unwrap()
    );

    let mut signed = open();
    signed.signed = true;
    restage_record(&mut signed, other);
    signed.unstage_record(other);
    assert_signed_refusal(signed);
}

#[test]
fn a_refused_staging_call_leaves_the_editor_unchanged() {
    let source = fixture();
    let mut editor = open();
    let other = other_persist_id(&editor);
    assert!(editor.reorder(&[7]).is_err());
    assert!(editor.replace_persisted_record(other, vec![0; 4]).is_err());
    assert!(editor.insert_persisted_record(0, vec![0; 8]).is_err());
    assert_eq!(editor.finish().unwrap(), source);
}
