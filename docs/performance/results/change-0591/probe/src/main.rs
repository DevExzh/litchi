//! Isolate the instruction cost of the DOCX edit-and-save *timed region*.
//!
//! `litchi-perf-baseline`'s `docx_semantic_*_edit_save` selectors time only
//! `edit_document` .. `to_stream`, but a callgrind run of the harness counts the
//! whole child, including corpus construction, the untimed open and the untimed
//! reopen/verify. This probe opens a constant number of packages *before* the
//! measured loop and then runs the timed region on `measured` of them, so a
//! callgrind isolation pair at `measured = 1` and `measured = 3` differences
//! exactly two timed regions and nothing else.
//!
//! usage: docx-edit-region <noop|one|one-percent> <paragraphs> <measured>
use std::hint::black_box;
use std::io::{Cursor, Seek, SeekFrom, Write};

use litchi_core::Position;
use litchi_docx::Package;
use litchi_docx::document::ParagraphTextReplacement;

const PREPARED: usize = 4;

#[derive(Default)]
struct CountingSeekSink {
    position: u64,
    accepted: u64,
}

impl Write for CountingSeekSink {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        self.position += buffer.len() as u64;
        self.accepted += buffer.len() as u64;
        Ok(buffer.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

impl Seek for CountingSeekSink {
    fn seek(&mut self, to: SeekFrom) -> std::io::Result<u64> {
        self.position = match to {
            SeekFrom::Start(offset) => offset,
            SeekFrom::Current(offset) => self.position.saturating_add_signed(offset),
            SeekFrom::End(offset) => self.accepted.saturating_add_signed(offset),
        };
        Ok(self.position)
    }
}

fn semantic_text(index: usize, updated: bool) -> String {
    let state = if updated { "updated" } else { "source" };
    format!("litchi-perf-baseline-docx-semantic-v1-{state}-{index:05}")
}

fn update_indices(count: usize) -> Vec<usize> {
    let updates = (count + 99) / 100;
    (0..updates).map(|index| index * count / updates).collect()
}

fn corpus(paragraphs: usize) -> Vec<u8> {
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        for index in 0..paragraphs {
            document.add_paragraph_with_text(&semantic_text(index, false));
        }
    }
    let mut output = Cursor::new(Vec::new());
    package.to_stream(&mut output).unwrap();
    output.into_inner()
}

fn main() {
    let arguments: Vec<String> = std::env::args().collect();
    let mode = arguments[1].as_str();
    let paragraphs: usize = arguments[2].parse().unwrap();
    let measured: usize = arguments[3].parse().unwrap();
    assert!(measured <= PREPARED);

    let archive = corpus(paragraphs);
    let updates = update_indices(paragraphs);
    let selected: Vec<usize> = match mode {
        "noop" => Vec::new(),
        "one" => vec![updates[0]],
        "one-percent" => updates,
        other => panic!("unknown mode {other}"),
    };
    let mut packages: Vec<Package> = (0..PREPARED)
        .map(|_| Package::from_reader(Cursor::new(archive.clone())).unwrap())
        .collect();

    let mut total = 0u64;
    for package in packages.iter_mut().take(measured) {
        // ---- exactly the harness's timed region ----
        let mut edit = package.edit_document().unwrap();
        if selected.len() > 1 {
            let replacements: Vec<ParagraphTextReplacement> = selected
                .iter()
                .map(|index| {
                    ParagraphTextReplacement::new(
                        Position::new(*index),
                        semantic_text(*index, true),
                    )
                })
                .collect();
            edit.replace_body_paragraph_texts(&replacements).unwrap();
        } else {
            for index in &selected {
                edit.replace_paragraph_text(Position::new(*index), semantic_text(*index, true))
                    .unwrap();
            }
        }
        let mut sink = CountingSeekSink::default();
        let commit = package.publish_document_edit(edit).unwrap();
        package.to_stream(&mut sink).unwrap();
        // ---- end of the timed region ----
        total += sink.accepted + u64::from(commit.diagnostics().changed());
    }
    black_box(total);
    println!("{mode} {paragraphs} measured={measured} accepted={total}");
}
