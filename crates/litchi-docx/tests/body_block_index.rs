//! The body block index (`/body/p[N]` and its siblings) over the repository's
//! DOCX corpus: both owners agree, the index agrees with the ordered block
//! view and with the document transaction's positions, repeated views reuse
//! the memoized projection and index, and paragraph identifiers read from the
//! start tag equal the full extension parse.

use std::path::{Path, PathBuf};
use std::sync::Arc;

use litchi_core::{OwnedSource, Position};
use litchi_docx::{Block, BlockKind, Blocks, IndexedBlock, Package, source_backed};

fn corpus() -> Vec<PathBuf> {
    fn walk(directory: &Path, files: &mut Vec<PathBuf>) {
        let Ok(entries) = std::fs::read_dir(directory) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                walk(&path, files);
            } else if path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| matches!(extension, "docx" | "docm" | "dotx" | "dotm"))
            {
                files.push(path);
            }
        }
    }
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
    let mut files = Vec::new();
    walk(&root, &mut files);
    files.sort();
    files
}

/// A comparable description of one indexed block.
fn describe(block: &IndexedBlock) -> String {
    format!(
        "{:?}#{}@{}:{}:{:?}:{}",
        block.kind(),
        block.ordinal().get(),
        block.position().get(),
        block.is_wrapped(),
        block.direct_position().map(Position::get),
        String::from_utf8_lossy(block.xml_bytes())
    )
}

fn signature(blocks: &Blocks) -> Vec<String> {
    blocks.iter().map(|block| describe(&block)).collect()
}

/// A block of the ordered view, described like the index describes it.
fn block_text(block: &Block) -> String {
    match block {
        Block::Paragraph(paragraph) => format!("p:{}", paragraph.text().unwrap()),
        Block::Table(table) => format!("t:{}", table.row_count().unwrap()),
        Block::Sdt(control) => format!("s:{}", control.text().unwrap()),
        Block::Alt(chunk) => format!("a:{}", chunk.relationship().as_str()),
        Block::Unknown(value) => format!("u:{}", value.xml_bytes().len()),
    }
}

fn indexed_text(block: &IndexedBlock) -> String {
    block_text(&block.to_block().unwrap())
}

struct Totals {
    files: usize,
    admitted: usize,
    blocks: usize,
    wrapped: usize,
    equations: usize,
    controls: usize,
    identifiers: usize,
}

fn check_transaction_positions(
    path: &Path,
    blocks: &Blocks,
    snapshot: &litchi_docx::document::Snapshot,
) {
    assert!(blocks.raw_positions_proven(), "{}", path.display());
    let mut paragraphs = 0;
    let mut tables = 0;
    let mut controls = 0;
    for block in blocks.iter() {
        let Ok(position) = block.transaction_position() else {
            assert!(block.is_wrapped() || block.direct_position().map(Position::get).is_none());
            continue;
        };
        if let Some(paragraph) = block.paragraph() {
            assert_eq!(position.get(), paragraphs, "{}", path.display());
            paragraphs += 1;
            assert_eq!(
                snapshot.paragraph(position).unwrap().text().unwrap(),
                paragraph.text().unwrap(),
                "{} paragraph {}",
                path.display(),
                position.get()
            );
        } else if block.kind() == BlockKind::Table {
            assert_eq!(position.get(), tables);
            tables += 1;
        } else {
            assert_eq!(block.kind(), BlockKind::ContentControl);
            assert_eq!(position.get(), controls);
            controls += 1;
        }
    }
    assert_eq!(paragraphs, snapshot.paragraph_count(), "{}", path.display());
    assert_eq!(tables, snapshot.table_count(), "{}", path.display());
    assert_eq!(
        controls,
        snapshot.block_content_control_count(),
        "{}",
        path.display()
    );
    assert!(snapshot.paragraph(Position::new(paragraphs)).is_none());
}

fn check_identifiers(blocks: &Blocks, totals: &mut Totals) {
    for block in blocks.iter() {
        let Some(paragraph) = block.paragraph() else {
            continue;
        };
        match (
            paragraph.extensions(),
            litchi_docx::document::raw::paragraph_ids(&paragraph),
        ) {
            (Ok(extensions), Ok(ids)) => {
                assert_eq!(extensions.ids(), ids);
                if let Some(id) = ids.para_id() {
                    totals.identifiers += 1;
                    let found = litchi_docx::document::raw::by_para_id(blocks, id)
                        .unwrap()
                        .unwrap();
                    assert!(found.position().get() <= block.position().get());
                    assert_eq!(found.paragraph().unwrap().para_id().unwrap(), Some(id));
                }
            },
            (Err(_), Err(_)) => {},
            (extensions, ids) => panic!(
                "start-tag identifiers disagree with extensions: {:?} vs {:?}",
                extensions.map(|value| value.ids()),
                ids
            ),
        }
    }
}

#[test]
fn body_index_agrees_with_both_owners_the_block_view_and_the_transaction() {
    let mut totals = Totals {
        files: 0,
        admitted: 0,
        blocks: 0,
        wrapped: 0,
        equations: 0,
        controls: 0,
        identifiers: 0,
    };
    for path in corpus() {
        let bytes = std::fs::read(&path).unwrap();
        let Ok(eager) = Package::from_reader(std::io::Cursor::new(bytes.clone())) else {
            continue;
        };
        let Ok(source) = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes)))
        else {
            continue;
        };
        let (Ok(eager_document), Ok(source_document)) = (eager.document(), source.document())
        else {
            continue;
        };
        let Ok(eager_blocks) = eager_document.body_blocks() else {
            // The ordered block view refuses the same documents.
            assert!(eager_document.blocks().is_err(), "{}", path.display());
            continue;
        };
        totals.files += 1;
        let source_blocks = source_document.body_blocks().unwrap();
        assert_eq!(
            signature(&eager_blocks),
            signature(&source_blocks),
            "{}",
            path.display()
        );
        assert_eq!(
            eager_blocks.raw_positions_proven(),
            source_blocks.raw_positions_proven()
        );
        totals.blocks += eager_blocks.len();
        totals.wrapped += eager_blocks.iter().filter(IndexedBlock::is_wrapped).count();
        totals.equations += eager_blocks.count(BlockKind::Equation);
        totals.controls += eager_blocks.count(BlockKind::ContentControl);

        // The ordered block view lists the same children whenever no
        // wrapper is flattened.
        if eager_blocks.iter().all(|block| !block.is_wrapped()) {
            let view: Vec<String> = eager_document
                .blocks()
                .unwrap()
                .iter()
                .map(block_text)
                .collect();
            let indexed: Vec<String> = eager_blocks
                .iter()
                .map(|block| indexed_text(&block))
                .collect();
            assert_eq!(view, indexed, "{}", path.display());
        }

        // Paragraph accessors agree with the index.
        for ordinal in 0..eager_blocks.count(BlockKind::Paragraph) {
            assert_eq!(
                eager_document
                    .body_paragraph(Position::new(ordinal))
                    .unwrap()
                    .unwrap()
                    .text()
                    .unwrap(),
                source_document
                    .body_paragraph(Position::new(ordinal))
                    .unwrap()
                    .unwrap()
                    .text()
                    .unwrap()
            );
        }
        assert!(
            eager_document
                .body_paragraph(Position::new(eager_blocks.count(BlockKind::Paragraph)))
                .unwrap()
                .is_none()
        );

        // Positions agree with the transaction wherever it admits the part.
        if let Ok(snapshot) = eager.document_snapshot() {
            totals.admitted += 1;
            check_transaction_positions(&path, &eager_blocks, &snapshot);
            let source_snapshot = source.document_snapshot().unwrap();
            check_transaction_positions(&path, &source_blocks, &source_snapshot);
        }
        check_identifiers(&eager_blocks, &mut totals);

        // A second view reuses the projection and the index.
        let again = source.document().unwrap();
        assert_eq!(
            again.extract_text().unwrap(),
            source_document.extract_text().unwrap()
        );
        assert_eq!(
            signature(&again.body_blocks().unwrap()),
            signature(&source_blocks)
        );
        let eager_again = eager.document().unwrap();
        assert_eq!(eager_again.text().unwrap(), eager_document.text().unwrap());
        assert_eq!(
            eager_again.part().xml_bytes(),
            eager_document.part().xml_bytes()
        );
    }
    println!(
        "files={} admitted={} blocks={} wrapped={} equations={} controls={} para_ids={}",
        totals.files,
        totals.admitted,
        totals.blocks,
        totals.wrapped,
        totals.equations,
        totals.controls,
        totals.identifiers
    );
    assert!(totals.files >= 50, "the corpus is present");
    assert!(totals.admitted >= 50);
    assert!(totals.identifiers > 0);
}

#[test]
fn body_content_controls_are_typed_and_their_content_is_indexed() {
    for path in corpus() {
        let bytes = std::fs::read(&path).unwrap();
        let Ok(package) = Package::from_reader(std::io::Cursor::new(bytes)) else {
            continue;
        };
        let Ok(document) = package.document() else {
            continue;
        };
        let Ok(blocks) = document.blocks() else {
            continue;
        };
        let Ok(body) = document.body_blocks() else {
            continue;
        };
        let typed = blocks
            .iter()
            .filter(|block| matches!(block, Block::Sdt(_)))
            .count();
        let unwrapped_controls = body
            .iter()
            .filter(|block| block.kind() == BlockKind::ContentControl && !block.is_wrapped())
            .count();
        assert_eq!(typed, unwrapped_controls, "{}", path.display());
        for ordinal in 0..body.count(BlockKind::ContentControl) {
            let control = document.body_sdt(Position::new(ordinal)).unwrap().unwrap();
            // The identity is read from the properties alone.
            litchi_docx::document::raw::content_control_id(&control).unwrap();
            control.tag().unwrap();
            control.alias().unwrap();
            let content = control.blocks().unwrap();
            assert_eq!(content.raw_positions_proven(), body.raw_positions_proven());
            for block in content.iter() {
                if let Some(paragraph) = block.paragraph() {
                    paragraph.text().unwrap();
                }
            }
        }
    }
}

#[test]
fn story_containers_index_their_direct_children_with_the_body_scanner() {
    let fixture = |name: &str| {
        Package::open(
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/ooxml/docx")
                .join(name),
        )
        .unwrap()
    };
    let paragraph_texts = |blocks: &Blocks| -> Vec<String> {
        (0..blocks.count(BlockKind::Paragraph))
            .map(|ordinal| {
                blocks
                    .paragraph(Position::new(ordinal))
                    .unwrap()
                    .text()
                    .unwrap()
            })
            .collect()
    };

    let package = fixture("footnotes.docx");
    let document = package.document().unwrap();
    let mut notes = 0;
    for note in document.footnotes().unwrap() {
        let blocks = note.blocks().unwrap();
        assert!(!blocks.raw_positions_proven());
        let direct: Vec<String> = paragraph_texts(&blocks);
        let legacy: Vec<String> = note
            .paragraphs()
            .unwrap()
            .iter()
            .map(|paragraph| paragraph.text().unwrap())
            .collect();
        // Notes in the fixture hold no tables, so both views agree.
        assert_eq!(direct, legacy);
        notes += 1;
    }
    assert!(notes > 0);

    let package = fixture("comment.docx");
    let document = package.document().unwrap();
    let comments = document.comments().unwrap();
    assert!(!comments.is_empty());
    for comment in comments {
        let blocks = comment.blocks().unwrap();
        assert_eq!(paragraph_texts(&blocks).concat(), comment.text().unwrap());
    }

    let package = fixture("headerFooter.docx");
    let document = package.document().unwrap();
    let mut stories = 0;
    for story in document
        .headers()
        .unwrap()
        .into_iter()
        .chain(document.footers().unwrap())
    {
        let blocks = story.blocks().unwrap();
        let legacy: Vec<String> = story
            .paragraphs()
            .unwrap()
            .iter()
            .map(|paragraph| paragraph.text().unwrap())
            .collect();
        assert_eq!(paragraph_texts(&blocks), legacy);
        stories += 1;
    }
    assert!(stories > 0);
}
