use litchi_markdown::reader::{BlockKind, Dialect, Error, ReadLimits, Snapshot};

#[test]
fn commonmark_block_corpus_is_classified_in_source_order() -> Result<(), Error> {
    let corpus = [
        ("plain *paragraph*", BlockKind::Paragraph),
        ("# atx", BlockKind::Heading { level: 1 }),
        ("setext\n------", BlockKind::Heading { level: 2 }),
        ("> quote\n> continuation", BlockKind::BlockQuote),
        ("    indented code", BlockKind::CodeBlock { fenced: false }),
        (
            "```rust\nfn main() {}\n```",
            BlockKind::CodeBlock { fenced: true },
        ),
        (
            "~~~\ncode with ```\n~~~",
            BlockKind::CodeBlock { fenced: true },
        ),
        ("<table>\n<tr><td>x</td></tr>\n</table>", BlockKind::Html),
        ("- one\n- two", BlockKind::List { start: None }),
        ("42. one\n43. two", BlockKind::List { start: Some(42) }),
        ("***", BlockKind::ThematicBreak),
        (
            "[label]: <https://example.test/a b> \"title\"",
            BlockKind::LinkDefinition,
        ),
        ("Escaped \\# marker and `code`", BlockKind::Paragraph),
        ("Unicode: lychee 🍋 é", BlockKind::Paragraph),
    ];

    for (source, expected) in corpus {
        let snapshot = Snapshot::read(source)?;
        assert_eq!(snapshot.source(), source);
        assert_eq!(snapshot.blocks().len(), 1, "{source:?}");
        assert_eq!(snapshot.block(0).map(|block| block.kind()), Some(expected));
        assert_eq!(snapshot.block(0).map(|block| block.source()), Some(source));
    }
    Ok(())
}

#[test]
fn exact_source_and_nonoverlapping_ranges_survive_mixed_line_endings() -> Result<(), Error> {
    let source = "# Head\r\n\r\nParagraph  🍋\n\n> quote\r\n\r\n[ref]: /target\r\n";
    let snapshot = Snapshot::read(source)?;
    assert_eq!(snapshot.source(), source);
    let blocks: Vec<_> = snapshot.blocks().collect();
    assert_eq!(blocks.len(), 4);
    assert_eq!(blocks[0].kind(), BlockKind::Heading { level: 1 });
    assert_eq!(blocks[1].kind(), BlockKind::Paragraph);
    assert_eq!(blocks[2].kind(), BlockKind::BlockQuote);
    assert_eq!(blocks[3].kind(), BlockKind::LinkDefinition);
    for pair in blocks.windows(2) {
        assert!(pair[0].range().end <= pair[1].range().start);
    }
    for block in blocks {
        assert_eq!(&source[block.range()], block.source());
    }
    Ok(())
}

#[test]
fn gfm_extensions_are_explicit_and_deterministic() -> Result<(), Error> {
    let table = "| a | b |\n| - | - |\n| x | y |";
    let commonmark = Snapshot::read(table)?;
    assert_eq!(
        commonmark.block(0).map(|block| block.kind()),
        Some(BlockKind::Paragraph)
    );

    let gfm = Snapshot::read_with(table, Dialect::GitHubFlavored, ReadLimits::DEFAULT)?;
    assert_eq!(
        gfm.block(0).map(|block| block.kind()),
        Some(BlockKind::Table)
    );

    let footnote = Snapshot::read_with(
        "body[^n]\n\n[^n]: note",
        Dialect::GitHubFlavored,
        ReadLimits::DEFAULT,
    )?;
    assert_eq!(footnote.blocks().len(), 2);
    assert_eq!(
        footnote.block(1).map(|block| block.kind()),
        Some(BlockKind::FootnoteDefinition)
    );
    Ok(())
}

#[test]
fn nested_containers_count_as_one_top_level_block() -> Result<(), Error> {
    let source = "1. first\n   - nested\n     > quoted\n     >\n     > ```\n     > code\n     > ```\n2. second";
    let snapshot = Snapshot::read(source)?;
    assert_eq!(snapshot.blocks().len(), 1);
    assert_eq!(
        snapshot.block(0).map(|block| block.kind()),
        Some(BlockKind::List { start: Some(1) })
    );
    Ok(())
}

#[test]
fn literal_edit_helpers_cannot_inject_markdown_blocks() -> Result<(), Error> {
    let source = Snapshot::read("original")?;
    let mut edit = source.edit();
    edit.replace_block_with_text(0, "# heading\n- list\n[link](target) <tag>")?;
    let commit = edit.commit()?;
    assert_eq!(
        commit.snapshot().source(),
        "\\# heading\n\\- list\n\\[link\\](target) \\<tag\\>"
    );
    assert_eq!(commit.snapshot().blocks().len(), 1);
    assert_eq!(
        commit.snapshot().block(0).map(|block| block.kind()),
        Some(BlockKind::Paragraph)
    );
    Ok(())
}

#[test]
fn reader_refuses_lossy_and_over_budget_inputs() {
    assert!(matches!(Snapshot::read_bytes(&[0xff]), Err(Error::Utf8(_))));
    assert!(matches!(
        Snapshot::read("a\0b"),
        Err(Error::NullByte { offset: 1 })
    ));

    let source_limit = ReadLimits {
        max_source_bytes: 3,
        ..ReadLimits::DEFAULT
    };
    assert!(matches!(
        Snapshot::read_with("four", Dialect::CommonMark, source_limit),
        Err(Error::SourceTooLarge {
            actual: 4,
            limit: 3
        })
    ));

    let line_limit = ReadLimits {
        max_line_bytes: 3,
        ..ReadLimits::DEFAULT
    };
    assert!(matches!(
        Snapshot::read_with("ok\nfour", Dialect::CommonMark, line_limit),
        Err(Error::LineTooLong {
            line: 2,
            actual: 4,
            limit: 3
        })
    ));

    let block_limit = ReadLimits {
        max_blocks: 1,
        ..ReadLimits::DEFAULT
    };
    assert!(matches!(
        Snapshot::read_with("one\n\ntwo", Dialect::CommonMark, block_limit),
        Err(Error::BlockLimitExceeded { limit: 1 })
    ));

    let event_limit = ReadLimits {
        max_events: 2,
        ..ReadLimits::DEFAULT
    };
    assert!(matches!(
        Snapshot::read_with("one", Dialect::CommonMark, event_limit),
        Err(Error::EventLimitExceeded { limit: 2 })
    ));

    let depth_limit = ReadLimits {
        max_nesting_depth: 2,
        ..ReadLimits::DEFAULT
    };
    assert!(matches!(
        Snapshot::read_with("- ***deep***", Dialect::CommonMark, depth_limit),
        Err(Error::NestingLimitExceeded { limit: 2, .. })
    ));

    let invalid_limits = ReadLimits {
        max_events: 0,
        ..ReadLimits::DEFAULT
    };
    assert!(matches!(
        Snapshot::read_with("", Dialect::CommonMark, invalid_limits),
        Err(Error::InvalidLimit { name: "max_events" })
    ));
}

#[test]
fn replacement_commit_preserves_untouched_bytes_and_is_reversible() -> Result<(), Error> {
    let source = Snapshot::read("# Original\r\n\r\nKeep  \r\n\r\n- one\r\n")?;
    let mut edit = source.edit();
    edit.replace_block(0, "## Replacement")?;
    let commit = edit.commit()?;
    assert_eq!(
        commit.snapshot().source(),
        "## Replacement\r\n\r\nKeep  \r\n\r\n- one\r\n"
    );
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(commit.diagnostics().touched_blocks(), 1);

    let restored = commit.snapshot().apply(&commit.patch().inverse())?;
    assert_eq!(restored.snapshot(), &source);
    assert_eq!(restored.snapshot().source(), source.source());
    Ok(())
}

#[test]
fn append_remove_noop_and_conflict_paths_are_typed() -> Result<(), Error> {
    let source = Snapshot::read("one")?;
    let mut append = source.edit();
    append.append_block("- two")?;
    let appended = append.commit()?;
    assert_eq!(appended.snapshot().source(), "one\n\n- two");

    let stale = Snapshot::read("different")?;
    assert!(matches!(
        stale.apply(appended.patch()),
        Err(Error::PatchConflict)
    ));

    let mut remove = appended.snapshot().edit();
    remove.remove_block(1)?;
    let removed = remove.commit()?;
    assert_eq!(removed.snapshot().source(), "one\n\n");

    let mut noop_edit = source.edit();
    noop_edit.replace_block(0, "one")?;
    let noop_commit = noop_edit.commit()?;
    assert!(noop_commit.patch().is_empty());
    assert!(!noop_commit.diagnostics().changed());
    assert!(!noop_commit.diagnostics().full_reparse_performed());

    let mut invalid = source.edit();
    assert!(matches!(
        invalid.append_block("one\n\ntwo"),
        Err(Error::ReplacementBlockCount { actual: 2 })
    ));
    assert!(matches!(
        source.edit().commit(),
        Err(Error::NoStagedOperation)
    ));
    Ok(())
}

#[test]
fn empty_and_large_deterministic_corpus_remains_bounded() -> Result<(), Error> {
    let empty = Snapshot::read("")?;
    assert!(empty.blocks().is_empty());

    let mut source = String::new();
    for index in 0..2_000 {
        use std::fmt::Write as _;
        writeln!(source, "## Heading {index}\n").map_err(|_format_error| {
            Error::SourceTooLarge {
                actual: usize::MAX,
                limit: ReadLimits::DEFAULT.max_source_bytes,
            }
        })?;
    }
    let first = Snapshot::read(&source)?;
    let second = Snapshot::read(&source)?;
    assert_eq!(first, second);
    assert_eq!(first.blocks().len(), 2_000);
    assert!(
        first
            .blocks()
            .all(|block| matches!(block.kind(), BlockKind::Heading { level: 2 }))
    );
    Ok(())
}

#[test]
fn snapshots_are_send_sync_and_cheap_to_clone() -> Result<(), Error> {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Snapshot>();

    let source = Snapshot::read("# shared")?;
    let cloned = source.clone();
    let joined = std::thread::spawn(move || cloned.source().to_owned())
        .join()
        .map_err(|_panic_payload| Error::PatchConflict)?;
    assert_eq!(joined, "# shared");
    assert_eq!(source.source(), "# shared");
    Ok(())
}
