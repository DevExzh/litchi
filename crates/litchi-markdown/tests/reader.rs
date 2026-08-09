use litchi_markdown::reader::{
    BlockKind, Dialect, Error, History, HistoryLimits, InlineKind, JoinError, Patch,
    PatchEnvelopeLimits, ProjectionCapabilities, ProjectionIssueKind, ReadLimits, ReferenceKind,
    Snapshot,
};

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
    let descendants: Vec<_> = snapshot
        .block(0)
        .ok_or(Error::BlockNotFound { position: 0 })?
        .descendants()
        .collect();
    assert!(
        descendants
            .iter()
            .any(|block| block.kind() == BlockKind::ListItem)
    );
    assert!(
        descendants
            .iter()
            .any(|block| block.kind() == BlockKind::BlockQuote)
    );
    assert!(descendants.iter().any(|block| block.depth() >= 3));
    for block in descendants {
        assert_eq!(&source[block.range()], block.source());
    }
    Ok(())
}

#[test]
fn nested_and_inline_structural_edits_are_exact_and_durable() -> Result<(), Error> {
    let source = Snapshot::read("> old *emphasis*\n\noutside **strong**")?;
    let quote = source
        .block(0)
        .ok_or(Error::BlockNotFound { position: 0 })?;
    let paragraph_position = quote
        .descendants()
        .position(|block| block.kind() == BlockKind::Paragraph)
        .ok_or(Error::NestedBlockNotFound {
            block_position: 0,
            nested_position: 0,
        })?;
    let strong_position = source
        .block(1)
        .ok_or(Error::BlockNotFound { position: 1 })?
        .inlines()
        .position(|inline| inline.kind() == InlineKind::Strong)
        .ok_or(Error::InlineNotFound {
            block_position: 1,
            inline_position: 0,
        })?;
    let mut edit = source.edit();
    edit.replace_nested_block(0, paragraph_position, "new `code`")?
        .replace_inline(1, strong_position, "_gentle_")?;
    let commit = edit.commit()?;
    assert_eq!(
        commit.snapshot().source(),
        "> new `code`\n\noutside _gentle_"
    );
    let json = commit.patch().to_json(PatchEnvelopeLimits::DEFAULT)?;
    let restored = Patch::from_json(&json, PatchEnvelopeLimits::DEFAULT)?;
    assert_eq!(source.apply(&restored)?.snapshot(), commit.snapshot());
    Ok(())
}

#[test]
fn dependency_aware_transfer_and_projection_are_preflighted() -> Result<(), Error> {
    let source = Snapshot::read_with(
        "[linked][id] and a note[^n]\n\n[id]: /target \"title\"\n\n[^n]: note body",
        Dialect::GitHubFlavored,
        ReadLimits::DEFAULT,
    )?;
    let destination =
        Snapshot::read_with("Destination", Dialect::GitHubFlavored, ReadLimits::DEFAULT)?;
    let plan = source.preflight_transfer_block(0, &destination)?;
    assert_eq!(plan.dependency_count(), 2);
    assert_eq!(destination.source(), "Destination");
    let transferred = plan.into_commit();
    assert!(transferred.snapshot().references().any(|reference| {
        reference.kind() == ReferenceKind::LinkDefinition && reference.label() == Some("id")
    }));
    assert!(transferred.snapshot().references().any(|reference| {
        reference.kind() == ReferenceKind::FootnoteDefinition && reference.label() == Some("n")
    }));

    let feature_source = Snapshot::read_with(
        "| a |\n| - |\n| b |\n\n- [x] done\n\nraw <i>x</i>[^n]\n\n[^n]: note",
        Dialect::GitHubFlavored,
        ReadLimits::DEFAULT,
    )?;
    let report = feature_source.preflight_projection(ProjectionCapabilities::default())?;
    assert!(!report.is_lossless());
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.kind() == ProjectionIssueKind::Table)
    );
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.kind() == ProjectionIssueKind::TaskList)
    );
    assert!(
        feature_source
            .preflight_projection(ProjectionCapabilities::LOSSLESS)?
            .is_lossless()
    );
    Ok(())
}

#[test]
fn checked_transfer_refuses_conflicting_destination_definition() -> Result<(), Error> {
    let source = Snapshot::read("[use][id]\n\n[id]: /source")?;
    let destination = Snapshot::read("[id]: /destination")?;
    assert!(matches!(
        source.preflight_transfer_block(0, &destination),
        Err(Error::TransferDependencyConflict { ref label }) if label == "id"
    ));
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
fn inline_ast_and_reference_graph_retain_exact_ranges() -> Result<(), Error> {
    let source = "Text *em **strong** [link][id]* and `code` ![alt](img.png \"t\") <b>x</b>.[^n]\n\n[id]: /target \"title\"\n\n[^n]: note";
    let snapshot = Snapshot::read_with(source, Dialect::GitHubFlavored, ReadLimits::DEFAULT)?;
    let paragraph = snapshot
        .block(0)
        .ok_or(Error::BlockNotFound { position: 0 })?;
    let inlines: Vec<_> = paragraph.inlines().collect();
    assert!(
        inlines
            .iter()
            .any(|inline| inline.kind() == InlineKind::Emphasis)
    );
    assert!(
        inlines
            .iter()
            .any(|inline| inline.kind() == InlineKind::Strong)
    );
    assert!(
        inlines
            .iter()
            .any(|inline| inline.kind() == InlineKind::Link)
    );
    assert!(
        inlines
            .iter()
            .any(|inline| inline.kind() == InlineKind::Image)
    );
    assert!(
        inlines
            .iter()
            .any(|inline| inline.kind() == InlineKind::Code)
    );
    assert!(
        inlines
            .iter()
            .any(|inline| inline.kind() == InlineKind::Html)
    );
    for inline in inlines {
        assert_eq!(&source[inline.range()], inline.source());
    }

    let references: Vec<_> = snapshot.references().collect();
    assert_eq!(references.len(), 5);
    assert!(references.iter().any(|reference| {
        reference.kind() == ReferenceKind::Link
            && reference.label() == Some("id")
            && reference.destination() == Some("/target")
    }));
    assert!(references.iter().any(|reference| {
        reference.kind() == ReferenceKind::Image
            && reference.destination() == Some("img.png")
            && reference.title() == Some("t")
    }));
    assert!(references.iter().any(|reference| {
        reference.kind() == ReferenceKind::FootnoteDefinition && reference.label() == Some("n")
    }));
    Ok(())
}

#[test]
fn multi_operation_edits_are_atomic_ordered_and_overlap_checked() -> Result<(), Error> {
    let source = Snapshot::read("one\n\ntwo\n\nthree")?;
    let mut edit = source.edit();
    edit.replace_block(2, "THREE")?
        .replace_block(0, "ONE")?
        .append_block("four")?
        .append_block("five")?;
    let commit = edit.commit()?;
    assert_eq!(
        commit.snapshot().source(),
        "ONE\n\ntwo\n\nTHREE\n\nfour\n\nfive"
    );
    assert_eq!(commit.patch().operation_count(), 4);
    assert_eq!(commit.diagnostics().touched_blocks(), 4);

    let mut overlap = source.edit();
    overlap.replace_block(1, "changed")?;
    assert!(matches!(
        overlap.remove_block(1),
        Err(Error::OverlappingOperation { position: 1 })
    ));
    Ok(())
}

#[test]
fn referenced_definitions_require_dependency_closure() -> Result<(), Error> {
    let source = Snapshot::read("[use][id]\n\n[id]: /old")?;
    let mut blocked = source.edit();
    blocked.remove_block(1)?;
    assert!(matches!(
        blocked.commit(),
        Err(Error::ReferenceDependency { ref label }) if label == "id"
    ));

    let mut update = source.edit();
    update.replace_block(1, "[id]: /new")?;
    let updated = update.commit()?;
    assert!(updated.snapshot().references().any(|reference| {
        reference.kind() == ReferenceKind::Link && reference.destination() == Some("/new")
    }));

    let mut closure = source.edit();
    closure.remove_block(0)?.remove_block(1)?;
    assert_eq!(closure.commit()?.snapshot().source(), "\n");
    Ok(())
}

#[test]
fn bounded_history_is_commit_coupled_and_reversible() -> Result<(), Error> {
    let source = Snapshot::read("one")?;
    let mut history = History::new(
        source.clone(),
        HistoryLimits {
            max_entries: 2,
            max_patch_bytes: 1_024,
        },
    )?;
    let mut edit = source.edit();
    edit.replace_block(0, "two")?;
    history.apply(edit.commit()?)?;
    assert_eq!(history.current().source(), "two");
    assert!(history.undo()?);
    assert_eq!(history.current().source(), "one");
    assert!(history.redo()?);
    assert_eq!(history.current().source(), "two");

    let stale = Snapshot::read("stale")?;
    let mut stale_edit = stale.edit();
    stale_edit.replace_block(0, "other")?;
    assert!(matches!(
        history.apply(stale_edit.commit()?),
        Err(Error::PatchConflict)
    ));
    Ok(())
}

#[test]
fn durable_patch_json_is_deterministic_bounded_and_semantically_verified() -> Result<(), Error> {
    let source = Snapshot::read("one\n\ntwo")?;
    let mut edit = source.edit();
    edit.replace_block(0, "ONE")?.append_block("three")?;
    let commit = edit.commit()?;
    let limits = PatchEnvelopeLimits::DEFAULT;
    let first = commit.patch().to_json(limits)?;
    let second = commit.patch().to_json(limits)?;
    assert_eq!(first, second);
    let decoded = Patch::from_json(&first, limits)?;
    assert_eq!(
        source.apply(&decoded)?.snapshot().source(),
        "ONE\n\ntwo\n\nthree"
    );

    let inverse_json = commit.patch().inverse().to_json(limits)?;
    let inverse = Patch::from_json(&inverse_json, limits)?;
    assert_eq!(commit.snapshot().apply(&inverse)?.snapshot(), &source);

    let tampered = first.replace("\"replacement\":\"ONE\"", "\"replacement\":\"wrong\"");
    assert!(matches!(
        Patch::from_json(&tampered, limits),
        Err(Error::InvalidPatchEnvelope { .. })
    ));
    assert!(matches!(
        Patch::from_json(
            &first,
            PatchEnvelopeLimits {
                max_json_bytes: 8,
                ..limits
            }
        ),
        Err(Error::PatchEnvelopeTooLarge { .. })
    ));
    Ok(())
}

#[test]
fn independent_patch_join_reports_structured_conflicts() -> Result<(), Error> {
    let source = Snapshot::read("one\n\ntwo\n\nthree")?;
    let mut left_edit = source.edit();
    left_edit
        .replace_block(0, "ONE")?
        .append_block("left tail")?;
    let left = left_edit.commit()?;
    let mut right_edit = source.edit();
    right_edit
        .replace_block(2, "THREE")?
        .append_block("right tail")?;
    let right = right_edit.commit()?;
    let joined = left
        .patch()
        .join(right.patch())
        .map_err(|join_error| match join_error {
            JoinError::Validation(validation_error) => validation_error,
            JoinError::Conflicts(_) => Error::PatchConflict,
        })?;
    assert_eq!(
        joined.snapshot().source(),
        "ONE\n\ntwo\n\nTHREE\n\nleft tail\n\nright tail"
    );

    let mut conflict_edit = source.edit();
    conflict_edit.replace_block(0, "different")?;
    let conflict_patch = conflict_edit.commit()?;
    let conflicts = match left.patch().join(conflict_patch.patch()) {
        Err(JoinError::Conflicts(conflicts)) => conflicts,
        Err(JoinError::Validation(error)) => return Err(error),
        Ok(_) => return Err(Error::PatchConflict),
    };
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(conflicts.conflicts()[0].position(), 0);
    Ok(())
}

#[test]
fn three_way_merge_plan_never_mutates_its_base() -> Result<(), Error> {
    let base = Snapshot::read("a\n\nb\n\nc")?;
    let mut left_edit = base.edit();
    left_edit.replace_block(0, "A")?;
    let left = left_edit.commit()?;
    let mut right_edit = base.edit();
    right_edit.replace_block(2, "C")?;
    let right = right_edit.commit()?;
    let plan = left.patch().plan_merge(right.patch())?;
    assert!(plan.conflicts().is_empty());
    assert_eq!(
        plan.merged_commit()
            .map(|commit| commit.snapshot().source()),
        Some("A\n\nb\n\nC")
    );
    assert_eq!(base.source(), "a\n\nb\n\nc");

    let mut overlap_edit = base.edit();
    overlap_edit.replace_block(0, "other")?;
    let overlap = overlap_edit.commit()?;
    let conflict_plan = left.patch().plan_merge(overlap.patch())?;
    assert!(conflict_plan.merged_commit().is_none());
    assert_eq!(conflict_plan.conflicts().conflicts()[0].position(), 0);
    assert_eq!(base.source(), "a\n\nb\n\nc");
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
