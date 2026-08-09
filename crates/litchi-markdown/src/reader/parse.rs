use std::sync::Arc;

use pulldown_cmark::{CodeBlockKind, Event, HeadingLevel, Options, Parser, Tag, TagEnd};

use super::model::{BlockKind, BlockRecord, Dialect, Error, ReadLimits, Snapshot, State};

pub(crate) fn read(source: &str, dialect: Dialect, limits: ReadLimits) -> Result<Snapshot, Error> {
    validate_input(source, limits)?;
    let options = parser_options(dialect);
    let parser = Parser::new_ext(source, options);
    let mut blocks = collect_reference_definitions(&parser, limits)?;
    collect_event_blocks(parser, &mut blocks, limits)?;
    for block in &mut blocks {
        expand_indented_start(source, block);
    }
    blocks.sort_unstable_by(|left, right| {
        left.range
            .start
            .cmp(&right.range.start)
            .then_with(|| left.range.end.cmp(&right.range.end))
    });
    if blocks.len() > limits.max_blocks {
        return Err(Error::BlockLimitExceeded {
            limit: limits.max_blocks,
        });
    }
    let retained_source = Arc::<str>::from(source);
    Ok(Snapshot {
        state: Arc::new(State {
            source: retained_source,
            blocks: blocks.into_boxed_slice(),
            dialect,
            limits,
        }),
    })
}

fn expand_indented_start(source: &str, block: &mut BlockRecord) {
    let line_start = source[..block.range.start]
        .rfind(['\r', '\n'])
        .map_or(0, |offset| offset.saturating_add(1));
    if source[line_start..block.range.start]
        .bytes()
        .all(|byte| matches!(byte, b' ' | b'\t'))
    {
        block.range.start = line_start;
    }
}

fn block_kind(tag: &Tag<'_>) -> Option<BlockKind> {
    match tag {
        Tag::Paragraph => Some(BlockKind::Paragraph),
        Tag::Heading { level, .. } => Some(BlockKind::Heading {
            level: heading_level(*level),
        }),
        Tag::BlockQuote(_) => Some(BlockKind::BlockQuote),
        Tag::CodeBlock(kind) => Some(BlockKind::CodeBlock {
            fenced: matches!(kind, CodeBlockKind::Fenced(_)),
        }),
        Tag::HtmlBlock => Some(BlockKind::Html),
        Tag::List(start) => Some(BlockKind::List { start: *start }),
        Tag::FootnoteDefinition(_) => Some(BlockKind::FootnoteDefinition),
        Tag::Table(_) => Some(BlockKind::Table),
        Tag::Item
        | Tag::DefinitionList
        | Tag::DefinitionListTitle
        | Tag::DefinitionListDefinition
        | Tag::TableHead
        | Tag::TableRow
        | Tag::TableCell
        | Tag::Emphasis
        | Tag::Strong
        | Tag::Strikethrough
        | Tag::Superscript
        | Tag::Subscript
        | Tag::Link { .. }
        | Tag::Image { .. }
        | Tag::MetadataBlock(_) => None,
    }
}

fn collect_event_blocks(
    parser: Parser<'_>,
    blocks: &mut Vec<BlockRecord>,
    limits: ReadLimits,
) -> Result<(), Error> {
    let mut event_count = 0usize;
    let mut tag_depth = 0usize;
    let mut block_depth = 0usize;
    let mut root_block: Option<(BlockKind, usize)> = None;
    for (event, range) in parser.into_offset_iter() {
        event_count = event_count.saturating_add(1);
        if event_count > limits.max_events {
            return Err(Error::EventLimitExceeded {
                limit: limits.max_events,
            });
        }
        match event {
            Event::Start(tag) => {
                tag_depth = tag_depth.saturating_add(1);
                if tag_depth > limits.max_nesting_depth {
                    return Err(Error::NestingLimitExceeded {
                        limit: limits.max_nesting_depth,
                        offset: range.start,
                    });
                }
                if let Some(kind) = block_kind(&tag) {
                    if block_depth == 0 {
                        root_block = Some((kind, range.start));
                    }
                    block_depth = block_depth.saturating_add(1);
                }
            },
            Event::End(end) => {
                if is_block_end(end) {
                    block_depth = block_depth.saturating_sub(1);
                    if block_depth == 0
                        && let Some((kind, start)) = root_block.take()
                    {
                        push_block(
                            blocks,
                            BlockRecord {
                                kind,
                                range: start..range.end,
                            },
                            limits,
                        )?;
                    }
                }
                tag_depth = tag_depth.saturating_sub(1);
            },
            Event::Rule if block_depth == 0 => push_block(
                blocks,
                BlockRecord {
                    kind: BlockKind::ThematicBreak,
                    range,
                },
                limits,
            )?,
            Event::Text(_)
            | Event::Code(_)
            | Event::InlineMath(_)
            | Event::DisplayMath(_)
            | Event::Html(_)
            | Event::InlineHtml(_)
            | Event::FootnoteReference(_)
            | Event::SoftBreak
            | Event::HardBreak
            | Event::Rule
            | Event::TaskListMarker(_) => {},
        }
    }
    Ok(())
}

fn collect_reference_definitions(
    parser: &Parser<'_>,
    limits: ReadLimits,
) -> Result<Vec<BlockRecord>, Error> {
    let definitions = parser.reference_definitions();
    let count = definitions.iter().count();
    if count > limits.max_blocks {
        return Err(Error::BlockLimitExceeded {
            limit: limits.max_blocks,
        });
    }
    let mut blocks = Vec::new();
    blocks
        .try_reserve_exact(count)
        .map_err(|allocation_error| Error::Allocation {
            resource: "Markdown block index",
            source: allocation_error,
        })?;
    blocks.extend(definitions.iter().map(|(_, definition)| BlockRecord {
        kind: BlockKind::LinkDefinition,
        range: definition.span.clone(),
    }));
    Ok(blocks)
}

const fn heading_level(level: HeadingLevel) -> u8 {
    match level {
        HeadingLevel::H1 => 1,
        HeadingLevel::H2 => 2,
        HeadingLevel::H3 => 3,
        HeadingLevel::H4 => 4,
        HeadingLevel::H5 => 5,
        HeadingLevel::H6 => 6,
    }
}

const fn is_block_end(end: TagEnd) -> bool {
    matches!(
        end,
        TagEnd::Paragraph
            | TagEnd::Heading(_)
            | TagEnd::BlockQuote(_)
            | TagEnd::CodeBlock
            | TagEnd::HtmlBlock
            | TagEnd::List(_)
            | TagEnd::FootnoteDefinition
            | TagEnd::Table
    )
}

const fn parser_options(dialect: Dialect) -> Options {
    match dialect {
        Dialect::CommonMark => Options::empty(),
        Dialect::GitHubFlavored => Options::ENABLE_TABLES
            .union(Options::ENABLE_FOOTNOTES)
            .union(Options::ENABLE_STRIKETHROUGH)
            .union(Options::ENABLE_TASKLISTS)
            .union(Options::ENABLE_GFM),
    }
}

fn push_block(
    blocks: &mut Vec<BlockRecord>,
    block: BlockRecord,
    limits: ReadLimits,
) -> Result<(), Error> {
    if blocks.len() == limits.max_blocks {
        return Err(Error::BlockLimitExceeded {
            limit: limits.max_blocks,
        });
    }
    blocks
        .try_reserve(1)
        .map_err(|allocation_error| Error::Allocation {
            resource: "Markdown block index",
            source: allocation_error,
        })?;
    blocks.push(block);
    Ok(())
}

fn validate_input(source: &str, limits: ReadLimits) -> Result<(), Error> {
    limits.validate()?;
    if source.len() > limits.max_source_bytes {
        return Err(Error::SourceTooLarge {
            actual: source.len(),
            limit: limits.max_source_bytes,
        });
    }
    if let Some(offset) = source.as_bytes().iter().position(|byte| *byte == 0) {
        return Err(Error::NullByte { offset });
    }
    for (line_index, physical_line) in source.split('\n').enumerate() {
        let content = physical_line.strip_suffix('\r').unwrap_or(physical_line);
        if content.len() > limits.max_line_bytes {
            return Err(Error::LineTooLong {
                line: line_index.saturating_add(1),
                actual: content.len(),
                limit: limits.max_line_bytes,
            });
        }
    }
    Ok(())
}
