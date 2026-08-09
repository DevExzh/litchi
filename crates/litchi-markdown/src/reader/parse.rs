use std::sync::Arc;

use pulldown_cmark::{CodeBlockKind, Event, HeadingLevel, Options, Parser, Tag, TagEnd};

use super::model::{
    BlockKind, BlockRecord, Dialect, Error, InlineKind, InlineRecord, NestedBlockRecord,
    ReadLimits, ReferenceKind, ReferenceRecord, Snapshot, State,
};

struct OpenBlock {
    kind: BlockKind,
    start: usize,
    inlines: Vec<InlineRecord>,
    descendants: Vec<NestedBlockRecord>,
    footnote_label: Option<String>,
}

struct OpenNestedBlock {
    kind: BlockKind,
    start: usize,
    depth: usize,
}

struct OpenInline {
    kind: InlineKind,
    start: usize,
    depth: usize,
    reference: Option<ReferenceRecord>,
}

pub(crate) fn read(source: &str, dialect: Dialect, limits: ReadLimits) -> Result<Snapshot, Error> {
    validate_input(source, limits)?;
    let options = parser_options(dialect);
    let parser = Parser::new_ext(source, options);
    let (mut blocks, mut references) = collect_reference_definitions(&parser, limits)?;
    collect_event_blocks(parser, &mut blocks, &mut references, limits)?;
    for block in &mut blocks {
        expand_indented_start(source, block);
    }
    nest_contained_link_definitions(&mut blocks)?;
    blocks.sort_unstable_by(|left, right| {
        left.range
            .start
            .cmp(&right.range.start)
            .then_with(|| left.range.end.cmp(&right.range.end))
    });
    normalize_adjacent_block_whitespace(source, &mut blocks);
    references.sort_by(|left, right| {
        left.range
            .start
            .cmp(&right.range.start)
            .then_with(|| left.range.end.cmp(&right.range.end))
            .then_with(|| reference_order(left.kind).cmp(&reference_order(right.kind)))
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
            references: references.into_boxed_slice(),
            dialect,
            limits,
        }),
    })
}

fn normalize_adjacent_block_whitespace(source: &str, blocks: &mut [BlockRecord]) {
    for current_index in 1..blocks.len() {
        let current_start = blocks[current_index].range.start;
        let previous = &mut blocks[current_index.saturating_sub(1)];
        if previous.range.end <= current_start {
            continue;
        }
        let structural_end = previous
            .descendants
            .iter()
            .map(|nested| nested.range.end)
            .chain(previous.inlines.iter().map(|inline| inline.range.end))
            .max()
            .unwrap_or(previous.range.start);
        if structural_end <= current_start
            && source
                .get(current_start..previous.range.end)
                .is_some_and(|overlap| overlap.bytes().all(|byte| byte.is_ascii_whitespace()))
        {
            previous.range.end = current_start;
        }
    }
}

fn nest_contained_link_definitions(blocks: &mut Vec<BlockRecord>) -> Result<(), Error> {
    let original = std::mem::take(blocks);
    let mut roots = Vec::new();
    let mut definitions = Vec::new();
    roots
        .try_reserve_exact(original.len())
        .map_err(|source| Error::Allocation {
            resource: "Markdown top-level block normalization",
            source,
        })?;
    definitions
        .try_reserve_exact(original.len())
        .map_err(|source| Error::Allocation {
            resource: "Markdown nested link-definition normalization",
            source,
        })?;
    for block in original {
        if block.kind == BlockKind::LinkDefinition {
            definitions.push(block);
        } else {
            roots.push(block);
        }
    }

    for definition in definitions {
        let candidate_owner_index = roots
            .iter()
            .enumerate()
            .filter(|(_, block)| {
                block.range.start <= definition.range.start
                    && definition.range.end <= block.range.end
            })
            .min_by_key(|(_, block)| block.range.end.saturating_sub(block.range.start))
            .map(|(index, _)| index);
        let Some(owner_index) = candidate_owner_index else {
            roots.push(definition);
            continue;
        };

        let root = &mut roots[owner_index];
        let depth = root
            .descendants
            .iter()
            .filter(|nested| {
                nested.range.start <= definition.range.start
                    && definition.range.end <= nested.range.end
            })
            .map(|nested| nested.depth.saturating_add(1))
            .max()
            .unwrap_or(1);
        let mut descendants = std::mem::take(&mut root.descendants).into_vec();
        descendants
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "Markdown nested link-definition index",
                source,
            })?;
        descendants.push(NestedBlockRecord {
            kind: BlockKind::LinkDefinition,
            range: definition.range,
            depth,
        });
        descendants.sort_by(|left, right| {
            left.range
                .start
                .cmp(&right.range.start)
                .then_with(|| left.depth.cmp(&right.depth))
                .then_with(|| right.range.end.cmp(&left.range.end))
        });
        root.descendants = descendants.into_boxed_slice();
    }
    *blocks = roots;
    Ok(())
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
        Tag::Item => Some(BlockKind::ListItem),
        Tag::TableHead => Some(BlockKind::TableHead),
        Tag::TableRow => Some(BlockKind::TableRow),
        Tag::TableCell => Some(BlockKind::TableCell),
        Tag::DefinitionList
        | Tag::DefinitionListTitle
        | Tag::DefinitionListDefinition
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
    references: &mut Vec<ReferenceRecord>,
    limits: ReadLimits,
) -> Result<(), Error> {
    let mut event_count = 0usize;
    let mut tag_depth = 0usize;
    let mut block_depth = 0usize;
    let mut root_block: Option<OpenBlock> = None;
    let mut open_nested_blocks = Vec::<OpenNestedBlock>::new();
    let mut open_inlines = Vec::<OpenInline>::new();
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
                        let footnote_label = match &tag {
                            Tag::FootnoteDefinition(label) => Some(label.to_string()),
                            Tag::Paragraph
                            | Tag::Heading { .. }
                            | Tag::BlockQuote(_)
                            | Tag::CodeBlock(_)
                            | Tag::HtmlBlock
                            | Tag::List(_)
                            | Tag::Item
                            | Tag::DefinitionList
                            | Tag::DefinitionListTitle
                            | Tag::DefinitionListDefinition
                            | Tag::Table(_)
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
                        };
                        root_block = Some(OpenBlock {
                            kind,
                            start: range.start,
                            inlines: Vec::new(),
                            descendants: Vec::new(),
                            footnote_label,
                        });
                    } else {
                        open_nested_blocks
                            .try_reserve(1)
                            .map_err(|allocation_error| Error::Allocation {
                                resource: "Markdown nested block parser stack",
                                source: allocation_error,
                            })?;
                        open_nested_blocks.push(OpenNestedBlock {
                            kind,
                            start: range.start,
                            depth: block_depth,
                        });
                    }
                    block_depth = block_depth.saturating_add(1);
                }
                if let Some(kind) = inline_kind(&tag) {
                    let inline_depth = open_inlines.len();
                    push_open_inline(
                        &mut open_inlines,
                        OpenInline {
                            kind,
                            start: range.start,
                            depth: inline_depth,
                            reference: inline_reference(&tag, range.clone()),
                        },
                    )?;
                }
            },
            Event::End(end) => {
                if is_inline_end(end)
                    && let Some(mut inline) = open_inlines.pop()
                {
                    let node_range = inline.start..range.end;
                    if let Some(reference) = inline.reference.as_mut() {
                        reference.range = node_range.clone();
                        push_reference(references, reference.clone())?;
                    }
                    if let Some(block) = root_block.as_mut() {
                        push_inline(
                            &mut block.inlines,
                            InlineRecord {
                                kind: inline.kind,
                                range: node_range,
                                depth: inline.depth,
                            },
                        )?;
                    }
                }
                if is_block_end(end) {
                    if block_depth > 1
                        && let Some(nested) = open_nested_blocks.pop()
                        && let Some(block) = root_block.as_mut()
                    {
                        push_nested_block(
                            &mut block.descendants,
                            NestedBlockRecord {
                                kind: nested.kind,
                                range: nested.start..range.end,
                                depth: nested.depth,
                            },
                        )?;
                    }
                    block_depth = block_depth.saturating_sub(1);
                    if block_depth == 0
                        && let Some(mut block) = root_block.take()
                    {
                        block.inlines.sort_by(|left, right| {
                            left.range
                                .start
                                .cmp(&right.range.start)
                                .then_with(|| left.depth.cmp(&right.depth))
                                .then_with(|| right.range.end.cmp(&left.range.end))
                        });
                        block.descendants.sort_by(|left, right| {
                            left.range
                                .start
                                .cmp(&right.range.start)
                                .then_with(|| left.depth.cmp(&right.depth))
                                .then_with(|| right.range.end.cmp(&left.range.end))
                        });
                        if let Some(label) = block.footnote_label.take() {
                            push_reference(
                                references,
                                ReferenceRecord {
                                    kind: ReferenceKind::FootnoteDefinition,
                                    range: block.start..range.end,
                                    label: Some(label),
                                    destination: None,
                                    title: None,
                                },
                            )?;
                        }
                        push_block(
                            blocks,
                            BlockRecord {
                                kind: block.kind,
                                range: block.start..range.end,
                                inlines: block.inlines.into_boxed_slice(),
                                descendants: block.descendants.into_boxed_slice(),
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
                    inlines: Box::new([]),
                    descendants: Box::new([]),
                },
                limits,
            )?,
            Event::Text(_) => {
                push_leaf(&mut root_block, InlineKind::Text, range, open_inlines.len())?;
            },
            Event::Code(_) => {
                push_leaf(&mut root_block, InlineKind::Code, range, open_inlines.len())?;
            },
            Event::Html(_) | Event::InlineHtml(_) => {
                push_leaf(&mut root_block, InlineKind::Html, range, open_inlines.len())?;
            },
            Event::FootnoteReference(label) => {
                push_leaf(
                    &mut root_block,
                    InlineKind::FootnoteReference,
                    range.clone(),
                    open_inlines.len(),
                )?;
                push_reference(
                    references,
                    ReferenceRecord {
                        kind: ReferenceKind::Footnote,
                        range,
                        label: Some(label.to_string()),
                        destination: None,
                        title: None,
                    },
                )?;
            },
            Event::SoftBreak => push_leaf(
                &mut root_block,
                InlineKind::SoftBreak,
                range,
                open_inlines.len(),
            )?,
            Event::HardBreak => push_leaf(
                &mut root_block,
                InlineKind::HardBreak,
                range,
                open_inlines.len(),
            )?,
            Event::TaskListMarker(checked) => push_leaf(
                &mut root_block,
                InlineKind::TaskListMarker { checked },
                range,
                open_inlines.len(),
            )?,
            Event::InlineMath(_) | Event::DisplayMath(_) | Event::Rule => {},
        }
    }
    Ok(())
}

fn collect_reference_definitions(
    parser: &Parser<'_>,
    limits: ReadLimits,
) -> Result<(Vec<BlockRecord>, Vec<ReferenceRecord>), Error> {
    let definitions = parser.reference_definitions();
    let count = definitions.iter().count();
    if count > limits.max_blocks {
        return Err(Error::BlockLimitExceeded {
            limit: limits.max_blocks,
        });
    }
    let mut blocks = Vec::new();
    let mut references = Vec::new();
    blocks
        .try_reserve_exact(count)
        .map_err(|allocation_error| Error::Allocation {
            resource: "Markdown block index",
            source: allocation_error,
        })?;
    references
        .try_reserve_exact(count)
        .map_err(|allocation_error| Error::Allocation {
            resource: "Markdown reference graph",
            source: allocation_error,
        })?;
    for (label, definition) in definitions.iter() {
        blocks.push(BlockRecord {
            kind: BlockKind::LinkDefinition,
            range: definition.span.clone(),
            inlines: Box::new([]),
            descendants: Box::new([]),
        });
        references.push(ReferenceRecord {
            kind: ReferenceKind::LinkDefinition,
            range: definition.span.clone(),
            label: Some(label.to_owned()),
            destination: Some(definition.dest.to_string()),
            title: definition.title.as_ref().map(ToString::to_string),
        });
    }
    Ok((blocks, references))
}

fn inline_kind(tag: &Tag<'_>) -> Option<InlineKind> {
    match tag {
        Tag::Emphasis => Some(InlineKind::Emphasis),
        Tag::Strong => Some(InlineKind::Strong),
        Tag::Strikethrough => Some(InlineKind::Strikethrough),
        Tag::Link { .. } => Some(InlineKind::Link),
        Tag::Image { .. } => Some(InlineKind::Image),
        Tag::Paragraph
        | Tag::Heading { .. }
        | Tag::BlockQuote(_)
        | Tag::CodeBlock(_)
        | Tag::HtmlBlock
        | Tag::List(_)
        | Tag::Item
        | Tag::FootnoteDefinition(_)
        | Tag::DefinitionList
        | Tag::DefinitionListTitle
        | Tag::DefinitionListDefinition
        | Tag::Table(_)
        | Tag::TableHead
        | Tag::TableRow
        | Tag::TableCell
        | Tag::Superscript
        | Tag::Subscript
        | Tag::MetadataBlock(_) => None,
    }
}

fn inline_reference(tag: &Tag<'_>, range: std::ops::Range<usize>) -> Option<ReferenceRecord> {
    match tag {
        Tag::Link {
            dest_url,
            title,
            id,
            ..
        } => Some(ReferenceRecord {
            kind: ReferenceKind::Link,
            range,
            label: (!id.is_empty()).then(|| id.to_string()),
            destination: Some(dest_url.to_string()),
            title: (!title.is_empty()).then(|| title.to_string()),
        }),
        Tag::Image {
            dest_url,
            title,
            id,
            ..
        } => Some(ReferenceRecord {
            kind: ReferenceKind::Image,
            range,
            label: (!id.is_empty()).then(|| id.to_string()),
            destination: Some(dest_url.to_string()),
            title: (!title.is_empty()).then(|| title.to_string()),
        }),
        Tag::Paragraph
        | Tag::Heading { .. }
        | Tag::BlockQuote(_)
        | Tag::CodeBlock(_)
        | Tag::HtmlBlock
        | Tag::List(_)
        | Tag::Item
        | Tag::FootnoteDefinition(_)
        | Tag::DefinitionList
        | Tag::DefinitionListTitle
        | Tag::DefinitionListDefinition
        | Tag::Table(_)
        | Tag::TableHead
        | Tag::TableRow
        | Tag::TableCell
        | Tag::Emphasis
        | Tag::Strong
        | Tag::Strikethrough
        | Tag::Superscript
        | Tag::Subscript
        | Tag::MetadataBlock(_) => None,
    }
}

const fn is_inline_end(end: TagEnd) -> bool {
    matches!(
        end,
        TagEnd::Emphasis | TagEnd::Strong | TagEnd::Strikethrough | TagEnd::Link | TagEnd::Image
    )
}

fn push_inline(inlines: &mut Vec<InlineRecord>, inline: InlineRecord) -> Result<(), Error> {
    inlines
        .try_reserve(1)
        .map_err(|allocation_error| Error::Allocation {
            resource: "Markdown inline index",
            source: allocation_error,
        })?;
    inlines.push(inline);
    Ok(())
}

fn push_nested_block(
    blocks: &mut Vec<NestedBlockRecord>,
    block: NestedBlockRecord,
) -> Result<(), Error> {
    blocks
        .try_reserve(1)
        .map_err(|allocation_error| Error::Allocation {
            resource: "Markdown nested block index",
            source: allocation_error,
        })?;
    blocks.push(block);
    Ok(())
}

fn push_leaf(
    root: &mut Option<OpenBlock>,
    kind: InlineKind,
    range: std::ops::Range<usize>,
    depth: usize,
) -> Result<(), Error> {
    let Some(block) = root.as_mut() else {
        return Ok(());
    };
    if matches!(block.kind, BlockKind::CodeBlock { .. } | BlockKind::Html) {
        return Ok(());
    }
    push_inline(&mut block.inlines, InlineRecord { kind, range, depth })
}

fn push_open_inline(inlines: &mut Vec<OpenInline>, inline: OpenInline) -> Result<(), Error> {
    inlines
        .try_reserve(1)
        .map_err(|allocation_error| Error::Allocation {
            resource: "Markdown inline parser stack",
            source: allocation_error,
        })?;
    inlines.push(inline);
    Ok(())
}

fn push_reference(
    references: &mut Vec<ReferenceRecord>,
    reference: ReferenceRecord,
) -> Result<(), Error> {
    references
        .try_reserve(1)
        .map_err(|allocation_error| Error::Allocation {
            resource: "Markdown reference graph",
            source: allocation_error,
        })?;
    references.push(reference);
    Ok(())
}

const fn reference_order(kind: ReferenceKind) -> u8 {
    match kind {
        ReferenceKind::Link => 0,
        ReferenceKind::Image => 1,
        ReferenceKind::Footnote => 2,
        ReferenceKind::LinkDefinition => 3,
        ReferenceKind::FootnoteDefinition => 4,
    }
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
            | TagEnd::Item
            | TagEnd::FootnoteDefinition
            | TagEnd::Table
            | TagEnd::TableHead
            | TagEnd::TableRow
            | TagEnd::TableCell
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
