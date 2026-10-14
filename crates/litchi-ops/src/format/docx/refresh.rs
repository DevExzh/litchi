//! DOCX field and table-of-contents refresh (record 0279): plan new
//! results over the leaf's field index, then rewrite them in one native
//! commit.

use std::collections::{BTreeMap, BTreeSet, HashMap};

use litchi_core::FileFormat;
use litchi_core::Metadata;
use litchi_docx::document::fields::{Fields, TocEntry, Update};

use super::write::{Commit, WorkEdit};
use super::{Docx, DocxEditor, DocxOwner};
use crate::error::{Code, Error, Result};
use crate::refresh::{
    Clock, Field, Instruction, Options, PageNumbers, Status, Summary, apply_case, format_number,
};
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

/// The default text width (US Letter, one-inch margins) for the page
/// number tab stop when the section does not declare one.
const DEFAULT_TEXT_WIDTH: u32 = 9360;
/// The fixed staging charge of one update, beside its text.
const UPDATE_CHARGE: usize = 64;

/// Refresh every field of the main story. `None` means nothing changed: no
/// native commit was made.
pub(crate) fn refresh(
    file: &Docx,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<Commit>, Summary)> {
    let result = refresh_inner(file, options, limits);
    file.owner().finish(result)
}

fn refresh_inner(
    file: &Docx,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<Commit>, Summary)> {
    let editor = super::merge::owning_editor(file)?;
    let levels = match &editor {
        DocxEditor::Unsigned(package) => style_levels(package)?,
        _ => match file.owner() {
            DocxOwner::Eager(package) => style_levels(package)?,
            DocxOwner::Source(source) => style_levels(&file.promoted(source)?)?,
        },
    };
    let WorkEdit::Eager(mut edit) = WorkEdit::from(editor)? else {
        return Err(Error::new(
            Code::InternalError,
            "a DOCX refresh needs the owning package editor",
        )
        .with_format(FileFormat::Docx));
    };
    let index = edit.fields()?;
    let metadata = file.metadata()?.unwrap_or_default();
    let (updates, summary) = plan(&index, &levels, &metadata, options, limits)?;
    if updates.is_empty() {
        return Ok((None, summary));
    }
    let outcome = edit.refresh_fields(&index, &updates)?;
    if outcome.is_empty() {
        return Ok((None, summary));
    }
    file.owner().check_source()?;
    let commit = WorkEdit::Eager(edit).commit(file)?;
    Ok((Some(commit), summary))
}

/// A paragraph style's outline level (1..=9) and name.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct StyleLevel {
    level: Option<u8>,
    name: Option<String>,
    /// The level of a table of contents entry style (`TOC1`, `toc 1`).
    toc: Option<u8>,
}

/// The outline level of every paragraph style: its own or its nearest
/// based-on style's `w:outlineLvl`, else the `HeadingN` / `heading N`
/// convention.
fn style_levels(package: &litchi_docx::Package) -> Result<HashMap<String, StyleLevel>> {
    let document = package.document()?;
    let mut styles = match document.styles() {
        Ok(styles) => styles,
        Err(litchi_docx::Error::PartNotFound(_)) => return Ok(HashMap::new()),
        Err(error) => return Err(error.into()),
    };
    let mut own: HashMap<String, (Option<u8>, bool, Option<String>, Option<String>)> =
        HashMap::new();
    for style in styles.iter()? {
        if style.style_type() != litchi_docx::styles::Type::Paragraph {
            continue;
        }
        own.insert(
            style.style_id().to_owned(),
            (
                style.outline().map(|outline| outline as u8 + 1),
                style.declares_outline(),
                style.based_on().map(str::to_owned),
                style.name().map(str::to_owned),
            ),
        );
    }
    let mut levels = HashMap::with_capacity(own.len());
    for (id, (_, _, _, name)) in &own {
        let mut level = None;
        let mut declared = false;
        let mut current = Some(id.as_str());
        for _ in 0..16 {
            let Some(style) = current.and_then(|id| own.get(id)) else {
                break;
            };
            if style.1 {
                level = style.0;
                declared = true;
                break;
            }
            current = style.2.as_deref();
        }
        let squeezed = name.as_deref().map(|name| name.replace(' ', ""));
        let level = level.or_else(|| {
            if declared {
                return None;
            }
            heading_number(id).or_else(|| squeezed.as_deref().and_then(heading_number))
        });
        let toc = numbered(id, "toc")
            .or_else(|| squeezed.as_deref().and_then(|name| numbered(name, "toc")));
        levels.insert(
            id.clone(),
            StyleLevel {
                level,
                name: name.clone(),
                toc,
            },
        );
    }
    Ok(levels)
}

/// `Heading1`..`Heading9` (any case) is level 1..9.
fn heading_number(id: &str) -> Option<u8> {
    numbered(id, "heading")
}

/// `{prefix}1`..`{prefix}9`, the prefix in any case, is 1..9.
fn numbered(id: &str, prefix: &str) -> Option<u8> {
    let rest = id
        .get(..prefix.len())?
        .eq_ignore_ascii_case(prefix)
        .then(|| &id[prefix.len()..])?;
    let level: u8 = rest.parse().ok()?;
    (1..=9).contains(&level).then_some(level)
}

/// The level a `TOCn` (or `toc n`-named) style gives an existing entry.
fn toc_level(style: Option<&str>, levels: &HashMap<String, StyleLevel>) -> u8 {
    style
        .and_then(|style| {
            levels
                .get(style)
                .and_then(|level| level.toc)
                .or_else(|| numbered(style, "toc"))
        })
        .unwrap_or(0)
}

/// A table of contents entry as compared for an exact no-op.
#[derive(Debug, PartialEq, Eq)]
struct Entry {
    level: u8,
    text: String,
    anchor: Option<String>,
    page: Option<String>,
}

/// Plan every field's new result.
pub(crate) fn plan(
    index: &Fields,
    levels: &HashMap<String, StyleLevel>,
    metadata: &Metadata,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Vec<Update>, Summary)> {
    let fields = index.fields();
    let paragraphs = index.paragraphs();
    let instructions: Vec<Instruction> = fields
        .iter()
        .map(|field| Instruction::parse(field.instruction()))
        .collect();
    let mut summary = Summary::default();
    let mut updates: Vec<Update> = Vec::new();
    let mut staged = 0usize;
    let mut charge = |updates: &[Update], bytes: usize| -> Result<()> {
        if updates.len() > limits.max_expanded_targets() {
            return Err(TransactionError::Limit {
                kind: LimitKind::ExpandedTargets,
                observed: updates.len(),
                maximum: limits.max_expanded_targets(),
            }
            .into());
        }
        staged = staged.saturating_add(bytes).saturating_add(UPDATE_CHARGE);
        if staged > limits.max_staged_bytes() {
            return Err(TransactionError::Limit {
                kind: LimitKind::StagedBytes,
                observed: staged,
                maximum: limits.max_staged_bytes(),
            }
            .into());
        }
        Ok(())
    };
    let location = |paragraph: usize| {
        paragraphs[paragraph]
            .body_position()
            .map(|position| format!("/body/p[{}]", position + 1))
    };
    // Paragraphs inside a table of contents' result are not headings.
    let mut in_toc = vec![false; paragraphs.len()];
    for (field, instruction) in fields.iter().zip(&instructions) {
        if instruction.kind == "TOC" {
            let (first, last) = field.paragraphs();
            for flag in &mut in_toc[first..=last.min(paragraphs.len() - 1)] {
                *flag = true;
            }
        }
    }
    // Bookmark names in use, and the `_Toc` bookmark each paragraph starts.
    let mut names: BTreeSet<String> = index
        .bookmarks()
        .iter()
        .map(|bookmark| bookmark.name().to_owned())
        .collect();
    let mut anchors: BTreeMap<usize, String> = BTreeMap::new();
    for bookmark in index.bookmarks() {
        if let Some(paragraph) = bookmark.paragraph()
            && bookmark.name().starts_with("_Toc")
        {
            anchors
                .entry(paragraph)
                .or_insert_with(|| bookmark.name().to_owned());
        }
    }
    let mut next_name = 100_000_000u64;
    // Tables of contents that are rebuilt: their nested fields go with them.
    let mut rebuilt: BTreeSet<usize> = BTreeSet::new();
    let mut seq: HashMap<String, (u64, bool)> = HashMap::new();
    let width = index.text_width().unwrap_or(DEFAULT_TEXT_WIDTH);
    for (position, (field, instruction)) in fields.iter().zip(&instructions).enumerate() {
        // A field under a rebuilt table of contents is replaced with it.
        let mut ancestor = field.parent();
        let mut replaced = false;
        while let Some(parent) = ancestor {
            if rebuilt.contains(&parent) {
                replaced = true;
                break;
            }
            ancestor = fields[parent].parent();
        }
        if replaced {
            continue;
        }
        let kind = if instruction.kind.is_empty() {
            "EMPTY".to_owned()
        } else {
            instruction.kind.clone()
        };
        let place = location(field.paragraphs().0);
        let report = |status| Field::new(kind.clone(), place.clone(), status);
        let layout = matches!(
            kind.as_str(),
            "PAGE" | "NUMPAGES" | "SECTIONPAGES" | "PAGEREF" | "SECTION"
        );
        if field.is_locked() {
            summary.fields.push(report(Status::Locked));
            continue;
        }
        if layout {
            summary.fields.push(report(Status::NeedsLayout));
            continue;
        }
        if field.parent().is_some() {
            summary
                .fields
                .push(report(Status::Unsupported).because("the field is nested in another field"));
            continue;
        }
        if field.is_in_nested_paragraph() {
            summary
                .fields
                .push(report(Status::Unsupported).because("the field is in a text box"));
            continue;
        }
        if kind == "TOC" {
            let plan = plan_toc(
                index,
                position,
                instruction,
                levels,
                &in_toc,
                &mut anchors,
                &mut names,
                &mut next_name,
                width,
            );
            let toc = match plan {
                Ok(toc) => toc,
                Err(reason) => {
                    summary
                        .fields
                        .push(report(Status::Unsupported).because(reason));
                    continue;
                },
            };
            rebuilt.insert(position);
            if toc.missing > 0 && options.page_number_policy() == PageNumbers::Require {
                return Err(Error::new(
                    Code::RefreshFailed,
                    "a table of contents needs page numbers Litchi cannot compute without layout",
                )
                .with_format(FileFormat::Docx)
                .with_suggestion(
                    "refresh with page numbers kept, or update the table in a word processor",
                ));
            }
            if toc.unchanged {
                summary.fields.push(report(Status::Unchanged));
                continue;
            }
            summary.missing_page_numbers += toc.missing;
            for (paragraph, name) in toc.bookmarks {
                updates.push(Update::Bookmark { paragraph, name });
                summary.bookmarks += 1;
                charge(&updates, 0)?;
            }
            let bytes = toc
                .entries
                .iter()
                .map(|entry| entry.text().len() + 160)
                .sum();
            updates.push(Update::Toc {
                field: position,
                entries: toc.entries,
                dirty: toc.missing > 0 && options.marks_dirty(),
            });
            charge(&updates, bytes)?;
            summary.fields.push(report(Status::Refreshed));
            continue;
        }
        let computed = compute(
            index,
            &kind,
            instruction,
            metadata,
            options.clock_value(),
            &mut seq,
        );
        let text = match computed {
            Ok(text) => text,
            Err((status, reason)) => {
                let mut entry = report(status);
                if let Some(reason) = reason {
                    entry = entry.because(reason);
                }
                summary.fields.push(entry);
                continue;
            },
        };
        if !index.accepts_text(position) {
            summary
                .fields
                .push(report(Status::Unsupported).because("the field result holds other content"));
            continue;
        }
        if field.has_result() && field.result() == text {
            summary.fields.push(report(Status::Unchanged));
            continue;
        }
        let bytes = text.len();
        updates.push(Update::Text {
            field: position,
            text,
        });
        charge(&updates, bytes)?;
        summary.fields.push(report(Status::Refreshed));
    }
    Ok((updates, summary))
}

/// A planned table of contents.
struct TocPlan {
    entries: Vec<TocEntry>,
    bookmarks: Vec<(usize, String)>,
    missing: usize,
    unchanged: bool,
}

fn parse_range(range: &str) -> Option<(u8, u8)> {
    let (start, end) = range.trim().split_once('-')?;
    let start: u8 = start.trim().parse().ok()?;
    let end: u8 = end.trim().parse().ok()?;
    (1 <= start && start <= end && end <= 9).then_some((start, end))
}

fn plan_toc(
    index: &Fields,
    position: usize,
    instruction: &Instruction,
    levels: &HashMap<String, StyleLevel>,
    in_toc: &[bool],
    anchors: &mut BTreeMap<usize, String>,
    names: &mut BTreeSet<String>,
    next_name: &mut u64,
    width: u32,
) -> std::result::Result<TocPlan, &'static str> {
    if let Some((name, _)) = instruction
        .switches
        .iter()
        .find(|(name, _)| matches!(name, 'b' | 'c' | 'a' | 'f' | 's' | 'd' | 'l' | '#' | '@'))
    {
        let _ = name;
        return Err(
            "the table of contents uses a switch Litchi does not compute (\\a, \\b, \\c, \\d, \\f, \\l or \\s)",
        );
    }
    let outline = match instruction.switch('o') {
        Some(range) => Some(parse_range(range).ok_or("the \\o level range is malformed")?),
        None if instruction.has('o') => Some((1, 9)),
        None => None,
    };
    let applied = instruction.has('u');
    let mut custom: Vec<(String, u8)> = Vec::new();
    if let Some(list) = instruction.switch('t') {
        let parts: Vec<&str> = list.split([',', ';']).collect();
        for pair in parts.chunks(2) {
            let [style, level] = pair else {
                return Err("the \\t style list is malformed");
            };
            let level: u8 = level
                .trim()
                .parse()
                .ok()
                .filter(|level| (1..=9).contains(level))
                .ok_or("the \\t style list is malformed")?;
            custom.push((style.trim().to_lowercase(), level));
        }
    }
    if outline.is_none() && !applied && custom.is_empty() {
        return Err("the table of contents has no \\o, \\u or \\t source");
    }
    if !index.accepts_paragraphs(position) {
        return Err("the table of contents' markup cannot be rebuilt");
    }
    let omit = match (instruction.has('n'), instruction.switch('n')) {
        (false, _) => None,
        (true, None) => Some((1, 9)),
        (true, Some(range)) => Some(parse_range(range).ok_or("the \\n level range is malformed")?),
    };
    let hyperlinks = instruction.has('h');
    let separator = instruction.switch('p').map(str::to_owned);
    let fields = index.fields();
    let paragraphs = index.paragraphs();
    let (first, last) = fields[position].paragraphs();
    // The cached entries and their page numbers.
    let mut pagerefs: HashMap<usize, (String, String)> = HashMap::new();
    for field in fields {
        if field.parent() == Some(position) {
            let nested = Instruction::parse(field.instruction());
            if nested.kind == "PAGEREF"
                && let Some(anchor) = nested.arguments.first()
            {
                pagerefs.insert(
                    field.paragraphs().0,
                    (anchor.clone(), field.result().to_owned()),
                );
            }
        }
    }
    let mut old: Vec<Entry> = Vec::new();
    for (offset, paragraph) in paragraphs[first..=last].iter().enumerate() {
        let mut text = paragraph.text().to_owned();
        let pageref = pagerefs.get(&(first + offset));
        let page = pageref.map(|(_, page)| page.clone());
        if let Some(page) = &page
            && let Some(stripped) = text.strip_suffix(page.as_str())
        {
            text = stripped.to_owned();
            match &separator {
                Some(separator) => {
                    if let Some(stripped) = text.strip_suffix(separator.as_str()) {
                        text = stripped.to_owned();
                    }
                },
                None => {
                    if text.ends_with('\t') {
                        text.pop();
                    }
                },
            }
        }
        if text.is_empty() && page.is_none() {
            continue;
        }
        let anchor = paragraph
            .hyperlink_anchor()
            .map(str::to_owned)
            .or_else(|| pageref.map(|(anchor, _)| anchor.clone()));
        old.push(Entry {
            level: toc_level(paragraph.style(), levels),
            text,
            anchor,
            page,
        });
    }
    let page_of = |anchor: Option<&str>, level: u8, text: &str| -> Option<String> {
        old.iter()
            .find(|entry| anchor.is_some() && entry.anchor.as_deref() == anchor)
            .or_else(|| {
                old.iter()
                    .find(|entry| entry.level == level && entry.text == text)
            })
            .and_then(|entry| entry.page.clone())
            .filter(|page| !page.is_empty())
    };
    let mut entries = Vec::new();
    let mut new: Vec<Entry> = Vec::new();
    let mut bookmarks = Vec::new();
    let mut missing = 0;
    for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
        if paragraph.is_nested() || in_toc[paragraph_index] {
            continue;
        }
        let style = paragraph
            .style()
            .and_then(|style| levels.get(style).map(|level| (style, level)));
        let mut level = None;
        if let Some((id, style)) = style
            && !custom.is_empty()
        {
            let id = id.to_lowercase();
            let name = style.name.as_deref().map(str::to_lowercase);
            // Word matches the style's (possibly localized) name; a
            // localized name is often the id with its spaces.
            level = custom
                .iter()
                .find(|(custom, _)| {
                    *custom == id || Some(custom) == name.as_ref() || custom.replace(' ', "") == id
                })
                .map(|(_, level)| *level);
        }
        let in_range =
            |level: u8| outline.is_none_or(|(start, end)| start <= level && level <= end);
        if level.is_none()
            && applied
            && let Some(direct) = paragraph.outline_level()
            && in_range(direct + 1)
        {
            level = Some(direct + 1);
        }
        if level.is_none()
            && outline.is_some()
            && let Some(style_level) = style.and_then(|(_, style)| style.level)
            && in_range(style_level)
        {
            level = Some(style_level);
        }
        let Some(level) = level else {
            continue;
        };
        let text = paragraph.text().replace(['\t', '\n'], " ");
        let text = text.trim();
        if text.is_empty() {
            continue;
        }
        let numbered = omit.is_none_or(|(start, end)| level < start || level > end);
        let anchor = if numbered || hyperlinks {
            let name = anchors.entry(paragraph_index).or_insert_with(|| {
                loop {
                    *next_name += 1;
                    let name = format!("_Toc{next_name}");
                    if names.insert(name.clone()) {
                        bookmarks.push((paragraph_index, name.clone()));
                        break name;
                    }
                }
            });
            Some(name.clone())
        } else {
            None
        };
        let page = numbered.then(|| {
            page_of(anchor.as_deref(), level, text).unwrap_or_else(|| {
                missing += 1;
                String::new()
            })
        });
        let mut entry = TocEntry::new(level, text).hyperlink(hyperlinks);
        if let Some(anchor) = &anchor {
            entry = entry.anchor(anchor.clone());
        }
        if let Some(page) = &page {
            entry = entry.page(page.clone()).tab(width);
            if let Some(separator) = &separator {
                entry = entry.separator(separator.clone());
            }
        }
        entries.push(entry);
        new.push(Entry {
            level,
            text: text.to_owned(),
            anchor: if hyperlinks || numbered { anchor } else { None },
            page,
        });
    }
    // Old entries carry an anchor only through a link or a page reference.
    let comparable = |entries: &[Entry]| -> Vec<(u8, String, Option<String>, Option<String>)> {
        entries
            .iter()
            .map(|entry| {
                (
                    entry.level,
                    entry.text.clone(),
                    entry.anchor.clone(),
                    entry.page.clone(),
                )
            })
            .collect()
    };
    let unchanged = bookmarks.is_empty() && comparable(&old) == comparable(&new);
    Ok(TocPlan {
        entries,
        bookmarks,
        missing,
        unchanged,
    })
}

type Skip = (Status, Option<&'static str>);

/// The new result of a non-TOC field.
fn compute(
    index: &Fields,
    kind: &str,
    instruction: &Instruction,
    metadata: &Metadata,
    clock: Option<Clock>,
    seq: &mut HashMap<String, (u64, bool)>,
) -> std::result::Result<String, Skip> {
    if instruction.has('#') {
        return Err((
            Status::Unsupported,
            Some("numeric pictures (\\#) are not applied"),
        ));
    }
    let text = match kind {
        "SEQ" => return sequence(instruction, seq),
        "DATE" | "TIME" => {
            let clock = clock.ok_or((Status::NeedsClock, None))?;
            let default = if kind == "DATE" {
                "M/d/yyyy"
            } else {
                "h:mm am/pm"
            };
            clock.format(instruction.switch('@').unwrap_or(default))
        },
        "CREATEDATE" | "SAVEDATE" | "PRINTDATE" => {
            let stamp = match kind {
                "CREATEDATE" => metadata.created,
                "SAVEDATE" => metadata.modified,
                _ => metadata.last_printed,
            };
            let clock = stamp.and_then(Clock::from_timestamp).ok_or((
                Status::Unresolved,
                Some("the document does not store this time"),
            ))?;
            clock.format(instruction.switch('@').unwrap_or("M/d/yyyy h:mm:ss am/pm"))
        },
        "REF" => {
            let name = instruction
                .arguments
                .first()
                .ok_or((Status::Unsupported, Some("REF names no bookmark")))?;
            if instruction
                .switches
                .iter()
                .any(|(switch, _)| !matches!(switch, 'h' | '*'))
            {
                return Err((
                    Status::Unsupported,
                    Some("REF paragraph-number and position switches are not computed"),
                ));
            }
            index
                .bookmarks()
                .iter()
                .find(|bookmark| bookmark.name().eq_ignore_ascii_case(name))
                .map(|bookmark| bookmark.text().trim_end_matches(['\n', '\r']).to_owned())
                .ok_or((Status::Unresolved, Some("the bookmark does not exist")))?
        },
        "TITLE" | "SUBJECT" | "AUTHOR" | "KEYWORDS" | "COMMENTS" | "LASTSAVEDBY" => {
            if !instruction.arguments.is_empty() {
                return Err((
                    Status::Unsupported,
                    Some("setting a property from a field is not done"),
                ));
            }
            property(metadata, kind).unwrap_or_default()
        },
        "DOCPROPERTY" => {
            let name = instruction
                .arguments
                .first()
                .ok_or((Status::Unsupported, Some("DOCPROPERTY names no property")))?;
            let key = name.replace(' ', "").to_ascii_uppercase();
            property(metadata, &key).ok_or((
                Status::Unresolved,
                Some("the property is not a known document property"),
            ))?
        },
        "NUMWORDS" | "NUMCHARS" => {
            let mut words = 0usize;
            let mut characters = 0usize;
            for paragraph in index.paragraphs() {
                words += paragraph.text().split_whitespace().count();
                characters += paragraph
                    .text()
                    .chars()
                    .filter(|character| !character.is_whitespace())
                    .count();
            }
            if kind == "NUMWORDS" {
                words
            } else {
                characters
            }
            .to_string()
        },
        _ => return Err((Status::Unsupported, None)),
    };
    let mut text = text;
    for format in instruction.formats() {
        text = apply_case(&text, format).ok_or((
            Status::Unsupported,
            Some("the \\* format is not applied to this field"),
        ))?;
    }
    Ok(text)
}

fn property(metadata: &Metadata, key: &str) -> Option<String> {
    let value = match key {
        "TITLE" => &metadata.title,
        "SUBJECT" => &metadata.subject,
        "AUTHOR" => &metadata.author,
        "KEYWORDS" => &metadata.keywords,
        "COMMENTS" => &metadata.description,
        "LASTSAVEDBY" => &metadata.last_modified_by,
        "CATEGORY" => &metadata.category,
        "COMPANY" => &metadata.company,
        "MANAGER" => &metadata.manager,
        "TEMPLATE" => &metadata.template,
        _ => return None,
    };
    Some(value.clone().unwrap_or_default())
}

/// `SEQ` numbering in document order, after OfficeCLI's evaluator.
fn sequence(
    instruction: &Instruction,
    counters: &mut HashMap<String, (u64, bool)>,
) -> std::result::Result<String, Skip> {
    let identifier = instruction
        .arguments
        .first()
        .ok_or((Status::Unsupported, Some("SEQ names no sequence")))?
        .clone();
    let counter = counters.entry(identifier).or_insert((0, false));
    if counter.1 {
        return Err((
            Status::Unsupported,
            Some("an earlier field of this sequence was not computed"),
        ));
    }
    if instruction.has('s') {
        counter.1 = true;
        return Err((
            Status::Unsupported,
            Some("heading-relative SEQ resets (\\s) need heading numbering"),
        ));
    }
    let value = if let Some(reset) = instruction.switch('r') {
        let Ok(reset) = reset.trim().parse::<u64>() else {
            counter.1 = true;
            return Err((
                Status::Unsupported,
                Some("the SEQ \\r reset is not a number"),
            ));
        };
        reset
    } else if instruction.has('c') {
        counter.0
    } else {
        counter.0 + 1
    };
    counter.0 = value;
    if instruction.has('h') {
        return Ok(String::new());
    }
    let mut formats = instruction.formats();
    let text = match formats.next() {
        None => value.to_string(),
        Some(format) => {
            if formats.next().is_some() {
                counter.1 = true;
                return Err((Status::Unsupported, Some("several SEQ formats")));
            }
            if let Some(text) = format_number(value, format) {
                text
            } else {
                counter.1 = true;
                return Err((
                    Status::Unsupported,
                    Some("the SEQ number format is not applied"),
                ));
            }
        },
    };
    Ok(text)
}

#[cfg(test)]
mod tests;
