//! The DOCX part elements (record 0319): styles and numbering (L-DOCX-R6),
//! headers, footers, footnotes and endnotes (L-DOCX-R5).
//!
//! | Path | Node | Leaf read |
//! |---|---|---|
//! | `/styles`, `/styles/<id or name>` | the styles part, one style | `styles::catalog::Catalog` |
//! | `/numbering`, `/numbering/abstractNum[N]`, `/numbering/num[N]`, `…/abstractNum[N]/level[L]` | the numbering part, a definition, an instance, a level (`level[L]` is the L-th level) | `numbering::catalog::Catalog` |
//! | `/header[N]`, `/footer[N]` (and `[@type=T]`) | the N-th header or footer part in relationship order | `header_footer::catalog::Entry` and the part payload |
//! | `/footnotes`, `/footnotes/footnote[N]`, `/footnote[N]` (note id N) | the notes part, a note | `footnote::Note` |
//!
//! Each part is read once per handle and kept in the handle's [`Cache`];
//! a source-backed owner's reads are still fenced against the source.
//! Values are local (ADR 0004): a style's properties are its own, not
//! inherited through `w:basedOn`. Story and note nodes have no listed
//! children, so unscoped `p` and `r` queries stay in the main story.

use std::borrow::Cow;
use std::sync::{Arc, OnceLock};

use litchi_docx::footnote::Note;
use litchi_docx::header_footer::catalog::Entry as StoryPart;
use litchi_docx::header_footer::{Kind as StoryKind, Role};
use litchi_docx::numbering::catalog::Catalog as NumberingCatalog;
use litchi_docx::properties as direct;
use litchi_docx::styles::catalog::{Catalog as StyleCatalog, Flag};

use super::node::{Address, Node, Step, Value};
use super::spec::{Kind, PARAGRAPH_PROPS, RUN_PROPS, index_of};
use super::{Docx, DocxOwner};
use crate::error::{Code, Error, Result};
use crate::node::spec::PropSpec;
use crate::node::{PathStep, Value as NodeValue};
use crate::path::navigate::{KindSet, Lookup, Resolved, Select};
use crate::registry::{Access, Constraint, Entry};
use crate::value::{Tokens, ValueType};

pub(super) mod write;

// ---------------------------------------------------------------------------
// The per-handle cache
// ---------------------------------------------------------------------------

/// The parts a handle has read, each on first use.
#[derive(Default)]
pub(crate) struct Cache {
    styles: OnceLock<Option<StyleCatalog>>,
    numbering: OnceLock<Option<NumberingCatalog>>,
    stories: OnceLock<Vec<Story>>,
    /// The section kind of each story, read from the section inventory
    /// when a `type` is first asked for.
    story_kinds: OnceLock<Vec<Option<StoryKind>>>,
    footnotes: OnceLock<Vec<Note>>,
    endnotes: OnceLock<Vec<Note>>,
    /// The main story's bookmarks and fields, read on first use (record
    /// 0325; the marks of record 0319).
    pub(super) bookmarks: OnceLock<Vec<litchi_docx::bookmark::Bookmark>>,
    pub(super) fields: OnceLock<Vec<litchi_docx::Field>>,
    /// The SmartArt inventory in document order, each graphic's outline
    /// read on first use (record 0335).
    pub(super) smartarts: OnceLock<Arc<[Arc<super::smartart::SmartArtValue>]>>,
    /// The chart parts read so far, by main-document relationship (record
    /// 0335): a chart's series and axes resolve against its projection.
    pub(super) charts:
        std::sync::Mutex<Vec<(Box<str>, Option<Arc<crate::format::chart::Summary>>)>>,
    /// Each comment's anchor path by comment id, from one walk of the main
    /// story on first use (record 0345).
    pub(super) comment_anchors: OnceLock<Arc<std::collections::HashMap<u32, String>>>,
    /// The `commentsExtended` entries, read on first use (record 0345).
    pub(super) comment_threads: OnceLock<Arc<[litchi_docx::modern_comments::Extended]>>,
}

/// One header or footer part, with its payload read on first use.
pub(crate) struct Story {
    part: StoryPart,
    xml: OnceLock<Arc<Vec<u8>>>,
    /// The story's block index, built on first use (record 0325).
    blocks: OnceLock<litchi_docx::Blocks>,
}

impl Story {
    pub(crate) fn part(&self) -> &StoryPart {
        &self.part
    }
}

fn source_or<T>(
    docx: &Docx,
    eager: impl FnOnce(&litchi_docx::Package) -> litchi_docx::Result<T>,
    source: impl FnOnce(&litchi_docx::source_backed::Package) -> litchi_docx::Result<T>,
) -> Result<T> {
    let result = match docx.owner() {
        DocxOwner::Eager(package) => eager(package).map_err(Error::from),
        DocxOwner::Source(package) => source(package).map_err(Error::from),
    };
    docx.owner().finish(result)
}

/// The styles part of `docx`, or `None` when it has none.
pub(super) fn styles(docx: &Docx) -> Result<Option<&StyleCatalog>> {
    let cache = docx.parts();
    if let Some(catalog) = cache.styles.get() {
        docx.owner().check_source()?;
        return Ok(catalog.as_ref());
    }
    let loaded = source_or(docx, litchi_docx::Package::style_catalog, |package| {
        package.style_catalog()
    })?;
    Ok(cache.styles.get_or_init(|| loaded).as_ref())
}

/// The numbering part of `docx`, or `None` when it has none.
pub(super) fn numbering(docx: &Docx) -> Result<Option<&NumberingCatalog>> {
    let cache = docx.parts();
    if let Some(catalog) = cache.numbering.get() {
        docx.owner().check_source()?;
        return Ok(catalog.as_ref());
    }
    let loaded = source_or(docx, litchi_docx::Package::numbering_catalog, |package| {
        package.numbering_catalog()
    })?;
    Ok(cache.numbering.get_or_init(|| loaded).as_ref())
}

/// The header and footer parts of `docx`, in relationship order.
pub(super) fn stories(docx: &Docx) -> Result<&[Story]> {
    let cache = docx.parts();
    if let Some(stories) = cache.stories.get() {
        docx.owner().check_source()?;
        return Ok(stories);
    }
    let parts = source_or(
        docx,
        litchi_docx::Package::header_footer_entries,
        |package| package.header_footer_entries(),
    )?;
    let stories = parts
        .into_iter()
        .map(|part| Story {
            part,
            xml: OnceLock::new(),
            blocks: OnceLock::new(),
        })
        .collect();
    Ok(cache.stories.get_or_init(|| stories))
}

/// The stories of one role, with their indices in [`stories`].
fn role_stories(docx: &Docx, role: Role) -> Result<Vec<usize>> {
    Ok(stories(docx)?
        .iter()
        .enumerate()
        .filter(|(_, story)| story.part.role() == role)
        .map(|(index, _)| index)
        .collect())
}

/// The payload of story `index`.
pub(super) fn story_xml(docx: &Docx, index: usize) -> Result<Arc<Vec<u8>>> {
    let story = stories(docx)?
        .get(index)
        .ok_or_else(|| internal("a DOCX story index is outside the story list"))?;
    if let Some(xml) = story.xml.get() {
        return Ok(Arc::clone(xml));
    }
    let name = story.part.name().to_owned();
    let xml = source_or(
        docx,
        |package| package.part_xml(&name),
        |package| package.part_xml(&name),
    )?
    .ok_or_else(|| {
        Error::new(
            Code::CorruptFile,
            "a related header or footer part is missing",
        )
        .with_format(litchi_core::FileFormat::Docx)
    })?;
    Ok(Arc::clone(story.xml.get_or_init(|| xml)))
}

/// The section kind (`default`, `first`, `even`) each story is referenced
/// as, from the first section that references it.
fn story_kinds(docx: &Docx) -> Result<&[Option<StoryKind>]> {
    let cache = docx.parts();
    if let Some(kinds) = cache.story_kinds.get() {
        docx.owner().check_source()?;
        return Ok(kinds);
    }
    let stories = stories(docx)?;
    let references = source_or(
        docx,
        litchi_docx::Package::header_footer_references,
        |package| package.header_footer_references(),
    )?;
    let kinds = stories
        .iter()
        .map(|story| {
            references
                .iter()
                .find(|(role, reference)| {
                    *role == story.part.role()
                        && reference.relationship_id == story.part.relationship_id()
                })
                .map(|(_, reference)| reference.kind)
        })
        .collect();
    Ok(cache.story_kinds.get_or_init(|| kinds))
}

/// The notes of `kind` of `docx`, in part order, separators excluded.
pub(super) fn notes(docx: &Docx, kind: litchi_docx::document::NoteKind) -> Result<&[Note]> {
    let cache = docx.parts();
    let lock = match kind {
        litchi_docx::document::NoteKind::Endnote => &cache.endnotes,
        _ => &cache.footnotes,
    };
    if let Some(notes) = lock.get() {
        docx.owner().check_source()?;
        return Ok(notes);
    }
    let endnotes = matches!(kind, litchi_docx::document::NoteKind::Endnote);
    let loaded = source_or(
        docx,
        |package| {
            let document = package.document()?;
            if endnotes {
                document.endnotes()
            } else {
                document.footnotes()
            }
        },
        |package| {
            if endnotes {
                package.endnotes()
            } else {
                package.footnotes()
            }
        },
    )?;
    Ok(lock.get_or_init(|| loaded))
}

/// The story the address step `kind[ordinal]` (a header, footer, footnote
/// or endnote) names: its leaf key and its part value (record 0325).
pub(super) fn story_of(
    docx: &Docx,
    kind: Kind,
    ordinal: usize,
) -> Result<(super::element::StoryKey, PartValue)> {
    let missing = || internal("a DOCX story step is outside its story list");
    match kind {
        Kind::Header | Kind::Footer => {
            let role = if kind == Kind::Header {
                Role::Header
            } else {
                Role::Footer
            };
            let index = *role_stories(docx, role)?.get(ordinal).ok_or_else(missing)?;
            let name = stories(docx)?
                .get(index)
                .ok_or_else(missing)?
                .part()
                .name()
                .to_owned();
            Ok((
                super::element::StoryKey::Part(name),
                PartValue::Story(ordinal_u32(index)),
            ))
        },
        _ => {
            let kind = note_kind(kind);
            let id = notes(docx, kind)?.get(ordinal).ok_or_else(missing)?.id();
            Ok((
                super::element::StoryKey::Note(kind, id),
                PartValue::Note(kind, ordinal_u32(ordinal)),
            ))
        },
    }
}

fn ordinal_u32(index: usize) -> u32 {
    u32::try_from(index).unwrap_or(u32::MAX)
}

/// The block index of a header, footer or note story (record 0325): its
/// paragraphs and tables are addressed as `/header[N]/p[M]`, ...; `None` for
/// other part values.
pub(super) fn story_blocks(docx: &Docx, value: PartValue) -> Result<Option<litchi_docx::Blocks>> {
    let blocks = match value {
        PartValue::Story(index) => {
            let story = stories(docx)?
                .get(index as usize)
                .ok_or_else(|| internal("a DOCX story index is outside the story list"))?;
            if let Some(blocks) = story.blocks.get() {
                docx.owner().check_source()?;
                return Ok(Some(blocks.clone()));
            }
            let xml = story_xml(docx, index as usize)?;
            let blocks =
                litchi_docx::header_footer::Story::from_shared_bytes(xml, StoryKind::Primary)
                    .and_then(|story| story.blocks())
                    .map_err(Error::from)?;
            story.blocks.get_or_init(|| blocks).clone()
        },
        PartValue::Note(kind, index) => notes(docx, kind)?
            .get(index as usize)
            .ok_or_else(|| internal("a DOCX note index is outside the note list"))?
            .blocks()
            .map_err(Error::from)?,
        _ => return Ok(None),
    };
    docx.owner().check_source()?;
    Ok(Some(blocks))
}

// ---------------------------------------------------------------------------
// Node values
// ---------------------------------------------------------------------------

/// What a part node addresses; indices are into the handle's cached
/// catalogs.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum PartValue {
    Styles,
    Style(u32),
    Numbering,
    AbstractNum(u32),
    Num(u32),
    /// A definition's index and the level's index in it.
    Level(u32, u32),
    /// An index into [`stories`].
    Story(u32),
    Notes(litchi_docx::document::NoteKind),
    Note(litchi_docx::document::NoteKind, u32),
}

/// Whether `kind` is one of this module's elements.
pub(super) const fn is_part_kind(kind: Kind) -> bool {
    matches!(
        kind,
        Kind::Styles
            | Kind::Style
            | Kind::Numbering
            | Kind::AbstractNum
            | Kind::Num
            | Kind::Level
            | Kind::Header
            | Kind::Footer
            | Kind::Footnotes
            | Kind::Footnote
            | Kind::Endnotes
            | Kind::Endnote
    )
}

const fn note_kind(kind: Kind) -> litchi_docx::document::NoteKind {
    match kind {
        Kind::Endnote | Kind::Endnotes => litchi_docx::document::NoteKind::Endnote,
        _ => litchi_docx::document::NoteKind::Footnote,
    }
}

fn ordinal(index: usize) -> u32 {
    u32::try_from(index).unwrap_or(u32::MAX)
}

fn node<'f>(docx: &'f Docx, kind: Kind, value: PartValue, address: Address) -> Node<'f> {
    Node::new(docx, kind, Value::Part(value), address)
}

/// The root address extended by one part step.
fn top(kind: Kind, index: usize) -> Address {
    Address::root().with(Step::new(kind, ordinal(index)))
}

/// The node of note `index` of `kind`, at its canonical
/// `/footnotes/footnote[N]` address.
fn note_node<'f>(docx: &'f Docx, kind: Kind, index: usize) -> Node<'f> {
    let notes = if kind == Kind::Endnote {
        Kind::Endnotes
    } else {
        Kind::Footnotes
    };
    node(
        docx,
        kind,
        PartValue::Note(note_kind(kind), ordinal(index)),
        top(notes, 0).with(Step::new(kind, ordinal(index))),
    )
}

/// The index `select` picks among `len` children.
fn pick(len: usize, select: &Select<'_>) -> Option<Option<usize>> {
    Some(match *select {
        Select::None => (len > 0).then_some(0),
        Select::Ordinal(position) => (position.get() < len).then(|| position.get()),
        Select::Last => len.checked_sub(1),
        _ => return None,
    })
}

/// The single identity value of `select` for `key`, `Some(None)` when the
/// stacked identities disagree.
fn identity<'s>(select: &Select<'s>, key: &str) -> Option<Option<&'s str>> {
    let Select::Identity(identities) = select else {
        return None;
    };
    let mut wanted = None;
    for (identity, value) in identities.iter() {
        if !identity.key.eq_ignore_ascii_case(key) {
            continue;
        }
        if wanted.is_some_and(|held| held != value) {
            return Some(None);
        }
        wanted = Some(value);
    }
    Some(wanted)
}

fn found<'f>(picked: Option<Option<Node<'f>>>) -> Lookup<Node<'f>, crate::node::spec::GapSpec> {
    match picked {
        Some(Some(node)) => Lookup::Found(node),
        Some(None) => Lookup::Missing,
        None => Lookup::NotApplicable,
    }
}

/// The child of `parent` (the root or a part node) that `step` selects.
pub(super) fn child<'f>(
    docx: &'f Docx,
    parent: &Node<'f>,
    step: &Resolved<'_, Kind>,
) -> Result<Lookup<Node<'f>, crate::node::spec::GapSpec>> {
    let kind = step.kind;
    let select = &step.select;
    let address = parent.child_address();
    let singleton = |present: bool| -> Option<Option<()>> {
        match select {
            Select::None | Select::Last => Some(present.then_some(())),
            Select::Ordinal(position) => Some((present && position.get() == 0).then_some(())),
            _ => None,
        }
    };
    Ok(match (&parent.value, kind) {
        (Value::Document, Kind::Styles) => found(
            singleton(styles(docx)?.is_some())
                .map(|hit| hit.map(|()| node(docx, kind, PartValue::Styles, top(kind, 0)))),
        ),
        (Value::Part(PartValue::Styles), Kind::Style) => {
            let Some(catalog) = styles(docx)? else {
                return Ok(Lookup::Missing);
            };
            match select {
                Select::Name { name, .. } => found(Some(
                    catalog
                        .entries()
                        .iter()
                        .position(|entry| entry.id() == *name || entry.name() == Some(*name))
                        .map(|index| {
                            node(
                                docx,
                                kind,
                                PartValue::Style(ordinal(index)),
                                address.with(Step::new(kind, ordinal(index))),
                            )
                        }),
                )),
                _ => Lookup::NotApplicable,
            }
        },
        (Value::Document, Kind::Numbering) => found(
            singleton(numbering(docx)?.is_some())
                .map(|hit| hit.map(|()| node(docx, kind, PartValue::Numbering, top(kind, 0)))),
        ),
        (Value::Part(PartValue::Numbering), Kind::AbstractNum | Kind::Num) => {
            let Some(catalog) = numbering(docx)? else {
                return Ok(Lookup::Missing);
            };
            let abstract_num = kind == Kind::AbstractNum;
            let ids: Vec<u32> = if abstract_num {
                catalog
                    .abstract_nums()
                    .iter()
                    .map(|entry| entry.id())
                    .collect()
            } else {
                catalog.nums().iter().map(|entry| entry.id()).collect()
            };
            let picked = match identity(select, "id") {
                Some(Some(value)) => {
                    let Ok(wanted) = value.trim().parse::<u32>() else {
                        return Err(Error::new(
                            Code::InvalidPath,
                            "invalid numbering identifier",
                        ));
                    };
                    Some(ids.iter().position(|id| *id == wanted))
                },
                Some(None) => Some(None),
                None => pick(ids.len(), select),
            };
            found(picked.map(|index| {
                index.map(|index| {
                    let value = if abstract_num {
                        PartValue::AbstractNum(ordinal(index))
                    } else {
                        PartValue::Num(ordinal(index))
                    };
                    node(
                        docx,
                        kind,
                        value,
                        address.with(Step::new(kind, ordinal(index))),
                    )
                })
            }))
        },
        (Value::Part(PartValue::AbstractNum(definition)), Kind::Level) => {
            let Some(catalog) = numbering(docx)? else {
                return Ok(Lookup::Missing);
            };
            let levels = catalog
                .abstract_nums()
                .get(*definition as usize)
                .map_or(&[][..], |definition| definition.levels());
            // `level[N]` is the N-th level (1-based), whose ilvl is N-1 in a
            // well-formed definition; `[@ilvl=L]` (raw paths) names the ilvl.
            let picked = match select {
                Select::Number(number) => {
                    Some(number.checked_sub(1).filter(|index| *index < levels.len()))
                },
                _ => match identity(select, "ilvl") {
                    Some(Some(value)) => {
                        let Ok(wanted) = value.trim().parse::<u8>() else {
                            return Err(Error::new(Code::InvalidPath, "invalid numbering level"));
                        };
                        Some(levels.iter().position(|level| level.ilvl() == wanted))
                    },
                    Some(None) => Some(None),
                    None => pick(levels.len(), select),
                },
            };
            let definition = *definition;
            found(picked.map(|index| {
                index.map(|index| {
                    node(
                        docx,
                        kind,
                        PartValue::Level(definition, ordinal(index)),
                        address.with(Step::new(kind, ordinal(index))),
                    )
                })
            }))
        },
        (Value::Document, Kind::Header | Kind::Footer) => {
            let role = if kind == Kind::Header {
                Role::Header
            } else {
                Role::Footer
            };
            let indices = role_stories(docx, role)?;
            let picked = match identity(select, "type") {
                Some(Some(value)) => {
                    let wanted = StoryKind::from_xml(value.trim());
                    let kinds = story_kinds(docx)?;
                    Some(indices.iter().position(|index| {
                        wanted.is_some() && kinds.get(*index).copied().flatten() == wanted
                    }))
                },
                Some(None) => Some(None),
                None => pick(indices.len(), select),
            };
            found(picked.map(|position| {
                position.map(|position| {
                    node(
                        docx,
                        kind,
                        PartValue::Story(ordinal(indices[position])),
                        top(kind, position),
                    )
                })
            }))
        },
        (Value::Document, Kind::Footnotes | Kind::Endnotes) => found(
            singleton(!notes(docx, note_kind(kind))?.is_empty()).map(|hit| {
                hit.map(|()| node(docx, kind, PartValue::Notes(note_kind(kind)), top(kind, 0)))
            }),
        ),
        (Value::Part(PartValue::Notes(_)), Kind::Footnote | Kind::Endnote) => {
            let notes = notes(docx, note_kind(kind))?;
            let picked = match identity(
                select,
                if kind == Kind::Footnote {
                    "footnoteId"
                } else {
                    "endnoteId"
                },
            ) {
                Some(Some(value)) => {
                    let Ok(wanted) = value.trim().parse::<u32>() else {
                        return Err(Error::new(Code::InvalidPath, "invalid note identifier"));
                    };
                    Some(notes.iter().position(|note| note.id() == wanted))
                },
                Some(None) => Some(None),
                None => pick(notes.len(), select),
            };
            found(picked.map(|index| index.map(|index| note_node(docx, kind, index))))
        },
        // `/footnote[N]` is note id N; `[@footnoteId=N]` too (raw paths).
        (Value::Document, Kind::Footnote | Kind::Endnote) => {
            let notes = notes(docx, note_kind(kind))?;
            let key = if kind == Kind::Footnote {
                "footnoteId"
            } else {
                "endnoteId"
            };
            let wanted = match (select, identity(select, key)) {
                (Select::Number(number), _) => Some(u32::try_from(*number).ok()),
                (_, Some(Some(value))) => Some(value.trim().parse::<u32>().ok()),
                (_, Some(None)) => Some(None),
                (Select::Last, None) => Some(notes.last().map(Note::id)),
                (Select::None, None) => Some(notes.first().map(Note::id)),
                _ => None,
            };
            found(wanted.map(|wanted| {
                wanted
                    .and_then(|wanted| notes.iter().position(|note| note.id() == wanted))
                    .map(|index| note_node(docx, kind, index))
            }))
        },
        _ => Lookup::NotApplicable,
    })
}

// ---------------------------------------------------------------------------
// Children
// ---------------------------------------------------------------------------

/// The part children of the root, or the children of a part node, whose
/// kinds are in a set.
pub(crate) struct Children {
    kinds: KindSet<Kind>,
    address: Address,
    state: State,
}

#[derive(Clone, Copy)]
enum State {
    /// The root: the styles part, headers, footers, the numbering part,
    /// then the notes parts, each when present.
    RootStyles,
    RootStories(Role, usize),
    RootNumbering,
    RootNotes(litchi_docx::document::NoteKind),
    Styles(usize),
    AbstractNums(usize),
    Nums(usize),
    Levels(u32, usize),
    Notes(Kind, usize),
    Done,
}

impl Children {
    /// The part children of the root.
    pub(super) fn root(kinds: KindSet<Kind>) -> Self {
        Self {
            kinds,
            address: Address::root(),
            state: State::RootStyles,
        }
    }

    /// The children of a part node.
    pub(super) fn of(parent: &Node<'_>, value: PartValue, kinds: KindSet<Kind>) -> Self {
        let state = match value {
            PartValue::Styles if kinds.contains(Kind::Style) => State::Styles(0),
            PartValue::Numbering if kinds.contains(Kind::AbstractNum) => State::AbstractNums(0),
            PartValue::Numbering if kinds.contains(Kind::Num) => State::Nums(0),
            PartValue::AbstractNum(definition) if kinds.contains(Kind::Level) => {
                State::Levels(definition, 0)
            },
            PartValue::Notes(litchi_docx::document::NoteKind::Endnote)
                if kinds.contains(Kind::Endnote) =>
            {
                State::Notes(Kind::Endnote, 0)
            },
            PartValue::Notes(_) if kinds.contains(Kind::Footnote) => {
                State::Notes(Kind::Footnote, 0)
            },
            _ => State::Done,
        };
        Self {
            kinds,
            address: parent.child_address(),
            state,
        }
    }

    /// Whether the root listing wants any part kind.
    pub(super) fn wants_root(kinds: KindSet<Kind>) -> bool {
        [
            Kind::Styles,
            Kind::Header,
            Kind::Footer,
            Kind::Numbering,
            Kind::Footnotes,
            Kind::Endnotes,
        ]
        .into_iter()
        .any(|kind| kinds.contains(kind))
    }

    fn produce<'f>(&mut self, docx: &'f Docx) -> Result<Option<Node<'f>>> {
        loop {
            match self.state {
                State::Done => return Ok(None),
                State::RootStyles => {
                    self.state = State::RootStories(Role::Header, 0);
                    if self.kinds.contains(Kind::Styles) && styles(docx)?.is_some() {
                        return Ok(Some(node(
                            docx,
                            Kind::Styles,
                            PartValue::Styles,
                            top(Kind::Styles, 0),
                        )));
                    }
                },
                State::RootStories(role, next) => {
                    let kind = if role == Role::Header {
                        Kind::Header
                    } else {
                        Kind::Footer
                    };
                    let indices = if self.kinds.contains(kind) {
                        role_stories(docx, role)?
                    } else {
                        Vec::new()
                    };
                    if let Some(index) = indices.get(next) {
                        self.state = State::RootStories(role, next + 1);
                        return Ok(Some(node(
                            docx,
                            kind,
                            PartValue::Story(ordinal(*index)),
                            top(kind, next),
                        )));
                    }
                    self.state = if role == Role::Header {
                        State::RootStories(Role::Footer, 0)
                    } else {
                        State::RootNumbering
                    };
                },
                State::RootNumbering => {
                    self.state = State::RootNotes(litchi_docx::document::NoteKind::Footnote);
                    if self.kinds.contains(Kind::Numbering) && numbering(docx)?.is_some() {
                        return Ok(Some(node(
                            docx,
                            Kind::Numbering,
                            PartValue::Numbering,
                            top(Kind::Numbering, 0),
                        )));
                    }
                },
                State::RootNotes(note) => {
                    let (kind, next) = match note {
                        litchi_docx::document::NoteKind::Endnote => (Kind::Endnotes, State::Done),
                        _ => (
                            Kind::Footnotes,
                            State::RootNotes(litchi_docx::document::NoteKind::Endnote),
                        ),
                    };
                    self.state = next;
                    if self.kinds.contains(kind) && !notes(docx, note)?.is_empty() {
                        return Ok(Some(node(docx, kind, PartValue::Notes(note), top(kind, 0))));
                    }
                },
                State::Styles(next) => {
                    let count = styles(docx)?.map_or(0, |catalog| catalog.entries().len());
                    if next >= count {
                        self.state = State::Done;
                        continue;
                    }
                    self.state = State::Styles(next + 1);
                    return Ok(Some(node(
                        docx,
                        Kind::Style,
                        PartValue::Style(ordinal(next)),
                        self.address.with(Step::new(Kind::Style, ordinal(next))),
                    )));
                },
                State::AbstractNums(next) => {
                    let count = numbering(docx)?.map_or(0, |catalog| catalog.abstract_nums().len());
                    if next >= count {
                        self.state = if self.kinds.contains(Kind::Num) {
                            State::Nums(0)
                        } else {
                            State::Done
                        };
                        continue;
                    }
                    self.state = State::AbstractNums(next + 1);
                    return Ok(Some(node(
                        docx,
                        Kind::AbstractNum,
                        PartValue::AbstractNum(ordinal(next)),
                        self.address
                            .with(Step::new(Kind::AbstractNum, ordinal(next))),
                    )));
                },
                State::Nums(next) => {
                    let count = numbering(docx)?.map_or(0, |catalog| catalog.nums().len());
                    if next >= count {
                        self.state = State::Done;
                        continue;
                    }
                    self.state = State::Nums(next + 1);
                    return Ok(Some(node(
                        docx,
                        Kind::Num,
                        PartValue::Num(ordinal(next)),
                        self.address.with(Step::new(Kind::Num, ordinal(next))),
                    )));
                },
                State::Levels(definition, next) => {
                    let count = numbering(docx)?
                        .and_then(|catalog| catalog.abstract_nums().get(definition as usize))
                        .map_or(0, |definition| definition.levels().len());
                    if next >= count {
                        self.state = State::Done;
                        continue;
                    }
                    self.state = State::Levels(definition, next + 1);
                    return Ok(Some(node(
                        docx,
                        Kind::Level,
                        PartValue::Level(definition, ordinal(next)),
                        self.address.with(Step::new(Kind::Level, ordinal(next))),
                    )));
                },
                State::Notes(kind, next) => {
                    if next >= notes(docx, note_kind(kind))?.len() {
                        self.state = State::Done;
                        continue;
                    }
                    self.state = State::Notes(kind, next + 1);
                    return Ok(Some(note_node(docx, kind, next)));
                },
            }
        }
    }
}

impl Children {
    /// The next child, read through `docx`.
    pub(super) fn next_node<'f>(&mut self, docx: &'f Docx) -> Option<Result<Node<'f>>> {
        match self.produce(docx) {
            Ok(node) => node.map(Ok),
            Err(error) => {
                self.state = State::Done;
                Some(Err(error))
            },
        }
    }
}

/// How many part children the root lists.
pub(super) fn root_count(docx: &Docx) -> Result<usize> {
    Ok(usize::from(styles(docx)?.is_some())
        + stories(docx)?.len()
        + usize::from(numbering(docx)?.is_some())
        + usize::from(!notes(docx, litchi_docx::document::NoteKind::Footnote)?.is_empty())
        + usize::from(!notes(docx, litchi_docx::document::NoteKind::Endnote)?.is_empty()))
}

/// How many children a part node lists.
pub(super) fn child_count(docx: &Docx, value: PartValue) -> Result<usize> {
    Ok(match value {
        PartValue::Styles => styles(docx)?.map_or(0, |catalog| catalog.entries().len()),
        PartValue::Numbering => numbering(docx)?.map_or(0, |catalog| {
            catalog.abstract_nums().len() + catalog.nums().len()
        }),
        PartValue::AbstractNum(definition) => numbering(docx)?
            .and_then(|catalog| catalog.abstract_nums().get(definition as usize))
            .map_or(0, |definition| definition.levels().len()),
        PartValue::Notes(kind) => notes(docx, kind)?.len(),
        _ => 0,
    })
}

/// The path step of a style node: the style's identifier.
pub(super) fn style_step<'f>(docx: &'f Docx, index: u32) -> Option<PathStep<'f>> {
    let catalog = docx.parts().styles.get()?.as_ref()?;
    catalog
        .entries()
        .get(index as usize)
        .map(|entry| PathStep::new(entry.id()))
}

// ---------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------

/// The text of a part node: a style's display name (its id when it has
/// none), a story's text, a note's text without its reference mark.
pub(super) fn text(docx: &Docx, value: PartValue) -> Result<Option<String>> {
    Ok(match value {
        PartValue::Style(index) => styles(docx)?
            .and_then(|catalog| catalog.entries().get(index as usize))
            .map(|entry| {
                entry
                    .name()
                    .filter(|name| !name.is_empty())
                    .unwrap_or(entry.id())
                    .to_owned()
            }),
        PartValue::Story(index) => Some(story_text(&story_xml(docx, index as usize)?)?),
        PartValue::Note(_, _) => match value {
            PartValue::Note(kind, index) => notes(docx, kind)?
                .get(index as usize)
                .map(|note| note.run_text().map_err(Error::from))
                .transpose()?,
            _ => None,
        },
        _ => None,
    })
}

// ---------------------------------------------------------------------------
// Property tables
// ---------------------------------------------------------------------------

/// A part property's write declaration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum PartSetter {
    /// A style attribute or `CT_Style` child.
    Style(StyleKey),
    /// A paragraph property of a style, a level or a story.
    Paragraph(super::setters::Setter),
    /// A run property of a style, a level or a story.
    Run(super::setters::Setter),
    /// A level value.
    Level(LevelKey),
    /// An abstract definition value.
    AbstractNum(AbstractKey),
    /// An instance value.
    Num(NumKey),
    /// A `level<N>.*` key of a new definition (OfficeCLI's "Mode A",
    /// record 0325): the value of level `level` (0–8).
    LevelSeed { level: u8, key: SeedKey },
    /// A story's or note's text.
    Text,
    /// A story's section kind.
    StoryKind,
    /// A field a new story holds.
    StoryField,
}

/// The style keys that are not paragraph or run properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum StyleKey {
    Id,
    Name,
    Type,
    BasedOn,
    Next,
    Linked,
    Flag(Flag),
    UiPriority,
    CustomStyle,
}

/// The level keys that are not run properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum LevelKey {
    Ilvl,
    Format,
    Text,
    Start,
    Indent,
    Hanging,
    Justification,
    Suffix,
    Restart,
    Legal,
}

/// The abstract definition keys.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum AbstractKey {
    Id,
    Type,
    Name,
    StyleLink,
    NumStyleLink,
    /// A level-0 seed of a new definition (`format`, `text`, `start`,
    /// `indent`).
    Seed(LevelKey),
}

/// The instance keys.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum NumKey {
    Id,
    AbstractNumId,
    /// `startOverride.N`, or `start` for level 0.
    StartOverride(u8),
    Continue,
    /// A key of the definition a new instance creates without an
    /// `abstractNumId` (OfficeCLI's "Mode A", record 0325).
    Template(AbstractKey),
}

/// What a `level<N>.*` key sets on its level (record 0325).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SeedKey {
    /// A level value (`format`, `text`, `start`, `indent`, `hanging`,
    /// `justification`, `suff`).
    Level(LevelKey),
    /// A run property of the level's marker (`font`, `size`, `color`,
    /// `bold`, `italic`).
    Run(super::setters::Setter),
}

/// A style's `type`: `ST_StyleType`.
pub(crate) static STYLE_TYPE: Tokens = Tokens::new(
    "style type",
    &["paragraph", "character", "table", "numbering"],
    &[("char", 1), ("para", 0)],
);

/// A header's or footer's `type`: `ST_HdrFtr`.
pub(crate) static STORY_TYPE: Tokens =
    Tokens::new("header type", &["default", "first", "even"], &[]);

/// A level's `justification`: `ST_Jc` as levels use it.
pub(crate) static LEVEL_JUSTIFICATION: Tokens = Tokens::new(
    "level justification",
    &["left", "center", "right", "start", "end"],
    &[],
);

/// A level's `suff`: `ST_LevelSuffix`.
pub(crate) static LEVEL_SUFFIX: Tokens =
    Tokens::new("level suffix", &["tab", "space", "nothing"], &[("none", 2)]);

/// A definition's `type`: `ST_MultiLevelType`.
pub(crate) static MULTI_LEVEL_TYPE: Tokens = Tokens::new(
    "multi-level type",
    &["hybridMultilevel", "multilevel", "singleLevel"],
    &[("hybrid", 0), ("multi", 1), ("single", 2)],
);

/// A new story's field: OfficeCLI's `field` values.
pub(crate) static STORY_FIELD: Tokens = Tokens::new(
    "story field",
    &[
        "page", "numpages", "date", "author", "title", "time", "filename",
    ],
    &[("pagenum", 0), ("pagenumber", 0)],
);

const fn meta(
    name: &'static str,
    ty: ValueType,
    access: Access<PartSetter>,
    constraint: Constraint,
    clear: bool,
) -> Entry<PartSetter> {
    let spec = PropSpec::new(name, ty);
    Entry {
        spec,
        access,
        constraint,
        clear,
    }
}

/// A read-only key.
const fn read(name: &'static str, ty: ValueType) -> Entry<PartSetter> {
    Entry::read(PropSpec::new(name, ty))
}

/// A key `get`, `set` and `add` all accept.
const fn rsa(
    name: &'static str,
    ty: ValueType,
    setter: PartSetter,
    constraint: Constraint,
) -> Entry<PartSetter> {
    meta(name, ty, Access::SetAndAdd(setter), constraint, true)
}

/// A key `get` and `add` accept.
const fn ra(
    name: &'static str,
    ty: ValueType,
    setter: PartSetter,
    constraint: Constraint,
) -> Entry<PartSetter> {
    meta(name, ty, Access::Add(setter), constraint, false)
}

/// A key only `add` accepts.
const fn add_only(
    name: &'static str,
    ty: ValueType,
    setter: PartSetter,
    constraint: Constraint,
) -> Entry<PartSetter> {
    let mut entry = meta(name, ty, Access::Add(setter), constraint, false);
    entry.spec.access.get = false;
    entry
}

/// A key `set` and `add` accept and `get` does not read.
const fn write_only(
    name: &'static str,
    ty: ValueType,
    setter: PartSetter,
    constraint: Constraint,
) -> Entry<PartSetter> {
    let mut entry = meta(name, ty, Access::SetAndAdd(setter), constraint, false);
    entry.spec.access.get = false;
    entry
}

/// A paragraph property of a style, named `name`, read and written as the
/// paragraph key `from`.
const fn paragraph(
    name: &'static str,
    from: &'static str,
    setter: super::setters::Setter,
) -> Entry<PartSetter> {
    let spec = PARAGRAPH_PROPS[index_of(PARAGRAPH_PROPS, from)];
    Entry {
        spec: PropSpec::new(name, spec.ty),
        access: Access::SetAndAdd(PartSetter::Paragraph(setter)),
        constraint: spec.admission.constraint,
        clear: true,
    }
}

/// A run property of a style, a level or a story, named `name`, read and
/// written as the run key `from`.
const fn run(
    name: &'static str,
    from: &'static str,
    setter: super::setters::Setter,
) -> Entry<PartSetter> {
    let spec = RUN_PROPS[index_of(RUN_PROPS, from)];
    Entry {
        spec: PropSpec::new(name, spec.ty),
        access: Access::SetAndAdd(PartSetter::Run(setter)),
        constraint: spec.admission.constraint,
        clear: true,
    }
}

use super::setters::Setter as S;

/// A `level<N>.*` level value of a new definition: add only.
const fn seed(
    name: &'static str,
    level: u8,
    key: LevelKey,
    ty: ValueType,
    constraint: Constraint,
) -> Entry<PartSetter> {
    add_only(
        name,
        ty,
        PartSetter::LevelSeed {
            level,
            key: SeedKey::Level(key),
        },
        constraint,
    )
}

/// A `level<N>.*` marker run property of a new definition: add only.
const fn seed_run(
    name: &'static str,
    level: u8,
    from: &'static str,
    setter: S,
) -> Entry<PartSetter> {
    let spec = RUN_PROPS[index_of(RUN_PROPS, from)];
    let mut spec = PropSpec::new(name, spec.ty);
    spec.access.get = false;
    Entry {
        spec,
        access: Access::Add(PartSetter::LevelSeed {
            level,
            key: SeedKey::Run(setter),
        }),
        constraint: RUN_PROPS[index_of(RUN_PROPS, from)].admission.constraint,
        clear: false,
    }
}

/// A numbering table whose entries are followed by OfficeCLI's
/// `level<N>.*` keys for levels 0–8 (record 0325).
macro_rules! seeded_registry {
    ($vis:vis $name:ident, $table:ident { $($entry:expr),* $(,)? }) => {
        seeded_registry!(@build $vis $name, $table [$($entry),*] [0 1 2 3 4 5 6 7 8]);
    };
    (@build $vis:vis $name:ident, $table:ident [$($entry:expr),*] [$($n:literal)*]) => {
        $vis const $table: crate::registry::Table<
            PartSetter,
            { [$(stringify!($entry)),*].len() + 12 * [$(stringify!($n)),*].len() },
            0,
        > = crate::registry::Table::new(
            [
                $($entry,)*
                $(
                    seed(concat!("level", $n, ".format"), $n, LevelKey::Format, ValueType::Text, Constraint::NotEmpty),
                    seed(concat!("level", $n, ".text"), $n, LevelKey::Text, ValueType::Text, Constraint::None),
                    seed(concat!("level", $n, ".start"), $n, LevelKey::Start, ValueType::Int, Constraint::Int { min: 0, max: 32_767 }),
                    seed(concat!("level", $n, ".indent"), $n, LevelKey::Indent, ValueType::Int, Constraint::Int { min: -31_680, max: 31_680 }),
                    seed(concat!("level", $n, ".hanging"), $n, LevelKey::Hanging, ValueType::Int, Constraint::Int { min: 0, max: 31_680 }),
                    seed(concat!("level", $n, ".justification"), $n, LevelKey::Justification, ValueType::Token(&LEVEL_JUSTIFICATION), Constraint::None),
                    seed(concat!("level", $n, ".suff"), $n, LevelKey::Suffix, ValueType::Token(&LEVEL_SUFFIX), Constraint::None),
                    seed_run(concat!("level", $n, ".font"), $n, "font.latin", S::RunFontLatin),
                    seed_run(concat!("level", $n, ".size"), $n, "size", S::RunSize),
                    seed_run(concat!("level", $n, ".color"), $n, "color", S::RunColor),
                    seed_run(concat!("level", $n, ".bold"), $n, "bold", S::RunBold),
                    seed_run(concat!("level", $n, ".italic"), $n, "italic", S::RunItalic),
                )*
            ],
            [],
        );
        $vis const $name: &[crate::node::spec::PropSpec] = &$table.readable();
    };
}

// The style node, in `WordHandler.Query` `/styles/<id>` order: identity
// and UI flags, then the run properties, then the paragraph properties.
crate::registry::registry! {
    pub(crate) STYLE_PROPS, STYLE_REGISTRY: PartSetter {
    ra("id", ValueType::Text, PartSetter::Style(StyleKey::Id), Constraint::NotEmpty),
    rsa("name", ValueType::Text, PartSetter::Style(StyleKey::Name), Constraint::None),
    ra("type", ValueType::Token(&STYLE_TYPE), PartSetter::Style(StyleKey::Type), Constraint::None),
    read("default", ValueType::Bool),
    rsa("basedOn", ValueType::Text, PartSetter::Style(StyleKey::BasedOn), Constraint::NotEmpty),
    read("basedOn.path", ValueType::Text),
    rsa("next", ValueType::Text, PartSetter::Style(StyleKey::Next), Constraint::NotEmpty),
    rsa("linked", ValueType::Text, PartSetter::Style(StyleKey::Linked), Constraint::NotEmpty),
    rsa("autoRedefine", ValueType::Bool, PartSetter::Style(StyleKey::Flag(Flag::AutoRedefine)), Constraint::None),
    rsa("hidden", ValueType::Bool, PartSetter::Style(StyleKey::Flag(Flag::Hidden)), Constraint::None),
    rsa("qFormat", ValueType::Bool, PartSetter::Style(StyleKey::Flag(Flag::QuickFormat)), Constraint::None),
    rsa("uiPriority", ValueType::Int, PartSetter::Style(StyleKey::UiPriority), Constraint::Int { min: 0, max: 99 }),
    rsa("semiHidden", ValueType::Bool, PartSetter::Style(StyleKey::Flag(Flag::SemiHidden)), Constraint::None),
    rsa("unhideWhenUsed", ValueType::Bool, PartSetter::Style(StyleKey::Flag(Flag::UnhideWhenUsed)), Constraint::None),
    rsa("locked", ValueType::Bool, PartSetter::Style(StyleKey::Flag(Flag::Locked)), Constraint::None),
    ra("customStyle", ValueType::Bool, PartSetter::Style(StyleKey::CustomStyle), Constraint::None),
    run("font", "font.latin", S::RunFontLatin),
    run("font.ascii", "font.ascii", S::RunFontAscii),
    run("font.hAnsi", "font.hAnsi", S::RunFontHAnsi),
    run("font.ea", "font.ea", S::RunFontEastAsia),
    run("font.cs", "font.cs", S::RunFontComplexScript),
    run("font.asciiTheme", "font.asciiTheme", S::RunFontAsciiTheme),
    run("font.hAnsiTheme", "font.hAnsiTheme", S::RunFontHAnsiTheme),
    run("font.eaTheme", "font.eaTheme", S::RunFontEastAsiaTheme),
    run("font.csTheme", "font.csTheme", S::RunFontComplexScriptTheme),
    run("size", "size", S::RunSize),
    run("size.cs", "size.cs", S::RunSizeComplexScript),
    run("bold", "bold", S::RunBold),
    run("bold.cs", "bold.cs", S::RunBoldComplexScript),
    run("italic", "italic", S::RunItalic),
    run("italic.cs", "italic.cs", S::RunItalicComplexScript),
    run("color", "color", S::RunColor),
    run("underline", "underline", S::RunUnderline),
    run("underline.color", "underline.color", S::RunUnderlineColor),
    run("strike", "strike", S::RunStrike),
    run("dstrike", "dstrike", S::RunDoubleStrike),
    run("caps", "caps", S::RunCaps),
    run("smallCaps", "smallcaps", S::RunSmallCaps),
    run("vanish", "vanish", S::RunVanish),
    run("highlight", "highlight", S::RunHighlight),
    run("superscript", "superscript", S::RunSuperscript),
    run("subscript", "subscript", S::RunSubscript),
    run("charSpacing", "charSpacing", S::RunCharacterSpacing),
    run("lang.latin", "lang.latin", S::RunLanguageLatin),
    run("lang.ea", "lang.ea", S::RunLanguageEastAsia),
    run("lang.cs", "lang.cs", S::RunLanguageComplexScript),
    paragraph("align", "align", S::ParagraphAlign),
    paragraph("direction", "direction", S::ParagraphDirection),
    paragraph("spaceBefore", "spaceBefore", S::ParagraphSpaceBefore),
    paragraph("spaceAfter", "spaceAfter", S::ParagraphSpaceAfter),
    paragraph("spaceBeforeAuto", "spaceBeforeAuto", S::ParagraphSpaceBeforeAuto),
    paragraph("spaceAfterAuto", "spaceAfterAuto", S::ParagraphSpaceAfterAuto),
    paragraph("lineSpacing", "lineSpacing", S::ParagraphLineSpacing),
    paragraph("lineRule", "lineRule", S::ParagraphLineRule),
    paragraph("spaceBeforeLines", "spaceBeforeLines", S::ParagraphSpaceBeforeLines),
    paragraph("spaceAfterLines", "spaceAfterLines", S::ParagraphSpaceAfterLines),
    paragraph("firstLineIndent", "firstLineIndent", S::ParagraphFirstLineIndent),
    paragraph("hangingIndent", "hangingIndent", S::ParagraphHangingIndent),
    paragraph("leftIndent", "indent", S::ParagraphIndent),
    paragraph("rightIndent", "rightIndent", S::ParagraphRightIndent),
    paragraph("firstLineChars", "firstLineChars", S::ParagraphFirstLineChars),
    paragraph("hangingChars", "hangingChars", S::ParagraphHangingChars),
    paragraph("leftChars", "leftChars", S::ParagraphLeftChars),
    paragraph("rightChars", "rightChars", S::ParagraphRightChars),
    paragraph("outlineLvl", "outlineLvl", S::ParagraphOutlineLevel),
    paragraph("keepNext", "keepNext", S::ParagraphKeepNext),
    paragraph("keepLines", "keepLines", S::ParagraphKeepLines),
    paragraph("pageBreakBefore", "pageBreakBefore", S::ParagraphPageBreakBefore),
    paragraph("widowControl", "widowControl", S::ParagraphWidowControl),
    paragraph("contextualSpacing", "contextualSpacing", S::ParagraphContextualSpacing),
    paragraph("numId", "numId", S::ParagraphNumId),
    paragraph("numLevel", "numLevel", S::ParagraphNumLevel),
    }
}

// `/styles`: the number of styles.
crate::registry::registry! {
    pub(crate) STYLES_PROPS, STYLES_REGISTRY: PartSetter {
    read("count", ValueType::Int),
    }
}

// `/numbering`: the numbers of definitions and instances.
crate::registry::registry! {
    pub(crate) NUMBERING_PROPS, NUMBERING_REGISTRY: PartSetter {
    read("abstractNumCount", ValueType::Int),
    read("numCount", ValueType::Int),
    }
}

// `/numbering/abstractNum[@id=N]`, in its order; the level-0 seeds are
// add-only, as OfficeCLI's bare `format`, `text`, `start` and `indent`, and
// so are its `level<N>.*` keys (record 0325).
seeded_registry! {
    pub(crate) ABSTRACT_NUM_PROPS, ABSTRACT_NUM_REGISTRY {
    ra("id", ValueType::Int, PartSetter::AbstractNum(AbstractKey::Id), Constraint::Int { min: 0, max: i32::MAX as i64 }),
    rsa("type", ValueType::Token(&MULTI_LEVEL_TYPE), PartSetter::AbstractNum(AbstractKey::Type), Constraint::None),
    rsa("name", ValueType::Text, PartSetter::AbstractNum(AbstractKey::Name), Constraint::NotEmpty),
    rsa("styleLink", ValueType::Text, PartSetter::AbstractNum(AbstractKey::StyleLink), Constraint::NotEmpty),
    rsa("numStyleLink", ValueType::Text, PartSetter::AbstractNum(AbstractKey::NumStyleLink), Constraint::NotEmpty),
    add_only("format", ValueType::Text, PartSetter::AbstractNum(AbstractKey::Seed(LevelKey::Format)), Constraint::NotEmpty),
    add_only("text", ValueType::Text, PartSetter::AbstractNum(AbstractKey::Seed(LevelKey::Text)), Constraint::None),
    add_only("start", ValueType::Int, PartSetter::AbstractNum(AbstractKey::Seed(LevelKey::Start)), Constraint::Int { min: 0, max: 32_767 }),
    add_only("indent", ValueType::Int, PartSetter::AbstractNum(AbstractKey::Seed(LevelKey::Indent)), Constraint::Int { min: -31_680, max: 31_680 }),
    }
}

// `/numbering/num[@id=N]`. Without an `abstractNumId`, the definition keys
// and `level<N>.*` create the definition too (OfficeCLI's "Mode A", record
// 0325).
seeded_registry! {
    pub(crate) NUM_PROPS, NUM_REGISTRY {
    ra("id", ValueType::Int, PartSetter::Num(NumKey::Id), Constraint::Int { min: 1, max: i32::MAX as i64 }),
    rsa("abstractNumId", ValueType::Int, PartSetter::Num(NumKey::AbstractNumId), Constraint::Int { min: 0, max: i32::MAX as i64 }),
    ra("startOverride.0", ValueType::Int, PartSetter::Num(NumKey::StartOverride(0)), Constraint::Int { min: 0, max: 32_767 }),
    ra("startOverride.1", ValueType::Int, PartSetter::Num(NumKey::StartOverride(1)), Constraint::Int { min: 0, max: 32_767 }),
    ra("startOverride.2", ValueType::Int, PartSetter::Num(NumKey::StartOverride(2)), Constraint::Int { min: 0, max: 32_767 }),
    ra("startOverride.3", ValueType::Int, PartSetter::Num(NumKey::StartOverride(3)), Constraint::Int { min: 0, max: 32_767 }),
    ra("startOverride.4", ValueType::Int, PartSetter::Num(NumKey::StartOverride(4)), Constraint::Int { min: 0, max: 32_767 }),
    ra("startOverride.5", ValueType::Int, PartSetter::Num(NumKey::StartOverride(5)), Constraint::Int { min: 0, max: 32_767 }),
    ra("startOverride.6", ValueType::Int, PartSetter::Num(NumKey::StartOverride(6)), Constraint::Int { min: 0, max: 32_767 }),
    ra("startOverride.7", ValueType::Int, PartSetter::Num(NumKey::StartOverride(7)), Constraint::Int { min: 0, max: 32_767 }),
    ra("startOverride.8", ValueType::Int, PartSetter::Num(NumKey::StartOverride(8)), Constraint::Int { min: 0, max: 32_767 }),
    add_only("start", ValueType::Int, PartSetter::Num(NumKey::StartOverride(0)), Constraint::Int { min: 0, max: 32_767 }),
    add_only("continue", ValueType::Bool, PartSetter::Num(NumKey::Continue), Constraint::None),
    add_only("type", ValueType::Token(&MULTI_LEVEL_TYPE), PartSetter::Num(NumKey::Template(AbstractKey::Type)), Constraint::None),
    add_only("name", ValueType::Text, PartSetter::Num(NumKey::Template(AbstractKey::Name)), Constraint::NotEmpty),
    add_only("styleLink", ValueType::Text, PartSetter::Num(NumKey::Template(AbstractKey::StyleLink)), Constraint::NotEmpty),
    add_only("numStyleLink", ValueType::Text, PartSetter::Num(NumKey::Template(AbstractKey::NumStyleLink)), Constraint::NotEmpty),
    add_only("format", ValueType::Text, PartSetter::Num(NumKey::Template(AbstractKey::Seed(LevelKey::Format))), Constraint::NotEmpty),
    add_only("text", ValueType::Text, PartSetter::Num(NumKey::Template(AbstractKey::Seed(LevelKey::Text))), Constraint::None),
    add_only("indent", ValueType::Int, PartSetter::Num(NumKey::Template(AbstractKey::Seed(LevelKey::Indent))), Constraint::Int { min: -31_680, max: 31_680 }),
    }
}

// `/numbering/abstractNum[@id=N]/level[L]`, in its order.
crate::registry::registry! {
    pub(crate) LEVEL_PROPS, LEVEL_REGISTRY: PartSetter {
    ra("ilvl", ValueType::Int, PartSetter::Level(LevelKey::Ilvl), Constraint::Int { min: 0, max: 8 }),
    rsa("start", ValueType::Int, PartSetter::Level(LevelKey::Start), Constraint::Int { min: 0, max: 32_767 }),
    rsa("format", ValueType::Text, PartSetter::Level(LevelKey::Format), Constraint::NotEmpty),
    rsa("lvlText", ValueType::Text, PartSetter::Level(LevelKey::Text), Constraint::None),
    rsa("justification", ValueType::Token(&LEVEL_JUSTIFICATION), PartSetter::Level(LevelKey::Justification), Constraint::None),
    rsa("suff", ValueType::Token(&LEVEL_SUFFIX), PartSetter::Level(LevelKey::Suffix), Constraint::None),
    rsa("lvlRestart", ValueType::Int, PartSetter::Level(LevelKey::Restart), Constraint::Int { min: 0, max: 9 }),
    rsa("isLgl", ValueType::Bool, PartSetter::Level(LevelKey::Legal), Constraint::None),
    rsa("indent", ValueType::Int, PartSetter::Level(LevelKey::Indent), Constraint::Int { min: -31_680, max: 31_680 }),
    rsa("hanging", ValueType::Int, PartSetter::Level(LevelKey::Hanging), Constraint::Int { min: 0, max: 31_680 }),
    run("font", "font.latin", S::RunFontLatin),
    run("size", "size", S::RunSize),
    run("color", "color", S::RunColor),
    run("bold", "bold", S::RunBold),
    run("italic", "italic", S::RunItalic),
    }
}

// `/header[N]` and `/footer[N]`, in `GetHeaderNode` order.
crate::registry::registry! {
    pub(crate) STORY_PROPS, STORY_REGISTRY: PartSetter {
    rsa("type", ValueType::Token(&STORY_TYPE), PartSetter::StoryKind, Constraint::None),
    rsa("text", ValueType::Text, PartSetter::Text, Constraint::None),
    run("font", "font.latin", S::RunFontLatin),
    run("size", "size", S::RunSize),
    run("bold", "bold", S::RunBold),
    run("italic", "italic", S::RunItalic),
    run("color", "color", S::RunColor),
    run("underline", "underline", S::RunUnderline),
    paragraph("align", "align", S::ParagraphAlign),
    write_only("direction", ValueType::Token(&super::spec::DIRECTION), PartSetter::Paragraph(S::ParagraphDirection), Constraint::None),
    add_only("field", ValueType::Token(&STORY_FIELD), PartSetter::StoryField, Constraint::None),
    }
}

// `/footnotes/footnote[N]` and `/endnotes/endnote[N]` (`FootnoteToNode`).
crate::registry::registry! {
    pub(crate) NOTE_PROPS, NOTE_REGISTRY: PartSetter {
    read("id", ValueType::Int),
    rsa("text", ValueType::Text, PartSetter::Text, Constraint::None),
    read("font", RUN_PROPS[index_of(RUN_PROPS, "font.latin")].ty),
    read("size", RUN_PROPS[index_of(RUN_PROPS, "size")].ty),
    read("bold", RUN_PROPS[index_of(RUN_PROPS, "bold")].ty),
    read("italic", RUN_PROPS[index_of(RUN_PROPS, "italic")].ty),
    read("color", RUN_PROPS[index_of(RUN_PROPS, "color")].ty),
    }
}

// ---------------------------------------------------------------------------
// Values
// ---------------------------------------------------------------------------

/// A part node's property set, read once per `format` request.
pub(crate) enum PartProps<'a> {
    None,
    Count(usize, usize),
    Style {
        docx: &'a Docx,
        index: usize,
        paragraph: Box<direct::Paragraph>,
        run: Box<direct::Run>,
    },
    AbstractNum {
        docx: &'a Docx,
        index: usize,
    },
    Num {
        docx: &'a Docx,
        index: usize,
    },
    Level {
        docx: &'a Docx,
        definition: usize,
        index: usize,
        paragraph: Box<direct::Paragraph>,
        run: Box<direct::Run>,
    },
    Story {
        kind: Option<StoryKind>,
        text: String,
        paragraph: Option<Box<direct::Paragraph>>,
        run: Option<Box<direct::Run>>,
    },
    Note {
        id: u32,
        text: String,
        run: Option<Box<direct::Run>>,
    },
}

impl std::fmt::Debug for PartProps<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("PartProps")
    }
}

/// The property set of a part node.
pub(super) fn props<'a>(docx: &'a Docx, value: PartValue) -> Result<PartProps<'a>> {
    Ok(match value {
        PartValue::Styles => PartProps::Count(child_count(docx, value)?, 0),
        PartValue::Numbering => match numbering(docx)? {
            Some(catalog) => PartProps::Count(catalog.abstract_nums().len(), catalog.nums().len()),
            None => PartProps::None,
        },
        PartValue::Style(index) => {
            let Some(catalog) = styles(docx)? else {
                return Ok(PartProps::None);
            };
            let Some(entry) = catalog.entries().get(index as usize) else {
                return Ok(PartProps::None);
            };
            PartProps::Style {
                docx,
                index: index as usize,
                paragraph: Box::new(catalog.paragraph_properties(entry).map_err(Error::from)?),
                run: Box::new(catalog.run_properties(entry).map_err(Error::from)?),
            }
        },
        PartValue::AbstractNum(index) => PartProps::AbstractNum {
            docx,
            index: index as usize,
        },
        PartValue::Num(index) => PartProps::Num {
            docx,
            index: index as usize,
        },
        PartValue::Level(definition, index) => {
            let Some(catalog) = numbering(docx)? else {
                return Ok(PartProps::None);
            };
            let Some(level) = catalog
                .abstract_nums()
                .get(definition as usize)
                .and_then(|definition| definition.levels().get(index as usize))
            else {
                return Ok(PartProps::None);
            };
            PartProps::Level {
                docx,
                definition: definition as usize,
                index: index as usize,
                paragraph: Box::new(
                    catalog
                        .level_paragraph_properties(level)
                        .map_err(Error::from)?,
                ),
                run: Box::new(catalog.level_run_properties(level).map_err(Error::from)?),
            }
        },
        PartValue::Story(index) => {
            let xml = story_xml(docx, index as usize)?;
            let stories = stories(docx)?;
            let story = &stories[index as usize];
            let kind = story_kinds(docx)?.get(index as usize).copied().flatten();
            let parsed = litchi_docx::header_footer::Story::from_shared_bytes(
                Arc::clone(&xml),
                kind.unwrap_or_default(),
            )
            .map_err(Error::from)?;
            let paragraphs = parsed.paragraphs().map_err(Error::from)?;
            let paragraph = match paragraphs.first() {
                Some(first) => Some(Box::new(first.direct_properties().map_err(Error::from)?)),
                None => None,
            };
            let mut run = None;
            for paragraph in &paragraphs {
                if let Some(first) = paragraph.addressable_runs().map_err(Error::from)?.first() {
                    run = Some(Box::new(first.direct_properties().map_err(Error::from)?));
                    break;
                }
            }
            let _ = story;
            PartProps::Story {
                kind,
                text: story_text(&xml)?,
                paragraph,
                run,
            }
        },
        PartValue::Note(kind, index) => {
            let notes = notes(docx, kind)?;
            let Some(note) = notes.get(index as usize) else {
                return Ok(PartProps::None);
            };
            let mut run = None;
            'paragraphs: for paragraph in note.paragraphs().map_err(Error::from)? {
                for candidate in paragraph.addressable_runs().map_err(Error::from)? {
                    let holds_text = candidate
                        .contents()
                        .map_err(Error::from)?
                        .iter()
                        .any(|content| matches!(content, litchi_docx::RunContent::Text(_)));
                    if holds_text {
                        run = Some(Box::new(
                            candidate.direct_properties().map_err(Error::from)?,
                        ));
                        break 'paragraphs;
                    }
                }
            }
            PartProps::Note {
                id: note.id(),
                text: note.run_text().map_err(Error::from)?,
                run,
            }
        },
        PartValue::Notes(_) => PartProps::None,
    })
}

/// The kind's property table.
pub(super) const fn table(kind: Kind) -> &'static [PropSpec] {
    match kind {
        Kind::Styles => STYLES_PROPS,
        Kind::Style => STYLE_PROPS,
        Kind::Numbering => NUMBERING_PROPS,
        Kind::AbstractNum => ABSTRACT_NUM_PROPS,
        Kind::Num => NUM_PROPS,
        Kind::Level => LEVEL_PROPS,
        Kind::Header | Kind::Footer => STORY_PROPS,
        Kind::Footnote | Kind::Endnote => NOTE_PROPS,
        _ => &[],
    }
}

fn text_value<'a>(text: &str) -> NodeValue<'a> {
    NodeValue::Text(Cow::Owned(text.to_owned()))
}

fn int_text(text: Option<&str>) -> Option<NodeValue<'static>> {
    text.and_then(|text| text.trim().parse::<i64>().ok())
        .map(NodeValue::Int)
}

/// Where a part table entry's value comes from: the part's own data, or
/// the run or paragraph table's entry at an index, computed at compile
/// time so a `format` read does no name lookup.
#[derive(Clone, Copy, PartialEq, Eq)]
enum From {
    /// The node's own key, by its position in the kind's own-key list
    /// (record 0326: a read matches the position, not the name).
    Own(u8),
    /// A run key; `font` falls back to the ASCII slot.
    Run(usize),
    Font,
    Paragraph(usize),
}

const fn same(left: &str, right: &str) -> bool {
    let (left, right) = (left.as_bytes(), right.as_bytes());
    if left.len() != right.len() {
        return false;
    }
    let mut index = 0;
    while index < left.len() {
        if left[index] != right[index] {
            return false;
        }
        index += 1;
    }
    true
}

const fn find(props: &[PropSpec], name: &str) -> Option<usize> {
    let mut index = 0;
    while index < props.len() {
        if same(props[index].name, name) {
            return Some(index);
        }
        index += 1;
    }
    None
}

/// The run table's `font.latin`, which a part's `font` reads first.
const FONT_LATIN: usize = match find(RUN_PROPS, "font.latin") {
    Some(at) => at,
    None => panic!("font.latin is a run key"),
};

/// The position of `name` in an own-key list, checked at compile time.
const fn own_at(names: &[&str], name: &str) -> u8 {
    let mut index = 0;
    while index < names.len() {
        if same(names[index], name) {
            #[allow(clippy::cast_possible_truncation, reason = "own-key lists are short")]
            return index as u8;
        }
        index += 1;
    }
    panic!("not an own key")
}

const fn listed(names: &[&str], name: &str) -> bool {
    let mut index = 0;
    while index < names.len() {
        if same(names[index], name) {
            return true;
        }
        index += 1;
    }
    false
}

/// The source of every entry of `table`: `own` keys first, then the run
/// table (`font` is `font.latin`, `smallCaps` is `smallcaps`), then the
/// paragraph table (`leftIndent` is `indent`; `direction` and `align` are
/// paragraph keys).
const fn sources<const N: usize>(table: &[PropSpec], own: &[&str]) -> [From; N] {
    let mut sources = [From::Own(u8::MAX); N];
    let mut index = 0;
    while index < N {
        let name = table[index].name;
        sources[index] = if listed(own, name) {
            From::Own(own_at(own, name))
        } else if same(name, "font") {
            From::Font
        } else if same(name, "direction") || same(name, "align") {
            match find(PARAGRAPH_PROPS, name) {
                Some(at) => From::Paragraph(at),
                None => From::Own(u8::MAX),
            }
        } else if same(name, "leftIndent") {
            match find(PARAGRAPH_PROPS, "indent") {
                Some(at) => From::Paragraph(at),
                None => From::Own(u8::MAX),
            }
        } else {
            let run_name = if same(name, "smallCaps") {
                "smallcaps"
            } else {
                name
            };
            match find(RUN_PROPS, run_name) {
                Some(at) => From::Run(at),
                None => match find(PARAGRAPH_PROPS, name) {
                    Some(at) => From::Paragraph(at),
                    None => From::Own(u8::MAX),
                },
            }
        };
        index += 1;
    }
    sources
}

const STYLE_OWN: &[&str] = &[
    "id",
    "name",
    "type",
    "default",
    "basedOn",
    "basedOn.path",
    "next",
    "linked",
    "autoRedefine",
    "hidden",
    "qFormat",
    "uiPriority",
    "semiHidden",
    "unhideWhenUsed",
    "locked",
    "customStyle",
];
const STYLE_SOURCES: [From; STYLE_PROPS.len()] = sources(STYLE_PROPS, STYLE_OWN);
const LEVEL_OWN: &[&str] = &[
    "ilvl",
    "start",
    "format",
    "lvlText",
    "justification",
    "suff",
    "lvlRestart",
    "isLgl",
    "indent",
    "hanging",
];
const LEVEL_SOURCES: [From; LEVEL_PROPS.len()] = sources(LEVEL_PROPS, LEVEL_OWN);
const STORY_SOURCES: [From; STORY_PROPS.len()] = sources(STORY_PROPS, &["type", "text", "field"]);
const NOTE_SOURCES: [From; NOTE_PROPS.len()] = sources(NOTE_PROPS, &["id", "text"]);

/// The value of a run or paragraph sourced entry.
fn sourced<'a>(
    from: From,
    paragraph: Option<&direct::Paragraph>,
    run: Option<&direct::Run>,
) -> Option<NodeValue<'a>> {
    match from {
        From::Own(_) => None,
        From::Run(at) => super::props::run(run?, at),
        From::Paragraph(at) => super::props::paragraph_value(paragraph?, None, at),
        From::Font => {
            let run = run?;
            super::props::run(run, FONT_LATIN).or_else(|| {
                // The ASCII slot when the ASCII and high-ANSI slots differ.
                let fonts = run.fonts.as_ref()?;
                fonts
                    .ascii
                    .as_deref()
                    .or(fonts.h_ansi.as_deref())
                    .filter(|font| !font.is_empty())
                    .map(text_value)
            })
        },
    }
}

impl<'a> PartProps<'a> {
    /// The value of the property at `index` of the node kind's table.
    pub(crate) fn get(&self, index: usize) -> Option<NodeValue<'a>> {
        match self {
            Self::None => None,
            Self::Count(first, second) => match index {
                0 => Some(NodeValue::Int(i64::try_from(*first).unwrap_or(i64::MAX))),
                1 => Some(NodeValue::Int(i64::try_from(*second).unwrap_or(i64::MAX))),
                _ => None,
            },
            Self::Style {
                docx,
                index: style,
                paragraph,
                run,
            } => style_value(docx, *style, paragraph, run, index),
            Self::AbstractNum { docx, index: at } => abstract_value(docx, *at, index),
            Self::Num { docx, index: at } => num_value(docx, *at, index),
            Self::Level {
                docx,
                definition,
                index: at,
                paragraph,
                run,
            } => level_value(docx, *definition, *at, paragraph, run, index),
            Self::Story {
                kind,
                text,
                paragraph,
                run,
            } => story_value(*kind, text, paragraph.as_deref(), run.as_deref(), index),
            Self::Note { id, text, run } => match NOTE_SOURCES.get(index)? {
                From::Own(_) => match NOTE_PROPS.get(index)?.name {
                    "id" => Some(NodeValue::Int(i64::from(*id))),
                    "text" => Some(text_value(text)),
                    _ => None,
                },
                from => sourced(*from, None, run.as_deref()),
            },
        }
    }
}

fn style_value<'a>(
    docx: &Docx,
    style: usize,
    paragraph: &direct::Paragraph,
    run: &direct::Run,
    index: usize,
) -> Option<NodeValue<'a>> {
    let From::Own(own) = *STYLE_SOURCES.get(index)? else {
        return sourced(*STYLE_SOURCES.get(index)?, Some(paragraph), Some(run));
    };
    let catalog = docx.parts().styles.get()?.as_ref()?;
    let entry = catalog.entries().get(style)?;
    let flag = |flag: Flag| entry.flag(flag).map(NodeValue::Bool);
    const ID: u8 = own_at(STYLE_OWN, "id");
    const NAME: u8 = own_at(STYLE_OWN, "name");
    const TYPE: u8 = own_at(STYLE_OWN, "type");
    const DEFAULT: u8 = own_at(STYLE_OWN, "default");
    const BASED_ON: u8 = own_at(STYLE_OWN, "basedOn");
    const BASED_ON_PATH: u8 = own_at(STYLE_OWN, "basedOn.path");
    const NEXT: u8 = own_at(STYLE_OWN, "next");
    const LINKED: u8 = own_at(STYLE_OWN, "linked");
    const AUTO_REDEFINE: u8 = own_at(STYLE_OWN, "autoRedefine");
    const HIDDEN: u8 = own_at(STYLE_OWN, "hidden");
    const Q_FORMAT: u8 = own_at(STYLE_OWN, "qFormat");
    const UI_PRIORITY: u8 = own_at(STYLE_OWN, "uiPriority");
    const SEMI_HIDDEN: u8 = own_at(STYLE_OWN, "semiHidden");
    const UNHIDE_WHEN_USED: u8 = own_at(STYLE_OWN, "unhideWhenUsed");
    const LOCKED: u8 = own_at(STYLE_OWN, "locked");
    const CUSTOM_STYLE: u8 = own_at(STYLE_OWN, "customStyle");
    match own {
        ID => Some(text_value(entry.id())),
        NAME => entry.name().map(text_value),
        TYPE => entry.kind().map(|kind| NodeValue::Token(kind.to_xml())),
        DEFAULT => entry.is_default().then_some(NodeValue::Bool(true)),
        BASED_ON => entry.based_on().map(text_value),
        BASED_ON_PATH => entry
            .based_on()
            .map(|parent| NodeValue::Text(Cow::Owned(format!("/styles/{parent}")))),
        NEXT => entry.next().map(text_value),
        LINKED => entry.link().map(text_value),
        AUTO_REDEFINE => flag(Flag::AutoRedefine),
        HIDDEN => flag(Flag::Hidden),
        Q_FORMAT => flag(Flag::QuickFormat),
        UI_PRIORITY => entry.ui_priority().map(NodeValue::Int),
        SEMI_HIDDEN => flag(Flag::SemiHidden),
        UNHIDE_WHEN_USED => flag(Flag::UnhideWhenUsed),
        LOCKED => flag(Flag::Locked),
        // OfficeCLI always emits `customStyle`.
        CUSTOM_STYLE => Some(NodeValue::Bool(entry.custom().unwrap_or(false))),
        _ => None,
    }
}

fn abstract_value<'a>(docx: &Docx, at: usize, index: usize) -> Option<NodeValue<'a>> {
    let catalog = docx.parts().numbering.get()?.as_ref()?;
    let definition = catalog.abstract_nums().get(at)?;
    match ABSTRACT_NUM_PROPS.get(index)?.name {
        "id" => Some(NodeValue::Int(i64::from(definition.id()))),
        "type" => definition
            .multi_level_type()
            .and_then(|value| MULTI_LEVEL_TYPE.get(value))
            .map(NodeValue::Token),
        "name" => definition.name().map(text_value),
        "styleLink" => definition.style_link().map(text_value),
        "numStyleLink" => definition.num_style_link().map(text_value),
        _ => None,
    }
}

fn num_value<'a>(docx: &Docx, at: usize, index: usize) -> Option<NodeValue<'a>> {
    let catalog = docx.parts().numbering.get()?.as_ref()?;
    let num = catalog.nums().get(at)?;
    let name = NUM_PROPS.get(index)?.name;
    match name {
        "id" => Some(NodeValue::Int(i64::from(num.id()))),
        "abstractNumId" => num
            .abstract_num_id()
            .map(|id| NodeValue::Int(i64::from(id))),
        _ => {
            let level = name.strip_prefix("startOverride.")?.parse::<u8>().ok()?;
            num.overrides()
                .iter()
                .find(|(ilvl, _)| *ilvl == level)
                .and_then(|(_, start)| int_text(start.as_deref()))
        },
    }
}

fn level_value<'a>(
    docx: &Docx,
    definition: usize,
    at: usize,
    paragraph: &direct::Paragraph,
    run: &direct::Run,
    index: usize,
) -> Option<NodeValue<'a>> {
    let catalog = docx.parts().numbering.get()?.as_ref()?;
    let level = catalog.abstract_nums().get(definition)?.levels().get(at)?;
    let indentation = paragraph.indentation.as_ref();
    let From::Own(own) = *LEVEL_SOURCES.get(index)? else {
        return sourced(*LEVEL_SOURCES.get(index)?, None, Some(run));
    };
    const ILVL: u8 = own_at(LEVEL_OWN, "ilvl");
    const START: u8 = own_at(LEVEL_OWN, "start");
    const FORMAT: u8 = own_at(LEVEL_OWN, "format");
    const LVL_TEXT: u8 = own_at(LEVEL_OWN, "lvlText");
    const JUSTIFICATION: u8 = own_at(LEVEL_OWN, "justification");
    const SUFF: u8 = own_at(LEVEL_OWN, "suff");
    const LVL_RESTART: u8 = own_at(LEVEL_OWN, "lvlRestart");
    const IS_LGL: u8 = own_at(LEVEL_OWN, "isLgl");
    const INDENT: u8 = own_at(LEVEL_OWN, "indent");
    const HANGING: u8 = own_at(LEVEL_OWN, "hanging");
    match own {
        ILVL => Some(NodeValue::Int(i64::from(level.ilvl()))),
        START => int_text(level.start()),
        FORMAT => level.format().map(text_value),
        LVL_TEXT => level.text().map(text_value),
        JUSTIFICATION => level
            .justification()
            .and_then(|value| LEVEL_JUSTIFICATION.get(value))
            .map(NodeValue::Token),
        SUFF => level
            .suffix()
            .and_then(|value| LEVEL_SUFFIX.get(value))
            .map(NodeValue::Token),
        LVL_RESTART => int_text(level.restart()),
        IS_LGL => level.legal().filter(|legal| *legal).map(NodeValue::Bool),
        INDENT => indentation?.leading().map(|value| NodeValue::Int(value.0)),
        HANGING => indentation?.hanging.map(|value| NodeValue::Int(value.0)),
        _ => None,
    }
}

fn story_value<'a>(
    kind: Option<StoryKind>,
    text: &str,
    paragraph: Option<&direct::Paragraph>,
    run: Option<&direct::Run>,
    index: usize,
) -> Option<NodeValue<'a>> {
    match STORY_PROPS.get(index)?.name {
        "type" => kind.map(|kind| NodeValue::Token(kind.to_xml())),
        "text" => Some(text_value(text)),
        "direction" | "field" => None,
        _ => sourced(*STORY_SOURCES.get(index)?, paragraph, run),
    }
}

/// A story's text as OfficeCLI reads a header: its runs' `w:t` text,
/// concatenated and trimmed.
fn story_text(xml: &Arc<Vec<u8>>) -> Result<String> {
    let story =
        litchi_docx::header_footer::Story::from_shared_bytes(Arc::clone(xml), StoryKind::Primary)
            .map_err(Error::from)?;
    Ok(story.run_text().map_err(Error::from)?.trim().to_owned())
}

fn internal(message: &'static str) -> Error {
    Error::new(Code::InternalError, message).with_format(litchi_core::FileFormat::Docx)
}

/// A paragraph's `styleName`, `numFmt`, `listStyle` or `start`, read through
/// the styles and numbering parts (record 0319). Only the paragraph's own
/// `w:pStyle` and `w:numPr` are followed, never the style's inherited
/// numbering (ADR 0004).
pub(super) fn paragraph_part_value<'a>(
    docx: &Docx,
    paragraph: &direct::Paragraph,
    index: usize,
) -> Option<NodeValue<'a>> {
    use super::props::{P_LIST_STYLE, P_NUM_FMT, P_START, P_STYLE_NAME};
    if index == P_STYLE_NAME {
        let id = paragraph.style.as_deref()?;
        let catalog = styles(docx).ok()??;
        return catalog
            .get(id)
            .and_then(|entry| entry.name())
            .filter(|name| !name.is_empty())
            .map(text_value);
    }
    let numbering_properties = paragraph.numbering?;
    let num_id = numbering_properties.id.filter(|id| *id != 0)?;
    let ilvl = numbering_properties.level.unwrap_or(0);
    let catalog = numbering(docx).ok()??;
    let level = catalog.level_of(num_id, ilvl)?;
    match index {
        P_NUM_FMT => level.format().map(text_value),
        P_LIST_STYLE => level.format().map(|format| {
            NodeValue::Token(if format.eq_ignore_ascii_case("bullet") {
                "bullet"
            } else {
                "ordered"
            })
        }),
        P_START => catalog
            .num(num_id)
            .and_then(|num| {
                num.overrides()
                    .iter()
                    .find(|(level, _)| *level == ilvl)
                    .and_then(|(_, start)| start.as_deref())
            })
            .or(level.start())
            .and_then(|start| int_text(Some(start))),
        _ => None,
    }
}
