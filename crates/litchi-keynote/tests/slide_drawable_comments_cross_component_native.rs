//! Native Keynote regression fixtures and opt-in candidate export.
//!
//! The exporter writes focused candidates only when explicitly enabled.  A
//! separate readback pass consumes checked-in native goldens by default (or
//! an explicitly overridden directory) after Keynote has opened and saved
//! them, then checks semantic selectors and document invariants rather than
//! physical object identifiers, which native Keynote may remap.

#[path = "support/drawable_comment_fixtures.rs"]
mod fixtures;

use std::{
    env, fs, io,
    path::{Path, PathBuf},
};

use litchi_keynote::{
    DrawableKind, DrawableSelector, MovieInfo, Package, ReplySelector, SlideSelector,
};

use fixtures::{RelocatedFixture, TestResult};

const OUTPUT_ENV: &str = "LITCHI_KEYNOTE_DRAWABLE_COMMENTS_CROSS_COMPONENT_NATIVE_OUTPUT_DIR";
const NATIVE_DIR_ENV: &str = "LITCHI_KEYNOTE_DRAWABLE_COMMENTS_CROSS_COMPONENT_NATIVE_DIR";

fn native_directory() -> PathBuf {
    env::var_os(NATIVE_DIR_ENV).map_or_else(
        || {
            PathBuf::from(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/iwork/keynote/drawable-comments-cross-component-native")
        },
        PathBuf::from,
    )
}

#[derive(Debug, Clone, Copy)]
enum Operation {
    SetRoot(&'static str),
    ClearRoot,
    AddReply(&'static str),
    SetReply(&'static str),
    RemoveReply,
    CreateRoot(&'static str),
    CreateRootAndReply {
        root: &'static str,
        reply: &'static str,
    },
}

struct Candidate {
    name: &'static str,
    fixture: RelocatedFixture,
    operation: Operation,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DrawableState {
    kind: DrawableKind,
    comment: Option<String>,
    comment_has_author: bool,
    replies: Vec<ReplyState>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ReplyState {
    text: String,
    has_author: bool,
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn candidates() -> TestResult<Vec<Candidate>> {
    let root = fixtures::root_foreign_fixture()?;
    let reply = fixtures::reply_foreign_fixture()?;
    let shared_root = fixtures::shared_foreign_root_fixture()?;
    let shared_reply = fixtures::shared_foreign_reply_fixture()?;
    let foreign_drawable = fixtures::foreign_drawable_fixture()?;
    Ok(vec![
        Candidate {
            name: "foreign-root-set.key",
            fixture: root.clone(),
            operation: Operation::SetRoot("native foreign root set"),
        },
        Candidate {
            name: "foreign-root-clear.key",
            fixture: root,
            operation: Operation::ClearRoot,
        },
        Candidate {
            name: "foreign-reply-add.key",
            fixture: reply.clone(),
            operation: Operation::AddReply("native foreign reply add"),
        },
        Candidate {
            name: "foreign-reply-set.key",
            fixture: reply.clone(),
            operation: Operation::SetReply("native foreign registered set"),
        },
        Candidate {
            name: "foreign-reply-remove.key",
            fixture: reply,
            operation: Operation::RemoveReply,
        },
        Candidate {
            name: "foreign-drawable-root-create.key",
            fixture: foreign_drawable.clone(),
            operation: Operation::CreateRoot("native foreign drawable root"),
        },
        Candidate {
            name: "shared-foreign-root-set.key",
            fixture: shared_root,
            operation: Operation::SetRoot("native shared foreign root set"),
        },
        Candidate {
            name: "shared-foreign-reply-clear.key",
            fixture: shared_reply,
            operation: Operation::ClearRoot,
        },
        Candidate {
            name: "foreign-drawable-root-and-reply.key",
            fixture: foreign_drawable,
            operation: Operation::CreateRootAndReply {
                root: "native foreign drawable root and reply",
                reply: "native foreign drawable reply",
            },
        },
    ])
}

fn apply_operation(candidate: &Candidate) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(&candidate.fixture.bytes)?;
    let slide = SlideSelector::index(0);
    let target = DrawableSelector::index(candidate.fixture.target.position);
    let bytes = match candidate.operation {
        Operation::SetRoot(text) => {
            let commit = package
                .edit_slide_drawable_comment(slide, target)?
                .set(text)?
                .commit()?;
            exact_bytes(commit.package())?
        },
        Operation::ClearRoot => {
            let commit = package
                .edit_slide_drawable_comment(slide, target)?
                .clear()?
                .commit()?;
            exact_bytes(commit.package())?
        },
        Operation::AddReply(text) => {
            let commit = package
                .edit_slide_drawable_comment(slide, target)?
                .add_reply(text)?
                .commit()?;
            exact_bytes(commit.package())?
        },
        Operation::SetReply(text) => {
            let commit = package
                .edit_slide_drawable_comment(slide, target)?
                .set_reply(ReplySelector::index(0), text)?
                .commit()?;
            exact_bytes(commit.package())?
        },
        Operation::RemoveReply => {
            let commit = package
                .edit_slide_drawable_comment(slide, target)?
                .remove_reply(ReplySelector::index(0))?
                .commit()?;
            exact_bytes(commit.package())?
        },
        Operation::CreateRoot(text) => {
            let commit = package
                .edit_slide_drawable_comment(slide, target)?
                .set(text)?
                .commit()?;
            exact_bytes(commit.package())?
        },
        Operation::CreateRootAndReply { root, reply } => {
            let root_commit = package
                .edit_slide_drawable_comment(slide, target)?
                .set(root)?
                .commit()?;
            let reply_commit = root_commit
                .package()
                .edit_slide_drawable_comment(slide, target)?
                .add_reply(reply)?
                .commit()?;
            exact_bytes(reply_commit.package())?
        },
    };
    Ok(bytes)
}

fn drawable_states(package: &Package) -> TestResult<Vec<DrawableState>> {
    let summaries = package.slide_drawables(SlideSelector::index(0))?;
    let mut states = Vec::with_capacity(summaries.len());
    for (index, summary) in summaries.iter().enumerate() {
        let selector = DrawableSelector::index(index);
        let comment = package.slide_drawable_comment(SlideSelector::index(0), selector)?;
        let comment_has_author = comment
            .as_ref()
            .is_some_and(|comment| comment.author().is_some());
        let comment = comment.map(|comment| comment.text().to_owned());
        if summary.has_comment() != comment.is_some() {
            return Err(io::Error::other("native inventory/comment projection disagrees").into());
        }
        let replies = package
            .slide_drawable_comment_replies(SlideSelector::index(0), selector)?
            .iter()
            .map(|reply| ReplyState {
                text: reply.text().to_owned(),
                has_author: reply.author().is_some(),
            })
            .collect();
        states.push(DrawableState {
            kind: summary.kind(),
            comment,
            comment_has_author,
            replies,
        });
    }
    Ok(states)
}

fn document_signature(
    package: &Package,
) -> TestResult<
    Vec<(
        Option<String>,
        Option<String>,
        Vec<String>,
        Option<String>,
        Vec<MovieInfo>,
    )>,
> {
    Ok(package
        .slides()?
        .iter()
        .map(|slide| {
            (
                slide.name().map(str::to_owned),
                slide.title().map(str::to_owned),
                slide.text_content().to_vec(),
                slide.notes().map(str::to_owned),
                slide.movies().to_vec(),
            )
        })
        .collect())
}

fn assert_operation_state(
    package: &Package,
    candidate: &Candidate,
    native: bool,
) -> TestResult<()> {
    let source = Package::from_bytes(&candidate.fixture.bytes)?;
    let before = drawable_states(&source)?;
    let after = drawable_states(package)?;
    assert_eq!(
        after.len(),
        before.len(),
        "native drawable inventory changed"
    );
    assert_eq!(
        document_signature(&source)?,
        document_signature(package)?,
        "document slide/title/movie semantics changed"
    );
    let target = candidate.fixture.target.position;
    for (index, (before, after)) in before.iter().zip(&after).enumerate() {
        if index != target {
            assert_eq!(after, before, "unselected drawable {index} changed");
        }
    }
    let expected = match candidate.operation {
        Operation::SetRoot(text) => DrawableState {
            kind: before[target].kind,
            comment: Some(text.to_owned()),
            comment_has_author: before[target].comment_has_author,
            replies: before[target].replies.clone(),
        },
        Operation::CreateRoot(text) => DrawableState {
            kind: before[target].kind,
            comment: Some(text.to_owned()),
            comment_has_author: true,
            replies: before[target].replies.clone(),
        },
        Operation::ClearRoot => DrawableState {
            kind: before[target].kind,
            comment: None,
            comment_has_author: false,
            replies: Vec::new(),
        },
        Operation::AddReply(text) => {
            let mut replies = before[target].replies.clone();
            replies.push(ReplyState {
                text: text.to_owned(),
                has_author: true,
            });
            DrawableState {
                kind: before[target].kind,
                comment: before[target].comment.clone(),
                comment_has_author: before[target].comment_has_author,
                replies,
            }
        },
        Operation::SetReply(text) => {
            let mut replies = before[target].replies.clone();
            if replies.is_empty() {
                return Err(io::Error::other("reply set candidate has no source reply").into());
            }
            replies[0].text = text.to_owned();
            DrawableState {
                kind: before[target].kind,
                comment: before[target].comment.clone(),
                comment_has_author: before[target].comment_has_author,
                replies,
            }
        },
        Operation::RemoveReply => {
            if before[target].replies.is_empty() {
                return Err(io::Error::other("reply remove candidate has no source reply").into());
            }
            DrawableState {
                kind: before[target].kind,
                comment: before[target].comment.clone(),
                comment_has_author: before[target].comment_has_author,
                replies: before[target].replies[1..].to_vec(),
            }
        },
        Operation::CreateRootAndReply { root, reply } => DrawableState {
            kind: before[target].kind,
            comment: Some(root.to_owned()),
            comment_has_author: true,
            replies: vec![ReplyState {
                text: reply.to_owned(),
                has_author: true,
            }],
        },
    };
    assert_eq!(
        after[target], expected,
        "unexpected native operation result"
    );
    if native {
        assert_eq!(
            package.slide_drawables(SlideSelector::index(0))?.len(),
            before.len(),
            "native inventory summary count changed"
        );
    }
    Ok(())
}

fn export_candidate(path: &Path, bytes: &[u8]) -> TestResult<()> {
    fs::write(path, bytes)?;
    eprintln!(
        "exported native cross-component candidate to {}",
        path.display()
    );
    Ok(())
}

#[test]
fn export_cross_component_native_candidates() -> TestResult<()> {
    let Some(directory) = env::var_os(OUTPUT_ENV) else {
        return Ok(());
    };
    let directory = Path::new(&directory);
    fs::create_dir_all(directory)?;
    for candidate in candidates()? {
        let bytes = apply_operation(&candidate)?;
        let package = Package::from_bytes(&bytes)?;
        let moved = candidate
            .fixture
            .root_identifier
            .or(candidate.fixture.reply_identifier)
            .unwrap_or(candidate.fixture.target.identifier);
        fixtures::assert_metadata_relocated(
            &candidate.fixture.metadata_source,
            &candidate.fixture.bytes,
            moved,
        )?;
        assert_operation_state(&package, &candidate, false)?;
        export_candidate(&directory.join(candidate.name), &bytes)?;
    }
    Ok(())
}

#[test]
fn native_resaved_cross_component_candidates_read_back() -> TestResult<()> {
    let directory = native_directory();
    for candidate in candidates()? {
        let path = directory.join(candidate.name);
        assert!(
            path.is_file(),
            "missing native candidate {}",
            path.display()
        );
        let package = Package::from_bytes(&fs::read(&path)?)?;
        assert!(
            !package.slides()?.is_empty(),
            "native document has no slides"
        );
        assert_operation_state(&package, &candidate, true)?;
    }
    Ok(())
}
