//! Keynote slide text-storage parity (record 0372): the host
//! `slide_text_storages`, `set_slide_text_storage`,
//! `replace_slide_text_storage` and `clear_slide_text_storage` against the
//! leaf drawable-list text and drawable `Edit::set_text`; and the host
//! `slide_drawables` against the leaf `slide_drawables`.
//!
//! Title and body placeholder storage is written by the leaf slide-text
//! `Edit` (`Package::edit_slide_title`/`edit_slide_body`), which also
//! invalidates the slide node's rendered preview where the host leaves it
//! stale (a Tier B violation by design), so the storage cases use text boxes.

use litchi_iwa::keynote;
use litchi_keynote::slide::drawable::Kind;
use litchi_keynote::{DrawableSelector, Package, SlideSelector};

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::keynote_drawable_cases::{
    TEXT_BOX, host_id, leaf_drawables, leaf_position, leaf_run, run,
};
use crate::leaf_bytes;

const SLIDE: usize = 0;
const COMMENTS: &str = "iwork/keynote/drawable-comments-source-native.key";

fn host_role(role: keynote::SlideTextRole) -> &'static str {
    match role {
        keynote::SlideTextRole::Title => "title",
        keynote::SlideTextRole::Body => "body",
        keynote::SlideTextRole::TextBox => "text box",
        keynote::SlideTextRole::Shape => "shape",
    }
}

fn leaf_role(kind: Kind) -> Option<&'static str> {
    match kind {
        Kind::TitlePlaceholder => Some("title"),
        Kind::BodyPlaceholder => Some("body"),
        Kind::TextBox => Some("text box"),
        Kind::Shape => Some("shape"),
        _ => None,
    }
}

/// Host and leaf agree on every text storage of every slide: role and text,
/// as a multiset (the host lists them in owned-drawable order, the leaf in
/// source order).
fn storages_agree(bytes: &[u8]) -> TestResult<usize> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let package = Package::from_bytes(bytes)?;
    let mut total = 0;
    for slide in 0..editor.slides()?.len() {
        let mut host = editor
            .slide_text_storages(slide)?
            .into_iter()
            .map(|info| {
                (
                    host_role(info.role).to_owned(),
                    info.storage.storage.text().to_owned(),
                )
            })
            .collect::<Vec<_>>();
        let mut leaf = package
            .slide_drawable_list(SlideSelector::index(slide))?
            .iter()
            .filter_map(|drawable| {
                Some((
                    leaf_role(drawable.kind())?.to_owned(),
                    drawable.text()?.to_owned(),
                ))
            })
            .collect::<Vec<_>>();
        host.sort();
        leaf.sort();
        assert_eq!(host, leaf, "slide {slide} text storages");
        total += leaf.len();
    }
    Ok(total)
}

/// Host and leaf agree on the commentable drawables of every slide: count
/// and how many carry a direct comment.
fn drawables_agree(bytes: &[u8]) -> TestResult<usize> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let package = Package::from_bytes(bytes)?;
    let mut total = 0;
    for slide in 0..editor.slides()?.len() {
        let host = editor.slide_drawables(slide)?;
        let leaf = package.slide_drawables(SlideSelector::index(slide))?;
        assert_eq!(host.len(), leaf.len(), "slide {slide} drawable count");
        assert_eq!(
            host.iter().filter(|info| info.comment_id.is_some()).count(),
            leaf.iter().filter(|summary| summary.has_comment()).count(),
            "slide {slide} commented drawables"
        );
        total += leaf.len();
    }
    Ok(total)
}

fn read_case(
    name: &'static str,
    path: &'static str,
    read: fn(&[u8]) -> TestResult<usize>,
) -> TestResult {
    let input = fixture(path)?;
    Case::new(name, Format::Keynote, Tier::A, Evidence::Queued, &input)
        .read_only()
        .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host and leaf agree", read)
        .run()?;
    Ok(())
}

#[test]
fn keynote_slide_text_storage_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    for path in [TEXT_BOX, COMMENTS] {
        read_case(
            "keynote_slide_text_storage_read_parity",
            path,
            storages_agree,
        )?;
    }
    Ok(())
}

#[test]
fn keynote_slide_drawables_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    for path in [TEXT_BOX, COMMENTS] {
        read_case("keynote_slide_drawables_read_parity", path, drawables_agree)?;
    }
    Ok(())
}

/// One text-box storage write: `host` on the host's text box id, the leaf
/// `set_text` of the resulting whole text.
fn text_box_case(
    name: &'static str,
    host: impl FnOnce(&mut keynote::Draft, u64) -> TestResult,
    text: &str,
) -> TestResult {
    run(
        name,
        TEXT_BOX,
        Tier::B,
        Evidence::Queued,
        |editor, source| {
            let id = host_id(editor, source, leaf_position(source, Kind::TextBox)?)?;
            host(editor, id)
        },
        |source| {
            let position = leaf_position(source, Kind::TextBox)?;
            leaf_run(source, |edit| {
                edit.set_text(
                    SlideSelector::index(SLIDE),
                    DrawableSelector::index(position),
                    text,
                )?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_text_storage_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    text_box_case(
        "keynote_slide_text_storage_set_parity",
        |editor, id| Ok(editor.set_slide_text_storage(SLIDE, id, "Quarterly review")?),
        "Quarterly review",
    )
}

#[test]
fn keynote_slide_text_storage_replace_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = fixture(TEXT_BOX)?;
    let current = leaf_drawables(&input.bytes)?
        .iter()
        .find(|drawable| drawable.kind() == Kind::TextBox)
        .and_then(|drawable| drawable.text().map(str::to_owned))
        .ok_or("fixture has no text box")?;
    let units = current.encode_utf16().collect::<Vec<_>>();
    let end = units.len().min(2);
    let mut spliced = String::from("Re");
    spliced.push_str(&String::from_utf16(&units[end..])?);
    // The leaf replaces a range by setting the spliced whole text.
    text_box_case(
        "keynote_slide_text_storage_replace_parity",
        move |editor, id| Ok(editor.replace_slide_text_storage(SLIDE, id, 0..end, "Re")?),
        &spliced,
    )
}

#[test]
fn keynote_slide_text_storage_clear_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    text_box_case(
        "keynote_slide_text_storage_clear_parity",
        |editor, id| Ok(editor.clear_slide_text_storage(SLIDE, id)?),
        "",
    )
}
