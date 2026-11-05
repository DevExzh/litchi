#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

//! Read accessors for story anchors and style/list references.

use litchi_rtf::Package;

const PNG: &str = "89504e470d0a1a0a";

#[test]
fn body_pictures_report_their_body_offset() {
    let source = Package::parse(&format!(
        r"{{\rtf1\ansi\pard Lead{{\pict\pngblip\picw10\pich10 {PNG}}}tail\par\trowd\cellx1000\pard\intbl {{\pict\pngblip {PNG}}}\cell\row\pard End{{\pict\pngblip {PNG}}}\par}}"
    ))
    .unwrap();
    assert_eq!(source.pictures().len(), 3);
    assert_eq!(source.text(), "Leadtail\nEnd\n");
    assert_eq!(source.picture_body_offset(0), Some(4));
    // Table-cell pictures are not anchored in the main story.
    assert_eq!(source.picture_body_offset(1), None);
    assert_eq!(source.picture_body_offset(2), Some(12));
    assert_eq!(source.document().picture_body_offset(2), Some(12));
    assert_eq!(source.picture_body_offset(3), None);
}

#[test]
fn paragraphs_and_runs_expose_style_and_list_references() {
    let source = Package::parse(
        r"{\rtf1\ansi{\stylesheet{\s0 Normal;}{\s1 heading 1;}{\*\cs10 Emph;}}{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc0\levelstartat1{\leveltext\'02\'00.;}{\levelnumbers\'01;}}\listid7}}{\*\listoverridetable{\listoverride\listid7\listoverridecount0\ls1}}\pard\s1 Title\par\pard\ls1\ilvl0 Item {\cs10 emph}\par}",
    )
    .unwrap();
    let paragraphs = source.body().paragraphs().collect::<Vec<_>>();
    assert_eq!(paragraphs[0].format().style(), Some(1));
    assert_eq!(paragraphs[0].format().list_override(), None);
    assert_eq!(paragraphs[1].format().style(), None);
    assert_eq!(paragraphs[1].format().list_override(), Some(1));
    assert_eq!(paragraphs[1].format().list_level(), Some(0));
    assert_eq!(paragraphs[1].text_range(), 6..15);
    let runs = paragraphs[1].runs().collect::<Vec<_>>();
    assert_eq!(runs.last().unwrap().text(), "emph");
    assert_eq!(runs.last().unwrap().format().character_style(), Some(10));
    assert_eq!(runs[0].format().character_style(), None);
}

#[test]
fn ftnalt_marks_an_endnote_and_round_trips() {
    let source =
        Package::parse(r"{\rtf1\ansi A{\footnote\chftn Foot}B{\footnote\ftnalt\chftn End}C{\footnote\ftnalt Plain}\par}")
            .unwrap();
    let kinds = source
        .notes()
        .iter()
        .map(|note| {
            (
                note.is_footnote,
                note.content.as_ref(),
                note.reference.as_ref(),
            )
        })
        .collect::<Vec<_>>();
    // `\chftn` carries an explicit reference; an unmarked endnote keeps the
    // endnote default.
    assert_eq!(
        kinds,
        [
            (true, "Foot", "1"),
            (false, "End", "1"),
            (false, "Plain", "i")
        ]
    );
    assert!(source.opaque().is_empty());
    assert_eq!(source.endnotes().count(), 2);
}
