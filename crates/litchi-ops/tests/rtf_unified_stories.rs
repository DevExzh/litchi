#![cfg(feature = "rtf")]
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure"
)]

//! Record 0275: unified RTF tables, header/footer, note and comment stories,
//! hyperlinks, pictures, styles and lists.

use litchi_core::selector::{Path, Scope};
use litchi_ops::query::Query;
use litchi_ops::{BatchItem, BatchMode, Code, File, Props, open_bytes, props};

const PNG: &str = "89504e470d0a1a0a0000000d49484452";

fn input() -> String {
    format!(
        r#"{{\rtf1\ansi{{\fonttbl{{\f0 Arial;}}}}{{\stylesheet{{\s0 Normal;}}{{\s1\sbasedon0\snext0 heading 1;}}{{\*\cs10 Emph;}}}}{{\*\listtable{{\list\listtemplateid1{{\listlevel\levelnfc0\levelstartat3{{\leveltext\'02\'00.;}}{{\levelnumbers\'01;}}}}{{\listlevel\levelnfc23\levelstartat1{{\leveltext\'01\u8226 ?;}}{{\levelnumbers;}}}}\listid7}}}}{{\*\listoverridetable{{\listoverride\listid7\listoverridecount0\ls1}}}}
{{\header \pard Running head\par}}
{{\footer \pard Page foot\par Second foot\par}}
\pard\s1 Title\par
\pard\ls1\ilvl0 Item {{\cs10 emph}}{{\field{{\*\fldinst HYPERLINK "https://example.com/x" \\o "Tip"}}{{\fldrslt site}}}} end{{\super\chftn{{\footnote\pard\plain\chftn Foot one}}}}\par
\pard Pic{{\pict\pngblip\picw20\pich10\picwgoal1440\pichgoal720 {PNG}}}tail{{\super\chftn{{\footnote\ftnalt\pard\plain\chftn End one}}}}{{\*\atnid AM}}{{\*\atnauthor Ada}}\chatn{{\*\annotation Remark}}\par
\trowd\trhdr\trleft0\cellx1440\clvertalc\cellx4320\pard\intbl A1\cell {{\b B1}}\cell\row
\trowd\trleft0\cellx1440\cellx4320\pard\intbl A2\cell\cell\row
\pard Last\par
}}"#
    )
}

/// A mutation query scoped to the document root.
fn scoped(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .unwrap()
        .with_scope(Scope::Path(Box::new(Path::parse("/").unwrap())))
}

fn file() -> File {
    open_bytes(input().into_bytes()).unwrap()
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap_or_default()
        .into_owned()
}

fn property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

fn paths(file: &File, selector: &str) -> Vec<String> {
    file.query(selector)
        .unwrap()
        .map(|node| node.unwrap().path().to_string())
        .collect()
}

#[test]
fn tables_rows_and_cells_read_through_paths_and_queries() {
    let file = file();
    assert_eq!(property(&file, "/body/tbl[1]", "rows").unwrap(), "2");
    assert_eq!(property(&file, "/body/tbl[1]", "cols").unwrap(), "2");
    assert_eq!(
        property(&file, "/body/tbl[1]/tr[1]", "header").unwrap(),
        "true"
    );
    assert_eq!(text(&file, "/body/tbl[1]/tr[1]/tc[1]"), "A1");
    assert_eq!(text(&file, "/body/tbl[1]/tr[1]/tc[2]"), "B1");
    assert_eq!(text(&file, "/body/tbl[1]/tr[2]/tc[2]"), "");
    assert_eq!(
        property(&file, "/body/tbl[1]/tr[1]/tc[1]", "width").unwrap(),
        "72pt"
    );
    assert_eq!(
        property(&file, "/body/tbl[1]/tr[1]/tc[2]", "valign").unwrap(),
        "center"
    );
    assert_eq!(paths(&file, "tbl").len(), 1);
    assert_eq!(paths(&file, "tc[text=A2]"), ["/body/tbl[1]/tr[2]/tc[1]"]);
    // Body children list paragraphs before tables.
    let body = file.get("/body").unwrap();
    assert_eq!(body.children().count(), 5);
    assert_eq!(
        file.get("/body/tbl[2]").unwrap_err().code(),
        Code::PathNotFound
    );
}

#[test]
fn header_footer_note_and_comment_stories_read() {
    let file = file();
    assert_eq!(text(&file, "/header[1]"), "Running head");
    assert_eq!(property(&file, "/header[1]", "type").unwrap(), "default");
    assert_eq!(text(&file, "/header[@type=default]"), "Running head");
    assert_eq!(text(&file, "/footer[1]"), "Page foot\nSecond foot");
    assert_eq!(text(&file, "/footer[1]/p[2]"), "Second foot");
    assert_eq!(file.get("/footer[1]").unwrap().children().count(), 2);
    assert_eq!(text(&file, "/footnote[1]"), "Foot one");
    assert_eq!(property(&file, "/footnote[1]", "number").unwrap(), "1");
    assert_eq!(text(&file, "/endnote[1]"), "End one");
    assert_eq!(
        file.get("/footnote[2]").unwrap_err().code(),
        Code::PathNotFound
    );
    assert_eq!(text(&file, "/comments/comment[1]"), "Remark");
    assert_eq!(
        property(&file, "/comments/comment[1]", "author").unwrap(),
        "Ada"
    );
    assert_eq!(paths(&file, "footnote"), ["/footnote[1]"]);
    assert_eq!(paths(&file, "comment"), ["/comments/comment[1]"]);
    // Story paragraphs stay out of the main-story paragraph axis.
    assert_eq!(paths(&file, "p").len(), 4);
}

#[test]
fn hyperlinks_and_pictures_anchor_in_their_paragraph() {
    let file = file();
    assert_eq!(text(&file, "/body/p[2]/hyperlink[1]"), "site");
    assert_eq!(
        property(&file, "/body/p[2]/hyperlink[1]", "url").unwrap(),
        "https://example.com/x"
    );
    assert_eq!(
        property(&file, "/body/p[2]/hyperlink[1]", "tooltip").unwrap(),
        "Tip"
    );
    assert_eq!(paths(&file, "hyperlink"), ["/body/p[2]/hyperlink[1]"]);
    assert_eq!(
        property(&file, "/body/p[3]/picture[1]", "type").unwrap(),
        "png"
    );
    assert_eq!(
        property(&file, "/body/p[3]/picture[1]", "width").unwrap(),
        "2.54cm"
    );
    assert_eq!(paths(&file, "picture"), ["/body/p[3]/picture[1]"]);
    assert!(file.find("/body/p[1]/picture[1]").unwrap().is_none());
}

#[test]
fn styles_lists_and_their_references_read() {
    let file = file();
    assert_eq!(property(&file, "/body/p[1]", "style").unwrap(), "heading 1");
    assert_eq!(
        file.get("/body/p[1]")
            .unwrap()
            .view()
            .style()
            .unwrap()
            .unwrap(),
        "heading 1"
    );
    assert_eq!(property(&file, "/body/p[2]", "numId").unwrap(), "1");
    assert_eq!(property(&file, "/body/p[2]", "numLevel").unwrap(), "0");
    assert_eq!(
        property(&file, "/body/p[2]/r[2]", "rStyle").unwrap(),
        "Emph"
    );
    assert_eq!(
        property(&file, "/styles/heading 1", "type").unwrap(),
        "paragraph"
    );
    assert_eq!(
        property(&file, "/styles/heading 1", "basedOn").unwrap(),
        "Normal"
    );
    assert_eq!(
        property(&file, "/styles/Emph", "type").unwrap(),
        "character"
    );
    assert_eq!(file.get("/styles").unwrap().children().count(), 3);
    assert_eq!(
        property(&file, "/numbering/abstractNum[1]", "id").unwrap(),
        "7"
    );
    assert_eq!(
        property(&file, "/numbering/abstractNum[1]/level[1]", "start").unwrap(),
        "3"
    );
    assert_eq!(
        property(&file, "/numbering/abstractNum[1]/level[2]", "numFmt").unwrap(),
        "bullet"
    );
    assert_eq!(
        property(&file, "/numbering/abstractNum[1]/level[2]", "ilvl").unwrap(),
        "1"
    );
    assert_eq!(
        file.get("/numbering/abstractNum[1]/level[2]")
            .unwrap()
            .path()
            .to_string(),
        "/numbering/abstractNum[1]/level[2]"
    );
    assert_eq!(
        property(&file, "/numbering/num[1]", "abstractNumId").unwrap(),
        "7"
    );
    assert_eq!(property(&file, "/numbering/num[1]", "id").unwrap(), "1");
}

#[test]
fn destination_text_sets_splice_the_source_exactly() {
    let source = file();
    let input = input();
    let cases: [(&str, &str, &str); 6] = [
        ("/body/tbl[1]/tr[1]/tc[1]", "Cell", "A1\\cell"),
        ("/body/tbl[1]/tr[1]/tc[2]", "Bold cell", "B1}\\cell"),
        ("/header[1]", "New head", "Running head\\par"),
        ("/footer[1]/p[2]", "Other foot", "Second foot\\par"),
        ("/footnote[1]", "Changed note", "Foot one}"),
        ("/comments/comment[1]", "New remark", "Remark}"),
    ];
    for (path, value, fragment) in cases {
        let changed = source.set(path, props! {"text" => value}).unwrap();
        assert_eq!(text(&changed, path), value, "{path}");
        let old = fragment.split(['\\', '}']).next().unwrap();
        let expected = input.replacen(fragment, &fragment.replacen(old, value, 1), 1);
        assert_eq!(
            String::from_utf8(changed.to_bytes().unwrap()).unwrap(),
            expected,
            "{path}"
        );
        assert_eq!(source.to_bytes().unwrap(), input.as_bytes());
    }
}

#[test]
fn empty_cells_and_unicode_round_trip_through_save() {
    let source = file();
    let changed = source
        .set("/body/tbl[1]/tr[2]/tc[2]", props! {"text" => "Zoë {x}"})
        .unwrap();
    let reopened = open_bytes(changed.to_bytes().unwrap()).unwrap();
    assert_eq!(text(&reopened, "/body/tbl[1]/tr[2]/tc[2]"), "Zoë {x}");
    assert_eq!(text(&reopened, "/body/tbl[1]/tr[2]/tc[1]"), "A2");
    assert_eq!(text(&reopened, "/body/p[4]"), "Last");
}

#[test]
fn query_fanout_publishes_one_native_transaction() {
    let source = file();
    let commit = source
        .edit(|edit| edit.set(scoped("tc"), props! {"text" => "x"}))
        .unwrap();
    assert_eq!(commit.diagnostics().operations(), 1);
    for path in [
        "/body/tbl[1]/tr[1]/tc[1]",
        "/body/tbl[1]/tr[1]/tc[2]",
        "/body/tbl[1]/tr[2]/tc[1]",
        "/body/tbl[1]/tr[2]/tc[2]",
    ] {
        assert_eq!(text(commit.file(), path), "x");
    }
    let restored = commit.patch().inverse().apply(commit.file()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), input().as_bytes());
}

#[test]
fn batches_mix_destination_kinds_atomically() {
    fn set(path: &'static str, value: &'static str) -> BatchItem<'static> {
        let mut item = BatchItem::new("set");
        item.path = Some(path.into());
        item.props = Some(Props::new().with("text", value));
        item
    }
    let source = file();
    let result = source
        .batch(
            vec![
                set("/body/tbl[1]/tr[1]/tc[1]", "one"),
                set("/footnote[1]", "two"),
                set("/endnote[1]", "three"),
            ],
            BatchMode::Atomic,
        )
        .unwrap();
    let file = result.file();
    assert_eq!(text(file, "/body/tbl[1]/tr[1]/tc[1]"), "one");
    assert_eq!(text(file, "/footnote[1]"), "two");
    assert_eq!(text(file, "/endnote[1]"), "three");

    let error = source
        .batch(
            vec![
                set("/body/tbl[1]/tr[1]/tc[1]", "one"),
                set("/footer[1]", "multi-paragraph footer"),
            ],
            BatchMode::Atomic,
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(source.to_bytes().unwrap(), input().as_bytes());
}

#[test]
fn unsupported_story_writes_are_typed_refusals() {
    let source = file();
    let refusal = |result: litchi_ops::Result<File>| result.unwrap_err().code();
    assert_eq!(
        refusal(source.set("/body/tbl[1]/tr[1]/tc[1]", props! {"valign" => "top"})),
        Code::UnsupportedCapability
    );
    assert_eq!(
        refusal(source.set("/body/p[2]/hyperlink[1]", props! {"url" => "x"})),
        Code::UnsupportedCapability
    );
    assert_eq!(
        refusal(source.remove("/body/tbl[1]/tr[1]/tc[1]")),
        Code::UnsupportedCapability
    );
    assert_eq!(
        refusal(source.set("/styles/Emph", props! {"name" => "x"})),
        Code::UnsupportedCapability
    );
    assert_eq!(
        refusal(source.set(scoped("p, tc"), props! {"text" => "x"})),
        Code::UnsupportedType
    );
    assert_eq!(source.to_bytes().unwrap(), input().as_bytes());
}

#[test]
fn pictures_remove_through_the_native_exact_source_editor() {
    let source = open_bytes(
        format!(
            r"{{\rtf1\ansi\pard Lead{{\pict\pngblip\picw10\pich10 {PNG}}}tail\par\pard Two{{\pict\pngblip {PNG}}}\par}}"
        )
        .into_bytes(),
    )
    .unwrap();
    let changed = source.remove("/body/p[1]/picture[1]").unwrap();
    assert_eq!(
        String::from_utf8(changed.to_bytes().unwrap()).unwrap(),
        format!(r"{{\rtf1\ansi\pard Leadtail\par\pard Two{{\pict\pngblip {PNG}}}\par}}")
    );
    let all = source.remove(scoped("picture")).unwrap();
    assert_eq!(all.query("picture").unwrap().count(), 0);
    assert_eq!(text(&all, "/body/p[2]"), "Two");
}

#[test]
fn capabilities_list_the_projected_elements_and_their_gaps() {
    let file = file();
    let capabilities = file.capabilities();
    for element in [
        "table",
        "row",
        "cell",
        "header",
        "footer",
        "footnote",
        "endnote",
        "comment",
        "hyperlink",
        "picture",
        "styles",
        "style",
        "numbering",
        "abstractNum",
        "num",
        "level",
    ] {
        let element = capabilities.element(element).expect(element);
        assert!(element.operations().get, "{}", element.name());
    }
    assert!(capabilities.element("cell").unwrap().operations().set);
    assert!(capabilities.element("footnote").unwrap().operations().set);
    assert!(!capabilities.element("hyperlink").unwrap().operations().set);
    assert!(capabilities.element("picture").unwrap().operations().remove);
    assert!(!capabilities.element("picture").unwrap().operations().add);
    let gap = |verb: &'static str, element: &'static str| {
        capabilities.gaps().any(|gap| {
            matches!(
                gap.subject(),
                litchi_ops::capabilities::Subject::Operation { verb: v, element: e, .. }
                    if v == verb && e == element
            )
        })
    };
    assert!(gap("add", "table"));
    assert!(gap("add", "picture"));
    assert!(gap("set", "hyperlink"));
    assert!(gap("set", "style"));
    assert!(gap("set", "abstractNum"));
}

#[test]
fn direct_destination_sets_coalesce_into_one_publication() {
    let source = file();
    let oracle = source
        .set("/body/tbl[1]/tr[1]/tc[1]", props! {"text" => "a"})
        .unwrap()
        .set("/body/tbl[1]/tr[2]/tc[1]", props! {"text" => "b"})
        .unwrap()
        .set("/footer[1]/p[1]", props! {"text" => "c"})
        .unwrap()
        .set("/footnote[1]", props! {"text" => "d"})
        .unwrap()
        .set("/comments/comment[1]", props! {"text" => "e"})
        .unwrap();
    let commit = source
        .edit(|edit| {
            edit.set("/body/tbl[1]/tr[1]/tc[1]", props! {"text" => "a"})?;
            edit.set("/body/tbl[1]/tr[2]/tc[1]", props! {"text" => "b"})?;
            edit.set("/footer[1]/p[1]", props! {"text" => "c"})?;
            edit.set("/footnote[1]", props! {"text" => "d"})?;
            edit.set("/comments/comment[1]", props! {"text" => "e"})
        })
        .unwrap();
    assert_eq!(commit.diagnostics().operations(), 5);
    assert_eq!(commit.diagnostics().commits(), 1);
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        oracle.to_bytes().unwrap()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        input().as_bytes()
    );

    // A second write to the same story closes the window.
    let commit = source
        .edit(|edit| {
            edit.set("/body/tbl[1]/tr[1]/tc[1]", props! {"text" => "a"})?;
            edit.set("/header[1]", props! {"text" => "h"})?;
            edit.set("/body/tbl[1]/tr[1]/tc[1]", props! {"text" => "b"})
        })
        .unwrap();
    assert_eq!(commit.diagnostics().commits(), 2);
    assert_eq!(text(commit.file(), "/body/tbl[1]/tr[1]/tc[1]"), "b");
    assert_eq!(text(commit.file(), "/header[1]"), "h");
}

#[test]
fn repeated_header_kinds_edit_only_the_addressable_story() {
    let input = br"{\rtf1\ansi{\header First head}{\header Second head}\pard Body\par}";
    let source = open_bytes(input.as_slice()).unwrap();
    assert_eq!(text(&source, "/header[2]"), "Second head");
    let error = source
        .set("/header[2]", props! {"text" => "x"})
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let changed = source.set("/header[1]", props! {"text" => "New"}).unwrap();
    assert_eq!(
        changed.to_bytes().unwrap(),
        br"{\rtf1\ansi{\header New}{\header Second head}\pard Body\par}"
    );
    assert_eq!(source.to_bytes().unwrap(), input);
}

/// Three paragraphs with three, one and two pictures.
fn pictures_input() -> Vec<u8> {
    let pict = format!(r"{{\pict\pngblip {PNG}}}");
    format!(
        r"{{\rtf1\ansi\pard A{pict}B{pict}C{pict}\par\pard D{pict}\par\pard E{pict}F{pict}\par}}"
    )
    .into_bytes()
}

/// Consecutive direct picture removals inside one edit publish one native
/// removal with the bytes, and the history, of the same removals made one
/// call at a time (record 0280).
#[test]
fn direct_picture_removals_coalesce_into_one_commit() {
    let source = open_bytes(pictures_input()).unwrap();
    let paths = [
        "/body/p[1]/picture[2]",
        "/body/p[1]/picture[1]",
        "/body/p[3]/picture[2]",
        "/body/p[1]/picture[1]",
        "/body/p[2]/image[1]",
    ];
    let mut sequential = source.clone();
    for path in paths {
        sequential = sequential.remove(path).unwrap();
    }
    let commit = source
        .edit(|edit| {
            for path in paths {
                edit.remove(path)?;
            }
            Ok(())
        })
        .unwrap();
    assert_eq!(commit.diagnostics().commits(), 1);
    assert_eq!(commit.diagnostics().operations(), paths.len());
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        sequential.to_bytes().unwrap()
    );
    assert_eq!(commit.file().query("picture").unwrap().count(), 1);
    let undone = commit.patch().inverse().apply(commit.file()).unwrap();
    assert_eq!(undone.to_bytes().unwrap(), pictures_input());
}

/// A read and a removal the run cannot place (a query) each close the run;
/// a missing picture is reported for its own command.
#[test]
fn picture_removal_runs_close_at_barriers_and_attribute_errors() {
    let source = open_bytes(pictures_input()).unwrap();
    let commit = source
        .edit(|edit| {
            edit.remove("/body/p[1]/picture[1]")?;
            assert_eq!(
                edit.get("/body/p[1]")?.view().text()?.as_deref(),
                Some("ABC")
            );
            edit.remove("/body/p[1]/picture[1]")?;
            edit.remove(scoped("picture[1]"))?;
            edit.remove("/body/p[3]/picture[1]")?;
            Ok(())
        })
        .unwrap();
    assert_eq!(commit.diagnostics().commits(), 4);
    let mut sequential = source.clone();
    sequential = sequential.remove("/body/p[1]/picture[1]").unwrap();
    sequential = sequential.remove("/body/p[1]/picture[1]").unwrap();
    sequential = sequential.remove(scoped("picture[1]")).unwrap();
    sequential = sequential.remove("/body/p[3]/picture[1]").unwrap();
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        sequential.to_bytes().unwrap()
    );

    let error = source
        .edit(|edit| {
            edit.remove("/body/p[2]/picture[1]")?;
            edit.remove("/body/p[2]/picture[1]")?;
            Ok(())
        })
        .unwrap_err();
    assert_eq!(error.item(), Some(1));
    assert_eq!(
        error.code(),
        source.remove("/body/p[2]/picture[2]").unwrap_err().code()
    );
}

/// More removals than one native removal accepts publish in several
/// commits with the sequential result.
#[test]
fn picture_removal_runs_split_at_the_native_cap() {
    let pict = format!(r"{{\pict\pngblip {PNG}}}");
    let body: String = (0..70).map(|_| pict.as_str()).collect();
    let input = format!(r"{{\rtf1\ansi\pard X{body}\par}}").into_bytes();
    let source = open_bytes(input).unwrap();
    let commit = source
        .edit(|edit| {
            for _ in 0..70 {
                edit.remove("/body/p[1]/picture[1]")?;
            }
            Ok(())
        })
        .unwrap();
    assert_eq!(commit.diagnostics().commits(), 2);
    assert_eq!(commit.file().query("picture").unwrap().count(), 0);
    let mut sequential = source.clone();
    for _ in 0..70 {
        sequential = sequential.remove("/body/p[1]/picture[1]").unwrap();
    }
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        sequential.to_bytes().unwrap()
    );
}
