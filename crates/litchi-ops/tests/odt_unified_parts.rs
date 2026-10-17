//! Unified ODT parts (record 0328): master-page headers and footers, list
//! styles as numbering, body fields, tables of contents, tab stops, shapes,
//! formula objects and table columns, read through the DOCX paths and
//! property names of records 0319 and 0324.

#![cfg(feature = "odt")]

use litchi_ops::{File, open_bytes};

fn fixture(path: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(path),
    )
    .expect("read fixture")
}

fn open(path: &str) -> File {
    open_bytes(fixture(path)).expect("open ODT")
}

fn parts() -> File {
    open("unified-ops/0294/parts.odt")
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file
        .get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"));
    let format = node.view().format().expect("format");
    format
        .into_iter()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .expect("text")
        .unwrap_or_default()
        .into_owned()
}

fn count(file: &File, query: &str) -> usize {
    file.query(query)
        .unwrap_or_else(|error| panic!("{query}: {error}"))
        .count()
}

#[test]
fn headers_and_footers_read_their_master_page_regions() {
    let file = parts();
    assert_eq!(count(&file, "header"), 3);
    assert_eq!(count(&file, "footer"), 2);
    assert_eq!(text(&file, "/header[1]"), "Top 1");
    assert_eq!(
        prop(&file, "/header[1]", "type").as_deref(),
        Some("default")
    );
    assert_eq!(
        prop(&file, "/header[1]", "masterPage").as_deref(),
        Some("Standard")
    );
    assert_eq!(text(&file, "/header[2]"), "Even top");
    assert_eq!(prop(&file, "/header[2]", "type").as_deref(), Some("even"));
    assert_eq!(prop(&file, "/header[3]", "type").as_deref(), Some("first"));
    assert_eq!(text(&file, "/header[@type=first]"), "First top");
    assert_eq!(text(&file, "/footer[2]"), "Hidden bottom");
    let header = open("odf/corpus/writer-header-footer.odt");
    assert_eq!(text(&header, "/header[1]"), "Header\t1\tTest");
    assert_eq!(text(&header, "/footer[1]"), "Footer\t1\tBottom");
}

#[test]
fn list_styles_read_as_numbering_definitions_and_levels() {
    let file = parts();
    assert_eq!(
        prop(&file, "/numbering", "abstractNumCount").as_deref(),
        Some("2")
    );
    let definition = "/numbering/abstractNum[1]";
    assert_eq!(
        prop(&file, definition, "name").as_deref(),
        Some("Numbering 123")
    );
    assert_eq!(
        prop(&file, definition, "styleName").as_deref(),
        Some("Numbering_20_123")
    );
    assert_eq!(
        prop(&file, definition, "type").as_deref(),
        Some("multilevel")
    );
    assert_eq!(count(&file, "/numbering/abstractNum[1]/level"), 2);
    let level = "/numbering/abstractNum[1]/level[1]";
    assert_eq!(prop(&file, level, "start").as_deref(), Some("3"));
    assert_eq!(prop(&file, level, "format").as_deref(), Some("decimal"));
    assert_eq!(prop(&file, level, "lvlText").as_deref(), Some("(%1)"));
    assert_eq!(
        prop(&file, level, "justification").as_deref(),
        Some("right")
    );
    assert_eq!(prop(&file, level, "suff").as_deref(), Some("space"));
    assert_eq!(prop(&file, level, "indent").as_deref(), Some("360"));
    assert_eq!(prop(&file, level, "hanging").as_deref(), Some("360"));
    assert_eq!(prop(&file, level, "bold").as_deref(), Some("true"));
    assert_eq!(prop(&file, level, "size").as_deref(), Some("12pt"));
    let second = "/numbering/abstractNum[1]/level[2]";
    assert_eq!(
        prop(&file, second, "format").as_deref(),
        Some("lowerLetter")
    );
    assert_eq!(prop(&file, second, "lvlText").as_deref(), Some("%1.%2."));
    assert_eq!(prop(&file, second, "indent").as_deref(), Some("1080"));
    let bullet = "/numbering/abstractNum[2]/level[1]";
    assert_eq!(prop(&file, bullet, "format").as_deref(), Some("bullet"));
    assert_eq!(prop(&file, bullet, "lvlText").as_deref(), Some("•"));
    assert_eq!(prop(&file, bullet, "font").as_deref(), Some("OpenSymbol"));
    assert_eq!(
        prop(&file, "/numbering/num[2]", "abstractNumId").as_deref(),
        Some("1")
    );
    // Native identities resolve through a raw path.
    let raw = litchi_ops::path::raw::path("/numbering/abstractNum[@id=1]/level[@ilvl=0]").unwrap();
    let level = file.get(raw).unwrap();
    assert_eq!(
        level.path().to_string(),
        "/numbering/abstractNum[2]/level[1]"
    );
}

#[test]
fn fields_and_tables_of_contents_read() {
    let file = parts();
    assert_eq!(count(&file, "field"), 5);
    assert_eq!(
        prop(&file, "/field[1]", "fieldType").as_deref(),
        Some("page")
    );
    assert_eq!(text(&file, "/field[1]"), "1");
    assert_eq!(
        prop(&file, "/field[2]", "fieldType").as_deref(),
        Some("numpages")
    );
    assert_eq!(prop(&file, "/field[3]", "fldLock").as_deref(), Some("true"));
    assert_eq!(prop(&file, "/field[3]", "element").as_deref(), Some("date"));
    assert_eq!(
        prop(&file, "/field[5]", "fieldType").as_deref(),
        Some("styleref")
    );
    assert_eq!(prop(&file, "/toc[1]", "levels").as_deref(), Some("1-3"));
    assert_eq!(prop(&file, "/toc[1]", "title").as_deref(), Some("Contents"));
    assert_eq!(
        prop(&file, "/toc[1]", "hyperlinks").as_deref(),
        Some("true")
    );
    assert_eq!(
        prop(&file, "/toc[1]", "pageNumbers").as_deref(),
        Some("true")
    );
    assert_eq!(text(&file, "/toc[1]"), "Parts\t1");
    let toc = open("odf/corpus/writer-table-of-contents.odt");
    assert_eq!(count(&toc, "toc"), 1);
}

#[test]
fn tabs_shapes_equations_and_columns_read() {
    let file = parts();
    let paragraph = file
        .query("p")
        .unwrap()
        .map(Result::unwrap)
        .find(|node| node.view().text().unwrap().as_deref() == Some("Name\tValue"))
        .unwrap()
        .path()
        .to_string();
    let tab = format!("{paragraph}/tab[2]");
    assert_eq!(prop(&file, &tab, "pos").as_deref(), Some("5669"));
    assert_eq!(prop(&file, &tab, "val").as_deref(), Some("right"));
    assert_eq!(prop(&file, &tab, "leader").as_deref(), Some("dot"));
    assert_eq!(count(&file, "tab"), 2);

    assert_eq!(count(&file, "/body/shape"), 2);
    assert_eq!(
        prop(&file, "/body/shape[1]", "geometry").as_deref(),
        Some("star5")
    );
    assert_eq!(
        prop(&file, "/body/shape[1]", "fill").as_deref(),
        Some("#FF0000")
    );
    assert_eq!(
        prop(&file, "/body/shape[1]", "line.color").as_deref(),
        Some("#0000FF")
    );
    assert_eq!(
        prop(&file, "/body/shape[1]", "wrap").as_deref(),
        Some("topAndBottom")
    );
    assert_eq!(text(&file, "/body/shape[1]"), "Starred");
    assert_eq!(
        prop(&file, "/body/shape[2]", "geometry").as_deref(),
        Some("rect")
    );
    assert_eq!(
        prop(&file, "/body/shape[2]", "width").as_deref(),
        Some("2cm")
    );

    assert_eq!(count(&file, "equation"), 2);
    assert_eq!(
        prop(&file, "/body/equation[1]", "formula").as_deref(),
        Some("a^2 + b^2 = c^2")
    );
    assert_eq!(
        prop(&file, "/body/equation[1]", "mode").as_deref(),
        Some("inline")
    );
    assert_eq!(text(&file, "/body/equation[2]"), "x + 1");

    assert_eq!(count(&file, "/body/tbl[1]/col"), 3);
    assert_eq!(
        prop(&file, "/body/tbl[1]/col[2]", "width").as_deref(),
        Some("3cm")
    );
    assert_eq!(
        prop(&file, "/body/tbl[1]/col[3]", "width").as_deref(),
        Some("4.5cm")
    );

    let flat = open("libreoffice-core/sw/qa/extras/uiwriter/data/text-with-formula.fodt");
    assert_eq!(
        prop(&flat, "/body/equation[1]", "formula").as_deref(),
        Some("{func e}^{i %pi} + 1 = 0")
    );
}

#[test]
fn views_list_only_the_structural_children() {
    let file = parts();
    let root: Vec<String> = file
        .get("/")
        .unwrap()
        .view()
        .children()
        .map(|child| child.unwrap().kind().to_owned())
        .collect();
    assert_eq!(root, ["body", "styles", "metadata"]);
    let body: Vec<String> = file
        .get("/body")
        .unwrap()
        .view()
        .children()
        .map(|child| child.unwrap().kind().to_owned())
        .collect();
    assert!(
        !body
            .iter()
            .any(|kind| kind == "shape" || kind == "equation")
    );
    // A part lists its own parts.
    let numbering = file.get("/numbering").unwrap();
    assert_eq!(numbering.view().children().count(), 4);
}

fn reopen(file: &File) -> File {
    open_bytes(file.to_bytes().expect("publish")).expect("reopen")
}

fn tab_paragraph(file: &File) -> String {
    file.query("p")
        .unwrap()
        .map(Result::unwrap)
        .find(|node| node.view().text().unwrap().as_deref() == Some("Name\tValue"))
        .unwrap()
        .path()
        .to_string()
}

#[test]
fn header_and_footer_text_sets_publish_styles_xml() {
    use litchi_ops::{Code, props};
    let file = parts();
    let edited = file
        .set("/header[2]", props! { "text" => "New even\ttop" })
        .unwrap();
    let edited = reopen(&edited);
    assert_eq!(text(&edited, "/header[2]"), "New even\ttop");
    assert_eq!(text(&edited, "/header[1]"), "Top 1");
    // The body is untouched.
    assert_eq!(text(&edited, "/body/p[1]"), text(&file, "/body/p[1]"));
    // Setting the shown text is an exact no-op.
    let again = edited
        .set("/header[2]", props! { "text" => "New even\ttop" })
        .unwrap();
    assert_eq!(again.to_bytes().unwrap(), edited.to_bytes().unwrap());
    // A region holding a field refuses the set instead of dropping it.
    let error = file
        .set("/header[1]", props! { "text" => "x" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    // Read-only keys are refused.
    let error = file
        .set("/footer[1]", props! { "type" => "even" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let footer = file
        .set("/footer[1]", props! { "text" => "Bottom line" })
        .unwrap();
    assert_eq!(text(&reopen(&footer), "/footer[1]"), "Bottom line");
}

#[test]
fn tab_stops_set_add_and_remove() {
    use litchi_ops::{Code, NewNode, props};
    let file = parts();
    let paragraph = tab_paragraph(&file);
    let tab = |n: usize| format!("{paragraph}/tab[{n}]");
    // Set the second stop's alignment and leader.
    let edited = reopen(
        &file
            .set(
                tab(2).as_str(),
                props! { "val" => "center", "leader" => "none" },
            )
            .unwrap(),
    );
    assert_eq!(prop(&edited, &tab(2), "val").as_deref(), Some("center"));
    assert_eq!(prop(&edited, &tab(2), "leader").as_deref(), Some("none"));
    assert_eq!(prop(&edited, &tab(1), "pos").as_deref(), Some("1417"));
    // Setting what a stop holds is an exact no-op.
    let same = file
        .set(tab(1).as_str(), props! { "pos" => 1417, "val" => "left" })
        .unwrap();
    assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());
    // Add a stop between the two: it takes position order.
    let added = file
        .add(
            paragraph.as_str(),
            NewNode::new("tab").prop("pos", 2880).prop("val", "decimal"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), tab(2));
    let added = reopen(&added.file);
    assert_eq!(count(&added, &format!("{paragraph}/tab")), 3);
    assert_eq!(prop(&added, &tab(2), "val").as_deref(), Some("decimal"));
    assert_eq!(prop(&added, &tab(3), "val").as_deref(), Some("right"));
    // A stop needs a position, and two stops cannot share one.
    let error = file
        .add(paragraph.as_str(), NewNode::new("tab").prop("val", "left"))
        .unwrap_err();
    assert_eq!(error.code(), Code::MissingProperty);
    let error = file
        .add(paragraph.as_str(), NewNode::new("tab").prop("pos", 1417))
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    // Remove the first stop.
    let removed = reopen(&file.remove(tab(1).as_str()).unwrap());
    assert_eq!(count(&removed, &format!("{paragraph}/tab")), 1);
    assert_eq!(prop(&removed, &tab(1), "val").as_deref(), Some("right"));
    // A tab stop on a paragraph without one starts its list.
    let first = file
        .add("/body/p[1]", NewNode::new("tab").prop("pos", 720))
        .unwrap()
        .file;
    assert_eq!(
        prop(&reopen(&first), "/body/p[1]/tab[1]", "pos").as_deref(),
        Some("720")
    );
    assert_eq!(
        text(&reopen(&first), "/body/p[1]"),
        text(&file, "/body/p[1]")
    );
}

#[test]
fn read_only_parts_refuse_writes() {
    use litchi_ops::{Code, props};
    let file = parts();
    for (path, key) in [
        ("/numbering/abstractNum[1]/level[1]", "start"),
        ("/field[1]", "fieldType"),
        ("/toc[1]", "levels"),
        ("/body/shape[1]", "width"),
        ("/body/equation[1]", "formula"),
        ("/body/tbl[1]/col[1]", "width"),
    ] {
        let error = file.set(path, props! { key => "1" }).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability, "{path}");
        let error = file.remove(path).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability, "{path}");
    }
}
