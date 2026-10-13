use super::*;
use crate::Package;

fn fixture(name: &str) -> Vec<u8> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/odf")
        .join(name);
    std::fs::read(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()))
}

fn built() -> Package {
    let mut builder = crate::Builder::new();
    builder.add_heading("Title", 1).unwrap();
    builder.add_paragraph("First  paragraph").unwrap();
    builder
        .add_rich_paragraph(vec![("plain ", None), ("styled", None)])
        .unwrap();
    builder.add_bulleted_list(vec!["one", "two"]).unwrap();
    builder.add_paragraph("Last").unwrap();
    builder.build().unwrap()
}

fn kinds(body: Body<'_>) -> Vec<Kind> {
    body.root().children().map(Node::kind).collect()
}

#[test]
fn indexes_builder_document_in_order() {
    let package = built();
    let document = package.document().unwrap();
    let body = document.body().unwrap();
    assert_eq!(
        kinds(body),
        vec![
            Kind::Heading,
            Kind::Paragraph,
            Kind::Paragraph,
            Kind::List,
            Kind::Paragraph
        ]
    );
    let children: Vec<_> = body.root().children().collect();
    assert_eq!(children[0].text().unwrap(), "Title");
    assert_eq!(
        children[0].attribute(Attribute::OutlineLevel).as_deref(),
        Some("1")
    );
    // ODF collapses white space in character data; the builder wrote two.
    assert_eq!(children[1].text().unwrap(), "First paragraph");
    let list_items: Vec<_> = children[3].children().collect();
    assert_eq!(list_items.len(), 2);
    assert_eq!(list_items[1].text().unwrap(), "two");
    // Every node's text reads.
    for node in body.nodes() {
        node.text().unwrap();
    }
}

#[test]
fn indexes_real_writer_documents() {
    for name in [
        "corpus/writer-paragraph-styles.odt",
        "corpus/writer-table.odt",
        "corpus/writer-images-frames.odt",
        "corpus/writer-user-fields.odt",
        "corpus/writer-definition-lists.odt",
        "corpus/writer-header-footer.odt",
        "corpus/writer-table-of-contents.odt",
        "corpus/writer-minimal-nasty.odt",
        "odt/table-cell-column-span.odt",
        "odt/shape-text-in-paragraph.odt",
    ] {
        // Fixtures the leaf itself refuses at open are out of scope here.
        let Ok(package) = Package::from_bytes(&fixture(name)) else {
            continue;
        };
        let document = package.document().unwrap();
        let body = document.body().unwrap();
        for node in body.nodes() {
            node.text()
                .unwrap_or_else(|error| panic!("{name}: {error}"));
        }
        // Paragraph text agrees with the leaf paragraph projection.
        let leaf: Vec<String> = document
            .paragraphs()
            .unwrap()
            .iter()
            .map(|paragraph| paragraph.text().unwrap())
            .collect();
        let indexed = body
            .nodes()
            .filter(|node| node.kind() == Kind::Paragraph)
            .count();
        assert!(indexed >= leaf.len().min(1), "{name}");
    }
}

#[test]
fn table_cells_and_spans() {
    let package = Package::from_bytes(&fixture("odt/table-cell-column-span.odt")).unwrap();
    let document = package.document().unwrap();
    let body = document.body().unwrap();
    let table = body
        .nodes()
        .find(|node| node.kind() == Kind::Table)
        .unwrap();
    let rows: Vec<_> = table.children().filter(|n| n.kind() == Kind::Row).collect();
    assert!(!rows.is_empty());
    assert!(
        rows[0]
            .children()
            .all(|cell| matches!(cell.kind(), Kind::Cell | Kind::CoveredCell))
    );
}

fn commit(
    package: &Package,
    stage: impl FnOnce(&mut crate::transaction::Edit, Body<'_>),
) -> Package {
    let document = package.document().unwrap();
    let body = document.body().unwrap();
    let mut edit = package.edit();
    stage(&mut edit, body);
    edit.commit().unwrap().into_package()
}

fn texts(package: &Package) -> Vec<String> {
    let document = package.document().unwrap();
    let body = document.body().unwrap();
    body.root()
        .children()
        .map(|node| node.text().unwrap().into_owned())
        .collect()
}

#[test]
fn set_text_keeps_other_bytes() {
    let package = built();
    let after = commit(&package, |edit, body| {
        let paragraph = body.root().children().nth(1).unwrap();
        edit.set_text(paragraph.id(), "New  text\twith tab")
            .unwrap();
    });
    assert_eq!(texts(&after)[1], "New  text\twith tab");
    assert_eq!(texts(&after)[0], "Title");
    assert_eq!(texts(&after)[4], "Last");
    // Exact no-op: the same text publishes nothing.
    let again = commit(&after, |edit, body| {
        let paragraph = body.root().children().nth(4).unwrap();
        edit.set_text(paragraph.id(), "Last").unwrap();
    });
    assert_eq!(again.as_bytes(), after.as_bytes());
}

#[test]
fn run_formatting_derives_automatic_styles() {
    let package = built();
    let after = commit(&package, |edit, body| {
        let paragraph = body.root().children().nth(1).unwrap();
        let run = paragraph.children().next().unwrap();
        assert_eq!(run.kind(), Kind::Text);
        let mut bold = TextPropertyEdit::new();
        bold.bold = Some(true);
        bold.size = Some("14pt".to_owned());
        edit.set_text_properties(run.id(), &bold).unwrap();
        let mut align = ParagraphPropertyEdit::new();
        align.align = Some("center".to_owned());
        edit.set_paragraph_properties(paragraph.id(), &align)
            .unwrap();
    });
    let document = after.document().unwrap();
    let body = document.body().unwrap();
    let paragraph = body.root().children().nth(1).unwrap();
    assert_eq!(paragraph.paragraph_properties().align(), Some("center"));
    let run = paragraph.children().next().unwrap();
    assert_eq!(run.kind(), Kind::Span);
    assert_eq!(run.text().unwrap(), "First paragraph");
    assert_eq!(run.text_properties().bold(), Some(true));
    assert_eq!(run.text_properties().size(), Some("14pt"));
}

#[test]
fn structural_edits() {
    let package = built();
    let after = commit(&package, |edit, body| {
        let children: Vec<_> = body.root().children().collect();
        edit.insert(
            Location::new(NodeId::BODY, Placement::End),
            &NewNode::Paragraph {
                text: "appended".to_owned(),
                style: None,
            },
        )
        .unwrap();
        edit.insert(
            Location::new(NodeId::BODY, Placement::Start),
            &NewNode::Table {
                rows: 2,
                columns: 3,
            },
        )
        .unwrap();
        edit.remove(children[2].id()).unwrap();
        edit.relocate(
            children[0].id(),
            Location::new(NodeId::BODY, Placement::After(children[1].id())),
        )
        .unwrap();
    });
    let document = after.document().unwrap();
    let body = document.body().unwrap();
    assert_eq!(
        kinds(body),
        vec![
            Kind::Table,
            Kind::Paragraph,
            Kind::Heading,
            Kind::List,
            Kind::Paragraph,
            Kind::Paragraph
        ]
    );
    assert_eq!(texts(&after)[5], "appended");
    let table = body.root().children().next().unwrap();
    assert_eq!(table.children().count(), 2);
    assert_eq!(table.children().next().unwrap().children().count(), 3);
    // Swap the two last paragraphs.
    let swapped = commit(&after, |edit, body| {
        let children: Vec<_> = body.root().children().collect();
        edit.swap(children[4].id(), children[5].id()).unwrap();
    });
    let texts = texts(&swapped);
    assert_eq!(&texts[4..], ["appended", "Last"]);
}

#[test]
fn overlapping_edits_are_refused_at_staging() {
    let package = built();
    let document = package.document().unwrap();
    let body = document.body().unwrap();
    let paragraph = body.root().children().nth(1).unwrap();
    let run = paragraph.children().next().unwrap();
    let mut edit = package.edit();
    edit.set_text(paragraph.id(), "x").unwrap();
    let mut bold = TextPropertyEdit::new();
    bold.bold = Some(true);
    assert!(edit.set_text_properties(run.id(), &bold).is_err());
    assert!(edit.remove(paragraph.id()).is_err());
    // The staged edit still commits.
    let after = edit.commit().unwrap().into_package();
    assert_eq!(texts(&after)[1], "x");
}

#[test]
fn metadata_edits() {
    let package = built();
    let mut edit = package.edit();
    edit.set_metadata(MetadataField::Title, Some("A & B"))
        .unwrap();
    edit.set_metadata(MetadataField::Keywords, Some("x, y"))
        .unwrap();
    let after = edit.commit().unwrap().into_package();
    let metadata = after.document().unwrap().metadata().unwrap();
    assert_eq!(metadata.title.as_deref(), Some("A & B"));
}

#[test]
fn edits_on_a_real_document() {
    let package = Package::from_bytes(&fixture("corpus/writer-paragraph-styles.odt")).unwrap();
    let after = commit(&package, |edit, body| {
        let paragraphs: Vec<_> = body
            .nodes()
            .filter(|node| node.kind() == Kind::Paragraph)
            .collect();
        let mut italic = TextPropertyEdit::new();
        italic.italic = Some(true);
        edit.set_text_properties(paragraphs[0].id(), &italic)
            .unwrap();
        edit.set_text(paragraphs[1].id(), "changed").unwrap();
    });
    let document = after.document().unwrap();
    let body = document.body().unwrap();
    let paragraphs: Vec<_> = body
        .nodes()
        .filter(|node| node.kind() == Kind::Paragraph)
        .collect();
    assert_eq!(paragraphs[0].text_properties().italic(), Some(true));
    assert_eq!(paragraphs[1].text().unwrap(), "changed");
}
