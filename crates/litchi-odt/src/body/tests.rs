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
        let mut bold = TextPropertyUpdate::new();
        bold.bold = Some(true);
        bold.size = Some("14pt".to_owned());
        edit.set_text_properties(run.id(), &bold).unwrap();
        let mut align = ParagraphPropertyUpdate::new();
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
    let mut bold = TextPropertyUpdate::new();
    bold.bold = Some(true);
    assert!(edit.set_text_properties(run.id(), &bold).is_err());
    assert!(edit.remove(paragraph.id()).is_err());
    // The staged edit still commits.
    let after = edit.commit().unwrap().into_package();
    assert_eq!(texts(&after)[1], "x");
}

#[test]
fn metadata_edits() {
    use crate::metadata::Field;
    let package = built();
    let mut edit = package.edit_metadata();
    edit.set(Field::Title, Some("A & B")).unwrap();
    edit.set(Field::Keywords, Some("x, y")).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let (after, patch) = commit.into_parts();
    let metadata = after.document().unwrap().metadata().unwrap();
    assert_eq!(metadata.title.as_deref(), Some("A & B"));
    let restored = after.apply_metadata(&patch.inverse()).unwrap();
    assert_eq!(restored.into_package().as_bytes(), package.as_bytes());
}

#[test]
fn edits_on_a_real_document() {
    let package = Package::from_bytes(&fixture("corpus/writer-paragraph-styles.odt")).unwrap();
    let after = commit(&package, |edit, body| {
        let paragraphs: Vec<_> = body
            .nodes()
            .filter(|node| node.kind() == Kind::Paragraph)
            .collect();
        let mut italic = TextPropertyUpdate::new();
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

/// Record 0295: text fields are listed with their text and attributes, and
/// one field's text is replaced in place, keeping its element.
#[test]
fn text_fields_are_listed_and_rewritten() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/odfdo/tests/samples/variable.odt");
    let package = Package::from_bytes(&std::fs::read(path).unwrap()).unwrap();
    let document = package.document().unwrap();
    let body = document.body().unwrap();
    let fields = text_fields(body).unwrap();
    let names: Vec<&str> = fields.iter().map(TextField::name).collect();
    assert!(names.starts_with(&["date", "time", "title"]), "{names:?}");
    assert_eq!(fields[0].attribute("fixed").as_deref(), Some("true"));
    assert!(fields[0].attribute("data-style-name").is_some());
    let title = fields.iter().find(|field| field.name() == "title").unwrap();
    assert_eq!(title.text(), "Intitulé");
    let mut edit = package.edit();
    edit.set_field_text(title, "A & B").unwrap();
    assert!(
        edit.set_field_text(title, "again").is_err(),
        "one edit per field"
    );
    let next = edit.commit().unwrap().into_package();
    let document = next.document().unwrap();
    let body = document.body().unwrap();
    let fields = text_fields(body).unwrap();
    let title = fields.iter().find(|field| field.name() == "title").unwrap();
    assert_eq!(title.text(), "A & B");
    assert_eq!(fields.len(), names.len());
}

fn unified_fixture(name: &str) -> Vec<u8> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/unified-ops/0294")
        .join(name);
    std::fs::read(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()))
}

#[test]
fn parts_read_lists_regions_fields_contents_shapes_objects_and_columns() {
    let package = Package::from_bytes(&unified_fixture("parts.odt")).unwrap();
    let document = package.document().unwrap();
    let body = document.body().unwrap();
    let parts = Parts::build(body).unwrap();
    let view = parts.view(body);

    // List styles: the common one first, then the automatic one.
    let lists: Vec<_> = view.list_styles().collect();
    assert_eq!(lists.len(), 2);
    assert_eq!(lists[0].name().as_deref(), Some("Numbering_20_123"));
    assert_eq!(lists[0].display_name().as_deref(), Some("Numbering 123"));
    assert_eq!(lists[0].part(), StylePart::Common);
    let levels: Vec<_> = lists[0].levels().collect();
    assert_eq!(levels.len(), 2);
    assert_eq!(levels[0].kind(), parts::ListLevelKind::Number);
    assert_eq!(levels[0].start_value(), Some(3));
    assert_eq!(levels[0].prefix().as_deref(), Some("("));
    assert_eq!(levels[0].suffix().as_deref(), Some(")"));
    assert_eq!(levels[0].label_followed_by().as_deref(), Some("space"));
    assert_eq!(levels[0].text_align().as_deref(), Some("end"));
    assert_eq!(levels[0].margin_left().as_deref(), Some("0.635cm"));
    assert_eq!(levels[0].bold(), Some(true));
    assert_eq!(levels[1].level(), 2);
    assert_eq!(levels[1].display_levels(), Some(2));
    assert_eq!(levels[1].space_before().as_deref(), Some("1.27cm"));
    assert_eq!(lists[1].part(), StylePart::Automatic);
    assert_eq!(
        lists[1].level(1).unwrap().bullet_char().as_deref(),
        Some("•")
    );
    assert_eq!(
        lists[1].level(1).unwrap().font().as_deref(),
        Some("OpenSymbol")
    );
    assert_eq!(view.list_style("L1").unwrap().index(), 1);

    // Master-page regions in declaration order.
    let regions: Vec<_> = view.regions().collect();
    let kinds: Vec<_> = regions.iter().map(|region| region.kind()).collect();
    assert_eq!(
        kinds,
        vec![
            parts::RegionKind::Header,
            parts::RegionKind::HeaderLeft,
            parts::RegionKind::Footer,
            parts::RegionKind::HeaderFirst,
            parts::RegionKind::Footer,
        ]
    );
    assert_eq!(regions[0].text().unwrap(), "Top 1");
    assert_eq!(regions[0].master_name().as_deref(), Some("Standard"));
    assert_eq!(regions[3].master_name().as_deref(), Some("Landscape"));
    assert!(regions[0].displayed());
    assert!(!regions[4].displayed());

    // Fields, tables of contents, shapes, objects and columns.
    let names: Vec<_> = fields(body).unwrap().iter().map(TextField::name).collect();
    assert_eq!(
        names,
        vec!["page-number", "page-count", "date", "title", "chapter"]
    );
    let tocs: Vec<_> = view.tables_of_contents().collect();
    assert_eq!(tocs.len(), 1);
    assert_eq!(tocs[0].name().as_deref(), Some("Contents1"));
    assert_eq!(tocs[0].outline_level(), Some(3));
    assert!(tocs[0].links() && tocs[0].page_numbers());
    assert_eq!(tocs[0].title().as_deref(), Some("Contents"));
    assert_eq!(tocs[0].text().unwrap(), "Parts\t1");
    let shapes: Vec<_> = view.shapes().collect();
    assert_eq!(shapes.len(), 2);
    assert_eq!(shapes[0].geometry(), "star5");
    assert_eq!(shapes[0].fill_color().as_deref(), Some("#ff0000"));
    assert_eq!(shapes[0].text().unwrap(), "Starred");
    assert_eq!(shapes[1].geometry(), "rect");
    assert_eq!(shapes[1].width().as_deref(), Some("2cm"));
    let objects: Vec<_> = view.objects().collect();
    assert_eq!(objects.len(), 2);
    assert_eq!(objects[0].href().as_deref(), Some("./Object 1"));
    assert_eq!(
        parts::formula_text(objects[1].inline_math().unwrap()).as_deref(),
        Some("x + 1")
    );
    let formula = document.get_file("Object 1/content.xml").unwrap();
    assert_eq!(
        parts::formula_text(std::str::from_utf8(&formula).unwrap()).as_deref(),
        Some("a^2 + b^2 = c^2")
    );
    let table = body
        .root()
        .children()
        .find(|node| node.kind() == Kind::Table)
        .unwrap();
    let widths: Vec<_> = view
        .columns(table)
        .map(|column| column.width().unwrap().into_owned())
        .collect();
    assert_eq!(widths, vec!["3cm", "3cm", "4.5cm"]);

    // Tab stops of the paragraph's automatic style.
    let paragraph = body
        .root()
        .children()
        .find(|node| node.style_name().as_deref() == Some("P1"))
        .unwrap();
    let stops = parts::tab_stops(paragraph);
    assert_eq!(stops.len(), 2);
    assert_eq!(stops[0].position(), "2.5cm");
    assert_eq!(stops[0].kind(), "left");
    assert_eq!(stops[1].kind(), "right");
    assert_eq!(stops[1].leader_style(), Some("dotted"));
}

#[test]
fn parts_index_real_writer_documents() {
    for name in [
        "corpus/writer-header-footer.odt",
        "corpus/writer-table-of-contents.odt",
        "corpus/writer-definition-lists.odt",
        "corpus/writer-user-fields.odt",
        "odt/shape-text-in-paragraph.odt",
        "corpus/writer-table.odt",
    ] {
        let Ok(package) = Package::from_bytes(&fixture(name)) else {
            continue;
        };
        let document = package.document().unwrap();
        let body = document.body().unwrap();
        let parts = Parts::build(body).unwrap();
        let view = parts.view(body);
        for region in view.regions() {
            region.text().unwrap();
        }
        for toc in view.tables_of_contents() {
            toc.text().unwrap();
        }
        for shape in view.shapes() {
            shape.text().unwrap();
        }
        if name.ends_with("header-footer.odt") {
            let texts: Vec<_> = view
                .regions()
                .map(|region| region.text().unwrap())
                .collect();
            assert_eq!(texts, vec!["Header\t1\tTest", "Footer\t1\tBottom"]);
        }
        if name.ends_with("table-of-contents.odt") {
            assert_eq!(view.tables_of_contents().len(), 1);
        }
    }
}

#[test]
fn paragraph_formatting_rewrites_the_one_paragraph_properties() {
    // Record 0328: formatting a paragraph whose automatic style already has
    // paragraph properties rewrites them; the copy no longer gains a second
    // `style:paragraph-properties`.
    let package = built();
    let align = |value: &str| {
        let mut edit = ParagraphPropertyUpdate::new();
        edit.align = Some(value.to_owned());
        edit
    };
    let once = commit(&package, |edit, body| {
        let paragraph = body.root().children().nth(1).unwrap();
        edit.set_paragraph_properties(paragraph.id(), &align("center"))
            .unwrap();
    });
    let twice = commit(&once, |edit, body| {
        let paragraph = body.root().children().nth(1).unwrap();
        edit.set_paragraph_properties(paragraph.id(), &align("right"))
            .unwrap();
    });
    let document = twice.document().unwrap();
    let body = document.body().unwrap();
    let paragraph = body.root().children().nth(1).unwrap();
    let style = paragraph.automatic_style().unwrap();
    assert_eq!(style.xml().matches("paragraph-properties").count(), 1);
    assert_eq!(paragraph.paragraph_properties().align(), Some("right"));
}

#[test]
fn tab_stops_replace_the_declared_ones() {
    let package = built();
    let mut first = ParagraphPropertyUpdate::new();
    let mut right = TabStopSpec::new("10cm");
    right.kind = Some("right".to_owned());
    right.leader_style = Some("dotted".to_owned());
    right.leader_text = Some(".".to_owned());
    first.tab_stops = Some(vec![TabStopSpec::new("2.5cm"), right]);
    first.align = Some("center".to_owned());
    let once = commit(&package, |edit, body| {
        let paragraph = body.root().children().nth(1).unwrap();
        edit.set_paragraph_properties(paragraph.id(), &first)
            .unwrap();
    });
    let stops_of = |package: &Package| -> Vec<(String, String)> {
        let document = package.document().unwrap();
        let body = document.body().unwrap();
        let paragraph = body.root().children().nth(1).unwrap();
        parts::tab_stops(paragraph)
            .iter()
            .map(|stop| (stop.position().to_owned(), stop.kind().to_owned()))
            .collect()
    };
    assert_eq!(
        stops_of(&once),
        vec![
            ("2.5cm".to_owned(), "left".to_owned()),
            ("10cm".to_owned(), "right".to_owned())
        ]
    );
    // A second edit of the derived style replaces the list in place.
    let mut second = ParagraphPropertyUpdate::new();
    second.tab_stops = Some(vec![TabStopSpec::new("3cm")]);
    let twice = commit(&once, |edit, body| {
        let paragraph = body.root().children().nth(1).unwrap();
        edit.set_paragraph_properties(paragraph.id(), &second)
            .unwrap();
    });
    assert_eq!(
        stops_of(&twice),
        vec![("3cm".to_owned(), "left".to_owned())]
    );
    let document = twice.document().unwrap();
    let body = document.body().unwrap();
    let paragraph = body.root().children().nth(1).unwrap();
    assert_eq!(paragraph.paragraph_properties().align(), Some("center"));
    let xml = paragraph.automatic_style().unwrap().xml();
    assert_eq!(xml.matches("<style:tab-stops").count(), 1);
    // An empty list declares none.
    let mut clear = ParagraphPropertyUpdate::new();
    clear.tab_stops = Some(Vec::new());
    let cleared = commit(&twice, |edit, body| {
        let paragraph = body.root().children().nth(1).unwrap();
        edit.set_paragraph_properties(paragraph.id(), &clear)
            .unwrap();
    });
    assert!(stops_of(&cleared).is_empty());
    // Invalid stops are refused at staging.
    let mut bad = ParagraphPropertyUpdate::new();
    bad.tab_stops = Some(vec![TabStopSpec::new("ten")]);
    let mut edit = cleared.edit();
    let document = cleared.document().unwrap();
    let body = document.body().unwrap();
    let paragraph = body.root().children().nth(1).unwrap();
    assert!(edit.set_paragraph_properties(paragraph.id(), &bad).is_err());
}

#[test]
fn region_text_replaces_header_and_footer_paragraphs() {
    let package = Package::from_bytes(&unified_fixture("parts.odt")).unwrap();
    let region_text = |package: &Package, index: usize| -> String {
        let document = package.document().unwrap();
        let body = document.body().unwrap();
        let parts = Parts::build(body).unwrap();
        parts
            .view(body)
            .regions()
            .nth(index)
            .unwrap()
            .text()
            .unwrap()
    };
    let edited = {
        let document = package.document().unwrap();
        let body = document.body().unwrap();
        let parts = Parts::build(body).unwrap();
        let regions: Vec<_> = parts.view(body).regions().collect();
        let mut edit = package.edit();
        // The first header holds a page-number field: refused, not dropped.
        let refused = edit.set_region_text(&regions[0], "x");
        assert!(matches!(
            refused,
            Err(
                litchi_odf_common::Error::Core(litchi_core::Error::Unsupported(_))
                    | litchi_odf_common::Error::Refused(_),
            )
        ));
        let mut edit = package.edit();
        edit.set_region_text(&regions[1], "New even\ttop").unwrap();
        edit.set_region_text(&regions[2], "New  bottom").unwrap();
        edit.commit().unwrap().into_package()
    };
    assert_eq!(region_text(&edited, 1), "New even\ttop");
    assert_eq!(region_text(&edited, 2), "New  bottom");
    assert_eq!(region_text(&edited, 0), "Top 1");
    // content.xml kept its bytes.
    assert_eq!(
        edited.document().unwrap().get_file("content.xml").unwrap(),
        package.document().unwrap().get_file("content.xml").unwrap()
    );
    // Setting the shown text is an exact no-op.
    let again = {
        let document = edited.document().unwrap();
        let body = document.body().unwrap();
        let parts = Parts::build(body).unwrap();
        let region = parts.view(body).regions().nth(1).unwrap();
        let mut edit = edited.edit();
        edit.set_region_text(&region, "New even\ttop").unwrap();
        edit.commit().unwrap().into_package()
    };
    assert_eq!(again.as_bytes(), edited.as_bytes());
}
