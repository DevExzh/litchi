use super::super::{TextIndex, TextIndexKind};
use super::xml::validated_indexes;
use super::*;

fn toc(name: &str) -> TextIndex {
    let mut source = TableOfContentsSource::new();
    source.outline_level = Some(10);
    source.title_template =
        Some(TextIndexTitleTemplate::new("Contents").with_style_name("Contents_20_Heading"));
    let mut entry = TextIndexEntryTemplate::new(1, "Contents_20_1");
    entry
        .push(TextIndexEntryToken::LinkStart { style_name: None })
        .push(TextIndexEntryToken::Text { style_name: None })
        .push(TextIndexEntryToken::TabStop(TextIndexTabStop::Right {
            leader: Some('.'),
            style_name: None,
        }))
        .push(TextIndexEntryToken::PageNumber { style_name: None })
        .push(TextIndexEntryToken::LinkEnd { style_name: None });
    source.push_entry_template(entry);
    let mut body = TextIndexBody::new();
    body.title = Some(TextIndexBodyTitle::new("Contents_Head", "Contents"));
    body.push_paragraph(TextIndexBodyParagraph::new("Cached heading  7"));
    TextIndex::table_of_contents(name, source, body).unwrap()
}

#[test]
fn typed_toc_is_canonical_and_round_trips() {
    let xml = toc("TOC1").with_protected(true).to_xml_fragment().unwrap();
    assert!(xml.starts_with("<text:table-of-content xmlns:style="));
    let document = format!(
        "<office:document-content xmlns:office=\"{}\"><office:body><office:text>{xml}</office:text></office:body></office:document-content>",
        std::str::from_utf8(OFFICE).unwrap()
    );
    let parsed = validated_indexes(&document).unwrap();
    assert_eq!(parsed[0].name(), "TOC1");
    assert_eq!(
        parsed[0].body().unwrap().all_text(),
        "ContentsCached heading  7"
    );
}

#[test]
fn mutation_preserves_unselected_xml_bytes() {
    let one = toc("one").to_xml_fragment().unwrap();
    let xml = format!(
        "<?xml version=\"1.0\"?><o:document-content xmlns:o=\"{}\"><o:body><o:text><!--keep--><text:p xmlns:text=\"{TEXT}\">before</text:p>{one}<text:p xmlns:text=\"{TEXT}\">after</text:p></o:text></o:body></o:document-content>",
        std::str::from_utf8(OFFICE).unwrap()
    );
    let replaced = replace_text_index_xml(&xml, "one", &toc("two")).unwrap();
    assert!(replaced.contains("<!--keep--><text:p xmlns:text="));
    assert!(replaced.contains(">before</text:p>"));
    assert!(replaced.contains(">after</text:p>"));
    let removed = remove_text_index_xml(&replaced, "two").unwrap();
    assert!(!removed.contains("table-of-content"));
}

#[test]
fn hostile_template_order_and_foreign_children_are_rejected() {
    let office = std::str::from_utf8(OFFICE).unwrap();
    let xml = format!(
        "<office:document-content xmlns:office=\"{office}\" xmlns:text=\"{TEXT}\" xmlns:u=\"urn:bad\"><office:body><office:text><text:table-of-content text:name=\"x\"><text:table-of-content-source><text:index-source-styles text:outline-level=\"1\"/><text:index-title-template>x</text:index-title-template></text:table-of-content-source><text:index-body/></text:table-of-content></office:text></office:body></office:document-content>"
    );
    assert!(validated_indexes(&xml).is_err());
    let foreign = xml
        .replace(
            "<text:index-source-styles text:outline-level=\"1\"/>",
            "<u:extension/>",
        )
        .replace(
            "<text:index-title-template>x</text:index-title-template>",
            "",
        );
    assert!(validated_indexes(&foreign).is_err());
}

#[test]
fn libreoffice_toc_fixture_round_trips_as_inert_markup() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(
        "../../test-data/libreoffice-core/sw/qa/extras/layout/data/toc-inline-heading-order.fodt",
    );
    let xml = std::fs::read_to_string(path).unwrap();
    let indexes = validated_indexes(&xml).unwrap();
    assert!(!indexes.is_empty());
    let fragment = indexes[0].to_xml_fragment().unwrap();
    assert!(fragment.contains("text:index-body"));
}

fn simple_template(style: &str) -> TextIndexSimpleEntryTemplate {
    let mut template = TextIndexSimpleEntryTemplate::new(style);
    template
        .push(TextIndexEntryToken::LinkStart { style_name: None })
        .push(TextIndexEntryToken::Text { style_name: None })
        .push(TextIndexEntryToken::PageNumber { style_name: None })
        .push(TextIndexEntryToken::LinkEnd { style_name: None });
    template
}

#[test]
fn shared_index_families_write_and_mutate_together() {
    let mut caption = IllustrationIndexSource::new();
    caption.use_caption = Some(true);
    caption.caption_sequence_name = Some("Figure".to_string());
    caption.caption_sequence_format = Some(TextIndexCaptionSequenceFormat::CategoryAndValue);
    caption.entry_template = Some(simple_template("Illustration_20_Index_20_1"));
    let illustration =
        TextIndex::illustration_index("figures", caption.clone(), TextIndexBody::new()).unwrap();
    let table = TextIndex::table_index("tables", caption, TextIndexBody::new()).unwrap();

    let mut objects = ObjectIndexSource::new();
    objects.use_chart_objects = Some(true);
    objects.use_math_objects = Some(false);
    objects.entry_template = Some(simple_template("Object_20_Index_20_1"));
    let object = TextIndex::object_index("objects", objects, TextIndexBody::new()).unwrap();

    let mut user = UserIndexSource::new("Glossary");
    user.use_index_marks = Some(true);
    let mut entry = TextIndexEntryTemplate::new(1, "User_20_Index_20_1");
    entry.push(TextIndexEntryToken::Text { style_name: None });
    user.push_entry_template(entry)
        .push_source_styles(TextIndexSourceStyles::new(1, vec!["Glossary".to_string()]));
    let user = TextIndex::user_index("user", user, TextIndexBody::new()).unwrap();

    let office = std::str::from_utf8(OFFICE).unwrap();
    let mut xml = format!(
        "<office:document-content xmlns:office=\"{office}\"><office:body><office:text/></office:body></office:document-content>"
    );
    for index in [&illustration, &table, &object, &user] {
        xml = insert_text_index_xml(&xml, index).unwrap();
    }
    let indexes = validated_indexes(&xml).unwrap();
    assert_eq!(
        indexes.iter().map(TextIndex::kind).collect::<Vec<_>>(),
        vec![
            TextIndexKind::Illustration,
            TextIndexKind::Table,
            TextIndexKind::Object,
            TextIndexKind::User
        ]
    );
    xml = replace_text_index_xml(
        &xml,
        "tables",
        &TextIndex::table_index(
            "new-tables",
            IllustrationIndexSource::new(),
            TextIndexBody::new(),
        )
        .unwrap(),
    )
    .unwrap();
    xml = remove_text_index_xml(&xml, "objects").unwrap();
    assert_eq!(validated_indexes(&xml).unwrap().len(), 3);
}

#[test]
fn hostile_shared_sources_are_rejected() {
    let mut source = IllustrationIndexSource::new();
    source.entry_template = Some(simple_template("Index"));
    let fragment = TextIndex::illustration_index("figures", source, TextIndexBody::new())
        .unwrap()
        .to_xml_fragment()
        .unwrap();
    let wrong_order = fragment.replace("<text:index-body/>", "").replace(
        "</text:illustration-index>",
        "<text:index-body/><text:illustration-index-source/></text:illustration-index>",
    );
    let office = std::str::from_utf8(OFFICE).unwrap();
    let document = format!(
        "<office:document-content xmlns:office=\"{office}\"><office:body><office:text>{wrong_order}</office:text></office:body></office:document-content>"
    );
    assert!(validated_indexes(&document).is_err());

    let mut user = UserIndexSource::new("Glossary");
    user.push_source_styles(TextIndexSourceStyles::new(1, vec![]));
    let user = TextIndex::user_index("user", user, TextIndexBody::new())
        .unwrap()
        .to_xml_fragment()
        .unwrap();
    let missing_name = user.replace(" text:index-name=\"Glossary\"", "");
    let document = format!(
        "<office:document-content xmlns:office=\"{office}\"><office:body><office:text>{missing_name}</office:text></office:body></office:document-content>"
    );
    assert!(validated_indexes(&document).is_err());
}

#[test]
fn libreoffice_table_index_source_round_trips_after_required_body_is_supplied() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/libreoffice-core/sw/qa/extras/uiwriter/data/IndexElementsInHiddenSections.fodt");
    let producer_xml = std::fs::read_to_string(path).unwrap();
    assert!(producer_xml.contains("<text:table-index-source text:use-caption=\"false\""));
    let start = producer_xml.find("<text:table-index>").unwrap();
    let closing = "</text:table-index>";
    let end = producer_xml[start..].find(closing).unwrap() + start + closing.len();
    let schema_complete_index = producer_xml[start..end]
        .replace(
            "<text:table-index>",
            "<text:table-index text:name=\"TableIndex1\">",
        )
        .replace(
            "</text:table-index>",
            "<text:index-body/></text:table-index>",
        );
    let office = std::str::from_utf8(OFFICE).unwrap();
    let schema_complete = format!(
        "<office:document-content xmlns:office=\"{office}\" xmlns:text=\"{TEXT}\" xmlns:style=\"{STYLE}\"><office:body><office:text>{schema_complete_index}</office:text></office:body></office:document-content>"
    );
    let indexes = validated_indexes(&schema_complete).unwrap();
    let table = indexes
        .iter()
        .find(|index| index.kind() == TextIndexKind::Table)
        .unwrap();
    let fragment = table.to_xml_fragment().unwrap();
    assert!(fragment.contains("text:caption-sequence-name=\"Table\""));
    assert!(fragment.contains("text:table-index-entry-template"));
}

#[test]
fn alphabetical_and_bibliography_write_and_mutate() {
    let mut alphabetical = AlphabeticalIndexSource::new();
    alphabetical.alphabetical_separators = Some(true);
    alphabetical.combine_entries = Some(true);
    alphabetical.language = Some("en".to_string());
    alphabetical.country = Some("US".to_string());
    alphabetical.rfc_language_tag = Some("en-US".to_string());
    let mut separator = TextAlphabeticalIndexEntryTemplate::new(
        TextAlphabeticalIndexLevel::Separator,
        "Alphabetical_20_Separator",
    );
    separator.push(TextIndexEntryToken::Text { style_name: None });
    let mut level = TextAlphabeticalIndexEntryTemplate::new(
        TextAlphabeticalIndexLevel::Level1,
        "Alphabetical_20_1",
    );
    level
        .push(TextIndexEntryToken::Text { style_name: None })
        .push(TextIndexEntryToken::TabStop(TextIndexTabStop::Right {
            leader: Some('.'),
            style_name: None,
        }))
        .push(TextIndexEntryToken::PageNumber { style_name: None });
    alphabetical
        .push_entry_template(separator)
        .push_entry_template(level);
    let alphabetical =
        TextIndex::alphabetical_index("alphabetical", alphabetical, TextIndexBody::new()).unwrap();

    let mut bibliography = BibliographyIndexSource::new();
    bibliography.title_template = Some(TextIndexTitleTemplate::new("Bibliography"));
    let mut book =
        TextBibliographyEntryTemplate::new(TextBibliographyType::Book, "Bibliography_20_1");
    book.push(TextBibliographyEntryToken::Field {
        field: crate::bibliography_configuration::Field::Identifier,
        style_name: None,
    })
    .push(TextBibliographyEntryToken::Span {
        style_name: None,
        text: ": ".to_string(),
    })
    .push(TextBibliographyEntryToken::Field {
        field: crate::bibliography_configuration::Field::Issn,
        style_name: None,
    });
    bibliography.push_entry_template(book);
    let bibliography =
        TextIndex::bibliography("bibliography", bibliography, TextIndexBody::new()).unwrap();
    let office = std::str::from_utf8(OFFICE).unwrap();
    let xml = format!(
        "<office:document-content xmlns:office=\"{office}\"><office:body><office:text/></office:body></office:document-content>"
    );
    let xml = insert_text_index_xml(&xml, &alphabetical).unwrap();
    let xml = insert_text_index_xml(&xml, &bibliography).unwrap();
    let parsed = validated_indexes(&xml).unwrap();
    assert_eq!(
        parsed.iter().map(TextIndex::kind).collect::<Vec<_>>(),
        vec![TextIndexKind::Alphabetical, TextIndexKind::Bibliography]
    );
    assert!(
        bibliography
            .to_xml_fragment()
            .unwrap()
            .contains("text:bibliography-data-field=\"issn\"")
    );
}

#[test]
fn hostile_distinct_index_grammars_are_rejected() {
    let mut alphabetical = AlphabeticalIndexSource::new();
    alphabetical.language = Some("bad_language".to_string());
    assert!(TextIndex::alphabetical_index("bad", alphabetical, TextIndexBody::new()).is_err());
    let mut alphabetical = AlphabeticalIndexSource::new();
    let mut template =
        TextAlphabeticalIndexEntryTemplate::new(TextAlphabeticalIndexLevel::Level1, "Index");
    template.push(TextIndexEntryToken::LinkStart { style_name: None });
    alphabetical.push_entry_template(template);
    assert!(TextIndex::alphabetical_index("bad", alphabetical, TextIndexBody::new()).is_err());

    let office = std::str::from_utf8(OFFICE).unwrap();
    let malformed = format!(
        "<office:document-content xmlns:office=\"{office}\" xmlns:text=\"{TEXT}\"><office:body><office:text><text:bibliography text:name=\"b\"><text:bibliography-source><text:bibliography-entry-template text:bibliography-type=\"invalid\" text:style-name=\"B\"><text:index-entry-text/></text:bibliography-entry-template></text:bibliography-source><text:index-body/></text:bibliography></office:text></office:body></office:document-content>"
    );
    assert!(validated_indexes(&malformed).is_err());
}

#[test]
fn libreoffice_bibliography_and_configuration_round_trip_inertly() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sw/qa/uibase/shells/data/protectedLinkCopy.fodt");
    let xml = std::fs::read_to_string(path).unwrap();
    assert!(xml.contains("<text:bibliography-configuration text:prefix=\"[\" text:suffix=\"]\""));
    let indexes = validated_indexes(&xml).unwrap();
    let bibliography = indexes
        .iter()
        .find(|index| index.kind() == TextIndexKind::Bibliography)
        .unwrap();
    let fragment = bibliography.to_xml_fragment().unwrap();
    assert!(fragment.contains("text:bibliography-type=\"article\""));
    assert!(fragment.contains("text:bibliography-data-field=\"author\""));
    assert!(fragment.contains("text:index-body"));
}
