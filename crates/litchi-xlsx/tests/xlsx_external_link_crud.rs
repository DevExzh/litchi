use litchi_opc::{PackURI, Part};
use litchi_xlsx::external_links::{
    ExternalDdeLink, ExternalLinkConformance, ExternalLinkKind, build_external_link_part,
    build_external_link_part_with_conformance,
};

fn dde(topic: &str) -> ExternalLinkKind {
    ExternalLinkKind::Dde(ExternalDdeLink {
        service: "Excel".into(),
        topic: topic.into(),
        items: Vec::new(),
    })
}

#[test]
fn external_link_parts_are_built_without_fetching_targets() {
    let kind = dde("https://127.0.0.1:9/never.xlsx");
    let part = build_external_link_part(
        PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
        &kind,
    )
    .unwrap();
    assert!(
        std::str::from_utf8(part.blob())
            .unwrap()
            .contains("never.xlsx")
    );
    assert!(part.rels().is_empty());
}

#[test]
fn strict_external_link_parts_keep_the_strict_namespace() {
    let part = build_external_link_part_with_conformance(
        PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
        &dde("strict"),
        ExternalLinkConformance::Strict,
    )
    .unwrap();
    assert!(
        std::str::from_utf8(part.blob())
            .unwrap()
            .contains("purl.oclc.org/ooxml/spreadsheetml/main")
    );
}
