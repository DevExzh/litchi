use litchi_docx::glossary::{Catalog, Conformance, Entry, Limits, Name, raw};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn entry(name: &str, text: &str) -> litchi_docx::Result<Entry> {
    Entry::new(
        name,
        format!(
            r#"<w:docPartBody xmlns:w="{W}"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:docPartBody>"#
        )
        .into_bytes(),
        &Limits::DEFAULT,
    )
}

#[test]
fn downstream_glossary_api_is_semantic_checked_and_move_first() -> litchi_docx::Result<()> {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Catalog>();

    let first = entry("First", "one")?;
    let body_pointer = first.body().expect("test entry body").as_ptr();
    let mut catalog = Catalog::new();
    assert_eq!(catalog.add(first)?, 0);
    assert_eq!(
        catalog
            .get("FIRST")?
            .and_then(Entry::body)
            .map(<[u8]>::as_ptr),
        Some(body_pointer)
    );

    let replaced = catalog
        .replace("first", entry("Second", "two")?)?
        .expect("selected entry");
    assert_eq!(replaced.name(), Some("First"));
    assert!(catalog.rename("second", Name::new("Renamed")?)?);
    assert_eq!(
        catalog.get("renamed")?.and_then(Entry::name),
        Some("Renamed")
    );
    assert!(catalog.at(usize::MAX).is_err());

    let xml = raw::write(&catalog, Conformance::Transitional, &Limits::DEFAULT)?;
    let (round_trip, conformance) = raw::read(&xml, &Limits::DEFAULT)?;
    assert_eq!(conformance, Conformance::Transitional);
    assert_eq!(
        round_trip.get("renamed")?.and_then(Entry::name),
        Some("Renamed")
    );
    Ok(())
}
