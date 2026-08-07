use std::collections::HashMap;
use std::env;

use litchi_iwa::raw::package::IWorkPackage;
use litchi_iwa_core::ArchiveObject;
use litchi_iwa_protos::tp::{
    DocumentArchive, SectionArchive, SectionTemplateArchive, SettingsArchive,
};
use litchi_iwa_protos::tswp::StorageArchive;
use litchi_pages::page_layout::Orientation;
use litchi_pages::section::{PageNumber, PageNumbering, Start};
use prost::Message;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args()
        .nth(1)
        .ok_or("usage: inspect_pages_structure <file>")?;
    let package = IWorkPackage::open(path)?;
    let mut objects: HashMap<u64, (String, ArchiveObject)> = HashMap::new();
    for name in package.entry_names().filter(|name| name.ends_with(".iwa")) {
        for object in package.archive(name)?.objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            objects.insert(identifier, (name.to_owned(), object));
        }
    }

    let (_, root) = objects.get(&1).ok_or("document object 1 is missing")?;
    let document = decode::<DocumentArchive>(root).ok_or("no Pages document payload")?;
    println!(
        "body={:?} section={:?} floating={:?}",
        document.body_storage.as_ref().map(|r| r.identifier),
        document.section.as_ref().map(|r| r.identifier),
        document.floating_drawables.as_ref().map(|r| r.identifier)
    );
    println!(
        "page={:?}x{:?} margins=({:?},{:?},{:?},{:?}) header_footer=({:?},{:?}) orientation={:?} single_header_footer={:?}",
        document.page_width,
        document.page_height,
        document.left_margin,
        document.right_margin,
        document.top_margin,
        document.bottom_margin,
        document.header_margin,
        document.footer_margin,
        document.orientation.map(Orientation::from_raw),
        document.uses_single_header_footer,
    );
    if let Some(reference) = &document.settings {
        let (_, object) = objects
            .get(&reference.identifier)
            .ok_or("settings object is missing")?;
        let settings = decode::<SettingsArchive>(object).ok_or("settings payload is invalid")?;
        println!(
            "settings={} body={:?} headers={:?} footers={:?} facing={:?} hyphenation={:?} ligatures={:?} language={:?} hyphenation_language={:?} footnotes=({:?},{:?},{:?},{:?})",
            reference.identifier,
            settings.body,
            settings.headers,
            settings.footers,
            settings.facing_pages,
            settings.hyphenation,
            settings.use_ligatures,
            settings.language,
            settings.hyphenation_language,
            settings.footnote_kind,
            settings.footnote_format,
            settings.footnote_numbering,
            settings.footnote_gap,
        );
    }

    if let Some(reference) = document.body_storage {
        let (name, object) = objects
            .get(&reference.identifier)
            .ok_or("body storage object is missing")?;
        let storage = decode::<StorageArchive>(object);
        println!(
            "body={} archive={} types={:?} text={:?}",
            reference.identifier,
            name,
            object.messages.iter().map(|m| m.type_).collect::<Vec<_>>(),
            storage.as_ref().map(|storage| &storage.text)
        );
        if let Some(storage) = storage {
            for entry in storage
                .table_section
                .into_iter()
                .flat_map(|table| table.entries)
            {
                if let Some(section) = entry.object {
                    inspect_section(&objects, section.identifier, entry.character_index)?;
                }
            }
        }
    }

    if let Some(reference) = document.section {
        let (name, object) = objects
            .get(&reference.identifier)
            .ok_or("section object is missing")?;
        let _ = (name, object);
        inspect_section(&objects, reference.identifier, 0)?;
    }

    Ok(())
}

fn inspect_section(
    objects: &HashMap<u64, (String, ArchiveObject)>,
    identifier: u64,
    character_index: u32,
) -> Result<(), Box<dyn std::error::Error>> {
    let (name, object) = objects
        .get(&identifier)
        .ok_or("section object is missing")?;
    let section = decode::<SectionArchive>(object).ok_or("section payload is invalid")?;
    let starting_page_number = section
        .section_page_number_start
        .map(PageNumber::new)
        .transpose()?;
    println!(
        "section={identifier} at={character_index} archive={name} name={:?} start={:?} \
         numbering={:?} starting_page={starting_page_number:?} first={:?} even={:?} odd={:?}",
        section.name,
        section.section_start_kind.map(Start::from_raw),
        section
            .section_page_number_kind
            .map(PageNumbering::from_raw),
        section
            .first_section_template_page
            .as_ref()
            .map(|reference| reference.identifier),
        section
            .even_section_template_page
            .as_ref()
            .map(|reference| reference.identifier),
        section
            .odd_section_template_page
            .as_ref()
            .map(|reference| reference.identifier),
    );
    for reference in [
        section.first_section_template_page,
        section.even_section_template_page,
        section.odd_section_template_page,
    ]
    .into_iter()
    .flatten()
    {
        inspect_template(objects, reference.identifier)?;
    }
    Ok(())
}

fn inspect_template(
    objects: &HashMap<u64, (String, ArchiveObject)>,
    identifier: u64,
) -> Result<(), Box<dyn std::error::Error>> {
    let (name, object) = objects
        .get(&identifier)
        .ok_or("section template object is missing")?;
    let template =
        decode::<SectionTemplateArchive>(object).ok_or("section template payload is invalid")?;
    println!(
        "  template={identifier} archive={name} headers={:?} footers={:?}",
        template
            .headers
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        template
            .footers
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>()
    );
    for (region, references) in [("header", template.headers), ("footer", template.footers)] {
        for (slot, reference) in references.into_iter().enumerate() {
            let (_, storage_object) = objects
                .get(&reference.identifier)
                .ok_or("header/footer storage object is missing")?;
            let storage = decode::<StorageArchive>(storage_object)
                .ok_or("header/footer storage payload is invalid")?;
            println!(
                "    {region}[{slot}]={} kind={:?} text={:?}",
                reference.identifier, storage.kind, storage.text
            );
        }
    }
    Ok(())
}

fn decode<T: Message + Default>(object: &ArchiveObject) -> Option<T> {
    object
        .messages
        .iter()
        .find_map(|message| T::decode(message.data.as_slice()).ok())
}
