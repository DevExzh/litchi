//! Keynote builder parity (wave W4-R, record 0417): the host builder
//! against `litchi_keynote::Builder`, every builder value (title, subtitle,
//! presenter notes, language, locale, slide size, slide-number visibility).
//!
//! The host builder draws a random document identity, template slide
//! identifier, theme UUID, hidden table-template UUID and row/column UIDs,
//! and object UUIDs. The leaf build takes the host output's document
//! identity, so the properties, document identifier and package metadata
//! revision compare exactly; the other random values are compared by form
//! (exact bytes with the one field masked, except the package metadata's
//! object UUIDs and the UID map), a documented policy. Record 0372 left the
//! ten builder rows host-only until this normalisation existed.

use litchi_iwa_protos::{kn, tswp};
use prost::Message as _;

use crate::archive_diff::{members, objects};
use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::known_fields::{Decoded, decoded};
use crate::leaf_bytes;

/// A hyphenated UUID string as its 16 bytes.
fn uuid_bytes(text: &str) -> TestResult<[u8; 16]> {
    let hex: String = text.chars().filter(|character| *character != '-').collect();
    Ok(u128::from_str_radix(&hex, 16)?.to_be_bytes())
}

/// The document, version and private UUIDs of a built package.
fn identity(bytes: &[u8]) -> TestResult<litchi_keynote::Identity> {
    let properties = members(bytes)?
        .into_iter()
        .find(|member| member.name == "Metadata/Properties.plist")
        .ok_or("no properties")?
        .data;
    let value = plist::Value::from_reader(std::io::Cursor::new(properties))?;
    let dictionary = value.as_dictionary().ok_or("properties")?;
    let field = |key: &str| -> TestResult<[u8; 16]> {
        uuid_bytes(
            dictionary
                .get(key)
                .and_then(plist::Value::as_string)
                .ok_or("missing UUID")?,
        )
    };
    Ok(litchi_keynote::Identity::from_bytes(
        field("documentUUID")?,
        field("versionUUID")?,
        field("privateUUID")?,
    )?)
}

/// `TSP.PackageMetadata` without its object UUIDs.
fn metadata_without_object_uuids(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut metadata = litchi_iwa_protos::tsp::PackageMetadata::decode(data)?;
    for component in metadata
        .components
        .iter_mut()
        .chain(metadata.versioned_components.iter_mut())
    {
        for entry in &mut component.object_uuid_map_entries {
            entry.uuid = Default::default();
        }
    }
    Ok(decoded(metadata, data))
}

/// Read one protobuf varint.
fn varint(data: &[u8], offset: &mut usize) -> Option<u64> {
    let mut value = 0_u64;
    for shift in (0..64).step_by(7) {
        let byte = *data.get(*offset)?;
        *offset += 1;
        value |= u64::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            return Some(value);
        }
    }
    None
}

/// The payload with the length-delimited top-level `field` replaced by an
/// empty value, every other byte kept. Comparing two such payloads is an
/// exact byte comparison with only that field compared by form (it must be
/// present exactly once, `length` bytes long when given). Unlike a
/// known-field decode it also covers unknown fields and extensions.
fn masked(data: &[u8], field: u32, length: Option<usize>) -> Result<Decoded, prost::DecodeError> {
    // A malformed payload fails the case: the builders only write
    // well-formed messages.
    let malformed = || -> prost::DecodeError { panic!("malformed payload") };
    let mut canonical = Vec::with_capacity(data.len());
    let mut offset = 0;
    let mut seen = 0;
    while offset < data.len() {
        let start = offset;
        let tag = varint(data, &mut offset).ok_or_else(malformed)?;
        let tag_end = offset;
        let skip = match tag & 7 {
            0 => varint(data, &mut offset).map(|_| 0),
            1 => Some(8),
            2 => varint(data, &mut offset).and_then(|value| usize::try_from(value).ok()),
            5 => Some(4),
            _ => None,
        }
        .ok_or_else(malformed)?;
        let end = offset.checked_add(skip).filter(|end| *end <= data.len());
        let end = end.ok_or_else(malformed)?;
        if tag >> 3 == u64::from(field) && tag & 7 == 2 {
            seen += 1;
            assert!(
                length.is_none_or(|length| skip == length),
                "field {field} form"
            );
            canonical.extend_from_slice(&data[start..tag_end]);
            canonical.push(0);
        } else {
            canonical.extend_from_slice(&data[start..end]);
        }
        offset = end;
    }
    assert_eq!(seen, 1, "field {field} present once");
    Ok(Decoded {
        debug: format!("{} bytes with field {field} masked", canonical.len()),
        canonical,
        unknown_fields: Default::default(),
    })
}

/// `TST.TableModelArchive` of the hidden table template with its table
/// UUID (field 1; the host's Numbers bootstrap draws it) masked.
fn model_without_table_uuid(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    masked(data, 1, Some(36))
}

/// `TST.ColumnRowUIDMapArchive` by form: its axis lengths, each index list
/// a permutation of the axis.
fn uid_map_by_form(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let map = litchi_iwa_protos::tst::ColumnRowUidMapArchive::decode(data)?;
    let permutation = |values: &[u32], length: usize| {
        let mut sorted = values.to_vec();
        sorted.sort_unstable();
        sorted == (0..u32::try_from(length).unwrap_or(u32::MAX)).collect::<Vec<_>>()
    };
    let columns = map.sorted_column_uids.len();
    let rows = map.sorted_row_uids.len();
    assert!(permutation(&map.column_index_for_uid, columns));
    assert!(permutation(&map.column_uid_for_index, columns));
    assert!(permutation(&map.row_index_for_uid, rows));
    assert!(permutation(&map.row_uid_for_index, rows));
    let form = litchi_iwa_protos::tst::ColumnRowUidMapArchive {
        sorted_column_uids: vec![Default::default(); columns],
        sorted_row_uids: vec![Default::default(); rows],
        ..Default::default()
    };
    Ok(decoded(form, data))
}

/// `KN.ThemeArchive` with its theme UUID string (field 3) masked.
fn theme_without_uuid(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    masked(data, 3, Some(36))
}

/// `KN.SlideNodeArchive` with its template slide identifier (field 29)
/// masked; the read `template_slide_ids` checks that both nodes share one.
fn node_without_template_id(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    masked(data, 29, None)
}

/// How many slide nodes carry a template slide identifier, and whether they
/// all carry the same one.
fn template_slide_ids(bytes: &[u8]) -> TestResult<(usize, bool)> {
    let mut identifiers = Vec::new();
    for member in members(bytes)? {
        if member.name != "Index/Document.iwa" {
            continue;
        }
        for object in objects(&member.data)? {
            for message in object.messages.iter().filter(|message| message.type_ == 4) {
                let node = kn::SlideNodeArchive::decode(message.data.as_slice())?;
                if let Some(identifier) = node.template_slide_id {
                    identifiers.push((identifier.lower, identifier.upper));
                }
            }
        }
    }
    let shared = identifiers.windows(2).all(|pair| pair[0] == pair[1]);
    Ok((identifiers.len(), shared))
}

/// `TSWP.StorageArchive` with the host's empty text (`text: [""]`) in the
/// form native Keynote writes for an empty storage: no text field at all.
/// Every native Keynote fixture under `test-data/iwork/keynote` stores its
/// empty storages that way (45–52 per presentation, none with `[""]`), and
/// the leaf text engine writes that form; the host builder's `[""]` is its
/// own encoding (record 0417, a documented difference).
fn storage_with_native_empty_text(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut storage = tswp::StorageArchive::decode(data)?;
    if storage.text.len() == 1 && storage.text[0].is_empty() {
        storage.text.clear();
    }
    Ok(decoded(storage, data))
}

/// The semantic reads of a built presentation: the first slide's title,
/// subtitle, notes and slide-number visibility, and the slide size.
fn semantic_reads(bytes: &[u8]) -> TestResult<String> {
    let package = litchi_keynote::Package::from_bytes(bytes)?;
    let first = litchi_keynote::SlideSelector::index(0);
    Ok(format!(
        "{:?} {:?} {:?} {:?} {:?}",
        package.slide_title(first)?,
        package.slide_body(first)?,
        package.slide_notes(first)?,
        package.slide_placeholder_visibility(
            first,
            litchi_keynote::slide::placeholder::Kind::SlideNumber
        )?,
        package.show_settings()?.size(),
    ))
}

/// The builder options one case sets (`None` keeps the default).
#[derive(Clone, Copy, Default)]
struct Options<'a> {
    title: Option<&'a str>,
    subtitle: Option<&'a str>,
    notes: Option<&'a str>,
    language: Option<&'a str>,
    locale: Option<&'a str>,
    size: Option<(f32, f32)>,
    slide_number_visible: Option<bool>,
    /// Compare an empty storage's text by its native form (see
    /// [`storage_with_native_empty_text`]).
    native_empty_text: bool,
}

fn run_options(name: &'static str, options: Options<'_>) -> TestResult {
    let Options {
        title,
        subtitle,
        notes,
        language,
        locale,
        size,
        slide_number_visible,
        native_empty_text,
    } = options;
    let mut host = litchi_iwa::keynote::Draft::builder();
    if let Some(title) = title {
        host = host.title(title);
    }
    if let Some(subtitle) = subtitle {
        host = host.subtitle(subtitle);
    }
    if let Some(notes) = notes {
        host = host.presenter_notes(notes);
    }
    if let Some(language) = language {
        host = host.language(language);
    }
    if let Some(locale) = locale {
        host = host.locale(locale);
    }
    if let Some((width, height)) = size {
        host = host.slide_size(width, height);
    }
    if let Some(visible) = slide_number_visible {
        host = host.slide_number_visible(visible);
    }
    let host_bytes = host.build()?.to_bytes()?;
    let identity = identity(&host_bytes)?;
    let leaf = move |_source: &[u8]| -> TestResult<LeafRun> {
        let mut builder = litchi_keynote::Builder::new().identity(identity);
        if let Some(title) = title {
            builder = builder.title(title);
        }
        if let Some(subtitle) = subtitle {
            builder = builder.subtitle(subtitle);
        }
        if let Some(notes) = notes {
            builder = builder.notes(notes);
        }
        if let Some(language) = language {
            builder = builder.language(language);
        }
        if let Some(locale) = locale {
            builder = builder.locale(locale);
        }
        if let Some((width, height)) = size {
            builder = builder.slide_size(litchi_keynote::show::Size::new(width, height)?);
        }
        if let Some(visible) = slide_number_visible {
            builder = builder.slide_number_visible(visible);
        }
        Ok(LeafRun::output(leaf_bytes!(&builder.build()?)))
    };
    let input = Fixture {
        path: "builder",
        bytes: leaf_bytes!(&litchi_keynote::Package::new()),
    };
    let host_output = host_bytes.clone();
    Case::new(name, Format::Keynote, Tier::B, Evidence::Queued, &input)
        .known_fields(|known| {
            let known = known
                .with(
                    11_006,
                    "TSP.PackageMetadata (object UUIDs by form)",
                    metadata_without_object_uuids,
                )
                .with(
                    6_001,
                    "TST.TableModelArchive (table UUID by form)",
                    model_without_table_uuid,
                )
                .with(
                    6_267,
                    "TST.ColumnRowUIDMapArchive (UIDs by form)",
                    uid_map_by_form,
                )
                .with(10, "KN.ThemeArchive (UUID by form)", theme_without_uuid)
                .with(
                    4,
                    "KN.SlideNodeArchive (template slide identifier by form)",
                    node_without_template_id,
                );
            if native_empty_text {
                known.with(
                    2_001,
                    "TSWP.StorageArchive (empty text in its native form)",
                    storage_with_native_empty_text,
                )
            } else {
                known
            }
        })
        .read("template_slide_ids", template_slide_ids)
        .read("semantic reads", semantic_reads)
        .host(move |_source| Ok(host_output.clone()))
        .leaf(leaf)
        .run()?;
    Ok(())
}

/// Every builder value at its default.
#[test]
fn keynote_builder_defaults_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    run_options("keynote_builder_defaults_parity", Options::default())
}

/// The first slide's title, subtitle and presenter notes.
#[test]
fn keynote_builder_text_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for (title, subtitle, notes) in [
        (
            "Quarterly review",
            "October 2026",
            "Open with the headline number.",
        ),
        ("Titre\nsur deux lignes", "Sous-titre", "Notes\n\nseconde"),
        ("日本語のタイトル", "サブタイトル", "発表者ノート"),
    ] {
        run_options(
            "keynote_builder_text_parity",
            Options {
                title: Some(title),
                subtitle: Some(subtitle),
                notes: Some(notes),
                ..Options::default()
            },
        )
        .map_err(|error| format!("{title:?}/{subtitle:?}/{notes:?}: {error}"))?;
    }
    Ok(())
}

/// An empty title, subtitle and presenter notes: the leaf writes empty
/// storages in their native form (no text field), the host as `[""]`.
#[test]
fn keynote_builder_empty_text_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    run_options(
        "keynote_builder_empty_text_parity",
        Options {
            title: Some(""),
            subtitle: Some(""),
            notes: Some(""),
            native_empty_text: true,
            ..Options::default()
        },
    )
}

/// The document language and locale, alone and with slide text.
#[test]
fn keynote_builder_language_locale_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for (language, locale) in [("fr", "fr_FR"), ("ja", "ja_JP"), ("de-CH", "de_CH")] {
        for text in [None, Some(("Titre", "Sous-titre", "Notes"))] {
            run_options(
                "keynote_builder_language_locale_parity",
                Options {
                    title: text.map(|text| text.0),
                    subtitle: text.map(|text| text.1),
                    notes: text.map(|text| text.2),
                    language: Some(language),
                    locale: Some(locale),
                    ..Options::default()
                },
            )
            .map_err(|error| format!("{language}/{locale}: {error}"))?;
        }
    }
    for (language, locale) in [(Some("fr"), None), (None, Some("fr_FR"))] {
        run_options(
            "keynote_builder_language_locale_parity",
            Options {
                language,
                locale,
                ..Options::default()
            },
        )
        .map_err(|error| format!("{language:?}/{locale:?}: {error}"))?;
    }
    Ok(())
}

/// The slide size, alone and with every other value.
#[test]
fn keynote_builder_slide_size_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for size in [(1024.0, 768.0), (1280.0, 720.0), (800.5, 600.25)] {
        run_options(
            "keynote_builder_slide_size_parity",
            Options {
                size: Some(size),
                ..Options::default()
            },
        )
        .map_err(|error| format!("{size:?}: {error}"))?;
    }
    run_options(
        "keynote_builder_slide_size_parity",
        Options {
            title: Some("Titre"),
            subtitle: Some("Sous-titre"),
            notes: Some("Notes"),
            language: Some("fr"),
            locale: Some("fr_FR"),
            size: Some((1024.0, 768.0)),
            slide_number_visible: Some(true),
            native_empty_text: false,
        },
    )
}

/// The first slide's slide-number placeholder visibility.
#[test]
fn keynote_builder_slide_number_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for visible in [true, false] {
        run_options(
            "keynote_builder_slide_number_parity",
            Options {
                slide_number_visible: Some(visible),
                title: Some("Numbered"),
                ..Options::default()
            },
        )
        .map_err(|error| format!("{visible}: {error}"))?;
    }
    Ok(())
}

/// `KEYNOTE_BUILDER_ROUTE`: `defaults`, `text` (title, subtitle, notes) or
/// `all` (text, `fr`/`fr_FR`, 1024×768, slide number).
#[test]
#[ignore = "callgrind cost probe (record 0417)"]
fn keynote_builder_cost_abba() -> TestResult {
    let route = std::env::var("KEYNOTE_BUILDER_ROUTE").unwrap_or_else(|_| "all".to_owned());
    let (text, all) = match route.as_str() {
        "defaults" => (false, false),
        "text" => (true, false),
        "all" => (true, true),
        other => return Err(format!("unknown route {other}").into()),
    };
    let host = |_source: &[u8]| -> TestResult<Vec<u8>> {
        let mut builder = litchi_iwa::keynote::Builder::new();
        if text {
            builder = builder
                .title("Quarterly review")
                .subtitle("October 2026")
                .presenter_notes("Open with the headline number.");
        }
        if all {
            builder = builder
                .language("fr")
                .locale("fr_FR")
                .slide_size(1024.0, 768.0)
                .slide_number_visible(true);
        }
        Ok(builder.build()?.to_bytes()?)
    };
    let leaf = |_source: &[u8]| -> TestResult<Vec<u8>> {
        let mut builder = litchi_keynote::Builder::new().identity(
            litchi_keynote::Identity::from_bytes([7; 16], [8; 16], [9; 16])?,
        );
        if text {
            builder = builder
                .title("Quarterly review")
                .subtitle("October 2026")
                .notes("Open with the headline number.");
        }
        if all {
            builder = builder
                .language("fr")
                .locale("fr_FR")
                .slide_size(litchi_keynote::show::Size::new(1024.0, 768.0)?)
                .slide_number_visible(true);
        }
        Ok(leaf_bytes!(&builder.build()?))
    };
    crate::pages_cost_probes::abba(
        &Fixture {
            path: "builder",
            bytes: Vec::new(),
        },
        host,
        leaf,
    )
}
