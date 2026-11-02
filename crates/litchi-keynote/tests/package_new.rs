//! Fresh Keynote authoring: `Package::new`, `Builder` and explicit
//! `Identity` (record 0298).

use litchi_keynote::builder::Error as BuilderError;
use litchi_keynote::identity::{Error as IdentityError, Role};
use litchi_keynote::show::Size;
use litchi_keynote::slide::creation::LayoutSelector;
use litchi_keynote::{Builder, Identity, Package, Position, SlideSelector};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const DEFAULT_TITLE: &str = "Presentation Title";
const DEFAULT_SUBTITLE: &str = "Presentation Subtitle";

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn fixed(seed: u8) -> TestResult<Identity> {
    Ok(Identity::from_bytes(
        [seed; 16],
        [seed.wrapping_add(1); 16],
        [seed.wrapping_add(2); 16],
    )?)
}

/// Return the stored ZIP member `name` from a package whose members are all
/// stored uncompressed (true for every package this crate authors).
fn stored_member(archive: &[u8], name: &str) -> Option<Vec<u8>> {
    let signature = [0x50, 0x4B, 0x03, 0x04];
    let mut offset = 0;
    while offset + 30 <= archive.len() && archive[offset..offset + 4] == signature {
        let field = |at: usize, width: usize| {
            archive[offset + at..offset + at + width]
                .iter()
                .rev()
                .fold(0_usize, |value, byte| (value << 8) | usize::from(*byte))
        };
        let size = field(18, 4);
        let name_length = field(26, 2);
        let extra_length = field(28, 2);
        let name_start = offset + 30;
        let data_start = name_start + name_length + extra_length;
        if &archive[name_start..name_start + name_length] == name.as_bytes() {
            return Some(archive[data_start..data_start + size].to_vec());
        }
        offset = data_start + size;
    }
    None
}

#[test]
fn new_package_opens_validates_and_has_one_default_slide() -> TestResult {
    let package = Package::new();
    package.validate()?;
    assert_eq!(package.slides()?.len(), 1);
    let first = SlideSelector::index(0);
    assert_eq!(package.slide_title(first)?.as_deref(), Some(DEFAULT_TITLE));
    assert_eq!(
        package.slide_body(first)?.as_deref(),
        Some(DEFAULT_SUBTITLE)
    );
    assert_eq!(package.show_settings()?.size(), Size::new(1920.0, 1080.0)?);
    assert_eq!(package.identity()?, Some(Identity::BLANK));
    assert_eq!(Package::default().identity()?, Some(Identity::BLANK));
    Ok(())
}

#[test]
fn new_package_is_deterministic() -> TestResult {
    assert_eq!(bytes(&Package::new())?, bytes(&Package::new())?);
    Ok(())
}

#[test]
fn new_package_metadata_reads() -> TestResult {
    let metadata = Package::new().metadata()?;
    assert!(metadata.is_some());
    Ok(())
}

#[test]
fn new_package_round_trips_through_write_to() -> TestResult {
    let package = Package::new();
    let written = bytes(&package)?;
    let reopened = Package::from_bytes(&written)?;
    reopened.validate()?;
    assert_eq!(bytes(&reopened)?, written);
    assert_eq!(reopened.identity()?, Some(Identity::BLANK));
    Ok(())
}

#[test]
fn builder_without_identity_is_a_typed_error() {
    assert!(matches!(
        Builder::new().title("x").build(),
        Err(BuilderError::MissingIdentity)
    ));
}

#[test]
fn builder_with_blank_identity_reproduces_new() -> TestResult {
    let built = Builder::new().identity(Identity::BLANK).build()?;
    assert_eq!(bytes(&built)?, bytes(&Package::new())?);
    Ok(())
}

#[test]
fn same_fixed_identity_builds_identical_bytes() -> TestResult {
    let build = || -> TestResult<Vec<u8>> {
        bytes(
            &Builder::new()
                .identity(fixed(7)?)
                .title("Title")
                .subtitle("Subtitle")
                .notes("Notes")
                .slide_size(Size::new(1024.0, 768.0)?)
                .build()?,
        )
    };
    assert_eq!(build()?, build()?);
    Ok(())
}

#[test]
fn different_identities_differ_in_the_document_identifier() -> TestResult {
    let first = bytes(&Builder::new().identity(fixed(1)?).build()?)?;
    let second = bytes(&Builder::new().identity(fixed(9)?).build()?)?;
    let first_identifier = stored_member(&first, "Metadata/DocumentIdentifier");
    let second_identifier = stored_member(&second, "Metadata/DocumentIdentifier");
    assert_eq!(
        first_identifier.as_deref(),
        Some(b"01010101-0101-0101-0101-010101010101".as_slice())
    );
    assert_eq!(
        second_identifier.as_deref(),
        Some(b"09090909-0909-0909-0909-090909090909".as_slice())
    );
    let package = Package::from_bytes(&second)?;
    package.validate()?;
    assert_eq!(package.identity()?, Some(fixed(9)?));
    Ok(())
}

#[test]
fn generated_identities_are_distinct() -> TestResult {
    let first = Identity::generate();
    let second = Identity::generate();
    assert_ne!(first, second);
    for identity in [first, second] {
        assert_ne!(identity.document(), identity.version());
        assert_ne!(identity.document(), identity.private());
        assert_ne!(identity.version(), identity.private());
        assert_eq!(identity.document()[6] >> 4, 4);
    }
    let package = Builder::new().identity(first).build()?;
    assert_eq!(package.identity()?, Some(first));
    Ok(())
}

#[test]
fn invalid_fixed_identities_are_rejected() {
    assert_eq!(
        Identity::from_bytes([0; 16], [1; 16], [2; 16]),
        Err(IdentityError::Nil {
            role: Role::DocumentUuid
        })
    );
    assert_eq!(
        Identity::from_bytes([1; 16], [2; 16], [2; 16]),
        Err(IdentityError::Repeated {
            first: Role::VersionUuid,
            second: Role::PrivateUuid
        })
    );
}

#[test]
fn built_values_read_back() -> TestResult {
    let package = Builder::new()
        .identity(fixed(3)?)
        .title("Quarterly review")
        .subtitle("October 2026")
        .notes("Open with the headline number.")
        .slide_size(Size::new(1024.0, 768.0)?)
        .build()?;
    package.validate()?;
    let first = SlideSelector::index(0);
    assert_eq!(
        package.slide_title(first)?.as_deref(),
        Some("Quarterly review")
    );
    assert_eq!(package.slide_body(first)?.as_deref(), Some("October 2026"));
    assert_eq!(
        package.slide_notes(first)?.as_deref(),
        Some("Open with the headline number.")
    );
    assert_eq!(package.show_settings()?.size(), Size::new(1024.0, 768.0)?);

    let reopened = Package::from_bytes(&bytes(&package)?)?;
    assert_eq!(
        reopened.slide_title(first)?.as_deref(),
        Some("Quarterly review")
    );
    assert_eq!(reopened.identity()?, Some(fixed(3)?));
    Ok(())
}

#[test]
fn slide_creation_works_on_a_fresh_package() -> TestResult {
    let package = Package::new();
    let layouts = package.slide_layouts()?;
    assert!(!layouts.is_empty());

    let mut edit = package.edit_slide_creation();
    edit.duplicate_slide(SlideSelector::index(0), Position::new(1))?;
    edit.add_slide(LayoutSelector::Position(Position::new(0)), Position::new(2))?;
    let created = edit.commit()?.into_package();
    created.validate()?;
    assert_eq!(created.slides()?.len(), 3);
    assert_eq!(
        created.slide_title(SlideSelector::index(1))?.as_deref(),
        Some(DEFAULT_TITLE)
    );
    Ok(())
}

#[test]
fn builder_writes_language_and_locale() -> Result<(), Box<dyn std::error::Error>> {
    use prost::Message as _;
    let package = Builder::new()
        .identity(Identity::BLANK)
        .language("fr")
        .locale("fr_FR")
        .build()?;
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    let catalog = litchi_iwa_archive::package::Catalog::from_bytes(&bytes, Default::default())?;
    let document = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Document.iwa")
        .ok_or("document member")?;
    let stream = litchi_iwa_core::SnappyStream::decompress(document.data(), Default::default())?;
    let archive = litchi_iwa_core::Archive::parse(&stream.into_bytes(), Default::default())?;
    let root = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find(|message| message.type_ == 1)
        .ok_or("document archive")?;
    #[allow(deprecated, reason = "native locale fields")]
    {
        let decoded = litchi_iwa_protos::kn::DocumentArchive::decode(root.data.as_slice())?;
        assert_eq!(decoded.super_.document_language.as_deref(), Some("fr"));
        assert_eq!(
            decoded.super_.super_.locale_identifier.as_deref(),
            Some("fr_FR")
        );
        assert_eq!(
            decoded.super_.super_.creation_locale_identifier.as_deref(),
            Some("fr_FR")
        );
    }
    assert!(matches!(
        Builder::new().identity(Identity::BLANK).locale(" ").build(),
        Err(litchi_keynote::builder::Error::InvalidLocale)
    ));
    Ok(())
}

#[test]
fn builder_shows_the_slide_number() -> TestResult {
    use litchi_keynote::slide::placeholder::{Kind, State};
    let first = SlideSelector::index(0);
    let shown = Builder::new()
        .identity(fixed(5)?)
        .slide_number_visible(true)
        .build()?;
    shown.validate()?;
    assert_eq!(
        shown.slide_placeholder_visibility(first, Kind::SlideNumber)?,
        Some(State::Visible)
    );
    let hidden = Builder::new()
        .identity(fixed(5)?)
        .slide_number_visible(false)
        .build()?;
    assert_eq!(
        hidden.slide_placeholder_visibility(first, Kind::SlideNumber)?,
        Some(State::Hidden)
    );
    assert_eq!(
        bytes(&hidden)?,
        bytes(&Builder::new().identity(fixed(5)?).build()?)?
    );
    Ok(())
}
