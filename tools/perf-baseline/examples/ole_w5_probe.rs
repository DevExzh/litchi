//! Paired native probe for the W5 OLE owner and user-defined-property waves.
//!
//! The probe deliberately keeps the fixture and timed operation boundaries
//! independent of the API revision.  On the pre-migration checkout, pass
//! `--no-default-features --features legacy-owner-api,legacy-property-api`;
//! on the current checkout leave both features disabled.  The output is one CSV row per measured
//! sample so an external ABBA runner can own CPU pinning, calibration, and
//! source/binary provenance.

use std::error::Error as StdError;
use std::hint::black_box;
use std::io::Cursor;
use std::time::Instant;

use litchi_cfb::{OleFile, OleWriter};
use litchi_ole_common::object::{Edit as ObjectEdit, Limits as ObjectLimits, Targets};
use litchi_ole_common::property_set::user_defined::{
    Edit as UserEdit, Hyperlink, Hyperlinks, Limits as UserLimits, LinkBase,
    Properties as UserProperties,
};
use litchi_ole_common::property_set::{
    CodePage, Section, USER_DEFINED_PROPERTIES_FMTID,
};
use litchi_ole_common::property_set::Value;

#[cfg(feature = "current-api")]
use litchi_ole_common::object::{OpenOptions, Package as ObjectPackage};

#[cfg(all(feature = "legacy-owner-api", not(feature = "current-api")))]
type OwnerSource = litchi_ole_common::object::Snapshot;
#[cfg(feature = "current-api")]
type OwnerSource = litchi_ole_common::object::Package;

const TARGET_NAME: &str = "ole_w5_edit_target.bin";
const ORIGINAL: &[u8] = b"ole-w5-original";
const REPLACEMENT: &[u8] = b"ole-w5-replacement";
const SIBLINGS: usize = 64;

struct Corpus {
    archive: Vec<u8>,
    limits: ObjectLimits,
    target_path: Vec<String>,
    user_section: Section,
    user_limits: UserLimits,
    link_base: LinkBase,
    hyperlinks: Hyperlinks,
}

fn main() -> Result<(), Box<dyn StdError>> {
    let (route, warmup, samples) = arguments()?;
    let corpus = build_corpus()?;
    let route_name = route.as_str();
    let mut prepared_owner = None;
    let mut prepared_section = None;

    if route_name == "begin_edit" {
        prepared_owner = Some(owner_open_source(&corpus)?);
    }
    if route_name == "read" {
        prepared_section = Some(corpus.user_section.clone());
    }

    println!("route,sample,nanoseconds,output_bytes");
    for iteration in 0..warmup.checked_add(samples).ok_or("sample count overflow")? {
        let started = Instant::now();
        let output_bytes = match route_name {
            "open" => {
                let editor = owner_open_editor(&corpus)?;
                black_box(editor.stream(&corpus.target_path));
                corpus.archive.len()
            },
            "begin_edit" => {
                let source = prepared_owner
                    .as_ref()
                    .ok_or("begin_edit source was not prepared")?;
                let editor = owner_start_edit(source)?;
                black_box(editor);
                corpus.archive.len()
            },
            "noop_commit" => {
                let editor = owner_open_editor(&corpus)?;
                let output = owner_commit_bytes(editor)?;
                verify_owner_output(&output, ORIGINAL)?;
                output.len()
            },
            "replace_stream" => {
                let mut editor = owner_open_editor(&corpus)?;
                editor.put_stream(&corpus.target_path, REPLACEMENT.to_vec())?;
                let output = owner_commit_bytes(editor)?;
                verify_owner_output(&output, REPLACEMENT)?;
                output.len()
            },
            "read" => {
                let section = prepared_section
                    .as_ref()
                    .ok_or("read section was not prepared")?;
                let properties = user_properties(section, corpus.user_limits)?;
                black_box((properties.link_base()?, properties.hyperlinks()?));
                section_size(section)
            },
            "edit" => {
                let mut section = corpus.user_section.clone();
                let link_base = corpus.link_base.clone();
                let hyperlinks = corpus.hyperlinks.clone();
                let mut edit = user_editor(&mut section, corpus.user_limits)?;
                edit.set_link_base(link_base)?;
                edit.set_hyperlinks(hyperlinks)?;
                black_box(&section);
                black_box(section_size(&section))
            },
            _ => return Err(format!("unknown route {route_name:?}").into()),
        };
        let elapsed = started.elapsed().as_nanos();
        if iteration >= warmup {
            println!(
                "{route_name},{},{elapsed},{output_bytes}",
                iteration - warmup
            );
        }
    }
    Ok(())
}

fn arguments() -> Result<(String, usize, usize), Box<dyn StdError>> {
    let mut route = None;
    let mut warmup = 10usize;
    let mut samples = 100usize;
    let mut arguments = std::env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--route" => route = Some(arguments.next().ok_or("--route needs a value")?),
            "--warmup" => warmup = arguments.next().ok_or("--warmup needs a value")?.parse()?,
            "--samples" => {
                samples = arguments.next().ok_or("--samples needs a value")?.parse()?
            },
            "--help" | "-h" => {
                println!(
                    "usage: ole_w5_probe --route <open|begin_edit|noop_commit|replace_stream|read|edit> [--warmup N] [--samples N]"
                );
                std::process::exit(0);
            },
            other => return Err(format!("unknown argument {other:?}").into()),
        }
    }
    let route = route.ok_or("--route is required")?;
    if samples == 0 {
        return Err("--samples must be nonzero".into());
    }
    Ok((route, warmup, samples))
}

fn build_corpus() -> Result<Corpus, Box<dyn StdError>> {
    let mut writer = OleWriter::new();
    for index in 0..SIBLINGS {
        let name = format!("ole_w5_sibling_{index:03}");
        let payload = format!("opaque sibling payload {index:03}");
        writer.create_stream(&[name.as_str()], payload.as_bytes())?;
    }
    writer.create_stream(&[TARGET_NAME], ORIGINAL)?;
    let mut archive = Cursor::new(Vec::new());
    writer.write_to(&mut archive)?;

    let stream_count = u64::try_from(SIBLINGS + 1)?;
    let total_size = u64::try_from(archive.get_ref().len())?;
    let limits = ObjectLimits {
        max_objects: 1,
        max_storage_depth: 1,
        max_streams_per_object: usize::try_from(stream_count)?,
        max_streams: usize::try_from(stream_count)?,
        max_stream_size: total_size.max(64),
        max_object_size: total_size.saturating_add(4096),
        max_total_size: total_size.saturating_add(4096),
    };

    let mut user_section = Section::new(USER_DEFINED_PROPERTIES_FMTID);
    user_section.set_page(CodePage::Utf16Le);
    let user_limits = UserLimits::default();
    let link_base = LinkBase::new("https://example.invalid/base/")?;
    let hyperlinks = Hyperlinks::new(vec![Hyperlink::new(
        7,
        11,
        0,
        "https://example.invalid/target",
        "bookmark",
    )?]);
    {
        let mut edit = user_editor(&mut user_section, user_limits)?;
        edit.set_link_base(link_base.clone())?;
        edit.set_hyperlinks(hyperlinks.clone())?;
    }
    // Keep one unrelated, opaque property so the typed overlay must preserve
    // the section's existing content while replacing only its reserved names.
    user_section.add(42, Value::Blob(vec![0x5a; 128]))?;

    Ok(Corpus {
        archive: archive.into_inner(),
        limits,
        target_path: vec![TARGET_NAME.to_owned()],
        user_section,
        user_limits,
        link_base,
        hyperlinks,
    })
}

fn owner_open_source(corpus: &Corpus) -> Result<OwnerSource, Box<dyn StdError>> {
    #[cfg(all(feature = "legacy-owner-api", not(feature = "current-api")))]
    {
        Ok(ObjectEdit::open(
            corpus.archive.clone(),
            Targets::default(),
            corpus.limits,
        )?
        .staged())
    }
    #[cfg(feature = "current-api")]
    {
        Ok(ObjectPackage::open_with_options(
            corpus.archive.clone(),
            OpenOptions::new()
                .targets(Targets::default())
                .limits(corpus.limits),
        )?)
    }
}

fn owner_open_editor(corpus: &Corpus) -> Result<ObjectEdit, Box<dyn StdError>> {
    let source = owner_open_source(corpus)?;
    owner_start_edit(&source)
}

fn owner_start_edit(source: &OwnerSource) -> Result<ObjectEdit, Box<dyn StdError>> {
    #[cfg(all(feature = "legacy-owner-api", not(feature = "current-api")))]
    {
        Ok(source.edit())
    }
    #[cfg(feature = "current-api")]
    {
        Ok(source.edit()?)
    }
}

fn owner_commit_bytes(editor: ObjectEdit) -> Result<Vec<u8>, Box<dyn StdError>> {
    let package = editor.commit()?.into_parts().0;
    #[cfg(all(feature = "legacy-owner-api", not(feature = "current-api")))]
    {
        Ok(package.finish()?)
    }
    #[cfg(feature = "current-api")]
    {
        let mut output = Vec::new();
        package.write_to(&mut output)?;
        Ok(output)
    }
}

fn verify_owner_output(output: &[u8], target: &[u8]) -> Result<(), Box<dyn StdError>> {
    let mut file = OleFile::open(Cursor::new(output))?;
    if file.open_stream(&[TARGET_NAME])? != target {
        return Err("owner probe output target differs from expected bytes".into());
    }
    Ok(())
}

fn user_properties<'a>(
    section: &'a Section,
    limits: UserLimits,
) -> Result<UserProperties<'a>, Box<dyn StdError>> {
    #[cfg(all(feature = "legacy-property-api", not(feature = "current-api")))]
    {
        Ok(UserProperties::with_limits(section, limits)?)
    }
    #[cfg(feature = "current-api")]
    {
        Ok(UserProperties::new(section, limits)?)
    }
}

fn user_editor<'a>(
    section: &'a mut Section,
    limits: UserLimits,
) -> Result<UserEdit<'a>, Box<dyn StdError>> {
    #[cfg(all(feature = "legacy-property-api", not(feature = "current-api")))]
    {
        Ok(UserEdit::with_limits(section, limits)?)
    }
    #[cfg(feature = "current-api")]
    {
        Ok(UserEdit::new(section, limits)?)
    }
}

fn section_size(section: &Section) -> usize {
    section.property_ids().count() + section.named_properties().count()
}
