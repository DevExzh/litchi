//! Record 0770: `checked_attributes` against quick-xml's checked iterator.
//!
//! `attribute_checks_equivalence --json PATH ROOT...` reads every XML member
//! of every OOXML package under the roots and, for every start tag, compares
//! what `BytesStartExt::checked_attributes` yields (from `litchi-opc` and from
//! the `litchi-ole-common` copy) with what quick-xml's own checked iterator
//! yields up to and including its first error. It does the same for mutated
//! copies of each tag: a repeated name at several positions, before and after
//! the 32nd name, malformed values and names. This binary needs the changed
//! library; it is not built for the base.
//!
//! `attribute_checks_equivalence --bench quick-xml|checked --rounds R ROOT...`
//! collects every start tag of the XML members under the roots once, then
//! iterates every tag's attributes `R` times with quick-xml's checked
//! iterator or with `checked_attributes`, stopping at the first error as a
//! fail-fast reader does. Two runs with different `R` under `perf stat` give
//! the instructions per pass of each iterator over the same tags.

#![forbid(unsafe_code)]

use std::collections::BTreeMap;
use std::error::Error;
use std::fmt::Write as _;
use std::path::{Path, PathBuf};

use quick_xml::events::attributes::{AttrError, Attribute};
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;
use soapberry_zip::office::ArchiveReader;

type Item = Result<(Vec<u8>, Vec<u8>), AttrError>;

fn owned(item: Result<Attribute<'_>, AttrError>) -> Item {
    item.map(|attribute| (attribute.key.as_ref().to_vec(), attribute.value.to_vec()))
}

fn quick_xml_until_error(tag: &BytesStart<'_>) -> Vec<Item> {
    let mut items = Vec::new();
    for item in tag.attributes() {
        let error = item.is_err();
        items.push(owned(item));
        if error {
            break;
        }
    }
    items
}

fn opc_checked(tag: &BytesStart<'_>) -> Vec<Item> {
    use litchi_opc::xml_attributes::BytesStartExt as _;
    tag.checked_attributes().map(owned).collect()
}

fn ole_checked(tag: &BytesStart<'_>) -> Vec<Item> {
    use litchi_ole_common::xml_attributes::BytesStartExt as _;
    tag.checked_attributes().map(owned).collect()
}

#[derive(Default)]
struct Tally {
    tags: usize,
    tags_over_32: usize,
    tags_with_error: usize,
    mutated: usize,
    mutated_over_32: usize,
    outcomes: BTreeMap<String, usize>,
    mismatches: Vec<String>,
}

impl Tally {
    fn compare(&mut self, content: &str, name_len: usize, mutated: bool) {
        let tag = BytesStart::from_content(content, name_len);
        let expected = quick_xml_until_error(&tag);
        let over_32 = expected.iter().filter(|item| item.is_ok()).count() > 32;
        if mutated {
            self.mutated += 1;
            self.mutated_over_32 += usize::from(over_32);
        } else {
            self.tags += 1;
            self.tags_over_32 += usize::from(over_32);
            self.tags_with_error += usize::from(expected.last().is_some_and(Result::is_err));
        }
        let label = match expected.last() {
            Some(Err(AttrError::Duplicated(..))) => "duplicate",
            Some(Err(AttrError::UnquotedValue(_))) => "unquoted value",
            Some(Err(AttrError::ExpectedValue(_))) => "no value",
            Some(Err(AttrError::ExpectedQuote(..))) => "no closing quote",
            Some(Err(AttrError::ExpectedEq(_))) => "no equals sign",
            _ => "accepted",
        };
        let phase = if over_32 { "after 32" } else { "within 32" };
        *self
            .outcomes
            .entry(format!(
                "{} / {label} / {phase}",
                if mutated { "mutated" } else { "real" }
            ))
            .or_default() += 1;
        for (implementation, actual) in [("opc", opc_checked(&tag)), ("ole", ole_checked(&tag))] {
            if actual != expected && self.mismatches.len() < 100 {
                let mut shown = content.to_owned();
                shown.truncate(400);
                self.mismatches.push(format!(
                    "{implementation}: {shown} -> {actual:?} != {expected:?}"
                ));
            }
        }
    }
}

/// Mutated copies of one start tag's content: text appended after its last
/// attribute, or inserted after its first.
fn mutations(content: &str, names: &[String], first_end: Option<usize>) -> Vec<String> {
    let mut variants = Vec::new();
    let first = names.first().cloned().unwrap_or_else(|| "a".to_owned());
    let last = names.last().cloned().unwrap_or_else(|| "a".to_owned());
    let mut padding = String::new();
    for index in 0..40 {
        let _ = write!(padding, " zz{index}=\"{index}\"");
    }
    for appended in [
        format!(" {first}=\"r\""),
        format!(" {last}=\"r\""),
        format!(" {first}=r"),
        format!(" {first}="),
        format!(" {first}=\"open"),
        format!(" {first} = 'r'"),
        format!(" {first}=\"x {last}='y'\" t=\"t\""),
        " zzflag".to_owned(),
        " zzq=v".to_owned(),
        " zzv=".to_owned(),
        " zzo=\"open".to_owned(),
        format!("{padding} zzv="),
        format!("{padding} zzo='open"),
        format!("{padding} {first}=\"r\""),
        format!("{padding} {last}=r"),
        format!("{padding} zz7=\"r\""),
        format!("{padding} zz39="),
        format!("{padding} zz0=\"open"),
        format!("{padding} zzq=v"),
        format!("{padding} zzflag"),
        padding.clone(),
    ] {
        variants.push(format!("{content}{appended}"));
    }
    if let Some(first_end) = first_end {
        for inserted in [format!(" {last}=\"r\""), format!(" {first}=\"r\"")] {
            let mut variant = content.to_owned();
            variant.insert_str(first_end, &inserted);
            variants.push(variant);
        }
    }
    variants
}

fn main() -> Result<(), Box<dyn Error>> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if let Some(index) = args.iter().position(|arg| arg == "--bench") {
        return bench(
            &args,
            args.get(index + 1).ok_or("--bench needs an iterator")?,
        );
    }
    let json = args
        .iter()
        .position(|arg| arg == "--json")
        .and_then(|index| args.get(index + 1))
        .ok_or("--json is required")?
        .clone();
    let mut roots = Vec::new();
    let mut skip = false;
    for arg in &args {
        if skip {
            skip = false;
            continue;
        }
        if arg.starts_with("--") {
            skip = true;
            continue;
        }
        roots.push(PathBuf::from(arg));
    }
    let mut packages = Vec::new();
    for root in &roots {
        collect_packages(root, &mut packages)?;
    }
    packages.sort();
    let mut tally = Tally::default();
    let mut members = 0usize;
    for path in &packages {
        let Ok(bytes) = std::fs::read(path) else {
            continue;
        };
        let Ok(archive) = ArchiveReader::new(&bytes) else {
            continue;
        };
        let names: Vec<String> = archive.file_names().map(str::to_owned).collect();
        for name in names {
            let lower = name.to_ascii_lowercase();
            if !(lower.ends_with(".xml") || lower.ends_with(".rels") || lower.ends_with(".vml")) {
                continue;
            }
            let Ok(member) = archive.read(&name) else {
                continue;
            };
            members += 1;
            let mut reader = Reader::from_reader(member.as_slice());
            let mut mutated_here = 0usize;
            loop {
                let event = match reader.read_event() {
                    Ok(Event::Eof) | Err(_) => break,
                    Ok(event) => event,
                };
                let (Event::Start(tag) | Event::Empty(tag)) = event else {
                    continue;
                };
                let name_len = tag.name().as_ref().len();
                let Ok(content) = std::str::from_utf8(&tag) else {
                    continue;
                };
                tally.compare(content, name_len, false);
                // Mutate the first 16 tags that carry attributes in each member.
                if mutated_here >= 16 {
                    continue;
                }
                let mut unchecked = tag.attributes();
                unchecked.with_checks(false);
                let mut attribute_names = Vec::new();
                let mut first_end = None;
                for attribute in unchecked.map_while(Result::ok) {
                    if first_end.is_none() {
                        let value = attribute.value.as_ref();
                        let end = (value.as_ptr() as usize) - (content.as_ptr() as usize)
                            + value.len()
                            + 1;
                        first_end = Some(end);
                    }
                    attribute_names
                        .push(String::from_utf8_lossy(attribute.key.as_ref()).into_owned());
                }
                if attribute_names.is_empty() {
                    continue;
                }
                mutated_here += 1;
                for variant in mutations(content, &attribute_names, first_end) {
                    tally.compare(&variant, name_len, true);
                }
            }
        }
    }
    let summary = serde_json::json!({
        "packages": packages.len(),
        "members": members,
        "tags": tally.tags,
        "tags_over_32_names": tally.tags_over_32,
        "tags_with_an_error": tally.tags_with_error,
        "mutated_tags": tally.mutated,
        "mutated_tags_over_32_names": tally.mutated_over_32,
        "outcomes": tally.outcomes,
        "mismatches": tally.mismatches.len(),
        "mismatch_examples": tally.mismatches,
    });
    std::fs::write(&json, serde_json::to_vec_pretty(&summary)?)?;
    if tally.mismatches.is_empty() {
        Ok(())
    } else {
        Err(format!("{} mismatches", tally.mismatches.len()).into())
    }
}

/// Every start tag of the XML members under the roots: its content and the
/// length of its name.
fn collect_tags(roots: &[PathBuf]) -> Result<Vec<(String, usize)>, Box<dyn Error>> {
    let mut packages = Vec::new();
    for root in roots {
        collect_packages(root, &mut packages)?;
    }
    packages.sort();
    let mut tags = Vec::new();
    for path in &packages {
        let Ok(bytes) = std::fs::read(path) else {
            continue;
        };
        let Ok(archive) = ArchiveReader::new(&bytes) else {
            continue;
        };
        let names: Vec<String> = archive.file_names().map(str::to_owned).collect();
        for name in names {
            let lower = name.to_ascii_lowercase();
            if !(lower.ends_with(".xml") || lower.ends_with(".rels") || lower.ends_with(".vml")) {
                continue;
            }
            let Ok(member) = archive.read(&name) else {
                continue;
            };
            let mut reader = Reader::from_reader(member.as_slice());
            loop {
                let event = match reader.read_event() {
                    Ok(Event::Eof) | Err(_) => break,
                    Ok(event) => event,
                };
                if let Event::Start(tag) | Event::Empty(tag) = event
                    && let Ok(content) = std::str::from_utf8(&tag)
                {
                    tags.push((content.to_owned(), tag.name().as_ref().len()));
                }
            }
        }
    }
    Ok(tags)
}

fn bench(args: &[String], iterator: &str) -> Result<(), Box<dyn Error>> {
    use litchi_opc::xml_attributes::BytesStartExt as _;
    let rounds: usize = args
        .iter()
        .position(|arg| arg == "--rounds")
        .and_then(|index| args.get(index + 1))
        .ok_or("--rounds is required")?
        .parse()?;
    let mut roots = Vec::new();
    let mut skip = false;
    for arg in args {
        if skip {
            skip = false;
            continue;
        }
        if arg.starts_with("--") {
            skip = true;
            continue;
        }
        roots.push(PathBuf::from(arg));
    }
    let tags: Vec<BytesStart<'_>> = collect_tags(&roots)?
        .into_iter()
        .map(|(content, name_len)| BytesStart::from_content(content, name_len))
        .collect();
    let mut checksum = 0usize;
    let started = std::time::Instant::now();
    for _ in 0..rounds {
        for tag in &tags {
            match iterator {
                "quick-xml" => {
                    for attribute in tag.attributes() {
                        let Ok(attribute) = attribute else { break };
                        checksum = checksum.wrapping_add(attribute.key.as_ref().len());
                    }
                },
                "checked" => {
                    for attribute in tag.checked_attributes() {
                        let Ok(attribute) = attribute else { break };
                        checksum = checksum.wrapping_add(attribute.key.as_ref().len());
                    }
                },
                _ => return Err(format!("unknown iterator {iterator}").into()),
            }
        }
    }
    let elapsed = started.elapsed();
    println!(
        "{}",
        serde_json::json!({
            "iterator": iterator,
            "tags": tags.len(),
            "rounds": rounds,
            "elapsed_ns": elapsed.as_nanos(),
            "checksum": std::hint::black_box(checksum),
        })
    );
    Ok(())
}

fn collect_packages(root: &Path, packages: &mut Vec<PathBuf>) -> Result<(), Box<dyn Error>> {
    let metadata = std::fs::symlink_metadata(root)?;
    if metadata.is_dir() {
        let mut entries: Vec<_> = std::fs::read_dir(root)?.filter_map(Result::ok).collect();
        entries.sort_by_key(std::fs::DirEntry::path);
        for entry in entries {
            collect_packages(&entry.path(), packages)?;
        }
    } else if metadata.is_file() {
        let Ok(bytes) = std::fs::read(root) else {
            return Ok(());
        };
        if bytes.starts_with(b"PK\x03\x04")
            && ArchiveReader::new(&bytes)
                .is_ok_and(|archive| archive.contains("[Content_Types].xml"))
        {
            packages.push(root.to_owned());
        }
    }
    Ok(())
}
