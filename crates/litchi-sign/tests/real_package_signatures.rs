use litchi_sign::xml::{self, Profile, Resolver, Transform};
use litchi_sign::{Coverage, Error, Limits, Policy, Status};
use quick_xml::{Reader, XmlVersion, events::Event};
use soapberry_zip::office::ArchiveReader;
use std::borrow::Cow;
use std::collections::BTreeMap;

const DOCX: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.docx");
const XLSX: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.xlsx");
const PPTX: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.pptx");

#[derive(Debug)]
struct FixtureResolver<'a> {
    archive: ArchiveReader<'a>,
    expected: usize,
    limits: &'a Limits,
}

impl<'a> FixtureResolver<'a> {
    fn new(archive: ArchiveReader<'a>, limits: &'a Limits) -> Self {
        let expected = archive
            .file_names()
            .filter(|member| eligible_member(member))
            .count();
        Self {
            archive,
            expected,
            limits,
        }
    }

    fn member_name(uri: &str) -> Result<&str, Error> {
        let (path, _) = uri
            .split_once("?ContentType=")
            .ok_or_else(|| Error::Container(format!("fixture URI lacks a content type: {uri}")))?;
        path.strip_prefix('/').ok_or_else(|| {
            Error::Container(format!(
                "fixture URI is not an absolute package path: {uri}"
            ))
        })
    }

    fn relationship_bytes(&self, member: &str, ids: &[String]) -> Result<Vec<u8>, Error> {
        let source = self.archive.read(member).map_err(|error| {
            Error::Container(format!("cannot read fixture relationship part: {error}"))
        })?;
        let relationships = parse_relationships(&source)?;
        let mut xml = String::from(
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
        );
        for id in ids {
            let relationship = relationships.get(id).ok_or_else(|| {
                Error::Container(format!("fixture relationship {id} does not exist"))
            })?;
            xml.push_str("<Relationship Id=\"");
            push_escaped(&mut xml, &relationship.id);
            xml.push_str("\" Target=\"");
            push_escaped(&mut xml, &relationship.target);
            xml.push_str("\" TargetMode=\"");
            push_escaped(&mut xml, &relationship.target_mode);
            xml.push_str("\" Type=\"");
            push_escaped(&mut xml, &relationship.relationship_type);
            xml.push_str("\"></Relationship>");
        }
        xml.push_str("</Relationships>");
        if xml.len() > self.limits.max_signature_bytes() {
            return Err(Error::Limit(
                "fixture relationship transform is too large".into(),
            ));
        }
        Ok(xml.into_bytes())
    }
}

impl Resolver for FixtureResolver<'_> {
    fn expected(&self) -> usize {
        self.expected
    }

    fn has(&self, uri: &str) -> bool {
        let Ok(member) = Self::member_name(uri) else {
            return false;
        };
        eligible_member(member) && self.archive.read(member).is_ok()
    }

    fn get<'a>(
        &'a self,
        uri: &str,
        transforms: &[Transform],
    ) -> Result<(Cow<'a, [u8]>, Coverage), Error> {
        let member = Self::member_name(uri)?;
        let bytes = if member.ends_with(".rels") {
            let (ids, canon) = match transforms {
                [Transform::Relationships(ids)] => (ids, None),
                [Transform::Relationships(ids), Transform::Canon(canon)] => (ids, Some(*canon)),
                _ => {
                    return Err(Error::Container(format!(
                        "fixture relationship part has an invalid transform chain: {uri}"
                    )));
                },
            };
            let data = self.relationship_bytes(member, ids)?;
            match canon {
                Some(canon) => xml::canonicalize(&data, canon, self.limits)?,
                None => data,
            }
        } else {
            let data = self.archive.read(member).map_err(|error| {
                Error::Container(format!("cannot read fixture package part: {error}"))
            })?;
            match transforms {
                [] => data,
                [Transform::Canon(canon)] => xml::canonicalize(&data, *canon, self.limits)?,
                _ => {
                    return Err(Error::Container(format!(
                        "fixture package part has an invalid transform chain: {uri}"
                    )));
                },
            }
        };
        Ok((Cow::Owned(bytes), Coverage::Complete))
    }
}

fn eligible_member(member: &str) -> bool {
    member != "[Content_Types].xml" && !member.starts_with("_xmlsignatures/")
}

#[derive(Debug)]
struct Relationship {
    id: String,
    target: String,
    target_mode: String,
    relationship_type: String,
}

fn parse_relationships(source: &[u8]) -> Result<BTreeMap<String, Relationship>, Error> {
    let mut reader = Reader::from_reader(source);
    let mut buffer = Vec::new();
    let mut output = BTreeMap::new();
    loop {
        match reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Container(format!("cannot parse fixture relationships: {error}"))
        })? {
            Event::Empty(element) | Event::Start(element)
                if element.local_name().as_ref() == b"Relationship" =>
            {
                let mut id = None;
                let mut target = None;
                let mut target_mode = None;
                let mut relationship_type = None;
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::Container(format!("invalid fixture relationship attribute: {error}"))
                    })?;
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| {
                            Error::Container(format!("invalid fixture relationship value: {error}"))
                        })?
                        .into_owned();
                    match attribute.key.as_ref() {
                        b"Id" => id = Some(value),
                        b"Target" => target = Some(value),
                        b"TargetMode" => target_mode = Some(value),
                        b"Type" => relationship_type = Some(value),
                        _ => {},
                    }
                }
                let id =
                    id.ok_or_else(|| Error::Container("fixture relationship lacks Id".into()))?;
                let relationship = Relationship {
                    id: id.clone(),
                    target: target.ok_or_else(|| {
                        Error::Container("fixture relationship lacks Target".into())
                    })?,
                    target_mode: target_mode.unwrap_or_else(|| "Internal".into()),
                    relationship_type: relationship_type.ok_or_else(|| {
                        Error::Container("fixture relationship lacks Type".into())
                    })?,
                };
                if output.insert(id, relationship).is_some() {
                    return Err(Error::Container(
                        "fixture relationship Id is duplicated".into(),
                    ));
                }
            },
            Event::Eof => return Ok(output),
            _ => {},
        }
        buffer.clear();
    }
}

fn push_escaped(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\t' => output.push_str("&#x9;"),
            '\n' => output.push_str("&#xA;"),
            '\r' => output.push_str("&#xD;"),
            other => output.push(other),
        }
    }
}

fn verify_fixture(bytes: &[u8]) -> litchi_sign::Report {
    let policy = Policy::compatible();
    let archive = ArchiveReader::new(bytes).expect("fixture ZIP is readable");
    let signature = archive
        .read("_xmlsignatures/sig1.xml")
        .expect("fixture signature part is readable");
    let resolver = FixtureResolver::new(archive, policy.limits());
    xml::verify(Profile::Package, &signature, &resolver, &policy)
        .expect("fixture signature verifies")
}

#[test]
fn real_microsoft_package_signatures_verify_through_public_xml_api() {
    for fixture in [DOCX, XLSX, PPTX] {
        let report = verify_fixture(fixture);
        assert_eq!(report.integrity(), Status::Valid);
        assert_eq!(report.signature(), Status::Valid);
        assert_eq!(report.coverage(), Coverage::Partial);
        assert!(report.uses_sha1());
        assert!(!report.certificates().is_empty());
    }
}

#[test]
fn real_microsoft_package_signatures_require_compatible_weak_algorithm_policy() {
    for fixture in [DOCX, XLSX, PPTX] {
        let archive = ArchiveReader::new(fixture).expect("fixture ZIP is readable");
        let signature = archive
            .read("_xmlsignatures/sig1.xml")
            .expect("fixture signature part is readable");
        let policy = Policy::strict();
        let resolver = FixtureResolver::new(archive, policy.limits());
        assert!(matches!(
            xml::verify(Profile::Package, &signature, &resolver, &policy),
            Err(Error::Sha1)
        ));
    }
}
