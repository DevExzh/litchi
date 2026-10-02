//! Hand-built ZIP fixtures for the ODF leniency repairs (LENIENCY_AUDIT
//! F9.1 `mimetype` container layouts, and member rewrites of corpus files for
//! the later F9.x findings).
//!
//! The builder writes every byte of the local headers, central directory, and
//! end record itself, so each fixture carries exactly one controlled
//! deviation and nothing a third-party archiver might add. Shared by the
//! `litchi-odf-common` and family-facade test suites through `#[path]`.

#![allow(
    dead_code,
    reason = "each including test suite uses a different subset of the helpers"
)]

/// How the re-zipped `mimetype` member deviates from ODF 1.2 Part 3 §3.3.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct MimetypeLayout {
    /// Store the member with Deflate (method 8) instead of Store.
    pub(crate) deflated: bool,
    /// Give the member's local header an Info-ZIP Extended Timestamp field.
    pub(crate) local_extra: bool,
    /// Write the member last instead of first.
    pub(crate) last: bool,
}

impl MimetypeLayout {
    /// The conforming layout: first, stored, no extra field.
    pub(crate) const CONFORMING: Self = Self {
        deflated: false,
        local_extra: false,
        last: false,
    };
}

/// One member of a hand-built archive.
struct Member<'a> {
    name: &'a str,
    data: &'a [u8],
    deflated: bool,
    local_extra: &'a [u8],
}

/// Info-ZIP Extended Timestamp extra field (`UT`, modification time only), as
/// `zip -r` leaves on a re-zipped package.
const EXTENDED_TIMESTAMP: [u8; 9] = [0x55, 0x54, 0x05, 0x00, 0x01, 0x00, 0x00, 0x00, 0x60];

fn crc32(data: &[u8]) -> u32 {
    let mut crc = 0xffff_ffff_u32;
    for &byte in data {
        crc ^= u32::from(byte);
        for _ in 0..8 {
            let mask = (crc & 1).wrapping_neg();
            crc = (crc >> 1) ^ (0xedb8_8320 & mask);
        }
    }
    !crc
}

/// A valid Deflate stream made of one final stored block.
fn deflate_stored_block(data: &[u8]) -> Vec<u8> {
    let length = u16::try_from(data.len()).expect("fixture member fits one stored block");
    let mut encoded = vec![0x01];
    encoded.extend_from_slice(&length.to_le_bytes());
    encoded.extend_from_slice(&(!length).to_le_bytes());
    encoded.extend_from_slice(data);
    encoded
}

fn u16_of(value: usize) -> [u8; 2] {
    u16::try_from(value)
        .expect("fixture field fits u16")
        .to_le_bytes()
}

fn u32_of(value: usize) -> [u8; 4] {
    u32::try_from(value)
        .expect("fixture field fits u32")
        .to_le_bytes()
}

fn build(members: &[Member<'_>]) -> Vec<u8> {
    let mut output = Vec::new();
    let mut central = Vec::new();
    for member in members {
        let payload = if member.deflated {
            deflate_stored_block(member.data)
        } else {
            member.data.to_vec()
        };
        let method: u16 = if member.deflated { 8 } else { 0 };
        let crc = crc32(member.data).to_le_bytes();
        let offset = output.len();

        output.extend_from_slice(b"PK\x03\x04");
        output.extend_from_slice(&20_u16.to_le_bytes());
        output.extend_from_slice(&0_u16.to_le_bytes());
        output.extend_from_slice(&method.to_le_bytes());
        output.extend_from_slice(&0_u16.to_le_bytes());
        output.extend_from_slice(&0x0021_u16.to_le_bytes());
        output.extend_from_slice(&crc);
        output.extend_from_slice(&u32_of(payload.len()));
        output.extend_from_slice(&u32_of(member.data.len()));
        output.extend_from_slice(&u16_of(member.name.len()));
        output.extend_from_slice(&u16_of(member.local_extra.len()));
        output.extend_from_slice(member.name.as_bytes());
        output.extend_from_slice(member.local_extra);
        output.extend_from_slice(&payload);

        central.extend_from_slice(b"PK\x01\x02");
        central.extend_from_slice(&20_u16.to_le_bytes());
        central.extend_from_slice(&20_u16.to_le_bytes());
        central.extend_from_slice(&0_u16.to_le_bytes());
        central.extend_from_slice(&method.to_le_bytes());
        central.extend_from_slice(&0_u16.to_le_bytes());
        central.extend_from_slice(&0x0021_u16.to_le_bytes());
        central.extend_from_slice(&crc);
        central.extend_from_slice(&u32_of(payload.len()));
        central.extend_from_slice(&u32_of(member.data.len()));
        central.extend_from_slice(&u16_of(member.name.len()));
        central.extend_from_slice(&0_u16.to_le_bytes());
        central.extend_from_slice(&0_u16.to_le_bytes());
        central.extend_from_slice(&0_u16.to_le_bytes());
        central.extend_from_slice(&0_u16.to_le_bytes());
        central.extend_from_slice(&0_u32.to_le_bytes());
        central.extend_from_slice(&u32_of(offset));
        central.extend_from_slice(member.name.as_bytes());
    }
    let directory_offset = output.len();
    output.extend_from_slice(&central);
    output.extend_from_slice(b"PK\x05\x06");
    output.extend_from_slice(&0_u16.to_le_bytes());
    output.extend_from_slice(&0_u16.to_le_bytes());
    output.extend_from_slice(&u16_of(members.len()));
    output.extend_from_slice(&u16_of(members.len()));
    output.extend_from_slice(&u32_of(central.len()));
    output.extend_from_slice(&u32_of(directory_offset));
    output.extend_from_slice(&0_u16.to_le_bytes());
    output
}

/// Build a package from `mimetype` content and the remaining members, with
/// the `mimetype` member written in `layout`. Other members are stored.
pub(crate) fn package(
    mimetype: &[u8],
    members: &[(&str, &[u8])],
    layout: MimetypeLayout,
) -> Vec<u8> {
    let local_extra: &[u8] = if layout.local_extra {
        &EXTENDED_TIMESTAMP
    } else {
        &[]
    };
    let mimetype_member = Member {
        name: "mimetype",
        data: mimetype,
        deflated: layout.deflated,
        local_extra,
    };
    let mut ordered = Vec::with_capacity(members.len() + 1);
    if !layout.last {
        ordered.push(mimetype_member);
    }
    ordered.extend(members.iter().map(|&(name, data)| Member {
        name,
        data,
        deflated: false,
        local_extra: &[],
    }));
    if layout.last {
        ordered.push(Member {
            name: "mimetype",
            data: mimetype,
            deflated: layout.deflated,
            local_extra,
        });
    }
    build(&ordered)
}

/// A minimal package of the given ODF MIME type and office body element
/// (for example `office:text`).
pub(crate) fn minimal_package(mimetype: &str, body: &str, layout: MimetypeLayout) -> Vec<u8> {
    minimal_package_with_mimetype_content(mimetype, mimetype.as_bytes(), body, layout)
}

/// [`minimal_package`] whose `mimetype` member holds `content` while the
/// manifest still declares `manifest_mimetype`.
pub(crate) fn minimal_package_with_mimetype_content(
    manifest_mimetype: &str,
    content: &[u8],
    body: &str,
    layout: MimetypeLayout,
) -> Vec<u8> {
    let manifest = format!(
        r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="{manifest_mimetype}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#
    );
    let document = format!(
        r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2"><office:body><{body}/></office:body></office:document-content>"#
    );
    package(
        content,
        &[
            ("META-INF/manifest.xml", manifest.as_bytes()),
            ("content.xml", document.as_bytes()),
        ],
        layout,
    )
}

/// Re-zip a conforming package's members, writing its `mimetype` member in
/// `layout` and every other member stored, in their original order.
pub(crate) fn rezip(conforming: &[u8], layout: MimetypeLayout) -> Vec<u8> {
    rezip_with(conforming, layout, |_name, data| data)
}

/// [`rezip`], passing every non-`mimetype` member through `transform` (given
/// the member name and its decoded bytes) before it is written.
pub(crate) fn rezip_with(
    conforming: &[u8],
    layout: MimetypeLayout,
    mut transform: impl FnMut(&str, Vec<u8>) -> Vec<u8>,
) -> Vec<u8> {
    let source = litchi_odf_common::OwnedPackage::from_bytes(conforming.to_vec())
        .expect("conforming fixture must open");
    let mimetype = source.get_file("mimetype").expect("fixture has a mimetype");
    let names = source.files().expect("fixture lists its members");
    let contents = names
        .iter()
        .filter(|name| name.as_str() != "mimetype" && !name.ends_with('/'))
        .map(|name| {
            let data = source.get_file(name).expect("fixture member is readable");
            let data = transform(name, data);
            (name.clone(), data)
        })
        .collect::<Vec<_>>();
    let members = contents
        .iter()
        .map(|(name, data)| (name.as_str(), data.as_slice()))
        .collect::<Vec<_>>();
    package(&mimetype, &members, layout)
}

/// Rewrite one UTF-8 XML member of a conforming package with `edit`, keeping
/// every other member and the conforming `mimetype` layout.
pub(crate) fn edit_member(
    conforming: &[u8],
    member: &str,
    edit: impl Fn(&str) -> String,
) -> Vec<u8> {
    let mut edited = false;
    let bytes = rezip_with(conforming, MimetypeLayout::CONFORMING, |name, data| {
        if name != member {
            return data;
        }
        edited = true;
        let text = String::from_utf8(data).expect("edited member is UTF-8");
        let rewritten = edit(&text);
        assert_ne!(rewritten, text, "the edit must change {member}");
        rewritten.into_bytes()
    });
    assert!(edited, "fixture has no {member} member");
    bytes
}

/// Remove the manifest's `/` root `file-entry` element (LENIENCY_AUDIT F9.2).
pub(crate) fn without_manifest_root(manifest: &str) -> String {
    let (start, end) = manifest_root_span(manifest);
    format!("{}{}", &manifest[..start], &manifest[end..])
}

/// Blank the `media-type` of the manifest's `/` root entry (LENIENCY_AUDIT
/// F9.2).
pub(crate) fn with_empty_manifest_root_media_type(manifest: &str) -> String {
    let (start, end) = manifest_root_span(manifest);
    let element = &manifest[start..end];
    let attribute = element
        .find("media-type=\"")
        .expect("root entry declares a media type")
        + "media-type=\"".len();
    let close = element[attribute..]
        .find('"')
        .expect("media type is quoted")
        + attribute;
    format!(
        "{}{}{}{}",
        &manifest[..start],
        &element[..attribute],
        &element[close..],
        &manifest[end..]
    )
}

/// Replace the manifest root entry's media type with `media_type`.
pub(crate) fn with_manifest_root_media_type(manifest: &str, media_type: &str) -> String {
    let blank = with_empty_manifest_root_media_type(manifest);
    let (start, end) = manifest_root_span(&blank);
    let element = &blank[start..end];
    let attribute =
        element.find("media-type=\"\"").expect("blanked media type") + "media-type=\"".len();
    format!(
        "{}{}{media_type}{}{}",
        &blank[..start],
        &element[..attribute],
        &element[attribute..],
        &blank[end..]
    )
}

/// Byte span of the self-closing `file-entry` element whose full path is `/`.
fn manifest_root_span(manifest: &str) -> (usize, usize) {
    let mut search = 0;
    while let Some(found) = manifest[search..].find("<manifest:file-entry") {
        let start = search + found;
        let end = start + manifest[start..].find("/>").expect("file entry closes") + 2;
        if manifest[start..end].contains("full-path=\"/\"") {
            return (start, end);
        }
        search = end;
    }
    panic!("manifest has no root file entry");
}
