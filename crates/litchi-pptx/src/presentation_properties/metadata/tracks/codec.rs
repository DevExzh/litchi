//! PowerPoint 2017 WebVTT Track parts.
//!
//! External tracks are retained as targets and are never fetched.

use super::model::*;
use crate::{Error, Result};
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use std::collections::HashSet;

pub const CONTENT_TYPE: &str = "text/vtt";
pub const RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2017/04/relationships/track";

const SLIDE: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const LAYOUT: &str = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
const MASTER: &str = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
const MAX_BYTES: usize = 16 * 1024 * 1024;
const MAX_BLOCKS: usize = 100_000;
const MAX_LINES: usize = 250_000;
const MAX_LINE: usize = 1024 * 1024;
const MAX_SETTINGS: usize = 64;

impl File {
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        parse_vtt(bytes)
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        validate(self)?;
        let mut out = b"WEBVTT".to_vec();
        if let Some(text) = &self.header_text {
            out.push(b' ');
            out.extend_from_slice(text.as_bytes());
        }
        out.push(b'\n');
        for header in &self.headers {
            out.extend_from_slice(header.name.as_bytes());
            out.push(b':');
            out.extend_from_slice(header.value.as_bytes());
            out.push(b'\n');
        }
        out.push(b'\n');
        for (index, block) in self.blocks.iter().enumerate() {
            if index != 0 {
                out.push(b'\n');
            }
            write_block(&mut out, block);
        }
        if !self.blocks.is_empty() {
            out.push(b'\n');
        }
        if out.len() > MAX_BYTES {
            return Err(limit("serialized WebVTT bytes"));
        }
        parse_vtt(&out)?;
        Ok(out)
    }
}

pub fn load(package: &OpcPackage) -> Result<Vec<Track>> {
    if package
        .rels()
        .iter()
        .any(|rel| rel.reltype() == RELATIONSHIP_TYPE)
    {
        return Err(invalid(
            "Track relationship cannot originate at the package root",
        ));
    }
    let mut values = Vec::new();
    let mut targets = HashSet::new();
    for source in package.iter_parts() {
        for rel in source
            .rels()
            .iter()
            .filter(|rel| rel.reltype() == RELATIONSHIP_TYPE)
        {
            if !source_type(source.content_type()) {
                return Err(invalid(format!(
                    "Track relationship has invalid source '{}'",
                    source.partname()
                )));
            }
            let target = if rel.is_external() {
                if rel.target_ref().is_empty() {
                    return Err(invalid("external Track target cannot be empty"));
                }
                Target::External {
                    target: rel.target_ref().to_owned(),
                }
            } else {
                let name = rel.target_partname()?;
                let part = package.get_part(&name)?;
                if part.content_type() != CONTENT_TYPE {
                    return Err(Error::ContentType {
                        expected: CONTENT_TYPE.into(),
                        actual: part.content_type().into(),
                    });
                }
                if !part.rels().is_empty() {
                    return Err(invalid(format!(
                        "Track part '{}' cannot have outbound relationships",
                        part.partname()
                    )));
                }
                targets.insert(name.to_string());
                Target::Internal {
                    part_name: name.to_string(),
                    track: File::parse(part.blob())?,
                }
            };
            values.push(Track {
                source_part_name: source.partname().to_string(),
                relationship_id: rel.r_id().to_owned(),
                target,
            });
        }
    }
    for part in package.iter_parts() {
        if part.content_type() == CONTENT_TYPE && !targets.contains(part.partname().as_str()) {
            return Err(invalid(format!(
                "package contains orphan Track part '{}'",
                part.partname()
            )));
        }
    }
    Ok(values)
}

pub fn store(package: &mut OpcPackage, value: &Track) -> Result<()> {
    load(package)?;
    valid_rel_id(&value.relationship_id)?;
    let source_name = PackURI::new(&value.source_part_name).map_err(Error::Uri)?;
    let source = package.get_part(&source_name)?;
    if !source_type(source.content_type()) {
        return Err(invalid(
            "Track source is not a Slide, Slide Layout, or Slide Master",
        ));
    }
    if source.rels().get(&value.relationship_id).is_some() {
        return Err(invalid("Track relationship ID already exists"));
    }
    match &value.target {
        Target::External { target } => {
            if target.is_empty() {
                return Err(invalid("external Track target cannot be empty"));
            }
            package
                .get_part_mut(&source_name)?
                .rels_mut()
                .add_relationship(
                    RELATIONSHIP_TYPE.into(),
                    target.clone(),
                    value.relationship_id.clone(),
                    true,
                );
        },
        Target::Internal { part_name, track } => {
            let name = PackURI::new(part_name).map_err(Error::Uri)?;
            if package.iter_parts().any(|part| part.partname() == &name) {
                return Err(invalid(format!("part '{name}' already exists")));
            }
            let bytes = track.to_bytes()?;
            let target = name.relative_ref(source_name.base_uri());
            package.try_add_part(Box::new(BlobPart::new(name, CONTENT_TYPE.into(), bytes)))?;
            package
                .get_part_mut(&source_name)?
                .rels_mut()
                .add_relationship(
                    RELATIONSHIP_TYPE.into(),
                    target,
                    value.relationship_id.clone(),
                    false,
                );
        },
    }
    Ok(())
}

fn parse_vtt(bytes: &[u8]) -> Result<File> {
    if bytes.len() > MAX_BYTES {
        return Err(limit("WebVTT part bytes"));
    }
    let bytes = bytes.strip_prefix(&[0xef, 0xbb, 0xbf]).unwrap_or(bytes);
    let text = std::str::from_utf8(bytes)
        .map_err(|error| invalid(format!("Track is not UTF-8: {error}")))?;
    if text.contains('\0') {
        return Err(invalid("WebVTT cannot contain U+0000"));
    }
    let normalized = text.replace("\r\n", "\n").replace('\r', "\n");
    let lines: Vec<&str> = normalized.split('\n').collect();
    if lines.len() > MAX_LINES || lines.iter().any(|line| line.len() > MAX_LINE) {
        return Err(limit("WebVTT lines"));
    }
    let first = lines.first().copied().unwrap_or_default();
    let header_text = if first == "WEBVTT" {
        None
    } else if let Some(value) = first
        .strip_prefix("WEBVTT ")
        .or_else(|| first.strip_prefix("WEBVTT\t"))
    {
        if value.contains("-->") {
            return Err(invalid("WebVTT header contains '-->'"));
        }
        Some(value.to_owned())
    } else {
        return Err(invalid("Track is missing the WEBVTT signature"));
    };
    if lines.len() == 1 {
        return Ok(File {
            header_text,
            headers: Vec::new(),
            blocks: Vec::new(),
        });
    }
    let separator = lines[1..]
        .iter()
        .position(|line| line.is_empty())
        .map(|index| index + 1)
        .ok_or_else(|| invalid("WebVTT header needs a blank separator"))?;
    let mut headers = Vec::new();
    for line in &lines[1..separator] {
        let (name, value) = line
            .split_once(':')
            .ok_or_else(|| invalid("invalid WebVTT header metadata"))?;
        if name.is_empty() || name.chars().any(char::is_whitespace) || line.contains("-->") {
            return Err(invalid("invalid WebVTT header metadata"));
        }
        headers.push(Header {
            name: name.to_owned(),
            value: value.to_owned(),
        });
    }
    let mut blocks = Vec::new();
    let mut index = separator + 1;
    let mut cue_seen = false;
    while index < lines.len() {
        while index < lines.len() && lines[index].is_empty() {
            index += 1;
        }
        if index == lines.len() {
            break;
        }
        let start = index;
        while index < lines.len() && !lines[index].is_empty() {
            index += 1;
        }
        if blocks.len() >= MAX_BLOCKS {
            return Err(limit("WebVTT blocks"));
        }
        let block = parse_block(&lines[start..index], cue_seen)?;
        cue_seen |= matches!(block, Block::Cue(_));
        blocks.push(block);
    }
    let value = File {
        header_text,
        headers,
        blocks,
    };
    validate(&value)?;
    Ok(value)
}

fn parse_block(lines: &[&str], cue_seen: bool) -> Result<Block> {
    let first = lines[0];
    if first == "NOTE" || first.starts_with("NOTE ") || first.starts_with("NOTE\t") {
        no_arrow(lines, "NOTE")?;
        return Ok(Block::Note {
            header: first[4..].trim_start_matches([' ', '\t']).to_owned(),
            lines: owned(&lines[1..]),
        });
    }
    if first == "STYLE" {
        if cue_seen || lines.len() == 1 {
            return Err(invalid("invalid WebVTT STYLE block position or body"));
        }
        no_arrow(lines, "STYLE")?;
        return Ok(Block::Style {
            lines: owned(&lines[1..]),
        });
    }
    if first == "REGION" {
        if cue_seen {
            return Err(invalid("WebVTT REGION block follows a cue"));
        }
        return Ok(Block::Region {
            settings: parse_region(&lines[1..])?,
        });
    }
    let (identifier, timing_index) = if first.contains("-->") {
        (None, 0)
    } else {
        (Some(first.to_owned()), 1)
    };
    let timing = lines
        .get(timing_index)
        .ok_or_else(|| invalid("cue is missing timings"))?;
    let (start, rest) = timing
        .split_once("-->")
        .ok_or_else(|| invalid("cue is missing '-->'"))?;
    if rest.contains("-->") {
        return Err(invalid("cue has multiple timing arrows"));
    }
    let mut tokens = rest.split_ascii_whitespace();
    let start = timestamp(start.trim())?;
    let end = timestamp(
        tokens
            .next()
            .ok_or_else(|| invalid("cue is missing end timestamp"))?,
    )?;
    if end <= start {
        return Err(invalid("cue end must be after start"));
    }
    let settings = parse_cue_settings(tokens)?;
    let payload = owned(&lines[timing_index + 1..]);
    if payload.iter().any(|line| line.contains("-->")) {
        return Err(invalid("cue payload contains '-->'"));
    }
    Ok(Block::Cue(Cue {
        identifier,
        start_milliseconds: start,
        end_milliseconds: end,
        settings,
        payload,
    }))
}

fn parse_cue_settings<'a>(tokens: impl Iterator<Item = &'a str>) -> Result<Vec<CueSetting>> {
    let mut result = Vec::new();
    let mut seen = HashSet::new();
    for token in tokens {
        if result.len() >= MAX_SETTINGS {
            return Err(limit("WebVTT cue settings"));
        }
        let (name, value) = token
            .split_once(':')
            .ok_or_else(|| invalid("invalid cue setting"))?;
        let kind = CueSettingKind::from(name)
            .ok_or_else(|| invalid(format!("unknown cue setting '{name}'")))?;
        if !seen.insert(kind) {
            return Err(invalid("duplicate cue setting"));
        }
        valid_cue_setting(kind, value)?;
        result.push(CueSetting {
            kind,
            value: value.to_owned(),
        });
    }
    Ok(result)
}

fn parse_region(lines: &[&str]) -> Result<Vec<RegionSetting>> {
    let mut result = Vec::new();
    let mut seen = HashSet::new();
    for token in lines.iter().flat_map(|line| line.split_ascii_whitespace()) {
        if result.len() >= MAX_SETTINGS {
            return Err(limit("WebVTT region settings"));
        }
        let (name, value) = token
            .split_once(':')
            .ok_or_else(|| invalid("invalid region setting"))?;
        let kind = RegionSettingKind::from(name)
            .ok_or_else(|| invalid(format!("unknown region setting '{name}'")))?;
        if !seen.insert(kind) {
            return Err(invalid("duplicate region setting"));
        }
        valid_region_setting(kind, value)?;
        result.push(RegionSetting {
            kind,
            value: value.to_owned(),
        });
    }
    Ok(result)
}

fn valid_cue_setting(kind: CueSettingKind, value: &str) -> Result<()> {
    let valid = match kind {
        CueSettingKind::Vertical => matches!(value, "rl" | "lr"),
        CueSettingKind::Align => {
            matches!(value, "start" | "center" | "end" | "left" | "right")
        },
        CueSettingKind::Size => percentage(value),
        CueSettingKind::Region => identifier(value),
        CueSettingKind::Line => {
            let (position, align) = comma(value);
            (position == "auto" || position.parse::<i64>().is_ok() || percentage(position))
                && align.is_none_or(|value| matches!(value, "start" | "center" | "end"))
        },
        CueSettingKind::Position => {
            let (position, align) = comma(value);
            percentage(position)
                && align.is_none_or(|value| {
                    matches!(value, "line-left" | "center" | "line-right" | "auto")
                })
        },
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(format!("invalid {} cue setting", kind.token())))
    }
}

fn valid_region_setting(kind: RegionSettingKind, value: &str) -> Result<()> {
    let valid = match kind {
        RegionSettingKind::Id => identifier(value),
        RegionSettingKind::Width => percentage(value),
        RegionSettingKind::Lines => value.parse::<u64>().is_ok_and(|value| value > 0),
        RegionSettingKind::RegionAnchor | RegionSettingKind::ViewportAnchor => value
            .split_once(',')
            .is_some_and(|(x, y)| percentage(x) && percentage(y)),
        RegionSettingKind::Scroll => value == "up",
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(format!("invalid {} region setting", kind.token())))
    }
}

fn validate(value: &File) -> Result<()> {
    if let Some(text) = &value.header_text {
        line(text)?;
        if text.contains("-->") {
            return Err(invalid("header contains '-->'"));
        }
    }
    for header in &value.headers {
        line(&header.name)?;
        line(&header.value)?;
        if header.name.is_empty()
            || header.name.contains(':')
            || header.name.chars().any(char::is_whitespace)
            || header.name.contains("-->")
            || header.value.contains("-->")
        {
            return Err(invalid("invalid WebVTT header metadata"));
        }
    }
    if value.blocks.len() > MAX_BLOCKS {
        return Err(limit("WebVTT blocks"));
    }
    let mut cue_seen = false;
    for block in &value.blocks {
        match block {
            Block::Cue(cue) => {
                cue_seen = true;
                if cue.end_milliseconds <= cue.start_milliseconds {
                    return Err(invalid("cue end must be after start"));
                }
                if let Some(id) = &cue.identifier {
                    line(id)?;
                    if id.contains("-->") {
                        return Err(invalid("cue id contains '-->'"));
                    }
                }
                let mut seen = HashSet::new();
                for setting in &cue.settings {
                    if !seen.insert(setting.kind) {
                        return Err(invalid("duplicate cue setting"));
                    }
                    valid_cue_setting(setting.kind, &setting.value)?;
                }
                raw_lines(&cue.payload)?;
            },
            Block::Note { header, lines } => {
                line(header)?;
                raw_lines(lines)?;
            },
            Block::Style { lines } => {
                if cue_seen || lines.is_empty() {
                    return Err(invalid("invalid STYLE block"));
                }
                raw_lines(lines)?;
            },
            Block::Region { settings } => {
                if cue_seen {
                    return Err(invalid("REGION follows cue"));
                }
                let mut seen = HashSet::new();
                for setting in settings {
                    if !seen.insert(setting.kind) {
                        return Err(invalid("duplicate region setting"));
                    }
                    valid_region_setting(setting.kind, &setting.value)?;
                }
            },
        }
    }
    Ok(())
}

fn write_block(out: &mut Vec<u8>, block: &Block) {
    match block {
        Block::Cue(cue) => {
            if let Some(id) = &cue.identifier {
                put_line(out, id);
            }
            write_time(out, cue.start_milliseconds);
            out.extend_from_slice(b" --> ");
            write_time(out, cue.end_milliseconds);
            for setting in &cue.settings {
                out.push(b' ');
                out.extend_from_slice(setting.kind.token().as_bytes());
                out.push(b':');
                out.extend_from_slice(setting.value.as_bytes());
            }
            out.push(b'\n');
            for line in &cue.payload {
                put_line(out, line);
            }
        },
        Block::Note { header, lines } => {
            out.extend_from_slice(b"NOTE");
            if !header.is_empty() {
                out.push(b' ');
                out.extend_from_slice(header.as_bytes());
            }
            out.push(b'\n');
            for line in lines {
                put_line(out, line);
            }
        },
        Block::Style { lines } => {
            out.extend_from_slice(b"STYLE\n");
            for line in lines {
                put_line(out, line);
            }
        },
        Block::Region { settings } => {
            out.extend_from_slice(b"REGION\n");
            for setting in settings {
                out.extend_from_slice(setting.kind.token().as_bytes());
                out.push(b':');
                put_line(out, &setting.value);
            }
        },
    }
}

fn timestamp(value: &str) -> Result<u64> {
    let parts: Vec<_> = value.split(':').collect();
    let (hours, minutes, seconds) = match parts.as_slice() {
        [m, s] if m.len() == 2 => (0u64, *m, *s),
        [h, m, s] if h.len() >= 2 && m.len() == 2 => {
            (h.parse().map_err(|_| invalid("invalid hours"))?, *m, *s)
        },
        _ => return Err(invalid("invalid WebVTT timestamp")),
    };
    let minutes: u64 = minutes.parse().map_err(|_| invalid("invalid minutes"))?;
    let (seconds, millis) = seconds
        .split_once('.')
        .ok_or_else(|| invalid("invalid timestamp fraction"))?;
    if seconds.len() != 2 || millis.len() != 3 {
        return Err(invalid("invalid timestamp precision"));
    }
    let seconds: u64 = seconds.parse().map_err(|_| invalid("invalid seconds"))?;
    let millis: u64 = millis
        .parse()
        .map_err(|_| invalid("invalid milliseconds"))?;
    if minutes > 59 || seconds > 59 {
        return Err(invalid("timestamp component out of range"));
    }
    hours
        .checked_mul(3_600_000)
        .and_then(|v| v.checked_add(minutes * 60_000 + seconds * 1_000 + millis))
        .ok_or_else(|| limit("WebVTT timestamp"))
}

fn write_time(out: &mut Vec<u8>, value: u64) {
    out.extend_from_slice(
        format!(
            "{:02}:{:02}:{:02}.{:03}",
            value / 3_600_000,
            value / 60_000 % 60,
            value / 1_000 % 60,
            value % 1_000
        )
        .as_bytes(),
    );
}
fn percentage(value: &str) -> bool {
    value.strip_suffix('%').is_some_and(|v| {
        !v.is_empty()
            && !v.contains(['e', 'E'])
            && v.parse::<f64>()
                .is_ok_and(|n| n.is_finite() && (0.0..=100.0).contains(&n))
    })
}
fn identifier(value: &str) -> bool {
    !value.is_empty() && !value.contains("-->") && !value.chars().any(char::is_whitespace)
}
fn comma(value: &str) -> (&str, Option<&str>) {
    value
        .split_once(',')
        .map_or((value, None), |(a, b)| (a, Some(b)))
}
fn owned(lines: &[&str]) -> Vec<String> {
    lines.iter().map(|line| (*line).to_owned()).collect()
}
fn no_arrow(lines: &[&str], label: &str) -> Result<()> {
    if lines.iter().any(|line| line.contains("-->")) {
        Err(invalid(format!("{label} contains '-->'")))
    } else {
        Ok(())
    }
}
fn line(value: &str) -> Result<()> {
    if value.len() > MAX_LINE {
        Err(limit("WebVTT line"))
    } else if value.contains(['\r', '\n', '\0']) {
        Err(invalid("WebVTT line contains forbidden character"))
    } else {
        Ok(())
    }
}
fn raw_lines(lines: &[String]) -> Result<()> {
    for value in lines {
        line(value)?;
        if value.contains("-->") {
            return Err(invalid("WebVTT block contains '-->'"));
        }
    }
    Ok(())
}
fn put_line(out: &mut Vec<u8>, value: &str) {
    out.extend_from_slice(value.as_bytes());
    out.push(b'\n');
}
fn source_type(value: &str) -> bool {
    matches!(value, SLIDE | LAYOUT | MASTER)
}
fn valid_rel_id(value: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_LINE || value.chars().any(char::is_whitespace) {
        Err(invalid("invalid Track relationship ID"))
    } else {
        Ok(())
    }
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn limit(label: &str) -> Error {
    invalid(format!("{label} exceeds implementation limit"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::Part;

    const W3C: &[u8] = b"WEBVTT\n\n00:11.000 --> 00:13.000\n<v Roger Bingham>We are in New York City\n\n00:13.000 --> 00:16.000\n<v Roger Bingham>We're actually at the Lucern Hotel\n";

    fn package(content_type: &str) -> OpcPackage {
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "ppt/presentation.xml".into(),
            "rId1".into(),
            false,
        );
        let mut presentation = BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                .into(),
            Vec::new(),
        );
        presentation.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide".into(),
            "slides/slide1.xml".into(),
            "rId1".into(),
            false,
        );
        package.add_part(Box::new(presentation));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/slides/slide1.xml").unwrap(),
            content_type.into(),
            Vec::new(),
        )));
        package
    }
    fn internal() -> Track {
        Track {
            source_part_name: "/ppt/slides/slide1.xml".into(),
            relationship_id: "rId9".into(),
            target: Target::Internal {
                part_name: "/ppt/media/captions1.vtt".into(),
                track: File::parse(W3C).unwrap(),
            },
        }
    }

    #[test]
    fn parses_w3c_specimen_and_round_trips() {
        let value = File::parse(W3C).unwrap();
        let Block::Cue(cue) = &value.blocks[0] else {
            panic!("cue")
        };
        assert_eq!(
            (cue.start_milliseconds, cue.end_milliseconds),
            (11_000, 13_000)
        );
        assert_eq!(File::parse(&value.to_bytes().unwrap()).unwrap(), value);
    }

    #[test]
    fn typed_blocks_and_internal_package_round_trip() {
        let bytes = b"WEBVTT captions\nKind:captions\n\nSTYLE\n::cue { color: lime; }\n\nREGION\nid:fred\nwidth:40%\nlines:3\nregionanchor:0%,100%\nviewportanchor:10%,90%\nscroll:up\n\nNOTE generated\nrelationship rId777 is inert\n\nintro\n00:00:01.000 --> 00:00:02.500 vertical:rl line:20%,center position:30%,line-left size:40% align:start region:fred\nHello\n";
        let track = File::parse(bytes).unwrap();
        let mut package = package(SLIDE);
        let mut value = internal();
        let Target::Internal { track: target, .. } = &mut value.target else {
            unreachable!()
        };
        *target = track.clone();
        store(&mut package, &value).unwrap();
        let Target::Internal { track: loaded, .. } = &load(&package).unwrap()[0].target else {
            panic!("internal")
        };
        assert_eq!(loaded, &track);
    }

    #[test]
    fn external_leaf_orphan_and_atomic_graph_rules() {
        let mut external_package = package(LAYOUT);
        let external = Track {
            source_part_name: "/ppt/slides/slide1.xml".into(),
            relationship_id: "rId8".into(),
            target: Target::External {
                target: "https://example.invalid/captions.vtt".into(),
            },
        };
        store(&mut external_package, &external).unwrap();
        assert_eq!(load(&external_package).unwrap(), vec![external]);
        let mut orphan = package(SLIDE);
        orphan.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/media/orphan.vtt").unwrap(),
            CONTENT_TYPE.into(),
            W3C.to_vec(),
        )));
        assert!(load(&orphan).is_err());
        let mut outbound = package(SLIDE);
        store(&mut outbound, &internal()).unwrap();
        outbound
            .get_part_mut(&PackURI::new("/ppt/media/captions1.vtt").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship("urn:no".into(), "x".into(), "rId1".into(), false);
        assert!(load(&outbound).is_err());
        let mut atomic = package(SLIDE);
        let mut bad = internal();
        let Target::Internal { track, .. } = &mut bad.target else {
            unreachable!()
        };
        let Block::Cue(cue) = &mut track.blocks[0] else {
            unreachable!()
        };
        cue.end_milliseconds = cue.start_milliseconds;
        let count = atomic.iter_parts().count();
        assert!(store(&mut atomic, &bad).is_err());
        assert_eq!(atomic.iter_parts().count(), count);
    }

    #[test]
    fn rejects_hostile_or_invalid_webvtt() {
        for bytes in [
            &b"not-vtt\n"[..],
            &b"WEBVTT\n00:00.000 --> 00:01.000\nx\n"[..],
            &b"WEBVTT\n\n00:61.000 --> 01:00.000\nx\n"[..],
            &b"WEBVTT\n\n00:02.000 --> 00:01.000\nx\n"[..],
            &b"WEBVTT\n\n00:00.000 --> 00:01.000 align:middle\nx\n"[..],
            &b"WEBVTT\n\n00:00.000 --> 00:01.000 align:start align:end\nx\n"[..],
            &b"WEBVTT\n\n00:00.000 --> 00:01.000\nx --> y\n"[..],
            &b"WEBVTT\0\n"[..],
        ] {
            assert!(
                File::parse(bytes).is_err(),
                "accepted {:?}",
                String::from_utf8_lossy(bytes)
            );
        }
        assert!(File::parse(&vec![b' '; MAX_BYTES + 1]).is_err());
    }
}
