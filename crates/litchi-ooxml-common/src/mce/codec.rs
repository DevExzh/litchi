//! Bounded ISO/IEC 29500-3 MCE preprocessing and active-offset selection.

use quick_xml::{
    Reader, XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
};
use std::{
    borrow::Cow,
    collections::{BTreeMap, HashSet},
    str,
};

use super::model::{
    Capabilities, Error, Limits, NAMESPACE, Name, OffsetLimits, Output, Report, XML_NS,
};
use crate::xml_name::{self, QualifiedName};

type R<T> = Result<T, Error>;

const ACTIVE_MARKER_TEMPLATE: &[u8; 38] = b"litchi-mce-active-0000000000000000-00:";
const ACTIVE_MARKER_HASH_START: usize = 18;
const ACTIVE_MARKER_HASH_END: usize = 34;
const ACTIVE_MARKER_SALT_START: usize = 35;
const ACTIVE_MARKER_SALT_END: usize = 37;
const ACTIVE_MARKER_WRAPPER_BYTES: usize = 7;
const DECIMAL_BUFFER_BYTES: usize = usize::BITS as usize;

/// Retain source byte offsets that survive semantic MCE branch selection.
///
/// Returned offsets always refer to `xml`; the marked preprocessing buffer is
/// an implementation detail. Input order and duplicate offsets are preserved.
/// Every offset must be less than `xml.len()` and identify a position where an
/// XML comment can be inserted (normally an element-start offset).
pub fn active_offsets(
    xml: &[u8],
    offsets: &[u32],
    capabilities: &Capabilities,
    limits: &OffsetLimits,
) -> R<Vec<u32>> {
    if xml.len() > limits.max_source_bytes {
        return Err(limit("active-offset source bytes"));
    }
    if offsets.len() > limits.max_offsets {
        return Err(limit("active-offset count"));
    }
    for &offset in offsets {
        if usize::try_from(offset).map_or(true, |offset| offset >= xml.len()) {
            return Err(bad("active offset is outside the source XML"));
        }
    }

    if offsets.is_empty() {
        return Ok(Vec::new());
    }
    if find_bytes(xml, NAMESPACE.as_bytes()).is_none() {
        return copy_offsets(offsets);
    }

    let marker = active_marker(xml)?;
    let max_index = offsets
        .len()
        .checked_sub(1)
        .ok_or_else(|| bad("active-offset count is invalid"))?;
    let decimal_digits = decimal_len(max_index)?;
    let marker_bytes = marker
        .len()
        .checked_add(ACTIVE_MARKER_WRAPPER_BYTES)
        .and_then(|bytes| bytes.checked_add(decimal_digits))
        .ok_or_else(|| limit("active-offset marker bytes"))?;
    let marked_len = offsets
        .len()
        .checked_mul(marker_bytes)
        .and_then(|extra| xml.len().checked_add(extra))
        .ok_or_else(|| limit("active-offset marked XML bytes"))?;
    if marked_len > limits.max_marked_bytes {
        return Err(limit("active-offset marked XML bytes"));
    }

    let mut positions = Vec::new();
    reserve_exact(&mut positions, offsets.len(), "active-offset positions")?;
    positions.extend(
        offsets
            .iter()
            .copied()
            .enumerate()
            .map(|(index, offset)| (offset, index)),
    );
    positions.sort_unstable_by_key(|&(offset, index)| (offset, index));

    let mut marked = Vec::new();
    reserve_exact(&mut marked, marked_len, "active-offset marked XML")?;
    let mut cursor = 0usize;
    let mut decimal = [0u8; DECIMAL_BUFFER_BYTES];
    for (offset, index) in positions {
        let offset = usize::try_from(offset)
            .map_err(|_| bad("active offset does not fit the platform address space"))?;
        marked.extend_from_slice(
            xml.get(cursor..offset)
                .ok_or_else(|| bad("active offsets are not valid source positions"))?,
        );
        marked.extend_from_slice(b"<!--");
        marked.extend_from_slice(&marker);
        marked.extend_from_slice(decimal_bytes(index, &mut decimal)?);
        marked.extend_from_slice(b"-->");
        cursor = offset;
    }
    marked.extend_from_slice(
        xml.get(cursor..)
            .ok_or_else(|| bad("active-offset source cursor is invalid"))?,
    );

    let processed = process_markup_compatibility(&marked, capabilities, &limits.processing)?;
    let processed = processed.xml.as_ref();
    let mut selected = Vec::new();
    reserve_exact(&mut selected, offsets.len(), "active-offset selection map")?;
    selected.resize(offsets.len(), false);

    let mut cursor = 0usize;
    while let Some(relative) = find_bytes(
        processed
            .get(cursor..)
            .ok_or_else(|| bad("active-offset output cursor is invalid"))?,
        &marker,
    ) {
        let marker_start = cursor
            .checked_add(relative)
            .ok_or_else(|| bad("active-offset marker position overflowed"))?;
        let opening_start = marker_start
            .checked_sub(4)
            .ok_or_else(|| bad("active-offset marker lacks a comment opening"))?;
        if processed.get(opening_start..marker_start) != Some(b"<!--".as_slice()) {
            return Err(bad("active-offset marker lacks a comment opening"));
        }
        let digits_start = marker_start
            .checked_add(marker.len())
            .ok_or_else(|| bad("active-offset marker position overflowed"))?;
        let tail = processed
            .get(digits_start..)
            .ok_or_else(|| bad("active-offset marker is truncated"))?;
        let digits_end =
            find_bytes(tail, b"-->").ok_or_else(|| bad("active-offset marker is unterminated"))?;
        let digits = tail
            .get(..digits_end)
            .ok_or_else(|| bad("active-offset marker range is invalid"))?;
        let index = parse_decimal(digits)?;
        let slot = selected
            .get_mut(index)
            .ok_or_else(|| bad("active-offset marker index is out of range"))?;
        if *slot {
            return Err(bad("active-offset marker is duplicated"));
        }
        *slot = true;
        cursor = digits_start
            .checked_add(digits_end)
            .and_then(|end| end.checked_add(3))
            .ok_or_else(|| bad("active-offset marker end overflowed"))?;
    }

    let retained = selected.iter().filter(|&&selected| selected).count();
    let mut active = Vec::new();
    reserve_exact(&mut active, retained, "active offsets")?;
    for (&offset, selected) in offsets.iter().zip(selected) {
        if selected {
            active.push(offset);
        }
    }
    Ok(active)
}

fn copy_offsets(offsets: &[u32]) -> R<Vec<u32>> {
    let mut copied = Vec::new();
    reserve_exact(&mut copied, offsets.len(), "active offsets")?;
    copied.extend_from_slice(offsets);
    Ok(copied)
}

pub(crate) fn reserve_exact<T>(
    values: &mut Vec<T>,
    additional: usize,
    resource: &'static str,
) -> R<()> {
    values
        .try_reserve_exact(additional)
        .map_err(|source| Error::Allocation { resource, source })
}

fn active_marker(xml: &[u8]) -> R<[u8; ACTIVE_MARKER_TEMPLATE.len()]> {
    let hash = xml.iter().fold(0xcbf2_9ce4_8422_2325u64, |hash, byte| {
        (hash ^ u64::from(*byte)).wrapping_mul(0x0000_0100_0000_01b3)
    });
    active_marker_with_hash(xml, hash)
}

pub(crate) fn active_marker_with_hash(
    xml: &[u8],
    hash: u64,
) -> R<[u8; ACTIVE_MARKER_TEMPLATE.len()]> {
    let mut base = *ACTIVE_MARKER_TEMPLATE;
    write_hex(
        hash,
        base.get_mut(ACTIVE_MARKER_HASH_START..ACTIVE_MARKER_HASH_END)
            .ok_or_else(|| bad("active-offset marker hash range is invalid"))?,
    )?;
    for salt in u8::MIN..=u8::MAX {
        let mut candidate = base;
        write_hex(
            u64::from(salt),
            candidate
                .get_mut(ACTIVE_MARKER_SALT_START..ACTIVE_MARKER_SALT_END)
                .ok_or_else(|| bad("active-offset marker salt range is invalid"))?,
        )?;
        if find_bytes(xml, &candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(bad("source XML collides with active-offset markers"))
}

fn write_hex(mut value: u64, output: &mut [u8]) -> R<()> {
    for byte in output.iter_mut().rev() {
        let nibble = u8::try_from(value & 0x0f)
            .map_err(|_| bad("active-offset marker nibble is invalid"))?;
        *byte = if nibble < 10 {
            b'0' + nibble
        } else {
            b'a' + (nibble - 10)
        };
        value >>= 4;
    }
    if value == 0 {
        Ok(())
    } else {
        Err(bad("active-offset marker value is too large"))
    }
}

fn decimal_len(mut value: usize) -> R<usize> {
    let mut len = 1usize;
    while value >= 10 {
        value /= 10;
        len = len
            .checked_add(1)
            .ok_or_else(|| bad("active-offset decimal length overflowed"))?;
    }
    Ok(len)
}

fn decimal_bytes(mut value: usize, buffer: &mut [u8]) -> R<&[u8]> {
    let mut cursor = buffer.len();
    loop {
        cursor = cursor
            .checked_sub(1)
            .ok_or_else(|| bad("active-offset decimal buffer is too small"))?;
        let digit =
            u8::try_from(value % 10).map_err(|_| bad("active-offset decimal digit is invalid"))?;
        let slot = buffer
            .get_mut(cursor)
            .ok_or_else(|| bad("active-offset decimal cursor is invalid"))?;
        *slot = b'0' + digit;
        value /= 10;
        if value == 0 {
            break;
        }
    }
    buffer
        .get(cursor..)
        .ok_or_else(|| bad("active-offset decimal range is invalid"))
}

fn parse_decimal(digits: &[u8]) -> R<usize> {
    if digits.is_empty() || !digits.iter().all(u8::is_ascii_digit) {
        return Err(bad("active-offset marker index is invalid"));
    }
    digits.iter().try_fold(0usize, |value, digit| {
        value
            .checked_mul(10)
            .and_then(|value| value.checked_add(usize::from(*digit - b'0')))
            .ok_or_else(|| bad("active-offset marker index overflowed"))
    })
}

pub(crate) fn find_bytes(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() {
        return Some(0);
    }
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}
#[derive(Clone, PartialEq, Eq, Hash)]
enum NamePattern {
    Exact(Name),
    Namespace(String),
}
#[derive(Clone)]
struct Ctx {
    ns: BTreeMap<String, String>,
    ign: HashSet<String>,
    process: HashSet<Name>,
    preserve_elements: HashSet<NamePattern>,
    preserve_attributes: HashSet<NamePattern>,
    opaque: bool,
}
enum Mode {
    Emit(String),
    Unwrap,
    Skip,
    Alt {
        choices: usize,
        selected: bool,
        fallback: bool,
    },
    Branch,
}
struct Frame {
    ctx: Ctx,
    mode: Mode,
    active: bool,
}
pub fn process_markup_compatibility<'a>(
    xml: &'a [u8],
    caps: &Capabilities,
    lim: &Limits,
) -> R<Output<'a>> {
    if !xml
        .windows(NAMESPACE.len())
        .any(|w| w == NAMESPACE.as_bytes())
    {
        return Ok(Output {
            xml: Cow::Borrowed(xml),
            report: Report::default(),
        });
    }
    if xml.len() > lim.max_input_bytes {
        return Err(limit("input bytes"));
    }
    let mut r = Reader::from_reader(xml);
    r.config_mut().trim_text(false);
    let (mut stack, mut out, mut rep, mut root, mut buf) = (
        Vec::new(),
        Vec::with_capacity(xml.len()),
        Report::default(),
        false,
        Vec::new(),
    );
    loop {
        let d = r.decoder();
        match r.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => start(
                &e, d, false, caps, lim, &mut stack, &mut out, &mut rep, &mut root,
            )?,
            Ok(Event::Empty(e)) => start(
                &e, d, true, caps, lim, &mut stack, &mut out, &mut rep, &mut root,
            )?,
            Ok(Event::End(_)) => {
                let f: Frame = stack.pop().ok_or_else(|| bad("unexpected end"))?;
                match f.mode {
                    Mode::Alt { choices: 0, .. } => {
                        return Err(bad("AlternateContent requires Choice"));
                    },
                    Mode::Emit(q) if f.active => {
                        out.extend_from_slice(b"</");
                        out.extend_from_slice(q.as_bytes());
                        out.push(b'>')
                    },
                    _ => {},
                }
            },
            Ok(Event::Text(e)) => {
                if visible(&stack) {
                    out.extend_from_slice(e.as_ref())
                }
            },
            Ok(Event::CData(e)) => {
                if visible(&stack) {
                    out.extend_from_slice(b"<![CDATA[");
                    out.extend_from_slice(e.as_ref());
                    out.extend_from_slice(b"]]>")
                }
            },
            Ok(Event::Comment(e)) => {
                if visible(&stack) {
                    out.extend_from_slice(b"<!--");
                    out.extend_from_slice(e.as_ref());
                    out.extend_from_slice(b"-->")
                }
            },
            Ok(Event::Decl(e)) => {
                if stack.is_empty() && !root {
                    out.extend_from_slice(b"<?");
                    out.extend_from_slice(e.as_ref());
                    out.extend_from_slice(b"?>")
                } else {
                    return Err(bad("late XML declaration"));
                }
            },
            Ok(Event::GeneralRef(e)) => {
                if visible(&stack) {
                    let n = e.decode().map_err(xerr)?;
                    if e.resolve_char_ref().map_err(xerr)?.is_none()
                        && !matches!(n.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot")
                    {
                        return Err(bad("custom entity"));
                    }
                    out.push(b'&');
                    out.extend_from_slice(e.as_ref());
                    out.push(b';')
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e.to_string())),
        }
        if out.len() > lim.max_output_bytes {
            return Err(limit("output bytes"));
        }
        buf.clear()
    }
    if !stack.is_empty() {
        return Err(bad("unterminated XML"));
    }
    Ok(Output {
        xml: Cow::Owned(out),
        report: rep,
    })
}
#[allow(clippy::too_many_arguments)]
fn start(
    e: &BytesStart<'_>,
    d: Decoder,
    empty: bool,
    caps: &Capabilities,
    lim: &Limits,
    st: &mut Vec<Frame>,
    out: &mut Vec<u8>,
    rep: &mut Report,
    root: &mut bool,
) -> R<()> {
    if st.len() >= lim.max_depth {
        return Err(limit("depth"));
    }
    let q = str::from_utf8(e.name().as_ref()).map_err(xerr)?.to_string();
    let mut raw = Vec::new();
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xerr)?;
        raw.push((
            str::from_utf8(a.key.as_ref()).map_err(xerr)?.to_string(),
            a.decoded_and_normalized_value(XmlVersion::Explicit1_0, d)
                .map_err(xerr)?
                .into_owned(),
        ))
    }
    let mut c = st.last().map(|f| f.ctx.clone()).unwrap_or_else(|| {
        let mut ns = BTreeMap::new();
        ns.insert("xml".into(), XML_NS.into());
        Ctx {
            ns,
            ign: HashSet::new(),
            process: HashSet::new(),
            preserve_elements: HashSet::new(),
            preserve_attributes: HashSet::new(),
            opaque: false,
        }
    });
    for (a, v) in &raw {
        if a == "xmlns" {
            c.ns.insert("".into(), v.clone());
        } else if let Some(p) = a.strip_prefix("xmlns:") {
            if !xml_name::is_ncname(p) || v.is_empty() {
                return Err(bad("invalid namespace"));
            }
            c.ns.insert(p.into(), v.clone());
        }
    }
    if c.ns.len() > lim.max_namespace_bindings {
        return Err(limit("namespace bindings"));
    }
    let name = expand(&q, &c.ns, true)?;
    let parent_active = st.last().is_none_or(|f| f.active);
    if c.opaque {
        let f = Frame {
            ctx: c.clone(),
            mode: Mode::Emit(q.clone()),
            active: parent_active,
        };
        if parent_active {
            write_start(
                out,
                &q,
                &c.ns,
                &raw,
                &c.ign,
                &c.preserve_attributes,
                caps,
                false,
                rep,
            )?
        }
        return close(st, f, empty, out);
    }

    let mut directives = Vec::new();
    let mut tokens = 0usize;
    for (a, v) in &raw {
        if a == "xmlns" || a.starts_with("xmlns:") {
            continue;
        }
        let n = expand(a, &c.ns, false)?;
        if n.namespace != NAMESPACE {
            continue;
        }
        if !matches!(
            n.local_name.as_str(),
            "Ignorable"
                | "ProcessContent"
                | "PreserveElements"
                | "PreserveAttributes"
                | "MustUnderstand"
        ) {
            return Err(bad("unknown MCE attribute"));
        }
        tokens = tokens
            .checked_add(v.split_whitespace().count())
            .ok_or_else(|| limit("directive tokens"))?;
        if tokens > lim.max_directive_tokens {
            return Err(limit("directive tokens"));
        }
        directives.push((n.local_name, v.as_str()));
    }

    let mut local_ign = HashSet::new();
    if let Some((_, value)) = directives.iter().find(|(name, _)| name == "Ignorable") {
        let mut seen = HashSet::new();
        for prefix in value.split_whitespace() {
            if !xml_name::is_ncname(prefix) || !seen.insert(prefix) {
                return Err(bad("invalid or duplicate Ignorable prefix"));
            }
            let uri =
                c.ns.get(prefix)
                    .ok_or_else(|| bad(format!("unbound Ignorable {prefix}")))?;
            if uri == NAMESPACE {
                return Err(bad("MCE cannot be ignorable"));
            }
            local_ign.insert(uri.clone());
        }
        for uri in &local_ign {
            c.process.retain(|name| &name.namespace != uri);
            c.preserve_elements
                .retain(|pattern| pattern_namespace(pattern) != uri);
            c.preserve_attributes
                .retain(|pattern| pattern_namespace(pattern) != uri);
            c.ign.insert(uri.clone());
        }
    }

    let mut local_process = HashSet::new();
    let mut local_preserve_elements = HashSet::new();
    let mut local_preserve_attributes = HashSet::new();
    for (name, value) in &directives {
        match name.as_str() {
            "Ignorable" => {},
            "ProcessContent" => {
                for token in value.split_whitespace() {
                    let target = parse_qname_target(token, &c.ns, false)?;
                    let NamePattern::Exact(target) = target else {
                        unreachable!()
                    };
                    if !local_ign.contains(&target.namespace) {
                        return Err(bad("ProcessContent target is not locally ignorable"));
                    }
                    if !local_process.insert(target) {
                        return Err(bad("duplicate ProcessContent target"));
                    }
                }
            },
            "PreserveElements" => {
                for token in value.split_whitespace() {
                    let target = parse_qname_target(token, &c.ns, true)?;
                    if !local_ign.contains(pattern_namespace(&target)) {
                        return Err(bad("PreserveElements target is not locally ignorable"));
                    }
                    if !local_preserve_elements.insert(target) {
                        return Err(bad("duplicate PreserveElements target"));
                    }
                }
            },
            "PreserveAttributes" => {
                for token in value.split_whitespace() {
                    let target = parse_qname_target(token, &c.ns, true)?;
                    if !local_ign.contains(pattern_namespace(&target)) {
                        return Err(bad("PreserveAttributes target is not locally ignorable"));
                    }
                    if !local_preserve_attributes.insert(target) {
                        return Err(bad("duplicate PreserveAttributes target"));
                    }
                }
            },
            "MustUnderstand" => {
                let mut seen = HashSet::new();
                for prefix in value.split_whitespace() {
                    if !xml_name::is_ncname(prefix) || !seen.insert(prefix) {
                        return Err(bad("invalid or duplicate MustUnderstand prefix"));
                    }
                    let uri =
                        c.ns.get(prefix)
                            .ok_or_else(|| bad(format!("unbound MustUnderstand {prefix}")))?;
                    if !caps.understands(uri) {
                        return Err(Error::MustUnderstand(uri.clone()));
                    }
                }
            },
            _ => unreachable!(),
        }
    }
    c.process.extend(local_process);
    c.preserve_elements.extend(local_preserve_elements);
    c.preserve_attributes.extend(local_preserve_attributes);
    c.opaque = caps.extensions.contains(&name);

    if let Some(parent) = st.last_mut()
        && let Mode::Alt {
            choices,
            selected,
            fallback,
        } = &mut parent.mode
    {
        if name.namespace != NAMESPACE {
            return Err(bad("non-MCE AlternateContent child"));
        }
        let (active, mode) = match name.local_name.as_str() {
            "Choice" => {
                if *fallback {
                    return Err(bad("Choice after Fallback"));
                }
                *choices += 1;
                if *choices > lim.max_choices_per_alternate {
                    return Err(limit("choices"));
                }
                let req = attr(&raw, "Requires")?.ok_or_else(|| bad("Choice lacks Requires"))?;
                let mut ok = true;
                let mut count = 0;
                for p in req.split_whitespace() {
                    if !xml_name::is_ncname(p) {
                        return Err(bad("invalid Requires prefix"));
                    }
                    count += 1;
                    ok &= caps.understands(
                        c.ns.get(p)
                            .ok_or_else(|| bad(format!("unbound Requires {p}")))?,
                    )
                }
                if count == 0 {
                    return Err(bad("empty Requires"));
                }
                let a = parent.active && !*selected && ok;
                if a {
                    *selected = true;
                    rep.selected_choices += 1
                }
                (a, Mode::Branch)
            },
            "Fallback" => {
                if *fallback {
                    return Err(bad("duplicate Fallback"));
                }
                *fallback = true;
                let a = parent.active && !*selected;
                if a {
                    *selected = true;
                    rep.selected_fallbacks += 1
                }
                (a, Mode::Branch)
            },
            _ => return Err(bad("invalid AlternateContent child")),
        };
        return close(
            st,
            Frame {
                ctx: c,
                mode,
                active,
            },
            empty,
            out,
        );
    }
    let mut active = parent_active;
    let mode = if name.namespace == NAMESPACE {
        match name.local_name.as_str() {
            "AlternateContent" => {
                rep.alternate_content_count += 1;
                Mode::Alt {
                    choices: 0,
                    selected: false,
                    fallback: false,
                }
            },
            _ => return Err(bad("Choice/Fallback outside AlternateContent")),
        }
    } else if c.ign.contains(&name.namespace) && !caps.understands(&name.namespace) {
        if matches_pattern(&c.preserve_elements, &name) {
            rep.preserved_elements += 1;
            Mode::Emit(q.clone())
        } else if c.process.contains(&name) {
            for (a, _) in &raw {
                let n = expand(a, &c.ns, false)?;
                if n.namespace == XML_NS
                    && matches!(n.local_name.as_str(), "base" | "lang" | "space")
                {
                    return Err(bad("xml context attribute on unwrapped element"));
                }
            }
            rep.unwrapped_elements += 1;
            Mode::Unwrap
        } else {
            rep.ignored_elements += 1;
            active = false;
            Mode::Skip
        }
    } else {
        Mode::Emit(q.clone())
    };
    if matches!(mode, Mode::Emit(_)) && active {
        write_start(
            out,
            &q,
            &c.ns,
            &raw,
            &c.ign,
            &c.preserve_attributes,
            caps,
            true,
            rep,
        )?
    }
    if st.is_empty() {
        if *root {
            return Err(bad("multiple roots"));
        }
        *root = true
    }
    close(
        st,
        Frame {
            ctx: c,
            mode,
            active,
        },
        empty,
        out,
    )
}

fn close(st: &mut Vec<Frame>, f: Frame, empty: bool, out: &mut Vec<u8>) -> R<()> {
    if empty {
        match f.mode {
            Mode::Emit(q) if f.active => {
                out.extend_from_slice(b"</");
                out.extend_from_slice(q.as_bytes());
                out.push(b'>')
            },
            Mode::Alt { .. } => return Err(bad("empty AlternateContent")),
            _ => {},
        }
    } else {
        st.push(f)
    }
    Ok(())
}
fn visible(s: &[Frame]) -> bool {
    s.last().is_some_and(|f| f.active)
}
fn bad(s: impl Into<String>) -> Error {
    Error::NonConformant(s.into())
}
fn limit(s: &str) -> Error {
    Error::LimitExceeded(s.into())
}
fn xerr(e: impl std::fmt::Display) -> Error {
    Error::Xml(e.to_string())
}
fn attr<'a>(r: &'a [(String, String)], n: &str) -> R<Option<&'a str>> {
    let mut v = None;
    for (a, x) in r {
        if a == n {
            if v.is_some() {
                return Err(bad("duplicate attribute"));
            }
            v = Some(x.as_str())
        }
    }
    Ok(v)
}
fn expand(q: &str, ns: &BTreeMap<String, String>, element: bool) -> R<Name> {
    let qualified = QualifiedName::try_from(q).map_err(|_| bad("invalid QName"))?;
    let p = qualified.prefix().unwrap_or_default();
    let l = qualified.local();
    let n = if p.is_empty() {
        if element {
            ns.get("").cloned().unwrap_or_default()
        } else {
            String::new()
        }
    } else {
        ns.get(p)
            .cloned()
            .ok_or_else(|| bad(format!("unbound prefix {p}")))?
    };
    Ok(Name {
        namespace: n,
        local_name: l.into(),
    })
}
fn pattern_namespace(pattern: &NamePattern) -> &str {
    match pattern {
        NamePattern::Exact(name) => &name.namespace,
        NamePattern::Namespace(namespace) => namespace,
    }
}
fn matches_pattern(patterns: &HashSet<NamePattern>, name: &Name) -> bool {
    patterns.contains(&NamePattern::Exact(name.clone()))
        || patterns.contains(&NamePattern::Namespace(name.namespace.clone()))
}
fn parse_qname_target(
    token: &str,
    ns: &BTreeMap<String, String>,
    wildcard: bool,
) -> R<NamePattern> {
    let (prefix, local) = token
        .split_once(':')
        .ok_or_else(|| bad("preservation and processing targets must be prefixed QNames"))?;
    if token.matches(':').count() != 1 || !xml_name::is_ncname(prefix) || local.is_empty() {
        return Err(bad("invalid compatibility target QName"));
    }
    let namespace = ns
        .get(prefix)
        .cloned()
        .ok_or_else(|| bad(format!("unbound compatibility target prefix {prefix}")))?;
    if namespace == NAMESPACE {
        return Err(bad("compatibility target cannot use the MCE namespace"));
    }
    if local == "*" {
        if wildcard {
            return Ok(NamePattern::Namespace(namespace));
        }
        return Err(bad("wildcard is not allowed in ProcessContent"));
    }
    if !xml_name::is_ncname(local) {
        return Err(bad("invalid compatibility target QName"));
    }
    Ok(NamePattern::Exact(Name {
        namespace,
        local_name: local.into(),
    }))
}
#[allow(clippy::too_many_arguments)]
fn write_start(
    o: &mut Vec<u8>,
    q: &str,
    ns: &BTreeMap<String, String>,
    raw: &[(String, String)],
    ign: &HashSet<String>,
    preserve: &HashSet<NamePattern>,
    caps: &Capabilities,
    filter: bool,
    rep: &mut Report,
) -> R<()> {
    o.push(b'<');
    o.extend_from_slice(q.as_bytes());
    for (p, u) in ns {
        if p == "xml" {
            continue;
        }
        o.extend_from_slice(if p.is_empty() { b" xmlns" } else { b" xmlns:" });
        if !p.is_empty() {
            o.extend_from_slice(p.as_bytes())
        }
        o.extend_from_slice(b"=\"");
        esc(o, u);
        o.push(b'\"')
    }
    for (a, v) in raw {
        if a == "xmlns" || a.starts_with("xmlns:") {
            continue;
        }
        let n = expand(a, ns, false)?;
        if filter && n.namespace == NAMESPACE {
            rep.ignored_attributes += 1;
            continue;
        }
        if filter
            && !n.namespace.is_empty()
            && ign.contains(&n.namespace)
            && !caps.understands(&n.namespace)
        {
            if matches_pattern(preserve, &n) {
                rep.preserved_attributes += 1
            } else {
                rep.ignored_attributes += 1;
                continue;
            }
        }
        o.push(b' ');
        o.extend_from_slice(a.as_bytes());
        o.extend_from_slice(b"=\"");
        esc(o, v);
        o.push(b'\"')
    }
    o.push(b'>');
    Ok(())
}
fn esc(o: &mut Vec<u8>, s: &str) {
    for c in s.chars() {
        match c {
            '&' => o.extend_from_slice(b"&amp;"),
            '<' => o.extend_from_slice(b"&lt;"),
            '"' => o.extend_from_slice(b"&quot;"),
            '\t' => o.extend_from_slice(b"&#x9;"),
            '\n' => o.extend_from_slice(b"&#xA;"),
            '\r' => o.extend_from_slice(b"&#xD;"),
            _ => {
                let mut b = [0; 4];
                o.extend_from_slice(c.encode_utf8(&mut b).as_bytes())
            },
        }
    }
}
/// Applies the baseline OOXML markup-compatibility profile.
pub fn process_ooxml(x: &[u8]) -> R<Cow<'_, [u8]>> {
    process_markup_compatibility(x, &Capabilities::default(), &Limits::default()).map(|x| x.xml)
}

/// Applies baseline preprocessing to one OPC part.
pub fn process_part(part: &dyn litchi_opc::Part) -> R<Cow<'_, [u8]>> {
    process_ooxml(part.blob())
}

/// Applies baseline preprocessing while retaining the part's shared blob on the fast path.
pub fn process_part_arc(part: &dyn litchi_opc::Part) -> R<std::sync::Arc<Vec<u8>>> {
    Ok(match process_part(part)? {
        Cow::Borrowed(_) => part.blob_arc(),
        Cow::Owned(v) => std::sync::Arc::new(v),
    })
}
/// Applies baseline preprocessing to UTF-8 OOXML.
pub fn process_str(x: &str) -> R<Cow<'_, str>> {
    match process_ooxml(x.as_bytes())? {
        Cow::Borrowed(_) => Ok(Cow::Borrowed(x)),
        Cow::Owned(v) => String::from_utf8(v)
            .map(Cow::Owned)
            .map_err(|error| Error::Xml(format!("MCE output is not UTF-8: {error}"))),
    }
}
