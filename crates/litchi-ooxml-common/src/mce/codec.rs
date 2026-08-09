//! Bounded ISO/IEC 29500-3 MCE preprocessing and active-offset selection.

use quick_xml::{
    Reader, XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
};
use std::{borrow::Cow, collections::HashSet, str, sync::Arc};

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

#[derive(Clone, Default)]
struct Namespaces {
    head: Option<Arc<NamespaceLayer>>,
    bindings: usize,
}

struct NamespaceLayer {
    parent: Option<Arc<NamespaceLayer>>,
    local: Vec<(String, String)>,
}

impl Namespaces {
    fn get(&self, prefix: &str) -> Option<&str> {
        if prefix == "xml" {
            return Some(XML_NS);
        }
        let mut layer = self.head.as_deref();
        while let Some(current) = layer {
            if let Some((_, namespace)) = current
                .local
                .iter()
                .rev()
                .find(|(candidate, _)| candidate == prefix)
            {
                return Some(namespace);
            }
            layer = current.parent.as_deref();
        }
        None
    }

    fn with_local(&self, local: Vec<(String, String)>, lim: &Limits) -> R<Self> {
        if local.is_empty() {
            return Ok(self.clone());
        }
        let mut bindings = self.bindings;
        for (index, (prefix, _)) in local.iter().enumerate() {
            if local[..index]
                .iter()
                .any(|(candidate, _)| candidate == prefix)
            {
                return Err(bad("duplicate namespace declaration"));
            }
            if self.get(prefix).is_none() {
                bindings = bindings
                    .checked_add(1)
                    .ok_or_else(|| limit("namespace bindings"))?;
            }
        }
        if bindings > lim.max_namespace_bindings {
            return Err(limit("namespace bindings"));
        }
        Ok(Self {
            head: Some(Arc::new(NamespaceLayer {
                parent: self.head.clone(),
                local,
            })),
            bindings,
        })
    }

    fn for_each_effective(&self, mut visit: impl FnMut(&str, &str) -> R<()>) -> R<()> {
        let mut layer = self.head.as_ref();
        while let Some(current) = layer {
            for (prefix, namespace) in &current.local {
                if prefix != "xml" && !self.shadowed_before(current, prefix) {
                    visit(prefix, namespace)?;
                }
            }
            layer = current.parent.as_ref();
        }
        Ok(())
    }

    fn shadowed_before(&self, target: &Arc<NamespaceLayer>, prefix: &str) -> bool {
        let mut layer = self.head.as_ref();
        while let Some(current) = layer {
            if Arc::ptr_eq(current, target) {
                return false;
            }
            if current
                .local
                .iter()
                .any(|(candidate, _)| candidate == prefix)
            {
                return true;
            }
            layer = current.parent.as_ref();
        }
        false
    }
}

struct DirectiveLayer {
    parent: Option<Arc<DirectiveLayer>>,
    ignorable: HashSet<String>,
    process: HashSet<NamePattern>,
    preserve_elements: HashSet<NamePattern>,
    preserve_attributes: HashSet<NamePattern>,
}

#[derive(Clone)]
struct Ctx {
    ns: Namespaces,
    directives: Option<Arc<DirectiveLayer>>,
    opaque: bool,
}

impl Ctx {
    fn root() -> Self {
        Self {
            ns: Namespaces {
                head: None,
                bindings: 1,
            },
            directives: None,
            opaque: false,
        }
    }

    fn is_ignorable(&self, namespace: &str) -> bool {
        let mut layer = self.directives.as_deref();
        while let Some(current) = layer {
            if current.ignorable.contains(namespace) {
                return true;
            }
            layer = current.parent.as_deref();
        }
        false
    }

    fn processes(&self, name: &Name) -> bool {
        pattern_directive_matches(&self.directives, name, |layer| &layer.process)
    }

    fn preserves_element(&self, name: &Name) -> bool {
        pattern_directive_matches(&self.directives, name, |layer| &layer.preserve_elements)
    }

    fn preserves_attribute(&self, name: &Name) -> bool {
        pattern_directive_matches(&self.directives, name, |layer| &layer.preserve_attributes)
    }
}

fn pattern_directive_matches(
    head: &Option<Arc<DirectiveLayer>>,
    name: &Name,
    select: impl Fn(&DirectiveLayer) -> &HashSet<NamePattern>,
) -> bool {
    let mut layer = head.as_deref();
    while let Some(current) = layer {
        if matches_pattern(select(current), name) {
            return true;
        }
        if current.ignorable.contains(&name.namespace) {
            return false;
        }
        layer = current.parent.as_deref();
    }
    false
}

struct BoundedOutput {
    bytes: Vec<u8>,
    max: usize,
}

impl BoundedOutput {
    fn new(hint: usize, max: usize) -> R<Self> {
        let mut bytes = Vec::new();
        reserve_exact(&mut bytes, hint.min(max), "MCE output")?;
        Ok(Self { bytes, max })
    }

    fn extend_from_slice(&mut self, value: &[u8]) -> R<()> {
        self.reserve(value.len())?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn push(&mut self, value: u8) -> R<()> {
        self.reserve(1)?;
        self.bytes.push(value);
        Ok(())
    }

    fn reserve(&mut self, additional: usize) -> R<()> {
        let len = self
            .bytes
            .len()
            .checked_add(additional)
            .ok_or_else(|| limit("output bytes"))?;
        if len > self.max {
            return Err(limit("output bytes"));
        }
        if additional > self.bytes.capacity().saturating_sub(self.bytes.len()) {
            self.bytes
                .try_reserve_exact(additional)
                .map_err(|source| Error::Allocation {
                    resource: "MCE output",
                    source,
                })?;
        }
        Ok(())
    }

    fn into_inner(self) -> Vec<u8> {
        self.bytes
    }
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
    if xml.len() > lim.max_input_bytes {
        return Err(limit("input bytes"));
    }
    if !xml
        .windows(NAMESPACE.len())
        .any(|w| w == NAMESPACE.as_bytes())
    {
        if xml.len() > lim.max_output_bytes {
            return Err(limit("output bytes"));
        }
        return Ok(Output {
            xml: Cow::Borrowed(xml),
            report: Report::default(),
        });
    }
    let mut r = Reader::from_reader(xml);
    r.config_mut().trim_text(false);
    let mut stack = Vec::new();
    let mut out = BoundedOutput::new(xml.len(), lim.max_output_bytes)?;
    let mut rep = Report::default();
    let mut root = false;
    let mut buf = Vec::new();
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
                        out.extend_from_slice(b"</")?;
                        out.extend_from_slice(q.as_bytes())?;
                        out.push(b'>')?;
                    },
                    _ => {},
                }
            },
            Ok(Event::Text(e)) => {
                if visible(&stack) {
                    out.extend_from_slice(e.as_ref())?;
                }
            },
            Ok(Event::CData(e)) => {
                if visible(&stack) {
                    out.extend_from_slice(b"<![CDATA[")?;
                    out.extend_from_slice(e.as_ref())?;
                    out.extend_from_slice(b"]]>")?;
                }
            },
            Ok(Event::Comment(e)) => {
                if visible(&stack) {
                    out.extend_from_slice(b"<!--")?;
                    out.extend_from_slice(e.as_ref())?;
                    out.extend_from_slice(b"-->")?;
                }
            },
            Ok(Event::Decl(e)) => {
                if stack.is_empty() && !root {
                    out.extend_from_slice(b"<?")?;
                    out.extend_from_slice(e.as_ref())?;
                    out.extend_from_slice(b"?>")?;
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
                    out.push(b'&')?;
                    out.extend_from_slice(e.as_ref())?;
                    out.push(b';')?;
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e.to_string())),
        }
        buf.clear();
    }
    if !stack.is_empty() {
        return Err(bad("unterminated XML"));
    }
    Ok(Output {
        xml: Cow::Owned(out.into_inner()),
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
    out: &mut BoundedOutput,
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
        reserve_exact(&mut raw, 1, "MCE attributes")?;
        raw.push((
            str::from_utf8(a.key.as_ref()).map_err(xerr)?.to_string(),
            a.decoded_and_normalized_value(XmlVersion::Explicit1_0, d)
                .map_err(xerr)?
                .into_owned(),
        ));
    }
    let mut c = st.last().map_or_else(Ctx::root, |f| f.ctx.clone());
    let mut local_namespaces = Vec::new();
    for (a, v) in &raw {
        if a == "xmlns" {
            reserve_exact(&mut local_namespaces, 1, "MCE namespace declarations")?;
            local_namespaces.push((String::new(), v.clone()));
        } else if let Some(p) = a.strip_prefix("xmlns:") {
            if !xml_name::is_ncname(p) || v.is_empty() {
                return Err(bad("invalid namespace"));
            }
            reserve_exact(&mut local_namespaces, 1, "MCE namespace declarations")?;
            local_namespaces.push((p.into(), v.clone()));
        }
    }
    c.ns = c.ns.with_local(local_namespaces, lim)?;
    let name = expand(&q, &c.ns, true)?;
    let parent_active = st.last().is_none_or(|f| f.active);
    if c.opaque {
        let f = Frame {
            ctx: c.clone(),
            mode: Mode::Emit(q.clone()),
            active: parent_active,
        };
        if parent_active {
            write_start(out, &q, &c, &raw, caps, false, rep)?;
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
        reserve_exact(&mut directives, 1, "MCE directives")?;
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
            local_ign
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "MCE Ignorable directives",
                    source,
                })?;
            local_ign.insert(uri.to_owned());
        }
    }
    let mut new_ignorable = HashSet::new();
    for namespace in &local_ign {
        if !c.is_ignorable(namespace) {
            new_ignorable
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "MCE effective Ignorable directives",
                    source,
                })?;
            new_ignorable.insert(namespace.clone());
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
                    let target = parse_qname_target(token, &c.ns, true)?;
                    let namespace = pattern_namespace(&target);
                    if !local_ign.contains(namespace) && !c.is_ignorable(namespace) {
                        return Err(bad("ProcessContent target is not effectively ignorable"));
                    }
                    local_process
                        .try_reserve(1)
                        .map_err(|source| Error::Allocation {
                            resource: "MCE ProcessContent directives",
                            source,
                        })?;
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
                    local_preserve_elements
                        .try_reserve(1)
                        .map_err(|source| Error::Allocation {
                            resource: "MCE PreserveElements directives",
                            source,
                        })?;
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
                    local_preserve_attributes.try_reserve(1).map_err(|source| {
                        Error::Allocation {
                            resource: "MCE PreserveAttributes directives",
                            source,
                        }
                    })?;
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
                        return Err(Error::MustUnderstand(uri.to_owned()));
                    }
                }
            },
            _ => unreachable!(),
        }
    }
    if !new_ignorable.is_empty()
        || !local_process.is_empty()
        || !local_preserve_elements.is_empty()
        || !local_preserve_attributes.is_empty()
    {
        c.directives = Some(Arc::new(DirectiveLayer {
            parent: c.directives.take(),
            ignorable: new_ignorable,
            process: local_process,
            preserve_elements: local_preserve_elements,
            preserve_attributes: local_preserve_attributes,
        }));
    }
    c.opaque = caps.extensions.contains(&name);

    if name.namespace == NAMESPACE {
        match name.local_name.as_str() {
            "AlternateContent" => {
                validate_alternate_attributes(&raw, &c, caps, AlternateKind::Container)?;
            },
            "Choice" => validate_alternate_attributes(&raw, &c, caps, AlternateKind::Choice)?,
            "Fallback" => validate_alternate_attributes(&raw, &c, caps, AlternateKind::Fallback)?,
            _ => {},
        }
    }

    if let Some(parent) = st.last_mut()
        && let Mode::Alt {
            choices,
            selected,
            fallback,
        } = &mut parent.mode
    {
        if name.namespace != NAMESPACE {
            if c.is_ignorable(&name.namespace) && !caps.understands(&name.namespace) {
                rep.ignored_elements += 1;
                return close(
                    st,
                    Frame {
                        ctx: c,
                        mode: Mode::Skip,
                        active: false,
                    },
                    empty,
                    out,
                );
            }
            return Err(bad("non-ignorable AlternateContent child"));
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
                    );
                }
                if count == 0 {
                    return Err(bad("empty Requires"));
                }
                let a = parent.active && !*selected && ok;
                if a {
                    *selected = true;
                    rep.selected_choices += 1;
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
                    rep.selected_fallbacks += 1;
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
    } else if c.is_ignorable(&name.namespace) && !caps.understands(&name.namespace) {
        if c.preserves_element(&name) {
            rep.preserved_elements += 1;
            Mode::Emit(q.clone())
        } else if c.processes(&name) {
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
        write_start(out, &q, &c, &raw, caps, true, rep)?;
    }
    if st.is_empty() {
        if *root {
            return Err(bad("multiple roots"));
        }
        *root = true;
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

fn close(st: &mut Vec<Frame>, f: Frame, empty: bool, out: &mut BoundedOutput) -> R<()> {
    if empty {
        match f.mode {
            Mode::Emit(q) if f.active => {
                out.extend_from_slice(b"</")?;
                out.extend_from_slice(q.as_bytes())?;
                out.push(b'>')?;
            },
            Mode::Alt { .. } => return Err(bad("empty AlternateContent")),
            _ => {},
        }
    } else {
        reserve_exact(st, 1, "MCE element stack")?;
        st.push(f);
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
            v = Some(x.as_str());
        }
    }
    Ok(v)
}

#[derive(Clone, Copy)]
enum AlternateKind {
    Container,
    Choice,
    Fallback,
}

fn validate_alternate_attributes(
    raw: &[(String, String)],
    ctx: &Ctx,
    caps: &Capabilities,
    kind: AlternateKind,
) -> R<()> {
    for (qualified, _) in raw {
        if qualified == "xmlns" || qualified.starts_with("xmlns:") {
            continue;
        }
        let name = expand(qualified, &ctx.ns, false)?;
        if name.namespace.is_empty() {
            if matches!(kind, AlternateKind::Choice) && name.local_name == "Requires" {
                continue;
            }
            return Err(bad("unexpected unprefixed AlternateContent attribute"));
        }
        if name.namespace == XML_NS && matches!(name.local_name.as_str(), "lang" | "space") {
            return Err(bad(
                "xml:lang and xml:space are forbidden on AlternateContent markup",
            ));
        }
        if name.namespace == NAMESPACE || name.namespace == XML_NS {
            continue;
        }
        if !caps.understands(&name.namespace) && !ctx.is_ignorable(&name.namespace) {
            return Err(bad(
                "AlternateContent attribute namespace is neither understood nor ignorable",
            ));
        }
    }
    Ok(())
}
fn expand(q: &str, ns: &Namespaces, element: bool) -> R<Name> {
    let qualified = QualifiedName::try_from(q).map_err(|_| bad("invalid QName"))?;
    let p = qualified.prefix().unwrap_or_default();
    let l = qualified.local();
    let n = if p.is_empty() {
        if element {
            ns.get("").unwrap_or_default().to_owned()
        } else {
            String::new()
        }
    } else {
        ns.get(p)
            .map(str::to_owned)
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
    patterns.iter().any(|pattern| match pattern {
        NamePattern::Exact(candidate) => candidate == name,
        NamePattern::Namespace(namespace) => namespace == &name.namespace,
    })
}
fn parse_qname_target(token: &str, ns: &Namespaces, wildcard: bool) -> R<NamePattern> {
    let (prefix, local) = token
        .split_once(':')
        .ok_or_else(|| bad("preservation and processing targets must be prefixed QNames"))?;
    if token.matches(':').count() != 1 || !xml_name::is_ncname(prefix) || local.is_empty() {
        return Err(bad("invalid compatibility target QName"));
    }
    let namespace = ns
        .get(prefix)
        .map(str::to_owned)
        .ok_or_else(|| bad(format!("unbound compatibility target prefix {prefix}")))?;
    if namespace == NAMESPACE {
        return Err(bad("compatibility target cannot use the MCE namespace"));
    }
    if local == "*" {
        return wildcard
            .then_some(NamePattern::Namespace(namespace))
            .ok_or_else(|| bad("wildcard is not allowed in this compatibility directive"));
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
    o: &mut BoundedOutput,
    q: &str,
    ctx: &Ctx,
    raw: &[(String, String)],
    caps: &Capabilities,
    filter: bool,
    rep: &mut Report,
) -> R<()> {
    o.push(b'<')?;
    o.extend_from_slice(q.as_bytes())?;
    ctx.ns.for_each_effective(|p, u| {
        o.extend_from_slice(if p.is_empty() { b" xmlns" } else { b" xmlns:" })?;
        if !p.is_empty() {
            o.extend_from_slice(p.as_bytes())?;
        }
        o.extend_from_slice(b"=\"")?;
        esc(o, u)?;
        o.push(b'\"')
    })?;
    for (a, v) in raw {
        if a == "xmlns" || a.starts_with("xmlns:") {
            continue;
        }
        let n = expand(a, &ctx.ns, false)?;
        if filter && n.namespace == NAMESPACE {
            rep.ignored_attributes += 1;
            continue;
        }
        if filter
            && !n.namespace.is_empty()
            && ctx.is_ignorable(&n.namespace)
            && !caps.understands(&n.namespace)
        {
            if ctx.preserves_attribute(&n) {
                rep.preserved_attributes += 1;
            } else {
                rep.ignored_attributes += 1;
                continue;
            }
        }
        o.push(b' ')?;
        o.extend_from_slice(a.as_bytes())?;
        o.extend_from_slice(b"=\"")?;
        esc(o, v)?;
        o.push(b'\"')?;
    }
    o.push(b'>')?;
    Ok(())
}
fn esc(o: &mut BoundedOutput, s: &str) -> R<()> {
    for c in s.chars() {
        match c {
            '&' => o.extend_from_slice(b"&amp;")?,
            '<' => o.extend_from_slice(b"&lt;")?,
            '"' => o.extend_from_slice(b"&quot;")?,
            '\t' => o.extend_from_slice(b"&#x9;")?,
            '\n' => o.extend_from_slice(b"&#xA;")?,
            '\r' => o.extend_from_slice(b"&#xD;")?,
            _ => {
                let mut b = [0; 4];
                o.extend_from_slice(c.encode_utf8(&mut b).as_bytes())?;
            },
        }
    }
    Ok(())
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
pub fn process_part_arc(part: &dyn litchi_opc::Part) -> R<Arc<Vec<u8>>> {
    Ok(match process_part(part)? {
        Cow::Borrowed(_) => part.blob_arc(),
        Cow::Owned(v) => Arc::new(v),
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
