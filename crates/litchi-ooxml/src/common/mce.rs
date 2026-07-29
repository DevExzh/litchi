//! Bounded ISO/IEC 29500-3:2015 semantic preprocessing.
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
use thiserror::Error;
pub const MCE_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ExpandedName {
    pub namespace: String,
    pub local_name: String,
}
#[derive(Debug, Clone)]
pub struct MceCapabilities {
    understood: HashSet<String>,
    extensions: HashSet<ExpandedName>,
}
impl MceCapabilities {
    pub fn new() -> Self {
        Self {
            understood: HashSet::new(),
            extensions: HashSet::new(),
        }
    }
    pub fn ooxml_baseline() -> Self {
        let mut s = Self::new();
        for n in [
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "http://purl.oclc.org/ooxml/spreadsheetml/main",
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "http://purl.oclc.org/ooxml/presentationml/main",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "http://purl.oclc.org/ooxml/drawingml/main",
            "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "http://purl.oclc.org/ooxml/drawingml/chart",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
            "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "http://purl.oclc.org/ooxml/officeDocument/math",
            "urn:schemas-microsoft-com:vml",
            "urn:schemas-microsoft-com:office:office",
            XML_NS,
        ] {
            s.understood.insert(n.into());
        }
        s
    }
    pub fn understand_namespace(&mut self, n: impl Into<String>) -> &mut Self {
        self.understood.insert(n.into());
        self
    }
    pub fn preserve_extension_element(&mut self, n: ExpandedName) -> &mut Self {
        self.extensions.insert(n);
        self
    }
    pub fn understands(&self, n: &str) -> bool {
        self.understood.contains(n)
    }
}
impl Default for MceCapabilities {
    fn default() -> Self {
        Self::ooxml_baseline()
    }
}
#[derive(Debug, Clone)]
pub struct MceLimits {
    pub max_input_bytes: usize,
    pub max_output_bytes: usize,
    pub max_depth: usize,
    pub max_namespace_bindings: usize,
    pub max_directive_tokens: usize,
    pub max_choices_per_alternate: usize,
}
impl Default for MceLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: 256 * 1024 * 1024,
            max_output_bytes: 512 * 1024 * 1024,
            max_depth: 256,
            max_namespace_bindings: 4096,
            max_directive_tokens: 4096,
            max_choices_per_alternate: 1024,
        }
    }
}
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MceReport {
    pub alternate_content_count: usize,
    pub selected_choices: usize,
    pub selected_fallbacks: usize,
    pub ignored_elements: usize,
    pub ignored_attributes: usize,
    pub preserved_elements: usize,
    pub preserved_attributes: usize,
    pub unwrapped_elements: usize,
}
#[derive(Debug)]
pub struct MceOutput<'a> {
    pub xml: Cow<'a, [u8]>,
    pub report: MceReport,
}
#[derive(Debug, Error, Clone, PartialEq, Eq)]
pub enum MceError {
    #[error("non-conformant markup compatibility XML: {0}")]
    NonConformant(String),
    #[error("unsupported namespace required by MustUnderstand: {0}")]
    MustUnderstand(String),
    #[error("markup compatibility resource limit exceeded: {0}")]
    LimitExceeded(String),
    #[error("markup compatibility XML error: {0}")]
    Xml(String),
}
type R<T> = std::result::Result<T, MceError>;
#[derive(Clone, PartialEq, Eq, Hash)]
enum NamePattern {
    Exact(ExpandedName),
    Namespace(String),
}
#[derive(Clone)]
struct Ctx {
    ns: BTreeMap<String, String>,
    ign: HashSet<String>,
    process: HashSet<ExpandedName>,
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
    caps: &MceCapabilities,
    lim: &MceLimits,
) -> R<MceOutput<'a>> {
    if !xml
        .windows(MCE_NAMESPACE.len())
        .any(|w| w == MCE_NAMESPACE.as_bytes())
    {
        return Ok(MceOutput {
            xml: Cow::Borrowed(xml),
            report: MceReport::default(),
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
        MceReport::default(),
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
            Err(e) => return Err(MceError::Xml(e.to_string())),
        }
        if out.len() > lim.max_output_bytes {
            return Err(limit("output bytes"));
        }
        buf.clear()
    }
    if !stack.is_empty() {
        return Err(bad("unterminated XML"));
    }
    Ok(MceOutput {
        xml: Cow::Owned(out),
        report: rep,
    })
}
#[allow(clippy::too_many_arguments)]
fn start(
    e: &BytesStart<'_>,
    d: Decoder,
    empty: bool,
    caps: &MceCapabilities,
    lim: &MceLimits,
    st: &mut Vec<Frame>,
    out: &mut Vec<u8>,
    rep: &mut MceReport,
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
            if !valid_ncname(p) || v.is_empty() {
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
        if n.namespace != MCE_NAMESPACE {
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
            if !valid_ncname(prefix) || !seen.insert(prefix) {
                return Err(bad("invalid or duplicate Ignorable prefix"));
            }
            let uri =
                c.ns.get(prefix)
                    .ok_or_else(|| bad(format!("unbound Ignorable {prefix}")))?;
            if uri == MCE_NAMESPACE {
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
                    if !valid_ncname(prefix) || !seen.insert(prefix) {
                        return Err(bad("invalid or duplicate MustUnderstand prefix"));
                    }
                    let uri =
                        c.ns.get(prefix)
                            .ok_or_else(|| bad(format!("unbound MustUnderstand {prefix}")))?;
                    if !caps.understands(uri) {
                        return Err(MceError::MustUnderstand(uri.clone()));
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

    if let Some(parent) = st.last_mut() {
        if let Mode::Alt {
            choices,
            selected,
            fallback,
        } = &mut parent.mode
        {
            if name.namespace != MCE_NAMESPACE {
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
                    let req =
                        attr(&raw, "Requires")?.ok_or_else(|| bad("Choice lacks Requires"))?;
                    let mut ok = true;
                    let mut count = 0;
                    for p in req.split_whitespace() {
                        if !valid_ncname(p) {
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
    }
    let mut active = parent_active;
    let mode = if name.namespace == MCE_NAMESPACE {
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

#[cfg(test)]
mod preservation_tests {
    use super::*;

    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    fn run_with(
        xml: &str,
        capabilities: &MceCapabilities,
        limits: &MceLimits,
    ) -> Result<(String, MceReport), MceError> {
        let output = process_markup_compatibility(xml.as_bytes(), capabilities, limits)?;
        Ok((
            String::from_utf8(output.xml.into_owned()).expect("MCE output must remain UTF-8"),
            output.report,
        ))
    }

    fn run(xml: &str) -> Result<(String, MceReport), MceError> {
        run_with(xml, &MceCapabilities::new(), &MceLimits::default())
    }

    #[test]
    fn preserves_exact_and_wildcard_attributes_by_expanded_name() {
        let exact = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" xmlns:y="urn:ext" mc:Ignorable="x" mc:PreserveAttributes="x:keep"><a y:keep="yes" y:drop="no"/></r>"#
        );
        let (xml, report) = run(&exact).unwrap();
        assert!(xml.contains(r#"y:keep="yes""#));
        assert!(!xml.contains("y:drop"));
        assert_eq!(report.preserved_attributes, 1);

        let wildcard = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" mc:Ignorable="x" mc:PreserveAttributes="x:*"><a x:one="1" x:two="2"/></r>"#
        );
        let (xml, report) = run(&wildcard).unwrap();
        assert!(xml.contains(r#"x:one="1""#));
        assert!(xml.contains(r#"x:two="2""#));
        assert_eq!(report.preserved_attributes, 2);
    }

    #[test]
    fn preserves_elements_but_still_processes_their_content_and_attributes() {
        let source = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" mc:Ignorable="x" mc:PreserveElements="x:keep" mc:PreserveAttributes="x:flag"><x:keep plain="yes" x:flag="yes" x:drop="no"><x:drop/><known/></x:keep><x:discard/></r>"#
        );
        let (xml, report) = run(&source).unwrap();
        assert!(xml.contains("<x:keep"));
        assert!(xml.contains(r#"plain="yes""#));
        assert!(xml.contains(r#"x:flag="yes""#));
        assert!(xml.contains("<known"));
        assert!(!xml.contains("x:drop"));
        assert!(!xml.contains("x:discard"));
        assert_eq!(report.preserved_elements, 1);
        assert_eq!(report.preserved_attributes, 1);
    }

    #[test]
    fn local_ignorable_redeclaration_resets_inherited_preservation() {
        let source = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" mc:Ignorable="x" mc:PreserveAttributes="x:*"><a mc:Ignorable="x" x:value="discarded"/></r>"#
        );
        let (xml, report) = run(&source).unwrap();
        assert!(!xml.contains("x:value"));
        assert_eq!(report.preserved_attributes, 0);
    }

    #[test]
    fn understood_attributes_are_not_discarded_and_spoofed_directives_do_not_apply() {
        let understood = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" mc:Ignorable="x"><a x:value="kept"/></r>"#
        );
        let mut capabilities = MceCapabilities::new();
        capabilities.understand_namespace("urn:ext");
        let (xml, _) = run_with(&understood, &capabilities, &MceLimits::default()).unwrap();
        assert!(xml.contains(r#"x:value="kept""#));

        let spoofed = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" xmlns:f="urn:fake" mc:Ignorable="x f" f:PreserveAttributes="x:*"><a x:value="discarded"/></r>"#
        );
        let (xml, _) = run(&spoofed).unwrap();
        assert!(!xml.contains("PreserveAttributes"));
        assert!(!xml.contains("x:value"));
    }

    #[test]
    fn rejects_invalid_preservation_tokens_and_duplicates() {
        for directive in ["keep", "missing:keep", "x:keep:extra", "x:keep x:keep"] {
            let source = format!(
                r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" mc:Ignorable="x" mc:PreserveAttributes="{directive}"/>"#
            );
            assert!(
                run(&source).is_err(),
                "accepted invalid token list: {directive}"
            );
        }

        let wildcard_process = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" mc:Ignorable="x" mc:ProcessContent="x:*"/>"#
        );
        assert!(run(&wildcard_process).is_err());

        let wrong_namespace = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" xmlns:y="urn:other" mc:Ignorable="x" mc:PreserveElements="y:keep"/>"#
        );
        assert!(run(&wrong_namespace).is_err());
    }

    #[test]
    fn preservation_tokens_respect_the_shared_directive_bound() {
        let source = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" mc:Ignorable="x" mc:PreserveAttributes="x:one x:two"/>"#
        );
        let limits = MceLimits {
            max_directive_tokens: 2,
            ..MceLimits::default()
        };
        assert!(matches!(
            run_with(&source, &MceCapabilities::new(), &limits),
            Err(MceError::LimitExceeded(_))
        ));
    }
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
fn bad(s: impl Into<String>) -> MceError {
    MceError::NonConformant(s.into())
}
fn limit(s: &str) -> MceError {
    MceError::LimitExceeded(s.into())
}
fn xerr(e: impl std::fmt::Display) -> MceError {
    MceError::Xml(e.to_string())
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
fn expand(q: &str, ns: &BTreeMap<String, String>, element: bool) -> R<ExpandedName> {
    let (p, l) = q.split_once(':').unwrap_or(("", q));
    if l.is_empty() || q.matches(':').count() > 1 {
        return Err(bad("invalid QName"));
    }
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
    Ok(ExpandedName {
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
fn matches_pattern(patterns: &HashSet<NamePattern>, name: &ExpandedName) -> bool {
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
    if token.matches(':').count() != 1 || !valid_ncname(prefix) || local.is_empty() {
        return Err(bad("invalid compatibility target QName"));
    }
    let namespace = ns
        .get(prefix)
        .cloned()
        .ok_or_else(|| bad(format!("unbound compatibility target prefix {prefix}")))?;
    if namespace == MCE_NAMESPACE {
        return Err(bad("compatibility target cannot use the MCE namespace"));
    }
    if local == "*" {
        if wildcard {
            return Ok(NamePattern::Namespace(namespace));
        }
        return Err(bad("wildcard is not allowed in ProcessContent"));
    }
    if !valid_ncname(local) {
        return Err(bad("invalid compatibility target QName"));
    }
    Ok(NamePattern::Exact(ExpandedName {
        namespace,
        local_name: local.into(),
    }))
}
fn valid_ncname(value: &str) -> bool {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    (first == '_' || first.is_alphabetic())
        && chars.all(|ch| ch == '_' || ch == '-' || ch == '.' || ch.is_alphanumeric())
}
#[allow(clippy::too_many_arguments)]
fn write_start(
    o: &mut Vec<u8>,
    q: &str,
    ns: &BTreeMap<String, String>,
    raw: &[(String, String)],
    ign: &HashSet<String>,
    preserve: &HashSet<NamePattern>,
    caps: &MceCapabilities,
    filter: bool,
    rep: &mut MceReport,
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
        if filter && n.namespace == MCE_NAMESPACE {
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
pub(crate) fn process_ooxml(x: &[u8]) -> crate::error::Result<Cow<'_, [u8]>> {
    process_markup_compatibility(x, &MceCapabilities::default(), &MceLimits::default())
        .map(|x| x.xml)
        .map_err(crate::error::OoxmlError::MarkupCompatibility)
}
pub(crate) fn process_part(part: &dyn litchi_opc::Part) -> crate::error::Result<Cow<'_, [u8]>> {
    process_ooxml(part.blob())
}
pub(crate) fn process_part_arc(
    part: &dyn litchi_opc::Part,
) -> crate::error::Result<std::sync::Arc<Vec<u8>>> {
    Ok(match process_part(part)? {
        Cow::Borrowed(_) => part.blob_arc(),
        Cow::Owned(v) => std::sync::Arc::new(v),
    })
}
pub(crate) fn process_str(x: &str) -> crate::error::Result<Cow<'_, str>> {
    match process_ooxml(x.as_bytes())? {
        Cow::Borrowed(_) => Ok(Cow::Borrowed(x)),
        Cow::Owned(v) => String::from_utf8(v).map(Cow::Owned).map_err(|e| {
            crate::error::OoxmlError::InvalidFormat(format!("MCE output is not UTF-8: {e}"))
        }),
    }
}
#[cfg(test)]
mod tests {
    use super::*;
    fn run(x: &str, c: &MceCapabilities) -> R<String> {
        Ok(String::from_utf8(
            process_markup_compatibility(x.as_bytes(), c, &MceLimits::default())?
                .xml
                .into_owned(),
        )
        .unwrap())
    }
    #[test]
    fn fast_borrowed() {
        assert!(matches!(
            process_markup_compatibility(b"<r/>", &MceCapabilities::new(), &MceLimits::default())
                .unwrap()
                .xml,
            Cow::Borrowed(_)
        ))
    }
    #[test]
    fn choice_fallback() {
        let x = r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a="urn:a"><mc:AlternateContent><mc:Choice Requires="a"><yes/></mc:Choice><mc:Fallback><no/></mc:Fallback></mc:AlternateContent></r>"#;
        let mut c = MceCapabilities::new();
        assert!(run(x, &c).unwrap().contains("<no"));
        c.understand_namespace("urn:a");
        assert!(run(x, &c).unwrap().contains("<yes"))
    }
    #[test]
    fn ignore_and_unwrap() {
        let x = r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x" mc:Ignorable="x" mc:ProcessContent="x:w"><x:no/><x:w><yes/></x:w></r>"#;
        let y = run(x, &MceCapabilities::new()).unwrap();
        assert!(!y.contains("<x:"));
        assert!(y.contains("<yes"))
    }
    #[test]
    fn security_and_limits() {
        let x = r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x" mc:MustUnderstand="x"/>"#;
        assert!(matches!(
            run(x, &MceCapabilities::new()),
            Err(MceError::MustUnderstand(_))
        ));
        let l = MceLimits {
            max_depth: 1,
            ..MceLimits::default()
        };
        assert!(process_markup_compatibility(b"<r xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><x/></r>",&MceCapabilities::new(),&l).is_err())
    }
}

#[cfg(test)]
mod fixture_tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    #[test]
    fn poi_styles_select_unsupported_vendor_fallbacks() {
        let package = OpcPackage::from_bytes(include_bytes!(
            "../../../../test-data/poi/test-data/spreadsheet/style-alternate-content.xlsx"
        ))
        .unwrap();
        let part = package
            .get_part(&PackURI::new("/xl/styles.xml").unwrap())
            .unwrap();
        let output = process_markup_compatibility(
            part.blob(),
            &MceCapabilities::default(),
            &MceLimits::default(),
        )
        .unwrap();
        let xml = std::str::from_utf8(output.xml.as_ref()).unwrap();
        assert!(!xml.contains("mc:AlternateContent"));
        assert!(!xml.contains("hs:extension"));
        assert!(output.report.selected_fallbacks > 10);
    }

    #[test]
    fn libreoffice_pptx_emits_only_fallback_shape() {
        let package = OpcPackage::from_bytes(include_bytes!(
            "../../../../test-data/libreoffice-core/oox/qa/unit/data/import-mce.pptx"
        ))
        .unwrap();
        let part = package
            .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
            .unwrap();
        let output = process_markup_compatibility(
            part.blob(),
            &MceCapabilities::default(),
            &MceLimits::default(),
        )
        .unwrap();
        let xml = std::str::from_utf8(output.xml.as_ref()).unwrap();
        assert!(!xml.contains("mc:AlternateContent"));
        assert!(!xml.contains("a14:m"));
        assert!(xml.contains("a:blipFill"));
        assert_eq!(output.report.selected_fallbacks, 1);
    }
}

#[cfg(test)]
mod adapter_tests {
    use crate::docx::enums::WdHeaderFooter;
    use crate::docx::header_footer::HeaderFooter;
    use crate::xlsx::SharedStrings;

    #[test]
    fn docx_header_uses_fallback_without_mutating_raw_xml() {
        let raw = br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><w:p><w:r><w:t>choice</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:hdr>"#;
        let header = HeaderFooter::from_xml_bytes(raw.to_vec(), WdHeaderFooter::Primary);
        assert_eq!(header.xml_bytes(), raw);
        assert_eq!(header.text().unwrap(), "fallback");
        assert_eq!(header.paragraph_count().unwrap(), 1);
    }

    #[test]
    fn xlsx_shared_strings_uses_fallback() {
        let xml = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported" count="1" uniqueCount="1"><mc:AlternateContent><mc:Choice Requires="x"><si><t>choice</t></si></mc:Choice><mc:Fallback><si><t>fallback</t></si></mc:Fallback></mc:AlternateContent></sst>"#;
        let strings = SharedStrings::parse(xml).unwrap();
        assert_eq!(strings.get(0), Some("fallback"));
    }

    #[test]
    fn generic_chart_reader_uses_fallback() {
        let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><x:chart/></mc:Choice><mc:Fallback><c:chart/></mc:Fallback></mc:AlternateContent></c:chartSpace>"#;
        crate::charts::reader::parse_chart(xml.as_slice()).unwrap();
    }

    #[test]
    fn alternate_content_picture_selects_fallback() {
        let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><x:picture/></mc:Choice><mc:Fallback><w:pict><w:t>fallback-picture</w:t></w:pict></mc:Fallback></mc:AlternateContent></w:r>"#;
        let output = super::process_ooxml(xml).unwrap();
        let semantic = std::str::from_utf8(output.as_ref()).unwrap();
        assert!(semantic.contains("w:pict"));
        assert!(!semantic.contains("x:picture"));
        assert!(!semantic.contains("AlternateContent"));
    }
}
