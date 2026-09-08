//! Bounded proof that a worksheet contains no web-extension bindings.
//!
//! The compaction pass already has to consume and successfully emit every XML
//! event.  This observer records only the finite state needed to decide
//! whether the later full web reader can be replaced by the known empty
//! result.  Any condition outside that proof boundary leaves the observer
//! ineligible and [`Probe::finish`] delegates to the unchanged reader.

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::Result;
use crate::web::Bindings;

use super::{MAX_DEPTH, MAX_XML_BYTES, read};

/// A finite, success-only proof accumulated while compacted XML is emitted.
///
/// `eligible` starts true and is cleared whenever this observer cannot prove
/// that the unchanged web reader would return an empty collection.  It never
/// constructs a replacement error: an ineligible proof always falls back to
/// [`super::read`] at [`Probe::finish`].
#[derive(Debug)]
pub(crate) struct Probe {
    depth: usize,
    saw_root: bool,
    closed_root: bool,
    eligible: bool,
}

impl Default for Probe {
    fn default() -> Self {
        Self {
            depth: 0,
            saw_root: false,
            closed_root: false,
            eligible: true,
        }
    }
}

impl Probe {
    /// Observe one event after the compaction writer has emitted it.
    ///
    /// The caller deliberately does not invoke this for whitespace text that
    /// compaction discards or for `Eof`.  Keeping those events out of the
    /// proof avoids claiming that an event was observed when it was not part
    /// of the published output.  `Eof` is represented by the final root and
    /// depth checks in [`Self::finish`].
    pub(crate) fn observe(&mut self, reader: &NsReader<&[u8]>, event: &Event<'_>) {
        if !self.eligible {
            return;
        }
        match event {
            Event::Start(element) => self.observe_start(reader, element, false),
            Event::Empty(element) => self.observe_start(reader, element, true),
            Event::End(element) => self.observe_end(reader, element.name()),
            Event::Text(text) => {
                if text.decode().is_err() {
                    self.eligible = false;
                }
            },
            Event::CData(text) => {
                if text.decode().is_err() {
                    self.eligible = false;
                }
            },
            Event::DocType(_) => self.eligible = false,
            Event::Decl(_) | Event::PI(_) | Event::Comment(_) | Event::GeneralRef(_) => {},
            Event::Eof => {},
        }
    }

    /// Finish the proof against the exact compacted bytes.
    ///
    /// A successful proof can only produce the empty binding collection.  All
    /// other states use the original reader so its errors, messages, limits,
    /// and phase ordering remain authoritative.
    pub(crate) fn finish(self, output: &[u8]) -> Result<Bindings> {
        if self.proven(output.len()) {
            Ok(Bindings::default())
        } else {
            read(output)
        }
    }

    fn observe_start(&mut self, reader: &NsReader<&[u8]>, element: &BytesStart<'_>, empty: bool) {
        if element.attributes_raw().contains(&b'\'') {
            self.eligible = false;
            return;
        }

        let Some((namespace, local)) = classify_name(reader, element.name()) else {
            self.eligible = false;
            return;
        };
        if is_web_name(namespace, local) {
            self.eligible = false;
            return;
        }

        if self.depth == 0 {
            if self.saw_root
                || self.closed_root
                || namespace != NamespaceKind::Spreadsheet
                || local != LocalName::Worksheet
            {
                self.eligible = false;
                return;
            }
            self.saw_root = true;
            if empty {
                self.closed_root = true;
                return;
            }
        } else if !self.saw_root || self.closed_root {
            self.eligible = false;
            return;
        }

        if !empty {
            self.depth = match self.depth.checked_add(1) {
                Some(depth) if depth <= MAX_DEPTH => depth,
                _ => {
                    self.eligible = false;
                    return;
                },
            };
        }
    }

    fn observe_end(&mut self, reader: &NsReader<&[u8]>, name: QName<'_>) {
        let Some((namespace, local)) = classify_name(reader, name) else {
            self.eligible = false;
            return;
        };
        if is_web_name(namespace, local) {
            self.eligible = false;
            return;
        }

        let Some(depth) = self.depth.checked_sub(1) else {
            self.eligible = false;
            return;
        };
        self.depth = depth;
        if depth == 0 {
            if !self.saw_root
                || self.closed_root
                || namespace != NamespaceKind::Spreadsheet
                || local != LocalName::Worksheet
            {
                self.eligible = false;
            } else {
                self.closed_root = true;
            }
        }
    }

    fn proven(&self, output_len: usize) -> bool {
        self.eligible
            && self.saw_root
            && self.closed_root
            && self.depth == 0
            && output_len <= MAX_XML_BYTES
    }

    #[cfg(test)]
    pub(crate) fn is_proven(&self) -> bool {
        self.proven(0)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Spreadsheet,
    X15,
    Xm,
    Other,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum LocalName {
    Worksheet,
    ExtLst,
    WebExtensions,
    WebExtension,
    F,
    Other,
}

fn classify_name(reader: &NsReader<&[u8]>, name: QName<'_>) -> Option<(NamespaceKind, LocalName)> {
    let (namespace, local) = reader.resolver().resolve_element(name);
    let namespace = match namespace {
        ResolveResult::Bound(namespace) => namespace.into_inner(),
        ResolveResult::Unbound => &[],
        ResolveResult::Unknown(_) => return None,
    };
    let namespace = if super::is_sml(namespace) {
        NamespaceKind::Spreadsheet
    } else if namespace == super::X15_NAMESPACE.as_bytes() {
        NamespaceKind::X15
    } else if namespace == super::XM_NAMESPACE.as_bytes() {
        NamespaceKind::Xm
    } else {
        NamespaceKind::Other
    };
    let local = match local.as_ref() {
        b"worksheet" => LocalName::Worksheet,
        b"extLst" => LocalName::ExtLst,
        b"webExtensions" => LocalName::WebExtensions,
        b"webExtension" => LocalName::WebExtension,
        b"f" => LocalName::F,
        _ => LocalName::Other,
    };
    Some((namespace, local))
}

fn is_web_name(namespace: NamespaceKind, local: LocalName) -> bool {
    matches!(
        (namespace, local),
        (NamespaceKind::Spreadsheet, LocalName::ExtLst)
            | (NamespaceKind::X15, LocalName::WebExtensions)
            | (NamespaceKind::X15, LocalName::WebExtension)
            | (NamespaceKind::Xm, LocalName::F)
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    fn worksheet(body: &str) -> Vec<u8> {
        format!(
            r#"<worksheet xmlns="{}">{body}</worksheet>"#,
            super::super::SML
        )
        .into_bytes()
    }

    fn compact_probe(input: &[u8]) -> (Vec<u8>, Probe) {
        let output = crate::raw::compact::changed(input, "test worksheet XML").unwrap();
        let mut reader = NsReader::from_reader(input);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut probe = Probe::default();
        // Every fixture passed here is deliberately free of formatting-only
        // text, so the observed stream is the stream emitted by compaction.
        loop {
            let event = reader
                .read_event()
                .expect("test worksheet XML should parse");
            if matches!(&event, Event::Eof) {
                break;
            }
            probe.observe(&reader, &event);
        }
        (output, probe)
    }

    fn assert_result_parity(actual: Result<Bindings>, expected: Result<Bindings>) {
        match (actual, expected) {
            (Ok(actual), Ok(expected)) => assert_eq!(actual, expected),
            (Err(actual), Err(expected)) => {
                assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
                assert_eq!(actual.to_string(), expected.to_string());
            },
            (actual, expected) => {
                panic!("probe and full web reader disagree: {actual:?} vs {expected:?}");
            },
        }
    }

    fn assert_compact_case(input: &[u8], expected_proven: bool) {
        let (output, probe) = compact_probe(input);
        assert_eq!(probe.is_proven(), expected_proven);
        let actual = probe.finish(&output);
        let expected = read(&output);
        assert_result_parity(actual, expected);
    }

    #[test]
    fn proves_ordinary_strict_alias_and_rebound_namespaces() {
        let transitional = worksheet("<sheetData><row r=\"1\"/></sheetData>");
        assert_compact_case(&transitional, true);

        let strict = format!(
            r#"<worksheet xmlns="{}"><sheetData><row r="1"/></sheetData></worksheet>"#,
            super::super::STRICT_SML
        );
        assert_compact_case(strict.as_bytes(), true);

        let alias = format!(
            r#"<s:worksheet xmlns:s="{}"><s:sheetData><s:row r="1"/></s:sheetData></s:worksheet>"#,
            super::super::SML
        );
        assert_compact_case(alias.as_bytes(), true);

        let rebound = format!(
            r#"<s:worksheet xmlns:s="{}"><item xmlns:s="urn:rebound"><s:value/></item></s:worksheet>"#,
            super::super::SML
        );
        assert_compact_case(rebound.as_bytes(), true);
    }

    #[test]
    fn falls_back_for_every_web_name_and_structural_form() {
        let x15 = super::super::X15_NAMESPACE;
        let xm = super::super::XM_NAMESPACE;
        let cases = [
            worksheet("<extLst></extLst>"),
            worksheet("<extLst/>"),
            worksheet(&format!(
                r#"<outer><inner><x15:webExtensions xmlns:x15="{x15}"></x15:webExtensions></inner></outer>"#
            )),
            worksheet(&format!(
                r#"<outer><inner><x15:webExtensions xmlns:x15="{x15}"/></inner></outer>"#
            )),
            worksheet(&format!(
                r#"<outer><inner><container><x15:webExtension xmlns:x15="{x15}"></x15:webExtension></container></inner></outer>"#
            )),
            worksheet(&format!(
                r#"<outer><inner><container><x15:webExtension xmlns:x15="{x15}"/></container></inner></outer>"#
            )),
            worksheet(&format!(
                r#"<outer><inner><container><holder><xm:f xmlns:xm="{xm}"></xm:f></holder></container></inner></outer>"#
            )),
            worksheet(&format!(
                r#"<outer><inner><container><holder><xm:f xmlns:xm="{xm}"/></holder></container></inner></outer>"#
            )),
            worksheet("<outer><extLst/></outer>"),
        ];

        for input in cases {
            assert_compact_case(&input, false);
        }
    }

    #[test]
    fn preserves_full_reader_behavior_for_unknown_or_unfinished_structure() {
        let unknown = worksheet("<unknown:item xmlns:unknown=\"urn:unknown\"/>");
        assert_compact_case(&unknown, true);

        let unknown_prefix = worksheet("<unknown:item/>");
        assert_compact_case(&unknown_prefix, false);

        let missing_root = format!(r#"<other xmlns="{}"/>"#, super::super::SML);
        assert_compact_case(missing_root.as_bytes(), false);

        let unclosed = format!(r#"<worksheet xmlns="{}">"#, super::super::SML);
        assert_compact_case(unclosed.as_bytes(), false);

        // The compactor deliberately checks matching end names and therefore
        // cannot emit this malformed tail.  Exercise the observer directly
        // with the same namespace reader contract to keep orphan ends inside
        // the fallback boundary as well.
        let orphan = format!(
            r#"<worksheet xmlns="{}"></worksheet></worksheet>"#,
            super::super::SML
        );
        let mut reader = NsReader::from_reader(orphan.as_bytes());
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = false;
        reader.config_mut().allow_unmatched_ends = true;
        let mut probe = Probe::default();
        loop {
            let event = reader
                .read_event()
                .expect("orphan test XML should tokenize");
            if matches!(&event, Event::Eof) {
                break;
            }
            probe.observe(&reader, &event);
        }
        assert!(!probe.is_proven());
        assert!(read(orphan.as_bytes()).is_err());
    }

    #[test]
    fn decodes_text_and_cdata_and_allows_safe_pi_declaration_and_references() {
        let input = format!(
            r#"<?xml version="1.0"?><worksheet xmlns="{}"><value>plain&amp;value&custom;</value><![CDATA[cdata]]><?pi value?></worksheet>"#,
            super::super::SML
        );
        assert_compact_case(input.as_bytes(), true);

        let mut malformed =
            format!(r#"<worksheet xmlns="{}"><value>"#, super::super::SML).into_bytes();
        malformed.push(0xff);
        malformed.extend_from_slice(b"</value></worksheet>");
        assert_compact_case(&malformed, false);

        let mut malformed_cdata = format!(
            r#"<worksheet xmlns="{}"><value><![CDATA["#,
            super::super::SML
        )
        .into_bytes();
        malformed_cdata.push(0xff);
        malformed_cdata.extend_from_slice(b"]]></value></worksheet>");
        assert_compact_case(&malformed_cdata, false);

        let dtd = format!(
            r#"<!DOCTYPE worksheet><worksheet xmlns="{}"/>"#,
            super::super::SML
        );
        assert_compact_case(dtd.as_bytes(), false);
    }

    #[test]
    fn apostrophe_attributes_force_fallback_before_normalization() {
        let quoted = format!(
            "<worksheet xmlns='{}'><item/></worksheet>",
            super::super::SML
        );
        assert_compact_case(quoted.as_bytes(), false);

        let value = format!(
            r#"<worksheet xmlns="{}"><item value="it's"/></worksheet>"#,
            super::super::SML
        );
        assert_compact_case(value.as_bytes(), false);
    }

    #[test]
    fn enforces_root_depth_and_output_size_at_finish() {
        let mut deep = format!(r#"<worksheet xmlns="{}">"#, super::super::SML);
        for _ in 0..MAX_DEPTH {
            deep.push_str("<node>");
        }
        for _ in 0..MAX_DEPTH {
            deep.push_str("</node>");
        }
        deep.push_str("</worksheet>");
        assert_compact_case(deep.as_bytes(), false);

        let (output, probe) = compact_probe(&worksheet(""));
        assert!(probe.is_proven());
        assert!(output.len() <= MAX_XML_BYTES);
        let oversized = vec![b' '; MAX_XML_BYTES + 1];
        assert_result_parity(probe.finish(&oversized), read(&oversized));
    }
}
