//! Repair reporting for non-structural defects in RTF documents.
//!
//! Two kinds of repair can happen while a document is read, and both are
//! *recorded*, so a caller can enumerate exactly what changed through
//! [`ToleranceReport`]. Nothing is swallowed silently, and nothing outside the
//! closed defect enumerations changes: resource limits, group-nesting depth,
//! allocation failures, and the root `{\rtf` framing remain hard errors under
//! every [`Leniency`] variant.
//!
//! **Word-compatible repairs** are applied under every [`Leniency`] variant,
//! including the default [`Leniency::Strict`]. They cover syntax RTF 1.9.1
//! leaves without reader behavior but real producers emit anyway — syntax
//! Word and LibreOffice read past without complaint, so rejecting it would
//! cost the caller the entire document. Each substitutes the interpretation
//! documented on the corresponding [`SyntaxDefect`], [`ParameterDefect`], or
//! [`UnicodeDefect`] variant.
//!
//! **Opt-in repairs** involve a judgment call — the deviation can hide real
//! corruption — and stay gated by named [`Leniency`] variants.
//!
//! This mirrors the contract `litchi-doc` and `litchi-xls` provide for legacy
//! binary documents so every reader behaves the same way.

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous document must not be able to make the report
/// grow without limit. Defects beyond this bound are counted, not stored; see
/// [`ToleranceReport::unrecorded`].
pub const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural document defects.
///
/// Structural defects — exceeded resource limits, excessive group nesting,
/// allocation failure, or a missing `{\rtf` root — are hard errors under every
/// variant. The Word-compatible repairs are applied and recorded under every
/// variant, including [`Leniency::Strict`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Leniency {
    /// Reject every defect whose repair requires opting in.
    ///
    /// This is the default: Word-compatible repairs are still applied and
    /// recorded in the [`ToleranceReport`].
    #[default]
    Strict,
}

/// A lexical or group-structure defect.
///
/// This enumeration is closed by design: together with [`ParameterDefect`]
/// and [`UnicodeDefect`] it is the exhaustive contract of what a read may
/// change about a document. The recorded [`ToleratedDefect::offset`] is the
/// byte offset of the offending syntax.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SyntaxDefect {
    /// A backslash was followed by a non-alphabetic character RTF 1.9.1 does
    /// not define as a control symbol (for example `\!` or `\;`).
    ///
    /// RTF 1.9.1 tells readers to ignore control symbols they do not
    /// recognize, and Word and LibreOffice do. The symbol is retained as an
    /// inert opaque control, exactly like an unknown control word, and is not
    /// interpreted as text. Repaired under every [`Leniency`] variant.
    UnknownControlSymbol,
    /// Content followed the root group's closing brace.
    ///
    /// Word and LibreOffice stop reading at the root group's closing brace
    /// (LibreOffice's tokenizer returns as soon as the group count reaches
    /// zero). The trailing tokens are ignored. Repaired under every
    /// [`Leniency`] variant.
    TrailingContentAfterRoot,
}

impl SyntaxDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::UnknownControlSymbol => {
                "ignored the undefined control symbol; it is retained as an inert opaque control"
            },
            Self::TrailingContentAfterRoot => {
                "ignored the content after the root group's closing brace"
            },
        }
    }
}

/// A control-word parameter whose shape the format does not allow.
///
/// This enumeration is closed by design. Each repair applies the
/// interpretation Word and LibreOffice give the same bytes, so a conforming
/// document never changes meaning. The recorded [`ToleratedDefect::offset`]
/// is the byte offset of the control word's backslash.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ParameterDefect {
    /// A toggle control carried a parameter other than 0 or 1 (`\hyphpar2`).
    ///
    /// RTF toggles are on unless their parameter is zero; LibreOffice passes
    /// the parameter through and treats every nonzero value as on. The
    /// parameter is normalized to 1.
    NonBinaryToggle,
    /// A flag, symbol, or destination control carried a numeric parameter
    /// (`\page2`).
    ///
    /// LibreOffice's tokenizer documents that "flags ignore any parameter by
    /// definition", and Word does the same. The parameter is dropped. The
    /// emphasis-mark (`\acc*`) and associated-underline (`\aul*`) selectors,
    /// which Word and LibreOffice treat as toggles, are turned off by a zero
    /// parameter instead.
    UnexpectedParameter,
    /// A value control that requires a numeric parameter had none (`\fs`).
    ///
    /// LibreOffice substitutes the keyword's default value when the parameter
    /// is absent. The documented default for the control is substituted
    /// (zero unless the control defines another default).
    MissingParameter,
}

impl ParameterDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::NonBinaryToggle => "treated the nonzero toggle parameter as on",
            Self::UnexpectedParameter => {
                "ignored the parameter the control does not define (a zero turns a toggled selector off)"
            },
            Self::MissingParameter => "substituted the control's default parameter",
        }
    }
}

/// A Unicode escape the reader decoded lossily.
///
/// This enumeration is closed by design. The recorded
/// [`ToleratedDefect::offset`] is the byte offset of the escape sequence.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum UnicodeDefect {
    /// A `\uN` sequence contained a UTF-16 surrogate without its partner.
    ///
    /// Word and LibreOffice substitute U+FFFD REPLACEMENT CHARACTER, as the
    /// Unicode lossy-conversion rules prescribe, and keep reading. One defect
    /// is recorded per substituted surrogate.
    UnpairedSurrogate,
}

impl UnicodeDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::UnpairedSurrogate => "substituted U+FFFD for the unpaired UTF-16 surrogate",
        }
    }
}

/// One class of defect repaired while a document was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// a read may change about a document.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A lexical or group-structure defect.
    Syntax(SyntaxDefect),
    /// A control-word parameter defect.
    Parameter(ParameterDefect),
    /// A lossily decoded Unicode escape.
    Unicode(UnicodeDefect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::Syntax(defect) => defect.repair(),
            Self::Parameter(defect) => defect.repair(),
            Self::Unicode(defect) => defect.repair(),
        }
    }
}

impl From<SyntaxDefect> for Defect {
    fn from(defect: SyntaxDefect) -> Self {
        Self::Syntax(defect)
    }
}

impl From<ParameterDefect> for Defect {
    fn from(defect: ParameterDefect) -> Self {
        Self::Parameter(defect)
    }
}

impl From<UnicodeDefect> for Defect {
    fn from(defect: UnicodeDefect) -> Self {
        Self::Unicode(defect)
    }
}

/// One repaired defect, with enough context to locate it in the file.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    /// Which defect class was repaired.
    pub defect: Defect,
    /// Byte offset of the offending syntax in the RTF source.
    ///
    /// For compressed RTF the offset is into the decompressed RTF stream.
    pub offset: u64,
}

/// Everything a read repaired, in the order the defects were found.
///
/// An empty report means the document conformed and no substitution was made.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

impl ToleranceReport {
    /// Whether the document parsed without any repair.
    #[inline]
    #[must_use]
    pub fn is_clean(&self) -> bool {
        self.defects.is_empty() && self.unrecorded == 0
    }

    /// The individually recorded defects, in discovery order.
    #[inline]
    #[must_use]
    pub fn defects(&self) -> &[ToleratedDefect] {
        &self.defects
    }

    /// Iterate over the individually recorded defects, in discovery order.
    #[inline]
    pub fn iter(&self) -> std::slice::Iter<'_, ToleratedDefect> {
        self.defects.iter()
    }

    /// Defects that were repaired but not stored, once the bound was reached.
    #[inline]
    #[must_use]
    pub fn unrecorded(&self) -> u32 {
        self.unrecorded
    }

    /// Total repairs performed, recorded or not.
    #[inline]
    #[must_use]
    pub fn total(&self) -> u64 {
        self.defects.len() as u64 + u64::from(self.unrecorded)
    }

    /// Record one repair, saturating at [`MAX_RECORDED_DEFECTS`] stored entries.
    pub(crate) fn record(&mut self, defect: impl Into<Defect>, offset: usize) {
        self.push(ToleratedDefect {
            defect: defect.into(),
            offset: offset as u64,
        });
    }

    fn push(&mut self, tolerated: ToleratedDefect) {
        if self.defects.len() < MAX_RECORDED_DEFECTS {
            self.defects.push(tolerated);
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }

    /// Merge another phase's report into this one, keeping discovery order.
    ///
    /// The stored-entry bound applies to the merged report as a whole; entries
    /// that no longer fit join the unrecorded count.
    pub(crate) fn append(&mut self, other: Self) {
        for tolerated in other.defects {
            self.push(tolerated);
        }
        self.unrecorded = self.unrecorded.saturating_add(other.unrecorded);
    }

    /// Translate offsets into the parser's one-scalar-per-byte transport view
    /// back into byte offsets of the original transport.
    ///
    /// Non-ASCII RTF bytes are tokenized through a Latin-1 view in which every
    /// byte at or above 0x80 occupies two UTF-8 bytes; the original byte
    /// offset is the number of scalars before the recorded position. The scan
    /// visits the view once regardless of how many defects were recorded.
    pub(crate) fn remap_latin1_offsets(&mut self, view: &str) {
        if self.defects.is_empty() {
            return;
        }
        let mut targets: Vec<u64> = self.defects.iter().map(|d| d.offset).collect();
        targets.sort_unstable();
        targets.dedup();
        let mut mapped = Vec::with_capacity(targets.len());
        let mut pending = targets.iter().copied().peekable();
        let mut scalars = 0u64;
        for (index, _) in view.char_indices() {
            while let Some(target) = pending.next_if(|target| *target <= index as u64) {
                mapped.push((target, scalars));
            }
            if pending.peek().is_none() {
                break;
            }
            scalars += 1;
        }
        for target in pending {
            mapped.push((target, scalars));
        }
        for tolerated in &mut self.defects {
            if let Ok(found) = mapped.binary_search_by_key(&tolerated.offset, |(from, _)| *from)
                && let Some((_, to)) = mapped.get(found)
            {
                tolerated.offset = *to;
            }
        }
    }
}

impl<'a> IntoIterator for &'a ToleranceReport {
    type Item = &'a ToleratedDefect;
    type IntoIter = std::slice::Iter<'a, ToleratedDefect>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strict_is_the_default() {
        assert_eq!(Leniency::default(), Leniency::Strict);
    }

    #[test]
    fn a_fresh_report_is_clean() {
        let report = ToleranceReport::default();
        assert!(report.is_clean());
        assert_eq!(report.total(), 0);
        assert!(report.defects().is_empty());
        assert_eq!(report.iter().count(), 0);
    }

    #[test]
    fn recorded_defects_are_returned_in_discovery_order() {
        let mut report = ToleranceReport::default();
        report.record(SyntaxDefect::UnknownControlSymbol, 15);
        report.record(ParameterDefect::NonBinaryToggle, 3);
        report.record(UnicodeDefect::UnpairedSurrogate, 42);

        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert_eq!(
            (&report)
                .into_iter()
                .map(|d| (d.defect, d.offset))
                .collect::<Vec<_>>(),
            vec![
                (Defect::Syntax(SyntaxDefect::UnknownControlSymbol), 15),
                (Defect::Parameter(ParameterDefect::NonBinaryToggle), 3),
                (Defect::Unicode(UnicodeDefect::UnpairedSurrogate), 42),
            ]
        );
    }

    /// A hostile document must not be able to grow the report without limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for offset in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(SyntaxDefect::UnknownControlSymbol, offset);
        }
        for _ in 0..3 {
            report.record(ParameterDefect::MissingParameter, 7);
        }

        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 8);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 8);
    }

    #[test]
    fn appended_reports_merge_within_the_bound() {
        let mut base = ToleranceReport::default();
        base.record(SyntaxDefect::UnknownControlSymbol, 7);

        let mut parsed = ToleranceReport::default();
        for offset in 0..MAX_RECORDED_DEFECTS {
            parsed.record(ParameterDefect::UnexpectedParameter, offset);
        }
        assert_eq!(parsed.unrecorded(), 0);

        base.append(parsed);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 1);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 1);
        assert_eq!(
            base.defects().first(),
            Some(&ToleratedDefect {
                defect: Defect::Syntax(SyntaxDefect::UnknownControlSymbol),
                offset: 7,
            })
        );

        let mut clean = ToleranceReport::default();
        clean.append(ToleranceReport::default());
        assert!(clean.is_clean());
    }

    #[test]
    fn latin1_view_offsets_map_back_to_transport_bytes() {
        // Transport bytes: `a`, 0xE9, `b`, 0xFF, `c`; each high byte occupies
        // two UTF-8 bytes in the view.
        let view = "a\u{e9}b\u{ff}c";
        let mut report = ToleranceReport::default();
        report.record(SyntaxDefect::UnknownControlSymbol, 6); // `c`
        report.record(SyntaxDefect::UnknownControlSymbol, 0); // `a`
        report.record(SyntaxDefect::UnknownControlSymbol, 3); // `b`
        report.record(SyntaxDefect::UnknownControlSymbol, 7); // end of input
        report.remap_latin1_offsets(view);
        assert_eq!(
            report.iter().map(|d| d.offset).collect::<Vec<_>>(),
            vec![4, 0, 2, 5]
        );
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        let defects = [
            Defect::from(SyntaxDefect::UnknownControlSymbol),
            Defect::from(SyntaxDefect::TrailingContentAfterRoot),
            Defect::from(ParameterDefect::NonBinaryToggle),
            Defect::from(ParameterDefect::UnexpectedParameter),
            Defect::from(ParameterDefect::MissingParameter),
            Defect::from(UnicodeDefect::UnpairedSurrogate),
        ];
        for defect in defects {
            assert!(!defect.repair().is_empty());
        }
        assert_eq!(
            Defect::Syntax(SyntaxDefect::UnknownControlSymbol).repair(),
            SyntaxDefect::UnknownControlSymbol.repair()
        );
        assert_eq!(
            Defect::Parameter(ParameterDefect::MissingParameter).repair(),
            ParameterDefect::MissingParameter.repair()
        );
        assert_eq!(
            Defect::Unicode(UnicodeDefect::UnpairedSurrogate).repair(),
            UnicodeDefect::UnpairedSurrogate.repair()
        );
    }
}
