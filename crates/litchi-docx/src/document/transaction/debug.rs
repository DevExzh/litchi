//! Content-free diagnostics for the main-document transaction family.

use super::*;
use std::fmt;

impl fmt::Debug for Snapshot {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Snapshot")
            .field("xml_bytes", &self.xml.bytes().len())
            .field("paragraphs", &self.paragraphs.len())
            .field("tables", &self.tables.len())
            .field("managed", &self.admission.is_some())
            .finish_non_exhaustive()
    }
}

impl fmt::Debug for Patch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Patch")
            .field("source_bytes", &self.before.xml.bytes().len())
            .field("target_bytes", &self.after.xml.bytes().len())
            .field("operations", &self.operations.as_slice().len())
            .finish_non_exhaustive()
    }
}

impl fmt::Debug for Edit {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Edit")
            .field("operations", &self.operations.len())
            .field("replacement_text_bytes", &self.replacement_text_bytes)
            .field("compaction", &self.compaction)
            .finish_non_exhaustive()
    }
}

impl fmt::Debug for ParagraphTextReplacement {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("ParagraphTextReplacement")
            .field("position", &self.position)
            .field("text_bytes", &self.text.len())
            .finish()
    }
}

impl fmt::Debug for HyperlinkTextReplacement {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("HyperlinkTextReplacement")
            .field("address", &self.address)
            .field("text_bytes", &self.text.len())
            .finish()
    }
}

impl fmt::Debug for ParagraphTransfer {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("ParagraphTransfer")
            .field("xml_bytes", &self.fragment.len())
            .field("parts", &self.graph.parts.len())
            .finish_non_exhaustive()
    }
}

impl fmt::Debug for GraphChange {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("GraphChange")
            .field("relationships", &self.relationships.len())
            .field("parts", &self.parts.len())
            .field("replaced", &self.replaced.len())
            .finish()
    }
}

impl fmt::Debug for TransferGraph {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("TransferGraph")
            .field("relationships", &self.main_relationships.len())
            .field("parts", &self.parts.len())
            .finish()
    }
}

impl fmt::Debug for Operation {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let kind = match self {
            Self::ReplaceParagraphText { .. } => "ReplaceParagraphText",
            Self::ReplaceHyperlinkText { .. } => "ReplaceHyperlinkText",
            Self::ReplaceRunText { .. } => "ReplaceRunText",
            Self::ReplaceRunTextNode { .. } => "ReplaceRunTextNode",
            Self::MarkCachedRunFieldDirty { .. } => "MarkCachedRunFieldDirty",
            Self::ReplaceSimpleFieldText { .. } => "ReplaceSimpleFieldText",
            Self::ReplaceComplexFieldText { .. } => "ReplaceComplexFieldText",
            Self::ReplaceRevisionText { .. } => "ReplaceRevisionText",
            Self::ApplyRevision { .. } => "ApplyRevision",
            Self::ReplaceContentControlText { .. } => "ReplaceContentControlText",
            Self::ReplaceNestedContentControlText { .. } => "ReplaceNestedContentControlText",
            Self::ReplaceNestedContentControlHyperlinkText { .. } => {
                "ReplaceNestedContentControlHyperlinkText"
            },
            Self::ReplaceBlockContentControlParagraphText { .. } => {
                "ReplaceBlockContentControlParagraphText"
            },
            Self::ReplaceBlockContentControlParagraphHyperlinkText { .. } => {
                "ReplaceBlockContentControlParagraphHyperlinkText"
            },
            Self::ReplaceCellText { .. } => "ReplaceCellText",
            Self::ReplaceCellParagraphText { .. } => "ReplaceCellParagraphText",
            Self::ReplaceNestedCellParagraphText { .. } => "ReplaceNestedCellParagraphText",
            Self::ReplaceNestedCellParagraphHyperlinkText { .. } => {
                "ReplaceNestedCellParagraphHyperlinkText"
            },
            Self::ReplaceRunProperties { .. } => "ReplaceRunProperties",
            Self::ReplaceParagraphProperties { .. } => "ReplaceParagraphProperties",
            Self::InsertParagraphFragment { .. } => "InsertParagraphFragment",
            Self::RemoveParagraphFragment { .. } => "RemoveParagraphFragment",
            Self::InsertRun { .. } => "InsertRun",
            Self::RemoveRun { .. } => "RemoveRun",
            Self::ReplaceElement { .. } => "ReplaceElement",
            Self::InsertElement { .. } => "InsertElement",
            Self::RemoveElement { .. } => "RemoveElement",
            Self::ChangeGraph { .. } => "ChangeGraph",
            Self::InsertParagraph { .. } => "InsertParagraph",
            Self::RemoveParagraph { .. } => "RemoveParagraph",
            Self::InsertTransferredParagraph { .. } => "InsertTransferredParagraph",
            Self::RemoveTransferredParagraph { .. } => "RemoveTransferredParagraph",
        };
        f.debug_struct(kind).finish_non_exhaustive()
    }
}

impl fmt::Debug for TransferRefusal {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let kind = match self {
            Self::MissingDonorRelationship(_) => "MissingDonorRelationship",
            Self::MissingEquivalentDependency { .. } => "MissingEquivalentDependency",
            Self::AmbiguousEquivalentDependency { .. } => "AmbiguousEquivalentDependency",
            Self::InvalidParagraphXml => "InvalidParagraphXml",
            Self::MarkupCompatibilityScope => "MarkupCompatibilityScope",
        };
        f.debug_struct(kind).finish_non_exhaustive()
    }
}

impl fmt::Debug for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::OutOfBounds { position, len } => f
                .debug_struct("OutOfBounds")
                .field("position", position)
                .field("len", len)
                .finish(),
            Self::Refused { position, reason } => f
                .debug_struct("Refused")
                .field("position", position)
                .field("reason", reason)
                .finish(),
            Self::Limit {
                kind,
                maximum,
                observed,
            } => f
                .debug_struct("Limit")
                .field("kind", kind)
                .field("observed", observed)
                .field("maximum", maximum)
                .finish(),
            Self::Document(_) => f.debug_struct("Document").finish_non_exhaustive(),
            Self::Durable(_) => f.debug_struct("Durable").finish_non_exhaustive(),
            Self::Composition(_) => f.debug_struct("Composition").finish_non_exhaustive(),
            Self::InvalidDurable(_) => f.debug_struct("InvalidDurable").finish_non_exhaustive(),
            Self::Transfer(reason) => f.debug_tuple("Transfer").field(reason).finish(),
            Self::PatchConflict => f.write_str("PatchConflict"),
            Self::SemanticPrecondition => f.write_str("SemanticPrecondition"),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn redacted(value: &impl fmt::Debug) {
        for rendered in [format!("{value:?}"), format!("{value:#?}")] {
            assert!(!rendered.contains("SECRET"), "{rendered}");
            assert!(!rendered.contains("83, 69, 67, 82, 69, 84"), "{rendered}");
            assert!(
                rendered.len() < 2048,
                "diagnostic expanded document content"
            );
        }
    }

    #[test]
    fn debug_redacts_document_operations_and_prepared_identifiers() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>SECRET source</w:t></w:r></w:p></w:body></w:document>"#;
        let snapshot = Snapshot::from_xml(xml.to_vec()).unwrap();
        redacted(&snapshot);
        let mut edit = snapshot.edit();
        edit.replace_paragraph_text(Position::new(0), "SECRET replacement")
            .unwrap();
        redacted(&edit);
        let commit = edit.commit().unwrap();
        redacted(&commit);
        redacted(commit.patch());
        redacted(&commit.patch().inverse());
        for operation in commit.patch().operations() {
            redacted(operation);
        }
        let mut edit = snapshot.edit();
        edit.replace_paragraph_text(Position::new(0), "SECRET replacement")
            .unwrap();
        let prepared = edit
            .prepare(CompositionLimits::new(8, 8, 32, 8), "SECRET identifier")
            .unwrap();
        redacted(&prepared);
        redacted(&ParagraphTextReplacement::new(Position::new(0), "SECRET"));
        redacted(&HyperlinkTextReplacement::new(
            ParagraphHyperlinkAddress::new(Position::new(0), Position::new(0)),
            "SECRET",
        ));
    }

    #[test]
    fn debug_redacts_transfer_payloads_and_error_details() {
        let graph = TransferGraph {
            main_relationships: vec![TransferRelationship {
                id: "SECRET id".into(),
                relationship_type: "SECRET type".into(),
                target: "SECRET target".into(),
                external: true,
            }]
            .into(),
            parts: vec![TransferPart {
                name: "SECRET part".into(),
                content_type: "SECRET content type".into(),
                blob: Arc::new(b"SECRET payload".to_vec()),
                relationships: vec![].into(),
            }]
            .into(),
        };
        redacted(&graph);
        let transfer = ParagraphTransfer::new(
            Arc::new(b"SECRET target".to_vec()),
            b"SECRET fragment".to_vec(),
            "SECRET digest".into(),
            "SECRET inverse".into(),
            graph,
        );
        redacted(&transfer);
        redacted(&Error::InvalidDurable("SECRET detail".into()));
        redacted(&Error::Transfer(
            TransferRefusal::MissingEquivalentDependency {
                relationship_type: "SECRET type".into(),
                target: "SECRET target".into(),
            },
        ));
    }
}
