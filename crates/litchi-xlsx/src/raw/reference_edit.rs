//! Lossless package-part surgery for local worksheet-name dependencies.
//!
//! The namespace-aware scanner lives in the shared OOXML crate so XLSB can
//! share its bounded XML and formula behavior. This module retains the XLSX
//! error context and identity policy at the format boundary.

use litchi_ooxml_common::xml::sheet_name::{self, BlockReason, Error as SheetNameError};

use crate::error::{Error, RenameBlock, Result};

/// One simultaneous source-to-final sheet spelling.
pub(crate) use litchi_ooxml_common::xml::sheet_name::Rename;

/// Semantic error context retained across the physical part boundary.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Context<'a> {
    pub(crate) sheet: &'a str,
    pub(crate) position: usize,
    pub(crate) part: &'a str,
    pub(crate) sheet_titles: &'a [&'a str],
}

/// Rewrite one XML part, preserving the XLSX leaf error context.
pub(crate) fn rewrite(
    content: &[u8],
    renames: &[Rename<'_>],
    context: Context<'_>,
) -> Result<Option<Vec<u8>>> {
    sheet_name::rewrite(
        content,
        renames,
        context.sheet_titles,
        crate::sheet::equivalent,
    )
    .map_err(|error| map_error(error, context))
}

fn map_error(error: SheetNameError, context: Context<'_>) -> Error {
    match error {
        SheetNameError::Xml(error) => Error::Xml(error),
        SheetNameError::Invalid(message) => Error::Invalid(message),
        SheetNameError::Blocked(reason) => Error::RenameBlocked {
            sheet: context.sheet.to_owned(),
            position: context.position,
            part: context.part.to_owned(),
            reason: match reason {
                BlockReason::MarkupCompatibility => RenameBlock::MarkupCompatibility,
                BlockReason::UnmodeledReference => RenameBlock::UnmodeledReference,
            },
        },
        SheetNameError::Allocation { resource, source } => Error::Allocation { resource, source },
        error => Error::SheetNameDependency(error),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    const EXTENDED_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    const PROPERTY_TYPES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
    const MAX_REFERENCE_DEPTH: usize = 256;

    fn context<'a>(part: &'a str) -> Context<'a> {
        Context {
            sheet: "Data",
            position: 0,
            part,
            sheet_titles: &["Data"],
        }
    }

    #[test]
    fn rewrites_formula_hyperlink_pivot_and_title_dependencies() {
        let source = format!(
            r#"<root xmlns:s="{S}" xmlns:ep="{}" xmlns:vt="{}"><s:f>Data!A1</s:f><s:formula1>'Data'!Named</s:formula1><s:hyperlink location="Data!$A$1" keep="x"/><s:worksheetSource sheet="data" ref="A1"/><s:dataRef sheet="Data" ref="A1"/><ep:TitlesOfParts><vt:vector><vt:lpstr>Data</vt:lpstr><vt:lpstr>Data</vt:lpstr><vt:lpstr>Data!Print_Area</vt:lpstr></vt:vector></ep:TitlesOfParts></root>"#,
            EXTENDED_PROPERTIES, PROPERTY_TYPES
        );
        let changed = rewrite(
            source.as_bytes(),
            &[Rename {
                before: "Data",
                after: "Input 2026",
            }],
            context("/xl/test.xml"),
        )
        .expect("rewrite")
        .expect("changed output");
        let changed = std::str::from_utf8(&changed).expect("UTF-8");
        assert!(changed.contains("<s:f>&apos;Input 2026&apos;!A1</s:f>"));
        assert!(changed.contains("<s:formula1>&apos;Input 2026&apos;!Named</s:formula1>"));
        assert!(changed.contains("location=\"&apos;Input 2026&apos;!$A$1\""));
        assert_eq!(changed.matches("sheet=\"Input 2026\"").count(), 2);
        assert!(changed.contains("<vt:lpstr>Input 2026</vt:lpstr>"));
        assert!(changed.contains("<vt:lpstr>Data</vt:lpstr>"));
        assert!(changed.contains("<vt:lpstr>&apos;Input 2026&apos;!Print_Area</vt:lpstr>"));
    }

    #[test]
    fn leaves_external_and_inert_values_exact() {
        let source = format!(
            r#"<root xmlns:s="{S}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><s:f>[1]Data!A1</s:f><value>Data!A1</value><s:f>Other!A1</s:f><s:worksheetSource sheet="Data" ref="A1" r:id="external-book"/><s:dataRef sheet="Data" ref="A1" r:id="external-data"/><s:hyperlink ref="A1" location="Data!A1" r:id="external-link"/></root>"#
        );
        let changed = rewrite(
            source.as_bytes(),
            &[Rename {
                before: "Data",
                after: "Input",
            }],
            context("/xl/test.xml"),
        )
        .expect("scan");
        assert!(changed.is_none());
    }

    #[test]
    fn blocks_unmodeled_formula_like_and_compatibility_references() {
        let rename = [Rename {
            before: "Data",
            after: "Input",
        }];
        let unknown = rewrite(
            br#"<root><futureFormulaCache>Data!A1</futureFormulaCache></root>"#,
            &rename,
            context("/xl/future.xml"),
        );
        assert!(matches!(
            unknown,
            Err(Error::RenameBlocked {
                reason: RenameBlock::UnmodeledReference,
                ..
            })
        ));

        let alternate = rewrite(
            br#"<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent><mc:Fallback><f>Data!A1</f></mc:Fallback></mc:AlternateContent></root>"#,
            &rename,
            context("/xl/alternate.xml"),
        );
        assert!(matches!(
            alternate,
            Err(Error::RenameBlocked {
                reason: RenameBlock::MarkupCompatibility,
                ..
            })
        ));

        let nested = rewrite(
            br#"<root><futureFormulaCache><value>Data!A1</value></futureFormulaCache></root>"#,
            &rename,
            context("/xl/future.xml"),
        );
        assert!(matches!(
            nested,
            Err(Error::RenameBlocked {
                reason: RenameBlock::UnmodeledReference,
                ..
            })
        ));

        let unrelated_nested = rewrite(
            br#"<root><futureFormulaCache><value>42</value></futureFormulaCache></root>"#,
            &rename,
            context("/xl/future.xml"),
        )
        .expect("unrelated nested extension");
        assert!(unrelated_nested.is_none());

        let malformed_known = rewrite(
            format!(r#"<s:f xmlns:s="{S}"><future/>Data!A1</s:f>"#).as_bytes(),
            &rename,
            context("/xl/malformed.xml"),
        );
        assert!(matches!(
            malformed_known,
            Err(Error::RenameBlocked {
                reason: RenameBlock::UnmodeledReference,
                ..
            })
        ));
    }

    #[test]
    fn classifies_namespaces_wrappers_and_vml_conservatively() {
        let rename = [Rename {
            before: "Data",
            after: "Input",
        }];
        let spoofed = rewrite(
            br#"<root xmlns:x="urn:not-a-formula"><x:f>Data!A1</x:f></root>"#,
            &rename,
            context("/xl/spoofed.xml"),
        );
        assert!(matches!(
            spoofed,
            Err(Error::RenameBlocked {
                reason: RenameBlock::UnmodeledReference,
                ..
            })
        ));
        let spoofed_titles = rewrite(
            br#"<TitlesOfParts><lpstr>Data</lpstr></TitlesOfParts>"#,
            &rename,
            context("/docProps/app.xml"),
        );
        assert!(matches!(
            spoofed_titles,
            Err(Error::RenameBlocked {
                reason: RenameBlock::UnmodeledReference,
                ..
            })
        ));

        let vml = rewrite(
            br#"<xml xmlns:x="urn:schemas-microsoft-com:office:excel"><x:FmlaRange>Data!A1:A3</x:FmlaRange></xml>"#,
            &rename,
            context("/xl/drawings/vmlDrawing1.vml"),
        );
        assert!(matches!(
            vml,
            Err(Error::RenameBlocked {
                reason: RenameBlock::UnmodeledReference,
                ..
            })
        ));

        let x14 = format!(
            r#"<root xmlns:x14="{}" xmlns:xm="{}"><x14:formula1><xm:f>Data!A1</xm:f></x14:formula1></root>"#,
            X14, "http://schemas.microsoft.com/office/excel/2006/main"
        );
        let changed = rewrite(
            x14.as_bytes(),
            &rename,
            context("/xl/worksheets/sheet1.xml"),
        )
        .expect("x14 scan")
        .expect("xm formula changed");
        assert!(
            std::str::from_utf8(&changed)
                .expect("UTF-8")
                .contains("<xm:f>Input!A1</xm:f>")
        );

        let numeric = [Rename {
            before: "1",
            after: "One",
        }];
        assert!(
            rewrite(
                format!(r#"<sheetProtection xmlns="{S}" sheet="1"/>"#).as_bytes(),
                &numeric,
                context("/xl/worksheets/sheet1.xml"),
            )
            .expect("boolean sheet attribute")
            .is_none()
        );

        let deep = format!(
            "{}Data!A1{}",
            "<futureFormula>".repeat(MAX_REFERENCE_DEPTH + 1),
            "</futureFormula>".repeat(MAX_REFERENCE_DEPTH + 1)
        );
        assert!(matches!(
            rewrite(deep.as_bytes(), &rename, context("/xl/deep.xml")),
            Err(Error::SheetNameDependency(SheetNameError::Limit(
                sheet_name::Limit::NestingDepth { .. }
            )))
        ));
    }
}
