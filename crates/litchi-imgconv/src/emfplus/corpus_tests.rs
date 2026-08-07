//! Framing fixtures for LibreOffice's representative EMF+ corpus.
//!
//! These tests deliberately exercise only the EMF envelope and EMF+ record
//! framing.  They do not require object decoding or playback support.

use std::{
    fs,
    path::{Path, PathBuf},
};

use super::renderer::{EmfPlusSvgRenderer, RendererLimits};
use super::{
    EMR_COMMENT, EmfPlusRecordIter, ParserLimits, RecordType, try_extract_emfplus_comment_body,
};

const EMF_RECORD_HEADER_SIZE: usize = 8;

const EMFPLUS_FIXTURES: &[&str] = &[
    "TestDrawString.emf",
    "TestDrawStringAlign.emf",
    "TestDrawStringTracking.emf",
    "TestDrawStringTransparent.emf",
    "TestDrawStringWithBrush.emf",
    "TestEmfPlusBrushPathGradientMultiSurroundColor.emf",
    "TestEmfPlusBrushPathGradientWithBlendColors.emf",
    "TestEmfPlusDrawBeziers.emf",
    "TestEmfPlusDrawCurve.emf",
    "TestEmfPlusDrawImagePointsWithMetafile.emf",
    "TestEmfPlusDrawLineWithCaps.emf",
    "TestEmfPlusDrawLineWithDash.emf",
    "TestEmfPlusDrawPathWithCustomCap.emf",
    "TestEmfPlusDrawPathWithMiterLimit.emf",
    "TestEmfPlusFillClosedCurve.emf",
    "TestEmfPlusFillClosedCurveWinding.emf",
    "TestEmfPlusFillRectsOverlap.emf",
    "TestEmfPlusFillRectsWithTextureBrush.emf",
    "TestEmfPlusGetDC.emf",
    "TestEmfPlusGetDC2.emf",
    "TestEmfPlusSave.emf",
    "TestEmfPlusSetPageTransform.emf",
];

const HATCHED_FIXTURES: &[&str] = &["TestHatchedBrush.emf", "TestHatchedPen.emf"];

const COVERED_RECORD_TYPES: &[RecordType] = &[
    RecordType::Header,
    RecordType::EndOfFile,
    RecordType::Comment,
    RecordType::GetDc,
    RecordType::Object,
    RecordType::FillRects,
    RecordType::DrawRects,
    RecordType::FillPolygon,
    RecordType::DrawLines,
    RecordType::FillPath,
    RecordType::DrawPath,
    RecordType::FillClosedCurve,
    RecordType::DrawClosedCurve,
    RecordType::DrawCurve,
    RecordType::DrawBeziers,
    RecordType::DrawImagePoints,
    RecordType::DrawString,
    RecordType::SetAntiAliasMode,
    RecordType::SetTextRenderingHint,
    RecordType::SetTextContrast,
    RecordType::SetInterpolationMode,
    RecordType::SetPixelOffsetMode,
    RecordType::SetCompositingQuality,
    RecordType::Save,
    RecordType::Restore,
    RecordType::SetWorldTransform,
    RecordType::ResetWorldTransform,
    RecordType::MultiplyWorldTransform,
    RecordType::TranslateWorldTransform,
    RecordType::ScaleWorldTransform,
    RecordType::RotateWorldTransform,
    RecordType::SetPageTransform,
    RecordType::ResetClip,
    RecordType::SetClipRect,
    RecordType::SetClipPath,
    RecordType::SetClipRegion,
    RecordType::OffsetClip,
];

/// A validated classic EMF record, retaining the body layout exposed by the
/// existing EMF parser (`EMR_COMMENT` `DataSize` starts at `body[0]`).
#[derive(Clone, Copy, Debug)]
struct ClassicEmfRecord<'a> {
    offset: usize,
    record_type: u32,
    body: &'a [u8],
}

#[derive(Debug, Default)]
struct CorpusCoverage {
    emf_comments: usize,
    emfplus_comments: usize,
    emfplus_records: usize,
    record_types: Vec<RecordType>,
}

fn corpus_dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/libreoffice-core/emfio/qa/cppunit/emf/data")
}

fn read_fixture(name: &str) -> Vec<u8> {
    let path = corpus_dir().join(name);
    assert!(
        path.is_file(),
        "missing LibreOffice corpus fixture: {path:?}"
    );
    fs::read(&path).unwrap_or_else(|error| panic!("failed to read {path:?}: {error}"))
}

fn parser_limits() -> ParserLimits {
    ParserLimits::default()
}

fn read_u32(bytes: &[u8], offset: usize) -> Result<u32, String> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| "classic EMF field offset overflow".to_owned())?;
    let field = bytes
        .get(offset..end)
        .ok_or_else(|| format!("truncated classic EMF u32 at offset {offset}"))?;
    let array: [u8; 4] = field
        .try_into()
        .map_err(|_| "classic EMF u32 has an invalid width".to_owned())?;
    Ok(u32::from_le_bytes(array))
}

/// Walk a classic EMF file without relying on record-specific decoding.
fn scan_classic_emf_records(bytes: &[u8]) -> Result<Vec<ClassicEmfRecord<'_>>, String> {
    let mut records = Vec::new();
    let mut offset = 0usize;

    while offset < bytes.len() {
        let remaining = bytes
            .get(offset..)
            .ok_or_else(|| "classic EMF record offset is outside the file".to_owned())?;
        if remaining.len() < EMF_RECORD_HEADER_SIZE {
            return Err(format!(
                "truncated classic EMF record header at offset {offset}: {} bytes remain",
                remaining.len()
            ));
        }

        let record_type = read_u32(bytes, offset)?;
        let declared_size = read_u32(bytes, offset + 4)?;
        let size = usize::try_from(declared_size).map_err(|_| {
            format!("classic EMF record size does not fit usize at offset {offset}")
        })?;
        if size < EMF_RECORD_HEADER_SIZE {
            return Err(format!(
                "classic EMF record at offset {offset} has undersized Size {size}"
            ));
        }
        if size % 4 != 0 {
            return Err(format!(
                "classic EMF record at offset {offset} is not 32-bit aligned"
            ));
        }

        let end = offset
            .checked_add(size)
            .ok_or_else(|| "classic EMF record end offset overflow".to_owned())?;
        let record = bytes.get(offset..end).ok_or_else(|| {
            format!(
                "classic EMF record at offset {offset} ends at {end}, beyond file length {}",
                bytes.len()
            )
        })?;
        let body = record
            .get(EMF_RECORD_HEADER_SIZE..)
            .ok_or_else(|| "classic EMF record body is outside its record".to_owned())?;
        records.push(ClassicEmfRecord {
            offset,
            record_type,
            body,
        });
        offset = end;
    }

    Ok(records)
}

fn collect_coverage(name: &str, bytes: &[u8]) -> CorpusCoverage {
    let records = scan_classic_emf_records(bytes)
        .unwrap_or_else(|error| panic!("{name}: invalid classic EMF framing: {error}"));
    let mut coverage = CorpusCoverage::default();

    for record in records {
        if record.record_type != EMR_COMMENT {
            continue;
        }
        coverage.emf_comments += 1;
        let payload = try_extract_emfplus_comment_body(record.body, parser_limits())
            .unwrap_or_else(|error| {
                panic!(
                    "{name}: invalid EMR_COMMENT body at offset {}: {error}",
                    record.offset
                )
            });
        let Some(payload) = payload else {
            continue;
        };
        coverage.emfplus_comments += 1;

        let iter = EmfPlusRecordIter::new(payload, parser_limits()).unwrap_or_else(|error| {
            panic!(
                "{name}: rejected EMF+ payload from comment at offset {}: {error}",
                record.offset
            )
        });
        for framed_record in iter {
            let framed_record = framed_record.unwrap_or_else(|error| {
                panic!(
                    "{name}: invalid EMF+ record in comment at offset {}: {error}",
                    record.offset
                )
            });
            coverage.emfplus_records += 1;
            coverage.record_types.push(framed_record.header.record_type);
        }
    }

    coverage
}

#[test]
fn frames_representative_libreoffice_emfplus_comments() {
    for name in EMFPLUS_FIXTURES {
        let bytes = read_fixture(name);
        let coverage = collect_coverage(name, &bytes);
        assert!(
            coverage.emfplus_comments > 0,
            "{name} should contain an EMR_COMMENT_EMFPLUS record"
        );
        assert!(
            coverage.emfplus_records > 0,
            "{name} should contain at least one EMF+ record"
        );
    }
}

#[test]
fn reports_stable_emfplus_framing_coverage() {
    let mut combined = CorpusCoverage::default();
    for name in EMFPLUS_FIXTURES {
        let bytes = read_fixture(name);
        let coverage = collect_coverage(name, &bytes);
        combined.emf_comments += coverage.emf_comments;
        combined.emfplus_comments += coverage.emfplus_comments;
        combined.emfplus_records += coverage.emfplus_records;
        combined.record_types.extend(coverage.record_types);
    }

    assert_eq!(combined.emf_comments, 178);
    assert_eq!(combined.emfplus_comments, 178);
    // Every scanner and nested iterator advances by its checked declared Size;
    // 178 comments contain 625 distinct, non-overlapping nested records.
    assert_eq!(combined.emfplus_records, 625);
    let covered_types: Vec<_> = RecordType::ALL
        .iter()
        .copied()
        .filter(|record_type| combined.record_types.contains(record_type))
        .collect();
    assert_eq!(covered_types, COVERED_RECORD_TYPES);
}

#[test]
fn identifies_hatched_samples_as_classic_emf_without_emfplus_comments() {
    for name in HATCHED_FIXTURES {
        let bytes = read_fixture(name);
        let coverage = collect_coverage(name, &bytes);
        assert_eq!(
            coverage.emf_comments, 0,
            "{name} unexpectedly has EMR_COMMENT"
        );
        assert_eq!(
            coverage.emfplus_comments, 0,
            "{name} unexpectedly has EMR_COMMENT_EMFPLUS"
        );
        assert_eq!(
            coverage.emfplus_records, 0,
            "{name} unexpectedly has EMF+ records"
        );
    }
}

#[test]
fn renders_representative_emfplus_streams_to_safe_svg() {
    for name in EMFPLUS_FIXTURES {
        // This LibreOffice unit fixture intentionally stops after DrawPath and
        // has no normative EndOfFile record; strict stream completion rejects it.
        if *name == "TestEmfPlusDrawPathWithCustomCap.emf" {
            continue;
        }
        let bytes = read_fixture(name);
        let records = scan_classic_emf_records(&bytes).unwrap();
        let mut renderer =
            EmfPlusSvgRenderer::new(1024.0, 1024.0, RendererLimits::default()).unwrap();
        for record in records {
            if record.record_type == EMR_COMMENT {
                renderer
                    .push_comment_body(record.body)
                    .unwrap_or_else(|error| {
                        panic!("{name}: EMF+ comment playback failed: {error}")
                    });
            }
        }
        let output = renderer
            .finish()
            .unwrap_or_else(|error| panic!("{name}: EMF+ SVG completion failed: {error}"));
        assert!(output.svg().starts_with("<svg"));
    }
}
