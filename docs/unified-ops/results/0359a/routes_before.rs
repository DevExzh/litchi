/// The leaf calls of each source-backed route on this build (record 0359
/// base: the former source-backed publishers).
#[cfg(all(feature = "doc", feature = "xls", feature = "rtf"))]
mod routes {
    use std::num::{NonZeroU64, NonZeroUsize};
    use std::sync::Arc;

    use litchi_core::{OwnedSource, Position, ReadAt};
    use litchi_xls::cell_values::{Reference, Selector};

    /// Whether this build has the separate plan-only XLS cell route.
    pub(super) const PLAN_ONLY_ROUTE: bool = true;

    /// The ops retention bound of the DOC 1k batch: the base retains one
    /// complete snapshot per isolated step (about 15 GB at 1,000 steps).
    pub(super) const DOC_BATCH_SNAPSHOT_BYTES: usize = 64 << 30;

    fn same_width(text: &str) -> String {
        text.replacen("holds", "keeps", 1)
    }

    pub(super) fn doc_source_set(bytes: &[u8]) -> Vec<u8> {
        let snapshot =
            litchi_doc::body_text::source::SourceSnapshot::from_bytes(bytes.to_vec()).expect("open");
        let mut edit = snapshot.edit(Position::new(500)).expect("edit");
        let text = same_width(edit.text());
        edit.set_text(text).expect("set");
        let commit = edit.commit().expect("commit");
        let mut output = Vec::new();
        commit.write_to(&mut output).expect("write");
        output
    }

    pub(super) fn xls_source_cells(bytes: &[u8], cells: usize, plan_only: bool) -> Vec<u8> {
        let snapshot =
            litchi_xls::cell_values::Snapshot::from_bytes(bytes.to_vec(), Default::default())
                .expect("open");
        let mut edit = snapshot.edit();
        for index in 0..cells {
            let reference = Reference::new((index / 10) as u32, (index % 10) as u32).expect("ref");
            edit.set_numeric(Selector::Name("Data"), reference, 7.0 + index as f64)
                .expect("set");
        }
        let mut output = Vec::new();
        if plan_only {
            let commit = edit.commit_source_backed_plan().expect("commit");
            commit.write_to(&mut output).expect("write");
        } else {
            let commit = edit.commit_source_backed().expect("commit");
            commit.write_to(&mut output).expect("write");
        }
        output
    }

    fn comment(index: usize) -> litchi_xls::comments::Value {
        litchi_xls::comments::Value::new(
            format!("Target {index:03}"),
            format!("comment target {index:03}"),
        )
        .expect("value")
    }

    pub(super) fn xls_source_comments(bytes: &[u8], count: usize) -> Vec<u8> {
        let mut edit = litchi_xls::Package::from_bytes(bytes)
            .expect("open")
            .edit_comments()
            .expect("edit");
        let updates = (0..count).map(|index| {
            litchi_xls::comments::Update::new(
                Reference::new(index as u32, 1).expect("ref"),
                comment(index),
            )
        });
        edit.replace_many(Selector::Position(Position::new(0)), updates)
            .expect("replace");
        let commit = edit.commit_source_backed().expect("commit");
        let mut output = Vec::new();
        commit.write_to(&mut output).expect("write");
        output
    }

    pub(super) fn xls_source_hide(bytes: &[u8], count: usize) -> Vec<u8> {
        let mut edit = litchi_xls::Package::from_bytes(bytes)
            .expect("open")
            .edit_sheet_visibility()
            .expect("edit");
        for position in 1..=count {
            edit.hide(position.into()).expect("hide");
        }
        let commit = edit.commit_source_backed().expect("commit");
        let mut output = Vec::new();
        commit.write_to(&mut output).expect("write");
        output
    }

    fn context() -> litchi_core::ExecutionContext {
        let budget = litchi_core::Budget::root(
            "legacy-w5-source-probe",
            litchi_core::Limits::new(1 << 30, 1 << 30, 1 << 30, 1 << 30, 64, 1 << 40),
        );
        let (_source, token) = litchi_core::CancellationSource::pair();
        let limits = litchi_core::ExecutionLimits::new(
            NonZeroUsize::new(1).expect("one"),
            NonZeroUsize::new(1).expect("one"),
            NonZeroU64::new(64 * 1024).expect("window"),
            0,
        )
        .expect("limits");
        litchi_core::ExecutionContext::new(budget, token, limits)
    }

    pub(super) fn rtf_source_append(document: &litchi_rtf::Package, bytes: &[u8]) -> Vec<u8> {
        use litchi_rtf::tail_append::{
            TailAppendPublicationLimits, TailAppendSourceEdit, TailAppendSourceProof, TailSelector,
        };

        let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes.to_vec()));
        let context = context();
        let proof = TailAppendSourceProof::from_document_for_source_with_context(
            document,
            TailSelector::Body,
            source.as_ref(),
            &context,
        )
        .expect("proof");
        let mut edit = TailAppendSourceEdit::new(source, proof, TailSelector::Body);
        edit.append_text_paragraphs(&["Appended paragraph."])
            .expect("append");
        let plan = edit
            .publication_plan_with_context(TailAppendPublicationLimits::default(), &context)
            .expect("plan");
        let mut output = Vec::new();
        plan.write_to(&mut output, &context).expect("write");
        output
    }
}
