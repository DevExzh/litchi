//! Logical owner boundary over a retained effective graph.

use super::*;
use litchi_opc::PreparedTopology;

impl Package {
    /// Adopt a validated effective topology derived from this physical source.
    ///
    /// The returned owner has fresh semantic caches. The original owner and
    /// its captured views remain unchanged. This is an OPC adoption boundary;
    /// it does not replace semantic DOCX edit validation. Operations not yet
    /// migrated to effective graphs return a typed refusal.
    ///
    /// # Errors
    /// Refuses a foreign physical lineage, stale source, cancellation or an
    /// invalid effective main-document relationship/content type.
    pub fn with_prepared_topology(&self, topology: PreparedTopology) -> Result<Self> {
        self.check_source()?;
        topology.check_source()?;
        let result = (|| {
            if topology.physical_source_lineage() != self.physical.source_lineage() {
                return Err(Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "with_prepared_topology",
                    reason: "the candidate belongs to a different physical source",
                });
            }
            let main = topology.main_document_part()?;
            validate_document_main_content_type(main.content_type())?;
            let mut owner = Self::with_execution(self.physical.clone(), self.execution.clone())?;
            owner.effective = Some(topology.clone());
            Ok(owner)
        })();
        self.check_source()?;
        topology.check_source()?;
        result
    }

    /// Recheck the physical archive and every retained transfer source/context.
    pub fn check_source(&self) -> Result<()> {
        match &self.effective {
            Some(topology) => Ok(topology.check_source()?),
            None => {
                self.physical.source_version()?;
                Ok(self.physical.check_execution()?)
            },
        }
    }

    /// Publish the effective package, preserving unchanged physical members.
    pub fn to_stream<W: Write>(&self, writer: W) -> Result<()> {
        match &self.effective {
            Some(topology) => Ok(topology.publish_to_stream(writer)?),
            None => Ok(self
                .physical
                .clone()
                .write_topology_to_stream(writer, SourceTopologyPlan::new())?),
        }
    }

    pub(super) fn effective_document(&self, topology: &PreparedTopology) -> Result<Document> {
        topology.check_source()?;
        let result = (|| {
            let main = topology.main_document_part()?;
            validate_document_main_content_type(main.content_type())?;
            let data = match main.data().map_err(Error::from) {
                Ok(data) => data,
                Err(error) if is_index_pressure_error(&error) => {
                    self.paragraph_index_cache.trim_clean();
                    main.data()?
                },
                Err(error) => return Err(error),
            };
            let source_version = topology.physical_source_version()?;
            let (xml, paragraph_memo, index_admission, paragraph_route, retained) =
                if let Some(context) = &self.execution {
                    let build = self.paragraph_index_cache.get_or_build_source(
                        source_version,
                        ParagraphIndexRoute::Managed,
                        || {
                            ensure_source_document_xml(data.as_bytes(), Some(context))?;
                            let admission =
                                DocumentIndexAdmission::new(context, data.as_bytes().len())?;
                            let _parser =
                                admit_document_query_parser(Some(context), data.as_bytes().len())?;
                            let index = ParagraphIndex::from_xml(data.as_bytes()).ok();
                            topology.check_source()?;
                            Ok(match index {
                                Some(index) => ParagraphIndexBuild::indexed(index, Some(admission)),
                                None => ParagraphIndexBuild::unavailable(Some(admission)),
                            })
                        },
                    )?;
                    (
                        DocumentPayload::Effective(data.clone()),
                        OnceLock::from(build.memo),
                        build.admission,
                        ParagraphIndexRoute::Managed,
                        Some(data),
                    )
                } else {
                    let mut bytes = Vec::new();
                    bytes
                        .try_reserve_exact(data.as_bytes().len())
                        .map_err(|source| litchi_opc::OpcError::Allocation {
                            resource: "effective DOCX visible XML",
                            source,
                        })?;
                    bytes.extend_from_slice(data.as_bytes());
                    let xml = visible_document_xml(Arc::new(bytes))?;
                    (
                        DocumentPayload::Owned(xml),
                        OnceLock::new(),
                        None,
                        ParagraphIndexRoute::Visible,
                        Some(data),
                    )
                };
            Ok(Document {
                xml,
                _retained: retained,
                paragraph_memo,
                _index_admission: index_admission,
                source_version,
                execution: self.execution.clone(),
                paragraph_index_cache: Arc::clone(&self.paragraph_index_cache),
                paragraph_route,
                tolerance: Arc::new(Mutex::new(self.opened_tolerance())),
            })
        })();
        match topology.check_source().map_err(Error::from) {
            Ok(()) => {
                if result.is_err() {
                    self.paragraph_index_cache.trim_clean();
                }
                result
            },
            Err(error) => {
                drop(result);
                self.paragraph_index_cache.trim_clean();
                Err(error)
            },
        }
    }
}
