//! Mail-merge settings, sources, recipients, and relationship publication.

use super::super::model::*;

use super::parts::settings_part_from_snapshot;
use super::validation::{validate_mail_merge_external_uri, validate_mail_merge_internal_source};

impl Package {
    /// Return the validated inert mail-merge settings, if configured.
    pub fn mail_merge_settings(&self) -> Result<Option<MailMergeSettings>> {
        let snapshot = self.settings_part_snapshot()?;
        let part = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        Ok(DocumentSettings::extract_from_part(&part)?
            .mail_merge()
            .cloned())
    }

    /// Resolve a mail-merge relationship without opening or fetching its target.
    ///
    /// The typed [`RelationshipId`] is validated and opaque: raw OPC
    /// relationship IDs are package plumbing (ADR-0004) and never appear on
    /// this semantic facade.
    pub fn mail_merge_target(&self, relationship_id: &RelationshipId) -> Result<Target> {
        self.mail_merge_target_str(relationship_id.as_str())
    }

    fn mail_merge_target_str(&self, relationship_id: &str) -> Result<Target> {
        let snapshot = self.settings_part_snapshot()?;
        let part = self.opc.get_part(&snapshot.target)?;
        let relationship = part.rels().get(relationship_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "mail-merge relationship '{relationship_id}' is missing"
            ))
        })?;
        if !is_mail_merge_relationship_type(relationship.reltype()) {
            return Err(Error::InvalidFormat(format!(
                "relationship '{relationship_id}' is not a mail-merge source"
            )));
        }
        if relationship.is_external() {
            return Ok(Target::External(relationship.target_ref().to_string()));
        }
        let target = relationship.target_partname()?;
        let target_part = self.opc.get_part(&target)?;
        Ok(Target::Internal {
            part_name: target,
            bytes: target_part.blob().to_vec(),
            content_type: target_part.content_type().to_string(),
        })
    }

    /// Set or replace the complete mail-merge graph atomically.
    pub fn set_mail_merge(
        &mut self,
        mut settings: MailMergeSettings,
        data_source: Option<Source>,
        header_source: Option<Source>,
        recipients: Option<Recipients>,
        conformance: mail_merge::Conformance,
    ) -> Result<()> {
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        DocumentSettings::extract_from_part(&original)?;
        let old_targets = self.mail_merge_internal_targets(&snapshot)?;
        let mut used_ids = snapshot
            .relationships
            .iter()
            .filter(|relationship| !is_mail_merge_relationship_type(&relationship.reltype))
            .map(|relationship| relationship.id.clone())
            .collect::<std::collections::HashSet<_>>();
        let mut staged_parts = Vec::new();
        let mut staged_relationships = Vec::new();

        let data_id = if let Some(source) = data_source {
            let (relationship, part) = self.stage_mail_merge_source(
                source,
                "Data",
                "mailMergeSource",
                conformance,
                &mut used_ids,
            )?;
            let id = relationship.id.clone();
            staged_relationships.push(relationship);
            if let Some(part) = part {
                staged_parts.push(part);
            }
            Some(id)
        } else {
            None
        };
        let header_id = if let Some(source) = header_source {
            let (relationship, part) = self.stage_mail_merge_source(
                source,
                "Header",
                "mailMergeHeaderSource",
                conformance,
                &mut used_ids,
            )?;
            let id = relationship.id.clone();
            staged_relationships.push(relationship);
            if let Some(part) = part {
                staged_parts.push(part);
            }
            Some(id)
        } else {
            None
        };
        let recipient_id = if let Some(recipients) = recipients {
            let xml = recipients
                .to_xml(conformance)
                .map_err(map_docx_error)?
                .into_bytes();
            let id = allocate_mail_merge_relationship_id("Recipients", &mut used_ids)?;
            let uri = self.allocate_mail_merge_part_name("recipientData", "xml")?;
            let target = uri.relative_ref(snapshot.target.base_uri());
            staged_parts.push(BlobPart::new(
                uri,
                Recipients::content_type().to_string(),
                xml,
            ));
            staged_relationships.push(StoredRelationship {
                reltype: mail_merge_relationship_type(conformance, "recipientData"),
                target,
                id: id.clone(),
                external: false,
            });
            Some(id)
        } else {
            None
        };
        settings.assign_package_relationships(data_id, header_id, recipient_id);
        let patched = patch_mail_merge(&snapshot.xml, Some(&settings), conformance)?;
        let mut replacement = settings_part_from_snapshot(&snapshot, patched, None);
        let old_ids = replacement
            .rels()
            .iter()
            .filter(|relationship| is_mail_merge_relationship_type(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_string())
            .collect::<Vec<_>>();
        for id in old_ids {
            replacement.rels_mut().remove(&id);
        }
        for relationship in staged_relationships {
            replacement.rels_mut().add_relationship(
                relationship.reltype,
                relationship.target,
                relationship.id,
                relationship.external,
            );
        }
        DocumentSettings::extract_from_part(&replacement)?;

        let mut installed = Vec::new();
        for part in staged_parts {
            let name = part.partname().clone();
            if let Err(error) = self.opc.try_add_part(Box::new(part)) {
                for installed_name in installed {
                    self.opc.remove_part(&installed_name);
                }
                return Err(error.into());
            }
            installed.push(name);
        }
        if let Err(error) = self.commit_settings_part(&snapshot, replacement) {
            for installed_name in installed {
                self.opc.remove_part(&installed_name);
            }
            return Err(error);
        }
        for old_target in old_targets {
            if !self.part_is_referenced(&old_target) {
                self.opc.remove_part(&old_target);
            }
        }
        self.opc.unsign();
        Ok(())
    }

    /// Update settings and sources using the same atomic replacement semantics.
    pub fn update_mail_merge(
        &mut self,
        settings: MailMergeSettings,
        data_source: Option<Source>,
        header_source: Option<Source>,
        recipients: Option<Recipients>,
        conformance: mail_merge::Conformance,
    ) -> Result<()> {
        self.set_mail_merge(
            settings,
            data_source,
            header_source,
            recipients,
            conformance,
        )
    }

    /// Replace recipient inclusion flags while retaining inert source targets and settings.
    pub fn update_mail_merge_recipients(
        &mut self,
        recipients: Recipients,
        conformance: mail_merge::Conformance,
    ) -> Result<()> {
        let settings = self
            .mail_merge_settings()?
            .ok_or_else(|| Error::InvalidFormat("document has no mail-merge settings".into()))?;
        let data_source = settings
            .data_source_relationship_id()
            .map(|id| {
                self.mail_merge_target_str(id)
                    .map(mail_merge_target_as_source)
            })
            .transpose()?;
        let header_source = settings
            .header_source_relationship_id()
            .map(|id| {
                self.mail_merge_target_str(id)
                    .map(mail_merge_target_as_source)
            })
            .transpose()?;
        self.set_mail_merge(
            settings,
            data_source,
            header_source,
            Some(recipients),
            conformance,
        )
    }

    /// Clear mail-merge XML, relationships, and unreachable owned targets.
    pub fn clear_mail_merge(&mut self) -> Result<bool> {
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        if DocumentSettings::extract_from_part(&original)?
            .mail_merge()
            .is_none()
        {
            return Ok(false);
        }
        let old_targets = self.mail_merge_internal_targets(&snapshot)?;
        let patched = patch_mail_merge(&snapshot.xml, None, mail_merge::Conformance::Transitional)?;
        let mut replacement = settings_part_from_snapshot(&snapshot, patched, None);
        let ids = replacement
            .rels()
            .iter()
            .filter(|relationship| is_mail_merge_relationship_type(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_string())
            .collect::<Vec<_>>();
        for id in ids {
            replacement.rels_mut().remove(&id);
        }
        DocumentSettings::extract_from_part(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        for old_target in old_targets {
            if !self.part_is_referenced(&old_target) {
                self.opc.remove_part(&old_target);
            }
        }
        self.opc.unsign();
        Ok(true)
    }

    fn stage_mail_merge_source(
        &self,
        source: Source,
        label: &str,
        relationship_suffix: &str,
        conformance: mail_merge::Conformance,
        used_ids: &mut std::collections::HashSet<String>,
    ) -> Result<(StoredRelationship, Option<BlobPart>)> {
        let id = allocate_mail_merge_relationship_id(label, used_ids)?;
        let settings_target = self.settings_part_snapshot()?.target;
        match source {
            Source::External(uri) => {
                validate_mail_merge_external_uri(&uri)?;
                Ok((
                    StoredRelationship {
                        reltype: mail_merge_relationship_type(conformance, relationship_suffix),
                        target: uri,
                        id,
                        external: true,
                    },
                    None,
                ))
            },
            Source::Internal {
                bytes,
                content_type,
                extension,
            } => {
                validate_mail_merge_internal_source(&bytes, &content_type, &extension)?;
                let uri = self.allocate_mail_merge_part_name(label, &extension)?;
                let target = uri.relative_ref(settings_target.base_uri());
                let part = BlobPart::new(uri, content_type, bytes);
                Ok((
                    StoredRelationship {
                        reltype: mail_merge_relationship_type(conformance, relationship_suffix),
                        target,
                        id,
                        external: false,
                    },
                    Some(part),
                ))
            },
        }
    }

    fn allocate_mail_merge_part_name(&self, stem: &str, extension: &str) -> Result<PackURI> {
        for number in 1usize.. {
            let candidate = PackURI::new(format!("/word/mailMerge/{stem}{number}.{extension}"))
                .map_err(Error::InvalidUri)?;
            if self.opc.iter_parts().all(|part| {
                !part
                    .partname()
                    .as_str()
                    .eq_ignore_ascii_case(candidate.as_str())
            }) {
                return Ok(candidate);
            }
        }
        unreachable!("the mail-merge part-name space is unbounded")
    }

    fn mail_merge_internal_targets(&self, snapshot: &SettingsPartSnapshot) -> Result<Vec<PackURI>> {
        let Ok(part) = self.opc.get_part(&snapshot.target) else {
            return Ok(Vec::new());
        };
        part.rels()
            .iter()
            .filter(|relationship| {
                is_mail_merge_relationship_type(relationship.reltype())
                    && !relationship.is_external()
            })
            .map(|relationship| relationship.target_partname().map_err(Into::into))
            .collect()
    }
}

fn mail_merge_relationship_type(conformance: mail_merge::Conformance, suffix: &str) -> String {
    let base = match conformance {
        mail_merge::Conformance::Transitional => {
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        },
        mail_merge::Conformance::Strict => {
            "http://purl.oclc.org/ooxml/officeDocument/relationships"
        },
    };
    format!("{base}/{suffix}")
}

fn allocate_mail_merge_relationship_id(
    label: &str,
    used: &mut std::collections::HashSet<String>,
) -> Result<String> {
    (1usize..=MAX_MAIL_MERGE_RELATIONSHIPS)
        .map(|number| format!("rIdMailMerge{label}{number}"))
        .find(|id| used.insert(id.clone()))
        .ok_or_else(|| Error::InvalidFormat("mail-merge relationship ID space is exhausted".into()))
}

fn mail_merge_target_as_source(target: Target) -> Source {
    match target {
        Target::External(uri) => Source::External(uri),
        Target::Internal {
            part_name,
            bytes,
            content_type,
        } => {
            let extension = part_name
                .as_str()
                .rsplit_once('.')
                .map(|(_, extension)| extension)
                .filter(|extension| {
                    !extension.is_empty()
                        && extension.len() <= 16
                        && extension.bytes().all(|byte| byte.is_ascii_alphanumeric())
                })
                .unwrap_or("bin")
                .to_string();
            Source::Internal {
                bytes,
                content_type,
                extension,
            }
        },
    }
}
