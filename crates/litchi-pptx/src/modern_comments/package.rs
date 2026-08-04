//! OPC graph lifecycle and transactional CRUD for modern comments.

use super::model::*;
use super::{
    MODERN_COMMENT_AUTHOR_CONTENT_TYPE, MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE,
    MODERN_COMMENT_CONTENT_TYPE, MODERN_COMMENT_RELATIONSHIP_TYPE, SLIDE_CONTENT_TYPE,
};
use crate::{Error, Result};
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use std::collections::HashSet;

mod comments {
    use super::*;

    fn invalid(message: impl Into<String>) -> Error {
        Error::Invalid(message.into())
    }

    fn limit(label: &str) -> Error {
        invalid(format!("{label} exceeds implementation limit"))
    }

    fn validate_relationship_id(value: &str) -> Result<()> {
        bounded(value)?;
        if value.is_empty() || value.chars().any(char::is_whitespace) {
            Err(invalid(
                "modern Comment relationship ID must be nonempty without whitespace",
            ))
        } else {
            Ok(())
        }
    }

    fn bounded(value: &str) -> Result<()> {
        if value.len() <= super::super::MAX_STRING_BYTES {
            Ok(())
        } else {
            Err(limit("modern Comment string bytes"))
        }
    }

    fn require_content_type(part: &dyn litchi_opc::Part, expected: &str) -> Result<()> {
        if part.content_type() == expected {
            Ok(())
        } else {
            Err(Error::ContentType {
                expected: expected.into(),
                actual: part.content_type().into(),
            })
        }
    }

    pub fn load_modern_comments(package: &OpcPackage) -> Result<Vec<ModernCommentPart>> {
        if package
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == MODERN_COMMENT_RELATIONSHIP_TYPE)
        {
            return Err(invalid(
                "modern Comment relationship cannot originate at the package root",
            ));
        }

        let mut relationships = Vec::new();
        let mut targets = HashSet::new();
        for source in package.iter_parts() {
            for relationship in source
                .rels()
                .iter()
                .filter(|relationship| relationship.reltype() == MODERN_COMMENT_RELATIONSHIP_TYPE)
            {
                if source.content_type() != SLIDE_CONTENT_TYPE {
                    return Err(invalid(format!(
                        "modern Comment relationship has non-Slide source '{}'",
                        source.partname()
                    )));
                }
                if relationship.is_external() {
                    return Err(invalid("modern Comment relationship cannot be external"));
                }
                let target = relationship.target_partname()?;
                let part = package.get_part(&target)?;
                require_content_type(part, MODERN_COMMENT_CONTENT_TYPE)?;
                targets.insert(target.to_string());
                relationships.push((
                    source.partname().to_string(),
                    relationship.r_id().to_string(),
                    target.to_string(),
                ));
            }
        }

        for part in package.iter_parts() {
            if part.content_type() == MODERN_COMMENT_CONTENT_TYPE
                && !targets.contains(part.partname().as_str())
            {
                return Err(invalid(format!(
                    "package contains orphan modern Comment part '{}'",
                    part.partname()
                )));
            }
        }

        relationships
            .into_iter()
            .map(|(slide_part_name, relationship_id, part_name)| {
                let uri = PackURI::new(&part_name).map_err(invalid)?;
                let comments = ModernCommentList::parse(package.get_part(&uri)?.blob())?;
                Ok(ModernCommentPart {
                    slide_part_name,
                    relationship_id,
                    part_name,
                    comments,
                })
            })
            .collect()
    }

    /// Add a new modern Comment part after validating the complete existing graph.
    /// Existing parts are deliberately not overwritten.
    pub fn store_modern_comment(package: &mut OpcPackage, value: &ModernCommentPart) -> Result<()> {
        load_modern_comments(package)?;
        validate_relationship_id(&value.relationship_id)?;
        let slide_name = PackURI::new(&value.slide_part_name).map_err(invalid)?;
        let slide = package.get_part(&slide_name)?;
        if slide.content_type() != SLIDE_CONTENT_TYPE {
            return Err(invalid(format!(
                "'{}' is not a Slide part",
                value.slide_part_name
            )));
        }
        if slide.rels().get(&value.relationship_id).is_some() {
            return Err(invalid("modern Comment relationship ID already exists"));
        }
        let part_name = PackURI::new(&value.part_name).map_err(invalid)?;
        if package
            .iter_parts()
            .any(|part| part.partname() == &part_name)
        {
            return Err(invalid(format!("part '{part_name}' already exists")));
        }
        let xml = value.comments.to_xml()?;
        let target = part_name.relative_ref(slide_name.base_uri());
        package.try_add_part(Box::new(BlobPart::new(
            part_name,
            MODERN_COMMENT_CONTENT_TYPE.into(),
            xml,
        )))?;
        package
            .get_part_mut(&slide_name)?
            .rels_mut()
            .add_relationship(
                MODERN_COMMENT_RELATIONSHIP_TYPE.into(),
                target,
                value.relationship_id.clone(),
                false,
            );
        Ok(())
    }

    pub fn find_modern_comment(
        package: &OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
    ) -> Result<Option<ModernComment>> {
        Ok(load_modern_comments(package)?
            .into_iter()
            .find(|part| part.slide_part_name == slide_part_name.to_string())
            .and_then(|part| {
                part.comments
                    .comments
                    .into_iter()
                    .find(|comment| comment.id == comment_id)
            }))
    }

    /// Add a modern comment to a slide, creating a collision-safe part when necessary.
    pub fn add_modern_comment(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment: ModernComment,
    ) -> Result<ModernCommentPart> {
        let mut parts = load_modern_comments(package)?;
        ensure_modern_comment_id_is_free(&parts, &comment.id)?;
        ensure_modern_reply_ids_are_free(&parts, &comment)?;
        if let Some(index) = parts
            .iter()
            .position(|part| part.slide_part_name == slide_part_name.to_string())
        {
            let mut staged = parts[index].clone();
            staged.comments.comments.push(comment);
            parts[index] = staged.clone();
            validate_and_commit_modern_comment_part(package, &staged, &parts)?;
            return Ok(staged);
        }

        package.get_part(slide_part_name)?;
        let part_name = next_modern_comment_part_name(package)?;
        let relationship_id = next_modern_comment_relationship_id(package, slide_part_name)?;
        let staged = ModernCommentPart {
            slide_part_name: slide_part_name.to_string(),
            relationship_id: relationship_id.clone(),
            part_name: part_name.to_string(),
            comments: ModernCommentList {
                root_prefix: "p188".into(),
                namespace_declarations: Vec::new(),
                comments: vec![comment],
            },
        };
        parts.push(staged.clone());
        validate_modern_comment_graph_for_mutation(package, &parts)?;
        let xml = staged.comments.to_xml()?;
        ModernCommentList::parse(&xml)?;
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part_name.clone(),
            MODERN_COMMENT_CONTENT_TYPE.into(),
            xml,
        )))?;
        package
            .get_part_mut(slide_part_name)?
            .rels_mut()
            .add_relationship(
                MODERN_COMMENT_RELATIONSHIP_TYPE.into(),
                part_name.relative_ref(slide_part_name.base_uri()),
                relationship_id,
                false,
            );
        package.unsign();
        Ok(staged)
    }

    /// Update a modern comment without permitting its stable GUID to change.
    pub fn update_modern_comment<F>(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut ModernComment),
    {
        let mut parts = load_modern_comments(package)?;
        let Some(part_index) = parts
            .iter()
            .position(|part| part.slide_part_name == slide_part_name.to_string())
        else {
            return Ok(false);
        };
        let Some(comment_index) = parts[part_index]
            .comments
            .comments
            .iter()
            .position(|comment| comment.id == comment_id)
        else {
            return Ok(false);
        };
        let reply_ids_before: Vec<String> = parts[part_index].comments.comments[comment_index]
            .replies
            .iter()
            .map(|reply| reply.id.clone())
            .collect();
        update(&mut parts[part_index].comments.comments[comment_index]);
        if parts[part_index].comments.comments[comment_index].id != comment_id {
            return Err(invalid("modern comment update cannot change its ID"));
        }
        let reply_ids_after: Vec<&str> = parts[part_index].comments.comments[comment_index]
            .replies
            .iter()
            .map(|reply| reply.id.as_str())
            .collect();
        if reply_ids_after.len() == reply_ids_before.len()
            && reply_ids_after
                .iter()
                .zip(&reply_ids_before)
                .any(|(after, before)| *after != before)
        {
            return Err(invalid("modern comment update cannot change a reply ID"));
        }
        ensure_all_modern_ids_are_unique(&parts)?;
        let staged = parts[part_index].clone();
        validate_and_commit_modern_comment_part(package, &staged, &parts)?;
        Ok(true)
    }

    /// Replace a modern comment without permitting its stable GUID to change.
    pub fn replace_modern_comment(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        replacement: ModernComment,
    ) -> Result<bool> {
        if replacement.id != comment_id {
            return Err(invalid("replacement modern comment ID must match"));
        }
        update_modern_comment(package, slide_part_name, comment_id, move |comment| {
            *comment = replacement;
        })
    }

    /// Remove a modern comment and remove an empty per-slide part unless another owner shares it.
    pub fn remove_modern_comment(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
    ) -> Result<bool> {
        let mut parts = load_modern_comments(package)?;
        let Some(part_index) = parts
            .iter()
            .position(|part| part.slide_part_name == slide_part_name.to_string())
        else {
            return Ok(false);
        };
        let Some(comment_index) = parts[part_index]
            .comments
            .comments
            .iter()
            .position(|comment| comment.id == comment_id)
        else {
            return Ok(false);
        };
        parts[part_index].comments.comments.remove(comment_index);
        if !parts[part_index].comments.comments.is_empty() {
            let staged = parts[part_index].clone();
            validate_and_commit_modern_comment_part(package, &staged, &parts)?;
            return Ok(true);
        }

        let removed = parts.remove(part_index);
        validate_modern_comment_graph_for_mutation(package, &parts)?;
        let part_name = PackURI::new(&removed.part_name).map_err(invalid)?;
        package
            .get_part_mut(slide_part_name)?
            .rels_mut()
            .remove(&removed.relationship_id);
        if !modern_comment_part_is_referenced(package, &part_name) {
            package.remove_part(&part_name);
        }
        package.unsign();
        Ok(true)
    }

    /// Reorder every modern comment in one slide part by a complete GUID list.
    pub fn reorder_modern_comments(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        ordered_comment_ids: &[String],
    ) -> Result<Vec<ModernComment>> {
        let mut parts = load_modern_comments(package)?;
        let Some(part_index) = parts
            .iter()
            .position(|part| part.slide_part_name == slide_part_name.to_string())
        else {
            if ordered_comment_ids.is_empty() {
                return Ok(Vec::new());
            }
            return Err(invalid("modern comment part is missing for slide"));
        };
        if ordered_comment_ids.len() != parts[part_index].comments.comments.len() {
            return Err(invalid("modern comment reorder must contain every comment"));
        }
        let mut remaining = std::collections::HashMap::new();
        for comment in parts[part_index].comments.comments.drain(..) {
            if remaining.insert(comment.id.clone(), comment).is_some() {
                return Err(invalid("duplicate modern comment ID"));
            }
        }
        let mut ordered = Vec::with_capacity(ordered_comment_ids.len());
        for id in ordered_comment_ids {
            let comment = remaining
                .remove(id)
                .ok_or_else(|| invalid(format!("unknown or duplicate modern comment ID {id}")))?;
            ordered.push(comment);
        }
        if !remaining.is_empty() {
            return Err(invalid("modern comment reorder must contain every comment"));
        }
        parts[part_index].comments.comments = ordered.clone();
        let staged = parts[part_index].clone();
        validate_and_commit_modern_comment_part(package, &staged, &parts)?;
        Ok(ordered)
    }

    /// Find a reply by its stable GUID within a modern comment thread.
    pub fn find_modern_comment_reply(
        package: &OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        reply_id: &str,
    ) -> Result<Option<ModernCommentReply>> {
        Ok(
            find_modern_comment(package, slide_part_name, comment_id)?.and_then(|comment| {
                comment
                    .replies
                    .into_iter()
                    .find(|reply| reply.id == reply_id)
            }),
        )
    }

    /// Add a reply to a modern comment thread.
    pub fn add_modern_comment_reply(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        reply: ModernCommentReply,
    ) -> Result<bool> {
        let reply_id = reply.id.clone();
        let parts = load_modern_comments(package)?;
        if modern_id_exists(&parts, &reply_id) {
            return Err(invalid(format!(
                "duplicate modern comment or reply ID {reply_id}"
            )));
        }
        update_modern_comment(package, slide_part_name, comment_id, move |comment| {
            comment.reply_list_present = true;
            comment.replies.push(reply);
        })
    }

    /// Update a reply without permitting its stable GUID to change.
    pub fn update_modern_comment_reply<F>(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        reply_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut ModernCommentReply),
    {
        if find_modern_comment_reply(package, slide_part_name, comment_id, reply_id)?.is_none() {
            return Ok(false);
        }
        let mut update = Some(update);
        let mut found = false;
        let result = update_modern_comment(package, slide_part_name, comment_id, |comment| {
            if let Some(reply) = comment
                .replies
                .iter_mut()
                .find(|reply| reply.id == reply_id)
            {
                if let Some(update) = update.take() {
                    update(reply);
                }
                found = true;
            }
        })?;
        if !result || !found {
            return Ok(false);
        }
        Ok(true)
    }

    /// Replace a reply without permitting its stable GUID to change.
    pub fn replace_modern_comment_reply(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        reply_id: &str,
        replacement: ModernCommentReply,
    ) -> Result<bool> {
        if replacement.id != reply_id {
            return Err(invalid("replacement modern reply ID must match"));
        }
        update_modern_comment_reply(
            package,
            slide_part_name,
            comment_id,
            reply_id,
            move |reply| {
                *reply = replacement;
            },
        )
    }

    /// Remove a reply from a modern comment thread.
    pub fn remove_modern_comment_reply(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        reply_id: &str,
    ) -> Result<bool> {
        if find_modern_comment_reply(package, slide_part_name, comment_id, reply_id)?.is_none() {
            return Ok(false);
        }
        let mut found = false;
        let result = update_modern_comment(package, slide_part_name, comment_id, |comment| {
            if let Some(index) = comment
                .replies
                .iter()
                .position(|reply| reply.id == reply_id)
            {
                comment.replies.remove(index);
                found = true;
            }
        })?;
        Ok(result && found)
    }

    fn validate_and_commit_modern_comment_part(
        package: &mut OpcPackage,
        staged: &ModernCommentPart,
        all_parts: &[ModernCommentPart],
    ) -> Result<()> {
        validate_modern_comment_graph_for_mutation(package, all_parts)?;
        let xml = staged.comments.to_xml()?;
        ModernCommentList::parse(&xml)?;
        let part_name = PackURI::new(&staged.part_name).map_err(invalid)?;
        package.get_part_mut(&part_name)?.set_blob(xml);
        package.unsign();
        Ok(())
    }

    fn validate_modern_comment_graph_for_mutation(
        package: &OpcPackage,
        parts: &[ModernCommentPart],
    ) -> Result<()> {
        ensure_all_modern_ids_are_unique(parts)?;
        let authors = super::authors::load_modern_comment_authors(package)?;
        super::authors::validate_modern_comment_author_references(authors.as_ref(), parts)
    }

    fn ensure_modern_comment_id_is_free(parts: &[ModernCommentPart], id: &str) -> Result<()> {
        if modern_id_exists(parts, id) {
            Err(invalid(format!(
                "duplicate modern comment or reply ID {id}"
            )))
        } else {
            Ok(())
        }
    }

    fn ensure_modern_reply_ids_are_free(
        parts: &[ModernCommentPart],
        comment: &ModernComment,
    ) -> Result<()> {
        let mut ids = std::collections::HashSet::new();
        for reply in &comment.replies {
            if reply.id == comment.id
                || !ids.insert(reply.id.clone())
                || modern_id_exists(parts, &reply.id)
            {
                return Err(invalid(format!(
                    "duplicate modern comment or reply ID {}",
                    reply.id
                )));
            }
        }
        Ok(())
    }

    fn ensure_all_modern_ids_are_unique(parts: &[ModernCommentPart]) -> Result<()> {
        let mut ids = std::collections::HashSet::new();
        for part in parts {
            for comment in &part.comments.comments {
                if !ids.insert(comment.id.clone()) {
                    return Err(invalid(format!(
                        "duplicate modern comment or reply ID {}",
                        comment.id
                    )));
                }
                for reply in &comment.replies {
                    if !ids.insert(reply.id.clone()) {
                        return Err(invalid(format!(
                            "duplicate modern comment or reply ID {}",
                            reply.id
                        )));
                    }
                }
            }
        }
        Ok(())
    }

    fn modern_id_exists(parts: &[ModernCommentPart], id: &str) -> bool {
        parts.iter().any(|part| {
            part.comments.comments.iter().any(|comment| {
                comment.id == id || comment.replies.iter().any(|reply| reply.id == id)
            })
        })
    }

    fn next_modern_comment_part_name(package: &OpcPackage) -> Result<PackURI> {
        for suffix in 1..=65_537u32 {
            let candidate = PackURI::new(format!("/ppt/comments/modernComment{suffix}.xml"))
                .map_err(invalid)?;
            if package.get_part(&candidate).is_err() {
                return Ok(candidate);
            }
        }
        Err(invalid("no free modern comment part name"))
    }

    fn next_modern_comment_relationship_id(
        package: &OpcPackage,
        slide_part_name: &PackURI,
    ) -> Result<String> {
        let relationships = package.get_part(slide_part_name)?.rels();
        for suffix in 1..=65_537u32 {
            let candidate = format!("rIdModernComments{suffix}");
            if relationships.get(&candidate).is_none() {
                return Ok(candidate);
            }
        }
        Err(invalid("no free modern comment relationship ID"))
    }

    fn modern_comment_part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
        package.iter_parts().any(|part| {
            part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship
                        .target_partname()
                        .is_ok_and(|name| name == *target)
            })
        }) || package.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|name| name == *target)
        })
    }
}

mod authors {
    use super::*;

    fn invalid(message: impl Into<String>) -> Error {
        Error::Invalid(message.into())
    }

    fn limit(label: &str) -> Error {
        invalid(format!("{label} exceeds implementation limit"))
    }

    fn bounded(value: &str) -> Result<()> {
        if value.len() <= super::super::MAX_STRING_BYTES {
            Ok(())
        } else {
            Err(limit("modern Comment Author string bytes"))
        }
    }

    fn validate_relationship_id(value: &str) -> Result<()> {
        bounded(value)?;
        if value.is_empty() || value.chars().any(char::is_whitespace) {
            Err(invalid(
                "modern Comment Author relationship ID must be nonempty without whitespace",
            ))
        } else {
            Ok(())
        }
    }

    fn require_presentation_content_type(value: &str) -> Result<()> {
        if matches!(
            value,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                | "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"
                | "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
                | "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
                | "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml"
                | "application/vnd.ms-powerpoint.template.macroEnabled.main+xml"
        ) {
            Ok(())
        } else {
            Err(invalid(format!(
                "'{value}' is not a PresentationML main content type"
            )))
        }
    }

    fn require_content_type(part: &dyn litchi_opc::Part, expected: &str) -> Result<()> {
        if part.content_type() == expected {
            Ok(())
        } else {
            Err(Error::ContentType {
                expected: expected.into(),
                actual: part.content_type().into(),
            })
        }
    }

    fn require_author_reference(ids: &HashSet<&str>, value: &str, label: &str) -> Result<()> {
        if ids.contains(value) {
            Ok(())
        } else {
            Err(invalid(format!(
                "{label} '{value}' does not resolve in the modern Comment Author part"
            )))
        }
    }

    pub fn load_modern_comment_authors(
        package: &OpcPackage,
    ) -> Result<Option<ModernCommentAuthorPart>> {
        let presentation = package.main_document_part()?;
        require_presentation_content_type(presentation.content_type())?;
        let presentation_name = presentation.partname().to_string();

        if package
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE)
        {
            return Err(invalid(
                "modern Comment Author relationship cannot originate at the package root",
            ));
        }
        for source in package.iter_parts() {
            if source.partname().as_str() != presentation_name.as_str()
                && source.rels().iter().any(|relationship| {
                    relationship.reltype() == MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE
                })
            {
                return Err(invalid(
                    "modern Comment Author relationship has non-Presentation source",
                ));
            }
        }

        let relationships: Vec<_> = presentation
            .rels()
            .iter()
            .filter(|relationship| {
                relationship.reltype() == MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE
            })
            .collect();
        if relationships.len() > 1 {
            return Err(invalid(
                "Presentation has multiple modern Comment Author relationships",
            ));
        }
        let Some(relationship) = relationships.first().copied() else {
            if package
                .iter_parts()
                .any(|part| part.content_type() == MODERN_COMMENT_AUTHOR_CONTENT_TYPE)
            {
                return Err(invalid(
                    "package contains an orphan modern Comment Author part",
                ));
            }
            return Ok(None);
        };
        if relationship.is_external() {
            return Err(invalid(
                "modern Comment Author relationship cannot be external",
            ));
        }
        let target = relationship.target_partname()?;
        let part = package.get_part(&target)?;
        require_content_type(part, MODERN_COMMENT_AUTHOR_CONTENT_TYPE)?;
        if !part.rels().is_empty() {
            return Err(invalid(
                "modern Comment Author part cannot have outbound relationships",
            ));
        }
        if package.iter_parts().any(|candidate| {
            candidate.content_type() == MODERN_COMMENT_AUTHOR_CONTENT_TYPE
                && candidate.partname() != &target
        }) {
            return Err(invalid(
                "package contains an orphan modern Comment Author part",
            ));
        }
        Ok(Some(ModernCommentAuthorPart {
            relationship_id: relationship.r_id().to_owned(),
            part_name: target.to_string(),
            authors: ModernCommentAuthorList::parse(part.blob())?,
        }))
    }

    pub fn load_modern_comment_graph(package: &OpcPackage) -> Result<ModernCommentGraph> {
        let authors = load_modern_comment_authors(package)?;
        let comments = super::comments::load_modern_comments(package)?;
        validate_modern_comment_author_references(authors.as_ref(), &comments)?;
        Ok(ModernCommentGraph { authors, comments })
    }

    /// Validate modeled comment, reply, and assignment author references.
    /// Author-looking values inside opaque extensions remain inert.
    pub fn validate_modern_comment_author_references(
        authors: Option<&ModernCommentAuthorPart>,
        comments: &[ModernCommentPart],
    ) -> Result<()> {
        // Every modeled comment carries an `authorId`, so any comment at all
        // references the Author part.
        let has_references = comments
            .iter()
            .any(|part| !part.comments.comments.is_empty());
        if !has_references {
            return Ok(());
        }
        let authors = authors.ok_or_else(|| {
            invalid("modern comments reference authors but the package has no Author part")
        })?;
        let ids: HashSet<_> = authors
            .authors
            .authors
            .iter()
            .map(|author| author.id.as_str())
            .collect();
        for part in comments {
            for comment in &part.comments.comments {
                require_author_reference(&ids, &comment.author_id, "comment authorId")?;
                if let Some(assigned) = &comment.assigned_to {
                    for author_id in assigned {
                        require_author_reference(&ids, author_id, "comment assignedTo")?;
                    }
                }
                for reply in &comment.replies {
                    require_author_reference(&ids, &reply.author_id, "reply authorId")?;
                }
            }
        }
        Ok(())
    }

    /// Add a new modern Comment Author part after validating the complete graph.
    /// Existing Author parts are deliberately not overwritten.
    pub fn store_modern_comment_authors(
        package: &mut OpcPackage,
        value: &ModernCommentAuthorPart,
    ) -> Result<()> {
        if load_modern_comment_authors(package)?.is_some() {
            return Err(invalid(
                "package already contains a modern Comment Author part",
            ));
        }
        let comments = super::comments::load_modern_comments(package)?;
        validate_modern_comment_author_references(Some(value), &comments)?;
        validate_relationship_id(&value.relationship_id)?;
        let presentation = package.main_document_part()?;
        require_presentation_content_type(presentation.content_type())?;
        if presentation.rels().get(&value.relationship_id).is_some() {
            return Err(invalid(
                "modern Comment Author relationship ID already exists",
            ));
        }
        let presentation_name = presentation.partname().clone();
        let part_name = PackURI::new(&value.part_name).map_err(Error::Invalid)?;
        if package
            .iter_parts()
            .any(|part| part.partname() == &part_name)
        {
            return Err(invalid(format!("part '{part_name}' already exists")));
        }
        let xml = value.authors.to_xml()?;
        let target = part_name.relative_ref(presentation_name.base_uri());
        package.try_add_part(Box::new(BlobPart::new(
            part_name,
            MODERN_COMMENT_AUTHOR_CONTENT_TYPE.into(),
            xml,
        )))?;
        package
            .get_part_mut(&presentation_name)?
            .rels_mut()
            .add_relationship(
                MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE.into(),
                target,
                value.relationship_id.clone(),
                false,
            );
        Ok(())
    }

    pub fn find_modern_comment_author(
        package: &OpcPackage,
        author_id: &str,
    ) -> Result<Option<ModernCommentAuthor>> {
        Ok(load_modern_comment_authors(package)?.and_then(|part| {
            part.authors
                .authors
                .into_iter()
                .find(|author| author.id == author_id)
        }))
    }

    /// Add a modern comment author, allocating a collision-safe part and relationship if needed.
    pub fn add_modern_comment_author(
        package: &mut OpcPackage,
        author: ModernCommentAuthor,
    ) -> Result<ModernCommentAuthorPart> {
        let mut graph = load_modern_comment_graph(package)?;
        if graph
            .authors
            .as_ref()
            .is_some_and(|part| part.authors.authors.iter().any(|item| item.id == author.id))
        {
            return Err(invalid(format!(
                "duplicate modern comment author ID {}",
                author.id
            )));
        }

        if let Some(mut part) = graph.authors.take() {
            part.authors.authors.push(author);
            commit_modern_comment_authors(package, &part, &graph.comments)?;
            return Ok(part);
        }

        let presentation_name = package.main_document_part()?.partname().clone();
        let part_name = next_modern_author_part_name(package)?;
        let relationship_id = next_modern_author_relationship_id(package, &presentation_name)?;
        let part = ModernCommentAuthorPart {
            relationship_id: relationship_id.clone(),
            part_name: part_name.to_string(),
            authors: ModernCommentAuthorList {
                root_prefix: "p188".into(),
                namespace_declarations: Vec::new(),
                authors: vec![author],
            },
        };
        validate_modern_comment_author_references(Some(&part), &graph.comments)?;
        let xml = part.authors.to_xml()?;
        ModernCommentAuthorList::parse(&xml)?;
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part_name.clone(),
            MODERN_COMMENT_AUTHOR_CONTENT_TYPE.into(),
            xml,
        )))?;
        package
            .get_part_mut(&presentation_name)?
            .rels_mut()
            .add_relationship(
                MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE.into(),
                part_name.relative_ref(presentation_name.base_uri()),
                relationship_id,
                false,
            );
        package.unsign();
        Ok(part)
    }

    /// Update a modern comment author while keeping its stable GUID unchanged.
    pub fn update_modern_comment_author<F>(
        package: &mut OpcPackage,
        author_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut ModernCommentAuthor),
    {
        let graph = load_modern_comment_graph(package)?;
        let Some(mut part) = graph.authors.clone() else {
            return Ok(false);
        };
        let Some(author) = part
            .authors
            .authors
            .iter_mut()
            .find(|item| item.id == author_id)
        else {
            return Ok(false);
        };
        update(author);
        if author.id != author_id {
            return Err(invalid("modern comment author update cannot change its ID"));
        }
        commit_modern_comment_authors(package, &part, &graph.comments)?;
        Ok(true)
    }

    /// Replace a modern comment author while keeping its stable GUID unchanged.
    pub fn replace_modern_comment_author(
        package: &mut OpcPackage,
        author_id: &str,
        replacement: ModernCommentAuthor,
    ) -> Result<bool> {
        if replacement.id != author_id {
            return Err(invalid("replacement modern comment author ID must match"));
        }
        update_modern_comment_author(package, author_id, move |author| *author = replacement)
    }

    /// Remove an unreferenced modern comment author.
    pub fn remove_modern_comment_author(package: &mut OpcPackage, author_id: &str) -> Result<bool> {
        let graph = load_modern_comment_graph(package)?;
        let Some(mut part) = graph.authors.clone() else {
            return Ok(false);
        };
        let Some(index) = part
            .authors
            .authors
            .iter()
            .position(|author| author.id == author_id)
        else {
            return Ok(false);
        };
        if modern_author_is_referenced(&graph.comments, author_id) {
            return Err(invalid(format!(
                "modern comment author {author_id} is still referenced"
            )));
        }
        part.authors.authors.remove(index);
        if !part.authors.authors.is_empty() {
            commit_modern_comment_authors(package, &part, &graph.comments)?;
            return Ok(true);
        }

        let presentation_name = package.main_document_part()?.partname().clone();
        let part_name = PackURI::new(&part.part_name).map_err(invalid)?;
        package
            .get_part_mut(&presentation_name)?
            .rels_mut()
            .remove(&part.relationship_id);
        if !modern_author_part_is_referenced(package, &part_name) {
            package.remove_part(&part_name);
        }
        package.unsign();
        Ok(true)
    }

    /// Reorder modern comment authors by a complete, duplicate-free GUID list.
    pub fn reorder_modern_comment_authors(
        package: &mut OpcPackage,
        ordered_author_ids: &[String],
    ) -> Result<Vec<ModernCommentAuthor>> {
        let graph = load_modern_comment_graph(package)?;
        let Some(mut part) = graph.authors.clone() else {
            if ordered_author_ids.is_empty() {
                return Ok(Vec::new());
            }
            return Err(invalid("modern comment author part is missing"));
        };
        if ordered_author_ids.len() != part.authors.authors.len() {
            return Err(invalid("modern author reorder must contain every author"));
        }
        let mut remaining = std::collections::HashMap::new();
        for author in part.authors.authors.drain(..) {
            if remaining.insert(author.id.clone(), author).is_some() {
                return Err(invalid("duplicate modern comment author ID"));
            }
        }
        let mut ordered = Vec::with_capacity(ordered_author_ids.len());
        for id in ordered_author_ids {
            let author = remaining.remove(id).ok_or_else(|| {
                invalid(format!(
                    "unknown or duplicate modern comment author ID {id}"
                ))
            })?;
            ordered.push(author);
        }
        if !remaining.is_empty() {
            return Err(invalid("modern author reorder must contain every author"));
        }
        part.authors.authors = ordered.clone();
        commit_modern_comment_authors(package, &part, &graph.comments)?;
        Ok(ordered)
    }

    fn commit_modern_comment_authors(
        package: &mut OpcPackage,
        part: &ModernCommentAuthorPart,
        comments: &[ModernCommentPart],
    ) -> Result<()> {
        validate_modern_comment_author_references(Some(part), comments)?;
        let xml = part.authors.to_xml()?;
        ModernCommentAuthorList::parse(&xml)?;
        let part_name = PackURI::new(&part.part_name).map_err(invalid)?;
        package.get_part_mut(&part_name)?.set_blob(xml);
        package.unsign();
        Ok(())
    }

    fn modern_author_is_referenced(comments: &[ModernCommentPart], author_id: &str) -> bool {
        comments.iter().any(|part| {
            part.comments.comments.iter().any(|comment| {
                comment.author_id == author_id
                    || comment
                        .assigned_to
                        .as_ref()
                        .is_some_and(|ids| ids.iter().any(|id| id == author_id))
                    || comment
                        .replies
                        .iter()
                        .any(|reply| reply.author_id == author_id)
            })
        })
    }

    fn next_modern_author_part_name(package: &OpcPackage) -> Result<PackURI> {
        for suffix in 1..=65_537u32 {
            let candidate =
                PackURI::new(format!("/ppt/authors/author{suffix}.xml")).map_err(invalid)?;
            if package.get_part(&candidate).is_err() {
                return Ok(candidate);
            }
        }
        Err(invalid("no free modern comment author part name"))
    }

    fn next_modern_author_relationship_id(
        package: &OpcPackage,
        presentation_name: &PackURI,
    ) -> Result<String> {
        let relationships = package.get_part(presentation_name)?.rels();
        for suffix in 1..=65_537u32 {
            let candidate = format!("rIdModernAuthors{suffix}");
            if relationships.get(&candidate).is_none() {
                return Ok(candidate);
            }
        }
        Err(invalid("no free modern comment author relationship ID"))
    }

    fn modern_author_part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
        package.iter_parts().any(|part| {
            part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship
                        .target_partname()
                        .is_ok_and(|name| name == *target)
            })
        }) || package.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|name| name == *target)
        })
    }
}

pub use authors::{
    add_modern_comment_author, find_modern_comment_author, load_modern_comment_authors,
    load_modern_comment_graph, remove_modern_comment_author, reorder_modern_comment_authors,
    replace_modern_comment_author, store_modern_comment_authors, update_modern_comment_author,
    validate_modern_comment_author_references,
};
pub use comments::{
    add_modern_comment, add_modern_comment_reply, find_modern_comment, find_modern_comment_reply,
    load_modern_comments, remove_modern_comment, remove_modern_comment_reply,
    reorder_modern_comments, replace_modern_comment, replace_modern_comment_reply,
    store_modern_comment, update_modern_comment, update_modern_comment_reply,
};

#[cfg(test)]
mod comment_tests {
    use super::*;
    use crate::modern_comments::{AC, MAX_BYTES, P188, PC};
    use litchi_opc::Part as _;

    const AUTHOR: &str = "{CD37207E-7903-4ED4-8AE8-017538D2DF7E}";
    const COMMENT: &str = "{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}";
    const REPLY: &str = "{E524A04C-CF22-45D7-A60D-09322EA5A80D}";

    fn sdk_xml() -> Vec<u8> {
        format!(r#"<p188:cmLst xmlns:p188="{P188}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Needs more cowbell</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>"#).into_bytes()
    }

    fn package() -> OpcPackage {
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "ppt/presentation.xml".into(),
            "rId1".into(),
            false,
        );
        let mut presentation = BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                .into(),
            Vec::new(),
        );
        presentation.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide".into(),
            "slides/slide1.xml".into(),
            "rId1".into(),
            false,
        );
        package.add_part(Box::new(presentation));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/slides/slide1.xml").unwrap(),
            SLIDE_CONTENT_TYPE.into(),
            Vec::new(),
        )));
        package
    }

    fn value() -> ModernCommentPart {
        ModernCommentPart {
            slide_part_name: "/ppt/slides/slide1.xml".into(),
            relationship_id: "rId9".into(),
            part_name: "/ppt/comments/modernComment1.xml".into(),
            comments: ModernCommentList::parse(&sdk_xml()).unwrap(),
        }
    }

    #[test]
    fn loads_microsoft_open_xml_sdk_documentation_specimen() {
        let parsed = ModernCommentList::parse(&sdk_xml()).unwrap();
        assert_eq!(parsed.comments.len(), 1);
        assert_eq!(parsed.comments[0].id, COMMENT);
        assert!(
            std::str::from_utf8(parsed.comments[0].text_body_xml.as_ref().unwrap())
                .unwrap()
                .contains("Needs more cowbell")
        );
        assert_eq!(
            ModernCommentList::parse(&parsed.to_xml().unwrap()).unwrap(),
            parsed
        );
    }

    #[test]
    fn package_round_trip_keeps_monikers_replies_and_extensions_inert() {
        let xml = format!(
            r#"<p188:cmLst xmlns:p188="{P188}" xmlns:pc="{PC}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:payload"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" status="resolved" created="2026-07-19T12:00:00+08:00" assignedTo="{AUTHOR}" complete="50%" title="Review"><pc:sldMkLst><pc:sldMk/></pc:sldMkLst><p188:pos x="10" y="-20"/><p188:replyLst><p188:reply id="{REPLY}" authorId="{AUTHOR}" created="2026-07-19T12:01:00+08:00"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p/></p188:txBody><p188:extLst><p:ext uri="{{A}}"><x:data relationship="rId999"/></p:ext></p188:extLst></p188:reply></p188:replyLst><p188:extLst><p:ext uri="{{B}}"><x:payload r:id="rId666" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></p:ext></p188:extLst></p188:cm></p188:cmLst>"#
        );
        let expected = ModernCommentList::parse(xml.as_bytes()).unwrap();
        assert_eq!(
            expected.comments[0].complete,
            Some(Progress::new(50).unwrap())
        );
        let mut package = package();
        let mut part = value();
        part.comments = expected.clone();
        store_modern_comment(&mut package, &part).unwrap();
        let loaded = load_modern_comments(&package).unwrap();
        assert_eq!(loaded.len(), 1);
        assert_eq!(loaded[0].comments, expected);
        assert!(
            loaded[0].comments.comments[0]
                .extension_xml
                .as_ref()
                .unwrap()
                .windows(6)
                .any(|window| window == b"rId666")
        );
        assert!(
            package
                .get_part(&PackURI::new("/ppt/comments/modernComment1.xml").unwrap())
                .unwrap()
                .rels()
                .is_empty()
        );
    }

    #[test]
    fn progress_is_bounded_typed_and_written_in_office_units() {
        assert_eq!(
            std::mem::size_of::<Option<Progress>>(),
            std::mem::size_of::<u32>()
        );
        assert_eq!(Progress::ZERO.thousandths(), 0);
        assert_eq!(Progress::FULL.thousandths(), 100_000);
        assert_eq!(Progress::new(25).unwrap().thousandths(), 25_000);
        assert_eq!(
            Progress::from_thousandths(50_250).unwrap().to_string(),
            "50250"
        );
        assert!(Progress::new(101).is_err());
        assert!(Progress::from_thousandths(100_001).is_err());

        let xml = format!(
            r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503" complete="50.25%"/></p188:cmLst>"#
        );
        let parsed = ModernCommentList::parse(xml.as_bytes()).unwrap();
        assert_eq!(
            parsed.comments[0].complete,
            Some(Progress::from_thousandths(50_250).unwrap())
        );
        let serialized = parsed.to_xml().unwrap();
        assert!(
            serialized
                .windows(b"complete=\"50250\"".len())
                .any(|window| window == b"complete=\"50250\"")
        );
        assert_eq!(ModernCommentList::parse(&serialized).unwrap(), parsed);

        for complete in ["-1%", "100.01%", "50.123%", "1e2%", "100001", ""] {
            let xml = format!(
                r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503" complete="{complete}"/></p188:cmLst>"#
            );
            assert!(
                ModernCommentList::parse(xml.as_bytes()).is_err(),
                "accepted invalid progress {complete:?}"
            );
        }
    }

    #[test]
    fn rejects_hostile_or_schema_invalid_comment_xml() {
        let cases = [
            format!(r#"<!DOCTYPE x><p188:cmLst xmlns:p188="{P188}"/>"#),
            r#"<x:cmLst xmlns:x="urn:wrong"/>"#.to_string(),
            format!(
                r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="bad" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503"/></p188:cmLst>"#
            ),
            format!(
                r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" status="pending" created="2024-12-30T20:26:06.503"/></p188:cmLst>"#
            ),
            format!(
                r#"<p188:cmLst xmlns:p188="{P188}" xmlns:pc="{PC}" xmlns:ac="{AC}"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503"><pc:sldMkLst/><ac:deMkLst/></p188:cm></p188:cmLst>"#
            ),
            format!(
                r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503"><p188:txBody/><p188:replyLst/></p188:cm></p188:cmLst>"#
            ),
        ];
        for xml in cases {
            assert!(
                ModernCommentList::parse(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
        assert!(ModernCommentList::parse(&vec![b' '; MAX_BYTES + 1]).is_err());
    }

    #[test]
    fn rejects_invalid_package_graphs_and_failed_store_is_atomic() {
        let mut external = package();
        external
            .get_part_mut(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                MODERN_COMMENT_RELATIONSHIP_TYPE.into(),
                "https://invalid.example/comments.xml".into(),
                "rId9".into(),
                true,
            );
        assert!(load_modern_comments(&external).is_err());

        let mut wrong_source = package();
        wrong_source
            .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                MODERN_COMMENT_RELATIONSHIP_TYPE.into(),
                "comments/modern.xml".into(),
                "rId9".into(),
                false,
            );
        assert!(load_modern_comments(&wrong_source).is_err());

        let mut orphan = package();
        orphan.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/comments/orphan.xml").unwrap(),
            MODERN_COMMENT_CONTENT_TYPE.into(),
            sdk_xml(),
        )));
        assert!(load_modern_comments(&orphan).is_err());

        let mut atomic = package();
        let mut invalid_value = value();
        invalid_value.comments.comments[0].id = "not-a-guid".into();
        let before_parts = atomic.iter_parts().count();
        let before_rels = atomic
            .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
            .unwrap()
            .rels()
            .len();
        assert!(store_modern_comment(&mut atomic, &invalid_value).is_err());
        assert_eq!(atomic.iter_parts().count(), before_parts);
        assert_eq!(
            atomic
                .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
                .unwrap()
                .rels()
                .len(),
            before_rels
        );
    }
}

#[cfg(test)]
mod author_tests {
    use super::*;
    use crate::modern_comments::{MAX_BYTES, P188};
    use crate::modern_comments::{ModernCommentList, store_modern_comment};
    use litchi_opc::Part as _;

    const AUTHOR: &str = "{CD37207E-7903-4ED4-8AE8-017538D2DF7E}";
    const OTHER: &str = "{0B2043D4-0908-4C42-8A79-51EA2CC309F7}";
    const COMMENT: &str = "{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}";

    fn sdk_author_xml() -> Vec<u8> {
        format!(r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="{AUTHOR}" name="Ada Lovelace" initials="AL" userId="ada@example.com::4b640067-2830-4c10-9c4f-5879bb2e41d1" providerId=""/></p188:authorLst>"#).into_bytes()
    }

    fn package() -> OpcPackage {
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "ppt/presentation.xml".into(),
            "rId1".into(),
            false,
        );
        let mut presentation = BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                .into(),
            Vec::new(),
        );
        presentation.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide".into(),
            "slides/slide1.xml".into(),
            "rId1".into(),
            false,
        );
        package.add_part(Box::new(presentation));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/slides/slide1.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
            Vec::new(),
        )));
        package
    }

    fn author_part() -> ModernCommentAuthorPart {
        ModernCommentAuthorPart {
            relationship_id: "rId8".into(),
            part_name: "/ppt/authors/author1.xml".into(),
            authors: ModernCommentAuthorList::parse(&sdk_author_xml()).unwrap(),
        }
    }

    fn comment_part(author: &str) -> ModernCommentPart {
        let xml = format!(
            r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="{COMMENT}" authorId="{author}" created="2024-12-30T20:26:06.503" assignedTo="{author}"/></p188:cmLst>"#
        );
        ModernCommentPart {
            slide_part_name: "/ppt/slides/slide1.xml".into(),
            relationship_id: "rId9".into(),
            part_name: "/ppt/comments/modernComment1.xml".into(),
            comments: ModernCommentList::parse(xml.as_bytes()).unwrap(),
        }
    }

    #[test]
    fn loads_open_xml_sdk_shaped_author_specimen() {
        let parsed = ModernCommentAuthorList::parse(&sdk_author_xml()).unwrap();
        assert_eq!(parsed.authors.len(), 1);
        assert_eq!(parsed.authors[0].name, "Ada Lovelace");
        assert_eq!(parsed.authors[0].provider_id, "");
        assert_eq!(
            ModernCommentAuthorList::parse(&parsed.to_xml().unwrap()).unwrap(),
            parsed
        );
    }

    #[test]
    fn author_and_comment_package_graph_round_trip_and_resolve() {
        let extension = format!(r#"<p188:extLst xmlns:p188="{P188}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:payload"><p:ext uri="{{A}}"><x:data authorId="{OTHER}" relationship="rId999"/></p:ext></p188:extLst>"#).into_bytes();
        let mut authors = author_part();
        authors.authors.authors[0].extension_xml = Some(extension.clone());
        let mut package = package();
        store_modern_comment_authors(&mut package, &authors).unwrap();
        store_modern_comment(&mut package, &comment_part(AUTHOR)).unwrap();
        let graph = load_modern_comment_graph(&package).unwrap();
        assert_eq!(
            graph.authors.unwrap().authors.authors[0].extension_xml,
            Some(extension)
        );
        assert_eq!(graph.comments.len(), 1);
    }

    #[test]
    fn rejects_hostile_author_grammar_and_unresolved_modeled_references() {
        let cases = [
            format!(r#"<!DOCTYPE x><p188:authorLst xmlns:p188="{P188}"/>"#),
            "<x:authorLst xmlns:x=\"urn:wrong\"/>".into(),
            format!(
                r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="bad" name="A" userId="u" providerId="p"/></p188:authorLst>"#
            ),
            format!(
                r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="{AUTHOR}" name="A" userId="u"/></p188:authorLst>"#
            ),
            format!(
                r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="{AUTHOR}" name="A" userId="u" providerId="p"><p188:extLst/><p188:extLst/></p188:author></p188:authorLst>"#
            ),
            format!(
                r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="{AUTHOR}" name="A" userId="u" providerId="p"/><p188:author id="{AUTHOR}" name="B" userId="v" providerId="p"/></p188:authorLst>"#
            ),
        ];
        for xml in cases {
            assert!(
                ModernCommentAuthorList::parse(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
        assert!(ModernCommentAuthorList::parse(&vec![b' '; MAX_BYTES + 1]).is_err());

        let authors = author_part();
        assert!(
            validate_modern_comment_author_references(Some(&authors), &[comment_part(OTHER)])
                .is_err()
        );
        assert!(validate_modern_comment_author_references(None, &[comment_part(AUTHOR)]).is_err());
    }

    #[test]
    fn rejects_author_package_graphs_and_failed_store_is_atomic() {
        let mut external = package();
        external
            .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE.into(),
                "https://invalid.example/authors.xml".into(),
                "rId8".into(),
                true,
            );
        assert!(load_modern_comment_authors(&external).is_err());

        let mut orphan = package();
        orphan.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/authors/orphan.xml").unwrap(),
            MODERN_COMMENT_AUTHOR_CONTENT_TYPE.into(),
            sdk_author_xml(),
        )));
        assert!(load_modern_comment_authors(&orphan).is_err());

        let mut outbound = package();
        store_modern_comment_authors(&mut outbound, &author_part()).unwrap();
        outbound
            .get_part_mut(&PackURI::new("/ppt/authors/author1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:forbidden".into(),
                "other.xml".into(),
                "rId1".into(),
                false,
            );
        assert!(load_modern_comment_authors(&outbound).is_err());

        let mut atomic = package();
        store_modern_comment(&mut atomic, &comment_part(OTHER)).unwrap();
        let before_parts = atomic.iter_parts().count();
        let before_rels = atomic
            .get_part(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .rels()
            .len();
        assert!(store_modern_comment_authors(&mut atomic, &author_part()).is_err());
        assert_eq!(atomic.iter_parts().count(), before_parts);
        assert_eq!(
            atomic
                .get_part(&PackURI::new("/ppt/presentation.xml").unwrap())
                .unwrap()
                .rels()
                .len(),
            before_rels
        );
    }
}
