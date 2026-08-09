//! OPC graph lifecycle and transactional CRUD for modern comments.

use super::model::{Author, AuthorPart, Authors, Comment, Graph, List, Part, Reply};
use super::semantic::extensions::List as Extensions;
use super::{
    MODERN_COMMENT_AUTHOR_CONTENT_TYPE, MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE,
    MODERN_COMMENT_CONTENT_TYPE, MODERN_COMMENT_RELATIONSHIP_TYPE, SLIDE_CONTENT_TYPE,
};
use crate::{Error, Result};
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use std::collections::HashSet;

mod comments {
    use super::{
        BlobPart, Comment, Error, Extensions, HashSet, List, MODERN_COMMENT_CONTENT_TYPE,
        MODERN_COMMENT_RELATIONSHIP_TYPE, OpcPackage, PackURI, Part, Reply, Result,
        SLIDE_CONTENT_TYPE, load_modern_comment_authors, validate_modern_comment_author_references,
    };

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

    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn load_modern_comments(package: &OpcPackage) -> Result<Vec<Part>> {
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
                let comments = List::parse(package.get_part(&uri)?.blob())?;
                Ok(Part {
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
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn store_modern_comment(package: &mut OpcPackage, value: &Part) -> Result<()> {
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

    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn find_modern_comment(
        package: &OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
    ) -> Result<Option<Comment>> {
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_modern_comment(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment: Comment,
    ) -> Result<Part> {
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
        let staged = Part {
            slide_part_name: slide_part_name.to_string(),
            relationship_id: relationship_id.clone(),
            part_name: part_name.to_string(),
            comments: List {
                root_prefix: "p188".into(),
                namespace_declarations: Vec::new(),
                comments: vec![comment],
            },
        };
        parts.push(staged.clone());
        validate_modern_comment_graph_for_mutation(package, &parts)?;
        let xml = staged.comments.to_xml()?;
        List::parse(&xml)?;
        package.try_add_part(Box::new(BlobPart::new(
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update_modern_comment<F>(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut Comment),
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_modern_comment(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        replacement: Comment,
    ) -> Result<bool> {
        if replacement.id != comment_id {
            return Err(invalid("replacement modern comment ID must match"));
        }
        update_modern_comment(package, slide_part_name, comment_id, move |comment| {
            *comment = replacement;
        })
    }

    /// Remove a modern comment and remove an empty per-slide part unless another owner shares it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reorder_modern_comments(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        ordered_comment_ids: &[String],
    ) -> Result<Vec<Comment>> {
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn find_modern_comment_reply(
        package: &OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        reply_id: &str,
    ) -> Result<Option<Reply>> {
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_modern_comment_reply(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        reply: Reply,
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update_modern_comment_reply<F>(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        reply_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut Reply),
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

    /// Read the typed extension envelope for one comment without exposing
    /// relationship or collaboration behavior.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn load_modern_comment_extensions(
        package: &OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
    ) -> Result<Option<Extensions>> {
        find_modern_comment(package, slide_part_name, comment_id)?
            .map(|comment| comment.extensions())
            .transpose()
    }

    /// Transactionally update one comment's typed task/reaction extension
    /// envelope while retaining all opaque extension entries.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update_modern_comment_extensions<F>(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut Extensions),
    {
        let mut update = Some(update);
        let mut failure = None;
        let result =
            update_modern_comment(
                package,
                slide_part_name,
                comment_id,
                |comment| match comment.extensions() {
                    Ok(mut extensions) => {
                        if let Some(update) = update.take() {
                            update(&mut extensions);
                        }
                        if let Err(error) = comment.set_extensions(extensions) {
                            failure = Some(error);
                        }
                    },
                    Err(error) => failure = Some(error),
                },
            )?;
        if let Some(error) = failure {
            return Err(error);
        }
        Ok(result)
    }

    /// Read the typed extension envelope for one reply.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn load_modern_comment_reply_extensions(
        package: &OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        reply_id: &str,
    ) -> Result<Option<Extensions>> {
        find_modern_comment_reply(package, slide_part_name, comment_id, reply_id)?
            .map(|reply| reply.extensions())
            .transpose()
    }

    /// Transactionally update one reply's typed reaction extension envelope.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update_modern_comment_reply_extensions<F>(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        reply_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut Extensions),
    {
        let mut update = Some(update);
        let mut failure = None;
        let result =
            update_modern_comment_reply(package, slide_part_name, comment_id, reply_id, |reply| {
                match reply.extensions() {
                    Ok(mut extensions) => {
                        if let Some(update) = update.take() {
                            update(&mut extensions);
                        }
                        if let Err(error) = reply.set_extensions(extensions) {
                            failure = Some(error);
                        }
                    },
                    Err(error) => failure = Some(error),
                }
            })?;
        if let Some(error) = failure {
            return Err(error);
        }
        Ok(result)
    }

    /// Replace a reply without permitting its stable GUID to change.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_modern_comment_reply(
        package: &mut OpcPackage,
        slide_part_name: &PackURI,
        comment_id: &str,
        reply_id: &str,
        replacement: Reply,
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
        staged: &Part,
        all_parts: &[Part],
    ) -> Result<()> {
        validate_modern_comment_graph_for_mutation(package, all_parts)?;
        let xml = staged.comments.to_xml()?;
        List::parse(&xml)?;
        let part_name = PackURI::new(&staged.part_name).map_err(invalid)?;
        package.get_part_mut(&part_name)?.set_blob(xml);
        package.unsign();
        Ok(())
    }

    fn validate_modern_comment_graph_for_mutation(
        package: &OpcPackage,
        parts: &[Part],
    ) -> Result<()> {
        ensure_all_modern_ids_are_unique(parts)?;
        let authors = load_modern_comment_authors(package)?;
        validate_modern_comment_author_references(authors.as_ref(), parts)
    }

    fn ensure_modern_comment_id_is_free(parts: &[Part], id: &str) -> Result<()> {
        if modern_id_exists(parts, id) {
            Err(invalid(format!(
                "duplicate modern comment or reply ID {id}"
            )))
        } else {
            Ok(())
        }
    }

    fn ensure_modern_reply_ids_are_free(parts: &[Part], comment: &Comment) -> Result<()> {
        let mut ids = HashSet::new();
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

    fn ensure_all_modern_ids_are_unique(parts: &[Part]) -> Result<()> {
        let mut ids = HashSet::new();
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

    fn modern_id_exists(parts: &[Part], id: &str) -> bool {
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
    use super::{
        Author, AuthorPart, Authors, BlobPart, Error, Graph, HashSet,
        MODERN_COMMENT_AUTHOR_CONTENT_TYPE, MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE, OpcPackage,
        PackURI, Part, Result, load_modern_comments,
    };

    /// # Errors
    ///
    /// Returns an error if the operation fails.
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

    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn load_modern_comment_authors(package: &OpcPackage) -> Result<Option<AuthorPart>> {
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
        Ok(Some(AuthorPart {
            relationship_id: relationship.r_id().to_owned(),
            part_name: target.to_string(),
            authors: Authors::parse(part.blob())?,
        }))
    }

    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn load_modern_comment_graph(package: &OpcPackage) -> Result<Graph> {
        let authors = load_modern_comment_authors(package)?;
        let comments = load_modern_comments(package)?;
        validate_modern_comment_author_references(authors.as_ref(), &comments)?;
        Ok(Graph { authors, comments })
    }

    /// Validate modeled comment, reply, and assignment author references.
    /// Author-looking values inside opaque extensions remain inert.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn validate_modern_comment_author_references(
        authors: Option<&AuthorPart>,
        comments: &[Part],
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
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn store_modern_comment_authors(
        package: &mut OpcPackage,
        value: &AuthorPart,
    ) -> Result<()> {
        if load_modern_comment_authors(package)?.is_some() {
            return Err(invalid(
                "package already contains a modern Comment Author part",
            ));
        }
        let comments = load_modern_comments(package)?;
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

    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn find_modern_comment_author(
        package: &OpcPackage,
        author_id: &str,
    ) -> Result<Option<Author>> {
        Ok(load_modern_comment_authors(package)?.and_then(|part| {
            part.authors
                .authors
                .into_iter()
                .find(|author| author.id == author_id)
        }))
    }

    /// Add a modern comment author, allocating a collision-safe part and relationship if needed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_modern_comment_author(
        package: &mut OpcPackage,
        author: Author,
    ) -> Result<AuthorPart> {
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
        let part = AuthorPart {
            relationship_id: relationship_id.clone(),
            part_name: part_name.to_string(),
            authors: Authors {
                root_prefix: "p188".into(),
                namespace_declarations: Vec::new(),
                authors: vec![author],
            },
        };
        validate_modern_comment_author_references(Some(&part), &graph.comments)?;
        let xml = part.authors.to_xml()?;
        Authors::parse(&xml)?;
        package.try_add_part(Box::new(BlobPart::new(
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update_modern_comment_author<F>(
        package: &mut OpcPackage,
        author_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut Author),
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_modern_comment_author(
        package: &mut OpcPackage,
        author_id: &str,
        replacement: Author,
    ) -> Result<bool> {
        if replacement.id != author_id {
            return Err(invalid("replacement modern comment author ID must match"));
        }
        update_modern_comment_author(package, author_id, move |author| *author = replacement)
    }

    /// Remove an unreferenced modern comment author.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reorder_modern_comment_authors(
        package: &mut OpcPackage,
        ordered_author_ids: &[String],
    ) -> Result<Vec<Author>> {
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
        part: &AuthorPart,
        comments: &[Part],
    ) -> Result<()> {
        validate_modern_comment_author_references(Some(part), comments)?;
        let xml = part.authors.to_xml()?;
        Authors::parse(&xml)?;
        let part_name = PackURI::new(&part.part_name).map_err(invalid)?;
        package.get_part_mut(&part_name)?.set_blob(xml);
        package.unsign();
        Ok(())
    }

    fn modern_author_is_referenced(comments: &[Part], author_id: &str) -> bool {
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
    load_modern_comment_extensions, load_modern_comment_reply_extensions, load_modern_comments,
    remove_modern_comment, remove_modern_comment_reply, reorder_modern_comments,
    replace_modern_comment, replace_modern_comment_reply, store_modern_comment,
    update_modern_comment, update_modern_comment_extensions, update_modern_comment_reply,
    update_modern_comment_reply_extensions,
};
