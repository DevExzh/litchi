"""Build the change 0653 witness pair for the DOCX text sink's cumulative
namespace-binding limit.

One `word/document.xml` whose root carries 33 declarations including
`xmlns:mc`, and the same bytes with the markup-compatibility URI replaced by an
inert URI of exactly the same length, so the two packages differ only in which
branch of `process_markup_compatibility` the main part takes. This is change
0664's marker/control technique, reproduced at the smallest size that shows the
limit.
"""
import sys, zipfile

MCE = "http://schemas.openxmlformats.org/markup-compatibility/2006"
INERT = "http://schemas.openxmlformats.org/markup-kompatibility/2006"
assert len(MCE) == len(INERT)
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""


def document(namespace, paragraphs):
    declarations = "".join(
        f' xmlns:n{index}="urn:litchi:perf:0653:{index}"' for index in range(31)
    )
    body = "".join(
        f"<w:p><w:r><w:t>paragraph {index}</w:t></w:r></w:p>" for index in range(paragraphs)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}" xmlns:mc="{namespace}"{declarations}'
        f' mc:Ignorable="n0"><w:body>{body}</w:body></w:document>'
    )


out_dir, paragraphs = sys.argv[1], int(sys.argv[2])
for label, namespace in (("marker", MCE), ("control", INERT)):
    path = f"{out_dir}/sink-{label}.docx"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/document.xml", document(namespace, paragraphs))
    print(label, path)
