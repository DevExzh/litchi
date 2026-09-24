# Local specification scope

The copied Markdown files are the local checked-in specifications used for this
PPTX `p14:bwMode` review. Their original repository paths and relevant ranges are
listed below.

| Bundle file | Original path | Relevant range |
| --- | --- | --- |
| `ms-pptx-bw-mode.md` | `3rdparty/specs/[MS-PPTX]/2 Structures/2.3 http---schemas.microsoft.com-office-powerpoint-2010-main.md` | §2.3.2.2, lines 553-561: p14 `bwMode`, `a:ST_BlackWhiteMode`, rendering interpretation |
| `ms-pptx-extensions.md` | `3rdparty/specs/[MS-PPTX]/2 Structures/2.2 Extensions.md` | §2.2.3, lines 118-128: `p:contentPart` owner; §2.2.4, lines 159-181: `p14:media` and `tracksInfo` |
| `ms-pptx-schema.md` | `3rdparty/specs/[MS-PPTX]/5 Appendix A - Full XML Schemas/5.1 http---schemas.microsoft.com-office-powerpoint-2010-main Schema.md` | lines 193-201: top-level qualified p14 attribute declaration and no explicit default |
| `ms-odrawxml-content-parts.md` | `3rdparty/specs/[MS-ODRAWXML]/3 Structure Examples/3.2 Content Parts and Ink.md` | lines 179-190: native `p14:bwMode` under `mc:Choice` |

The codec scope is limited to inert metadata. It does not claim playback,
rendering, signature trust, or native producer acceptance.
