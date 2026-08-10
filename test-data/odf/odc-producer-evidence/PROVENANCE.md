# Standalone ODC producer evidence: ODFDOM 0.13.0

Generated and verified on 2026-08-10. These files are genuine standalone
OpenDocument Chart packages created and saved through the public ODFDOM API.
They are not extracted embedded charts, renamed XML, hand-built ZIP packages,
or evidence inferred from a `meta:generator` string.

## Artifacts

| File | SHA-256 | Meaning |
|---|---|---|
| `odfdom-created.odc` | `628e75dd2a9072c102d06412d9353fabb40efc24fa160d33055188b525a081d6` | Created by `OdfChartDocument.newChartDocument()`, populated with generated typed chart DOM elements, and saved by ODFDOM. |
| `odfdom-resaved.odc` | `985efb04560d41d3fe542d1592bd3266979eb5092f08c2ac4f162dd8a3e5ab33` | Independently loaded from the first artifact by a fresh JVM with `OdfChartDocument.loadDocument(File)`, changed, and saved to a new package. |
| `OdfdomChartProducer.java` | `f66a9bb52b1b3d29a35dec61cf17018d91ff10963093c095bd6d365d2eadeae4` | Exact producer, resave, and semantic-reopen source. Licensed Apache-2.0 as marked in the file. |
| `validator-created.txt` | `51459341d741441a040cfe2b8be43c29d20e48e0c3cfefff405c8b376374f315` | Exact verbose ODF Validator transcript for the created package. |
| `validator-resaved.txt` | `6996d2327be8d8dd5accc8deb9602fc0a23fa97c4835cffeca6fbf4481108fd8` | Exact verbose ODF Validator transcript for the independently resaved package. |

The producer source deliberately disables automatic metadata updates and
removes legacy creator/generator fields inherited from ODFDOM's bundled chart
template. Producer identity comes from the pinned library, executable source,
commands, and hashes below—not document metadata.

## Producer and validator provenance

The official [ODF Toolkit](https://odftoolkit.org/) describes ODFDOM as its
Java creation/manipulation API and the ODF Validator as its conformance checker.
The release inputs came directly from Maven Central under the
`org.odftoolkit` group:

| Input | SHA-256 |
|---|---|
| `odfdom-java-0.13.0-jar-with-dependencies.jar` | `6580b36c9e6b03e3f97343bab01e9118d2f303935eea0c977ed2ae6b3155384a` |
| `odfdom-java-0.13.0-sources.jar` | `4f48ac069da4eb02ea14ec714eb92058c29310263b24db3287e33d2c27393295` |
| `odfdom-java-0.13.0.pom` | `108ffc8649de6289be17a55abf19a2314f512d7f5b9e5035b567616399daec5d` |
| `odfvalidator-0.13.0-jar-with-dependencies.jar` | `5684feec5cbdcd5783998978c096ac9ccea53a454e2d6ae803ce482d2336d1dc` |

Every hash matched the adjacent `.sha256` published by Maven Central. Git tag
`v0.13.0` resolves to annotated tag object
`fe9697a2e9a33e3d3a576522e7346283826eb9ff` and commit
`b926a6134a2fee782076500dfc02c47c2d651cff`; the latter also appears as the
official validator's embedded SCM revision.

ODFDOM and ODF Validator declare the Apache License 2.0 in their Maven
metadata and source archives. The generated package retains the license notice
from ODFDOM's bundled chart template. No dependency binary is checked in here.

The runtime was OpenJDK 17.0.19+10 from Ubuntu 24.04. Its source package file
`openjdk-17-jdk-headless_17.0.19+10-1~24.04.2_amd64.deb` had SHA-256
`dcdeb373cc2b174e7b6ae64a9af14c1494e29a9a5b8a01523a57c9b89ea47de1`.

## Exact command route

The working directory was `/tmp/litchi-odc-odfdom.mMiZBQ`. Downloads used the
following Maven Central URLs, each followed by its adjacent `.sha256` URL:

```text
https://repo1.maven.org/maven2/org/odftoolkit/odfdom-java/0.13.0/odfdom-java-0.13.0-jar-with-dependencies.jar
https://repo1.maven.org/maven2/org/odftoolkit/odfdom-java/0.13.0/odfdom-java-0.13.0-sources.jar
https://repo1.maven.org/maven2/org/odftoolkit/odfdom-java/0.13.0/odfdom-java-0.13.0.pom
https://repo1.maven.org/maven2/org/odftoolkit/odfvalidator/0.13.0/odfvalidator-0.13.0-jar-with-dependencies.jar
```

After checksum verification, these commands were run. Each `java` invocation
is a separate JVM process, so creation, changed resave, and semantic reopen are
independent operations.

```sh
javac -cp odfdom.jar OdfdomChartProducer.java
java -cp odfdom.jar:. OdfdomChartProducer create odfdom-created.odc
java -cp odfdom.jar:. OdfdomChartProducer verify odfdom-created.odc 'ODFDOM standalone chart' axis-x
java -cp odfdom.jar:. OdfdomChartProducer change odfdom-created.odc odfdom-resaved.odc
java -cp odfdom.jar:. OdfdomChartProducer verify odfdom-resaved.odc 'ODFDOM independently resaved chart' axis-x-resaved
java -jar odfvalidator.jar -v odfdom-created.odc
java -jar odfvalidator.jar -v odfdom-resaved.odc
```

The two semantic reopen commands printed:

```text
application/vnd.oasis.opendocument.chart | ODFDOM standalone chart | axis-x
application/vnd.oasis.opendocument.chart | ODFDOM independently resaved chart | axis-x-resaved
```

Both validator commands exited 0. The retained transcripts report ODF 1.2,
media type `application/vnd.oasis.opendocument.chart`, and no errors or
warnings for the manifest, MIME member, metadata, content, or overall package.

## Package and member diff

Both ZIPs contain exactly `mimetype`, `meta.xml`, `content.xml`, `META-INF/`,
and `META-INF/manifest.xml`. In each package, `mimetype` is the first local
entry at byte offset 0, is stored without compression, and contains exactly
`application/vnd.oasis.opendocument.chart`.

| Member | Created SHA-256 | Resaved SHA-256 | Result |
|---|---|---|---|
| `mimetype` | `9a1659a2e29fb47c9cf64c39f8f17ceb279b6fa51853a4d7eb087f2ec9e48f7b` | same | byte-identical |
| `content.xml` | `8818a7d8de73f69d5ef5ebaf83575fb016a7360f9671279426f3f595c5682517` | `ff25dbfd40ce5636508fe5ca524fec1cfed8280cf08cca48523fc90f4f6b4c15` | changed |
| `meta.xml` | `ff9ce74fd6eda7d642dd81c0e0683f5675eeec7e1f19d188b6870f2952bfc09b` | same | byte-identical |
| `META-INF/manifest.xml` | `125eaafae989390050da3941f28dff32cc6e4e63122380229ef2f24c7954cdfc` | same | byte-identical |

The semantic `content.xml` change is limited to these two values:

```diff
-<text:p>ODFDOM standalone chart</text:p>
+<text:p>ODFDOM independently resaved chart</text:p>
-chart:name="axis-x"
+chart:name="axis-x-resaved"
```

No standalone FODC route was found in ODFDOM 0.13.0. This evidence therefore
claims only packaged standalone ODC creation and changed resave.
