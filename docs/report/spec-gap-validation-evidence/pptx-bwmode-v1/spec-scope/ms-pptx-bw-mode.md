## 2.3 http://schemas.microsoft.com/office/powerpoint/2010/main

### 2.3.1 Elements

#### 2.3.1.1 bmkTgt

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_MediaBookmarkTarget** element (section 2.3.3.13) that specifies a media bookmark that triggers an animation. See section 2.2.2 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

1.  \<xsd:element name="bmkTgt" type="CT_MediaBookmarkTarget"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.2 browseMode

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_BrowseMode** element (section 2.3.3.1) that specifies the status bar visibility when **slide show** is in browse mode. See section 2.2.6 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

2.  \<xsd:element name="browseMode" type="CT_BrowseMode"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.3 conveyor

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_LeftRightDirectionTransition** element (section 2.3.3.10) that specifies a conveyor transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following is a sample with **dir** equal to "l" (left):

![Sample conveyor transition](assets/image1.png)

*Figure 1: Sample conveyor transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

3.  \<xsd:element name="conveyor" type="CT_LeftRightDirectionTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.4 creationId

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_RandomId** element (section 2.3.3.21) that specifies an identifier that is generated when a slide is created. See section 2.2.9 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

4.  \<xsd:element name="creationId" type="CT_RandomId"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.5 defaultImageDpi

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_DefaultImageDpi** element (section 2.3.3.3) that specifies the resolution, in dots per inch (dpi), to use when saving images in the document. This element value is used only when the **autoCompressPictures** attribute ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.2.1.26) is set to **true**, and the **useLocalDpi** element ([\[MS-ODRAWXML\]](%5bMS-ODRAWXML%5d.pdf#Section_06cff208c6e14db7bb68665135e5f0de) section 2.3.1.13) is set to **false**. See section 2.2.7 for how this element integrates with \[ISO/IEC29500-1:2016\].

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

5.  \<xsd:element name="defaultImageDpi" type="CT_DefaultImageDpi"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.6 discardImageEditData

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_DiscardImageEditData** element (section 2.3.3.4) that specifies that all **imgProps** elements ([\[MS-ODRAWXML\]](%5bMS-ODRAWXML%5d.pdf#Section_06cff208c6e14db7bb68665135e5f0de) section 2.3.1.9) and cropped out areas of images are not to be saved. See section 2.2.7 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

6.  \<xsd:element name="discardImageEditData" type="CT_DiscardImageEditData"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.7 doors

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **p:CT_OrientationTransition** element ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4) that specifies a doors transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following is a sample with **dir** equal to "vert" (vertical):

![Sample doors transition](assets/image2.png)

*Figure 2: Sample doors transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

7.  \<xsd:element name="doors" type="p:CT_OrientationTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.8 extLst

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **p:CT_ExtensionListModify** element ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4) that specifies the extension list with modification ability. All future extensions to the **contentPart** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.14) will be stored within this extension list. See section 2.2.3 for how this element integrates with \[ISO/IEC29500-1:2016\].

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

8.  \<xsd:element name="extLst" type="p:CT_ExtensionListModify"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.9 ferris

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_LeftRightDirectionTransition** element (section 2.3.3.10) that specifies a ferris transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following is a sample with **dir** equal to "l" (left):

![Sample ferris transition](assets/image3.png)

*Figure 3: Sample ferris transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

9.  \<xsd:element name="ferris" type="CT_LeftRightDirectionTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.10 flash

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **p:CT_Empty** element ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4) that specifies a flash transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following figure is a sample:

![Sample flash transition](assets/image4.png)

*Figure 4: Sample flash transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

10. \<xsd:element name="flash" type="p:CT_Empty"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.11 flip

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_LeftRightDirectionTransition** element (section 2.3.3.10) that specifies a flip transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following figure is a sample with **dir** equal to "r" (right).

![Sample flip transition](assets/image5.png)

*Figure 5: Sample flip transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

11. \<xsd:element name="flip" type="CT_LeftRightDirectionTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.12 flythrough

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_FlyThroughTransition** element (section 2.3.3.5) that specifies a fly through transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following figure is a sample with **dir** equal to "in", **hasBounce** equal to **false**:

![Sample flythrough transition](assets/image6.png)

*Figure 6: Sample flythrough transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

12. \<xsd:element name="flythrough" type="CT_FlyThroughTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.13 gallery

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_LeftRightDirectionTransition** element (section 2.3.3.10) that specifies a gallery transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following figure is a sample with **dir** equal to "l" (left).

![Sample gallery transition](assets/image7.png)

*Figure 7: Sample gallery transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

13. \<xsd:element name="gallery" type="CT_LeftRightDirectionTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.14 glitter

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_GlitterTransition** element (section 2.3.3.6) that specifies a glitter transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following is a sample with **dir** equal to "l" (left), **pattern** equal to "diamond":

![Sample glitter transition](assets/image8.png)

*Figure 8: Sample glitter transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

14. \<xsd:element name="glitter" type="CT_GlitterTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.15 honeycomb

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **p:CT_Empty** element ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4) that specifies a honeycomb transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following figure is a sample:

![Sample honeycomb transition](assets/image9.png)

*Figure 9: Sample honeycomb transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

15. \<xsd:element name="honeycomb" type="p:CT_Empty"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.16 laserClr

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

An **a:CT_Color** element ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.5.1) that specifies the color to be used to render the laser dot. See section 2.2.6 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

16. \<xsd:element name="laserClr" type="a:CT_Color"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.17 laserTraceLst

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_LaserTraceList** element (section 2.3.3.8) that specifies a list of laser traces on the **presentation slide**. See section 2.2.6 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

17. \<xsd:element name="laserTraceLst" type="CT_LaserTraceList"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.18 media

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_Media** element (section 2.3.3.14) that specifies a media object. See section 2.2.4 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

18. \<xsd:element name="media" type="CT_Media"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.19 modId

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_RandomId** element (section 2.3.3.21) that specifies an identifier that is updated each time a **shape** is modified. Each identifier specified MUST be unique to the **presentation slide** that contains it. See section 2.2.9 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

19. \<xsd:element name="modId" type="CT_RandomId"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.20 nvContentPartPr

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_ContentPartNonVisual** element (section 2.3.3.2) that specifies the non-visual drawing properties for a **contentPart** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.14). These non-visual properties are properties that the generating application uses when rendering and editing the content part. See section 2.2.3 for how this element integrates with \[ISO/IEC29500-1:2016\].

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

20. \<xsd:element name="nvContentPartPr" type="CT_ContentPartNonVisual"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.21 pan

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **p:CT_SideDirectionTransition** element ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4) that specifies a pan transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following figure is a sample with **dir** equal to "u" (up).

![Sample pan transition](assets/image10.png)

*Figure 10: Sample pan transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

21. \<xsd:element name="pan" type="p:CT_SideDirectionTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.22 prism

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_PrismTransition** element (section 2.3.3.20) that specifies a prism transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following figure is a sample with **dir** equal to "l" (left), **isContent** equal to **false**, **isInverted** equal to **false**:

![Sample prism transition 1](assets/image11.png)

*Figure 11: Sample prism transition 1*

The following figure is a sample with **dir** equal to "r" (right), **isContent** equal to **false**, **isInverted** equal to **true**:

![Sample prism transition 2](assets/image12.png)

*Figure 12: Sample prism transition 2*

The following figure is a sample with **dir** equal to "u" (up), **isContent** equal to **true**, **isInverted** equal to **false**:

![Sample prism transition 3](assets/image13.png)

*Figure 13: Sample prism transition 3*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

22. \<xsd:element name="prism" type="CT_PrismTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.23 reveal

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_RevealTransition** element (section 2.3.3.22) that specifies a reveal transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following is a sample with **thruBlk** equal to **true**, **dir** equal to "l" (left):

![Sample reveal transition](assets/image14.png)

*Figure 14: Sample reveal transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

23. \<xsd:element name="reveal" type="CT_RevealTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.24 ripple

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_RippleTransition** element (section 2.3.3.23) that specifies a ripple transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following figure is a sample with **dir** equal to "center".

![Sample ripple transition](assets/image15.png)

*Figure 15: Sample ripple transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

24. \<xsd:element name="ripple" type="CT_RippleTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.25 sectionLst

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_SectionList** element (section 2.3.3.25) that specifies section properties for the document. See section 2.2.5 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

25. \<xsd:element name="sectionLst" type="CT_SectionList"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.26 showEvtLst

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_ShowEventRecordList** element (section 2.3.3.28) that specifies a list of **slide show** events that occur during slide show. See section 2.2.6 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

26. \<xsd:element name="showEvtLst" type="CT_ShowEventRecordList"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.27 showMediaCtrls

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_ShowMediaControls** element (section 2.3.3.29) that specifies the visibility of media user interface controls during **slide show**. See section 2.2.4 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

27. \<xsd:element name="showMediaCtrls" type="CT_ShowMediaControls"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.28 shred

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_ShredTransition** element (section 2.3.3.30) that specifies a shred transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following figure is a sample with **pattern** equal to "strip", **dir** equal to "in":

![Sample shred transition](assets/image16.png)

*Figure 16: Sample shred transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

28. \<xsd:element name="shred" type="CT_ShredTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.29 switch

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **CT_LeftRightDirectionTransition** element (section 2.3.3.10) that specifies a switch transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following figure is a sample with **dir** equal to "r" (right):

![Sample switch transition](assets/image17.png)

*Figure 17: Sample switch transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

29. \<xsd:element name="switch" type="CT_LeftRightDirectionTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.30 vortex

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **p:CT_SideDirectionTransition** element ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4) that specifies a vortex transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following figure is a sample with **dir** equal to "r" (right):

![Sample vortex transition](assets/image18.png)

*Figure 18: Sample vortex transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

30. \<xsd:element name="vortex" type="p:CT_SideDirectionTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.31 warp

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **p:CT_InOutTransition** element ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.3) that specifies a warp transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following is a sample with **dir** equal to "in":

![Sample warp transition](assets/image19.png)

*Figure 19: Sample warp transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

31. \<xsd:element name="warp" type="p:CT_InOutTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.32 wheelReverse

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **p:CT_WheelTransition** element ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4) that specifies a reverse wheel transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following figure is a sample with **spokes** equal to 1.

![Sample wheelReverse transition](assets/image20.png)

*Figure 20: Sample wheelReverse transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

32. \<xsd:element name="wheelReverse" type="p:CT_WheelTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.33 window

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

A **p:CT_OrientationTransition** element ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4) that specifies a window transition. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following is a sample with **dir** equal to "vert" (vertical):

![Sample window transition](assets/image21.png)

*Figure 21: Sample window transition*

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

33. \<xsd:element name="window" type="p:CT_OrientationTransition"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.1.34 xfrm

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

An **a:CT_Transform2D** element ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.5.1) that specifies the 2-D transform for a content part. See section 2.2.3 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this element.

34. \<xsd:element name="xfrm" type="a:CT_Transform2D"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

### 2.3.2 Attributes

#### 2.3.2.1 bounceEnd

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

An **s:ST_PositiveFixedPercentage** attribute ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.8.9) that specifies the percentage of the **time node**'s duration to do a bounce at the end of the animation. The bounce emulates a mass-spring-damper system. See section 2.2.2 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this attribute.

35. \<xsd:attribute name="bounceEnd" type="s:ST_PositiveFixedPercentage"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.2.2 bwMode

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

An **a:ST_BlackWhiteMode** attribute ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.5.1) that specifies how to interpret color information contained within a content part to achieve a color, black and white, or grayscale rendering of the content part. This attribute specifies only the rendering mode applied to the content part; it does not affect how the actual color information is persisted. See section 2.2.3 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this attribute.

36. \<xsd:attribute name="bwMode" type="a:ST_BlackWhiteMode"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.2.3 dur

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

An **ST_UniversalTimeOffset** attribute (section 2.3.4.6) that specifies the time a transition takes to display from start to finish. See section 2.2.1 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this attribute.

37. \<xsd:attribute name="dur" type="p14:ST_UniversalTimeOffset"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.2.4 presetBounceEnd

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

An **s:ST_PositiveFixedPercentage** attribute ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.8.9) that specifies a preset percentage of a **time node**'s duration to do a bounce at the end of the animation. This value is used only by the user interface. The actual bounce animations are specified by the **bounceEnd** (section 2.3.2.1) attributes. See section 2.2.2 for how this element integrates with [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this attribute.

38. \<xsd:attribute name="presetBounceEnd" type="s:ST_PositiveFixedPercentage"/\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

### 2.3.3 Complex Types

#### 2.3.3.1 CT_BrowseMode

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* browseMode

A complex type that specifies the visibility of the status bar in **slide show** browse mode.

*Attributes:*

**showStatus:** An **xsd:boolean** attribute ([\[XMLSCHEMA2/2\]](https://go.microsoft.com/fwlink/?LinkId=90609) section 3.2.2) that specifies the visibility of status bar when slide show is in browse mode.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

39. \<xsd:complexType name="CT_BrowseMode"\>

    \<xsd:attribute name="showStatus" type="xsd:boolean" use="optional" default="true"/\>

    \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.2 CT_ContentPartNonVisual

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* nvContentPartPr

A complex type which specifies non-visual properties for a **contentPart** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.14).

*Child Elements:*

**cNvPr:** An **a:CT_NonVisualDrawingProps** element ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.5.1) that specifies the non-visual drawing properties for the content part. This element allows for additional information that does not affect the appearance of the content part to be stored.

**cNvContentPartPr:** An **a14:CT_NonVisualInkContentPartProperties** element ([\[MS-ODRAWXML\]](%5bMS-ODRAWXML%5d.pdf#Section_06cff208c6e14db7bb68665135e5f0de) section 2.3.3.7) that specifies non-visual **ink** properties for the content part. This element allows for additional information that does not affect the appearance of ink in the content part to be stored.

**nvPr:** A **p:CT_ApplicationNonVisualDrawingProps** element (\[ISO/IEC29500-4:2016\] section A.4) that specifies PresentationML Drawing specific non-visual properties for the content part.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

42. \<xsd:complexType name="CT_ContentPartNonVisual"\>

    \<xsd:sequence\>

    \<xsd:element name="cNvPr" type="a:CT_NonVisualDrawingProps" minOccurs="1" maxOccurs="1"/\>

    \<xsd:element name="cNvContentPartPr" type="a14:CT_NonVisualInkContentPartProperties" minOccurs="0" maxOccurs="1"/\>

    \<xsd:element name="nvPr" type="p:CT_ApplicationNonVisualDrawingProps" minOccurs="1" maxOccurs="1"/\>

    \</xsd:sequence\>

    \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.3 CT_DefaultImageDpi

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* defaultImageDpi

A complex type that specifies the default dpi to use when saving each image in the document.

*Attributes:*

**val:** An **xsd:unsignedInt** attribute ([\[XMLSCHEMA2/2\]](https://go.microsoft.com/fwlink/?LinkId=90609) section 3.3.22) that specifies the dpi.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

49. \<xsd:complexType name="CT_DefaultImageDpi"\>

    \<xsd:attribute name="val" type="xsd:unsignedInt" use="required"/\>

    \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.4 CT_DiscardImageEditData

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* discardImageEditData

A complex type that specifies whether to discard editing data for images when saving.

*Attributes:*

**val:** An **xsd:boolean** attribute ([\[XMLSCHEMA2/2\]](https://go.microsoft.com/fwlink/?LinkId=90609) section 3.2.2) that specifies whether to discard image editing data when saving.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

52. \<xsd:complexType name="CT_DiscardImageEditData"\>

    \<xsd:attribute name="val" type="xsd:boolean" use="required"/\>

    \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.5 CT_FlyThroughTransition

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* flythrough

A complex type that specifies the parameters for a fly-through transition.

*Attributes:*

**dir:** A **p:ST_TransitionInOutDirectionType** attribute ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4) that specifies the direction **presentation slides** move during the transition.

**hasBounce:** An **xsd:boolean** attribute ([\[XMLSCHEMA2/2\]](https://go.microsoft.com/fwlink/?LinkId=90609) section 3.2.2) that specifies that the movement of the presentation slides during the transition includes a bounce.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

55. \<xsd:complexType name="CT_FlyThroughTransition"\>

    \<xsd:attribute name="dir" type="p:ST_TransitionInOutDirectionType" use="optional" default="in"/\>

    \<xsd:attribute name="hasBounce" type="xsd:boolean" use="optional" default="false"/\>

    \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.6 CT_GlitterTransition

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* glitter

A complex type that specifies the parameters for a glitter transition.

*Attributes:*

**dir:** A **p:ST_TransitionSideDirectionType** attribute ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4) that specifies the direction **presentation slides** move during the transition.

**pattern:** An **ST_TransitionPattern** (section 2.3.4.4) attribute that specifies the shape of the visuals used during the transition.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

59. \<xsd:complexType name="CT_GlitterTransition"\>

    \<xsd:attribute name="dir" type="p:ST_TransitionSideDirectionType" use="optional" default="l"/\>

    \<xsd:attribute name="pattern" type="ST_TransitionPattern" use="optional" default="diamond"/\>

    \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.7 CT_LaserTrace

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_LaserTraceList

A complex type that specifies a list of laser trace points. The first point in the list of points specifies when and where the laser point appears. The last point in the list of points specifies when and where the laser point disappears. Other trace points in the list of points specify laser point appearance time and locations.

*Child Elements:*

**tracePt:** A **CT_LaserTracePoint** element (section 2.3.3.9) that specifies a location of the laser point.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

63. \<xsd:complexType name="CT_LaserTrace"\>

    \<xsd:sequence\>

    \<xsd:element name="tracePt" type="CT_LaserTracePoint" minOccurs="0" maxOccurs="unbounded"/\>

    \</xsd:sequence\>

    \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.8 CT_LaserTraceList

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* laserTraceLst

A complex type that specifies a list of laser traces.

*Child Elements:*

**tracePtLst:** A **CT_LaserTrace** element (section 2.3.3.7) that specifies a laser trace.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

68. \<xsd:complexType name="CT_LaserTraceList"\>

    \<xsd:sequence\>

    \<xsd:element name="tracePtLst" type="CT_LaserTrace" minOccurs="0" maxOccurs="unbounded"/\>

    \</xsd:sequence\>

    \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.9 CT_LaserTracePoint

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_LaserTrace

A complex type that specifies the time and location of the laser pointer on a **presentation slide**.

*Attributes:*

**t:** An **ST_UniversalTimeOffset** attribute (section 2.3.4.6) that specifies the time relative to the beginning of the slide timeline and not including transition time.

**x:** An **a:ST_Coordinate** attribute ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.5.1) that specifies the horizontal location relative to the top-left corner of the presentation slide.

**y:** An **a:ST_Coordinate** attribute (\[ISO/IEC29500-4:2016\] section A.5.1) that specifies the vertical location relative to the top-left corner of the presentation slide.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

73. \<xsd:complexType name="CT_LaserTracePoint"\>

    \<xsd:attribute name="t" type="p14:ST_UniversalTimeOffset" use="required"/\>

    \<xsd:attribute name="x" type="a:ST_Coordinate" use="required"/\>

    \<xsd:attribute name="y" type="a:ST_Coordinate" use="required"/\>

    \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.10 CT_LeftRightDirectionTransition

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* switch, flip, ferris, gallery, conveyor

A complex type that specifies a transition restricted to left and right directions.

*Attributes:*

**dir:** An **ST_TransitionLeftRightDirectionType** attribute (section 2.3.4.3) that specifies the direction **presentation slides** move during the transition.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

78. \<xsd:complexType name="CT_LeftRightDirectionTransition"\>

    \<xsd:attribute name="dir" type="ST_TransitionLeftRightDirectionType"/\>

    \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.11 CT_MediaBookmark

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_MediaBookmarkList

A complex type that specifies a media bookmark on the media.

*Attributes:*

**name:** An xsd:string ([\[XMLSCHEMA2/2\]](https://go.microsoft.com/fwlink/?LinkId=90609) section 3.2.1) attribute that specifies the name for the bookmark. The bookmark name MUST be unique among all bookmarks in its parent **CT_MediaBookmarkList** (section 2.3.3.12).

**time:** An **ST_UniversalTimeOffset** attribute (section 2.3.4.6) that specifies the time of the bookmark relative to the beginning of the media. It MUST be unique among all bookmarks in its parent **CT_MediaBookmarkList** (section 2.3.3.12). It MUST be less than or equal to the length of the media.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

81. \<xsd:complexType name="CT_MediaBookmark"\>

    \<xsd:attribute name="name" type="xsd:string"/\>

    \<xsd:attribute name="time" type="p14:ST_UniversalTimeOffset"/\>

    \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.12 CT_MediaBookmarkList

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_Media

A complex type that specifies a list of media bookmarks on the media.

*Child Elements:*

**bmk:** A **CT_MediaBookmark** element (section 2.3.3.11) that specifies a bookmark name and time.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

85. \<xsd:complexType name="CT_MediaBookmarkList"\>

    \<xsd:sequence\>

    \<xsd:element name="bmk" type="CT_MediaBookmark" minOccurs="0" maxOccurs="unbounded"/\>

    \</xsd:sequence\>

    \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.13 CT_MediaBookmarkTarget

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* bmkTgt

A complex type that specifies a media bookmark that triggers an animation.

*Attributes:*

**spid:** An **a:ST_DrawingElementId** attribute ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.5.1) that specifies the **shape identifier** of the **shape** that contains video or audio data.

**bmkName:** An **xsd:string** attribute ([\[XMLSCHEMA2/2\]](https://go.microsoft.com/fwlink/?LinkId=90609) section 3.2.1) that specifies the name of the bookmark for the media that is contained in the shape specified by **spid**. The name MUST exist in the **CT_MediaBookmarkList** (section 2.3.3.12) contained by the media.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

90. \<xsd:complexType name="CT_MediaBookmarkTarget"\>

    \<xsd:attribute name="spid" type="a:ST_DrawingElementId" use="required"/\>

    \<xsd:attribute name="bmkName" type="xsd:string" use="required"/\>

    \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.14 CT_Media

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* media

A complex type that specifies the media.

*Child Elements:*

**trim:** A **CT_MediaTrim** element (section 2.3.3.18) that specifies the amount of time to be trimmed from the start and end of the media during playback.

**fade:** A **CT_MediaFade** element (section 2.3.3.15) that specifies the starting and ending fade durations for the media.

**bmkLst:** A **CT_MediaBookmarkList** element (section 2.3.3.12) that specifies a list of bookmarks on the media.

**extLst:** A **p:CT_ExtensionList** ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.3) element that specifies the extension list. All future extensions to the media will be stored in the extension list.

*Attributes:*

**r:embed:** A **r:ST_RelationshipId** (\[ISO/IEC29500-4:2016\] section A.8.8) attribute that specifies the relationship identifier that is used to determine the location of the media if it is embedded in the document. The **r:embed** attribute MUST be present if the **r:link** attribute is not present. If both the **r:embed** and **r:link** attributes are present, the **r:link** attribute takes precedence.

**r:link:** A **r:ST_RelationshipId** (\[ISO/IEC29500-4:2016\] section A.8.8) attribute that specifies the relationship identifier that is used to determine the location of the media if it is linked from the document. The **r:link** attribute MUST be present if the **r:embed** attribute is not present. If both the **r:link** and **r:embed** attributes are present, the **r:link** attribute takes precedence.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

94. \<xsd:complexType name="CT_Media"\>

    \<xsd:sequence\>

    \<xsd:element name="trim" type="CT_MediaTrim" minOccurs="0" maxOccurs="1"/\>

    \<xsd:element name="fade" type="CT_MediaFade" minOccurs="0" maxOccurs="1"/\>

    \<xsd:element name="bmkLst" type="CT_MediaBookmarkList" minOccurs="0" maxOccurs="1"/\>

    \<xsd:element name="extLst" type="p:CT_ExtensionList" minOccurs="0" maxOccurs="1"/\>

    \</xsd:sequence\>

    \<xsd:attributeGroup ref="a:AG_Blob"/\>

    \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.15 CT_MediaFade

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_Media

A complex type that specifies the media fade.

*Attributes:*

**in:** An **ST_UniversalTimeOffset** (section 2.3.4.6) attribute that specifies the duration of the starting fade. It MUST be less than or equal to the total length of the media minus the **out**.

**out:** An **ST_UniversalTimeOffset** attribute that specifies the duration of the ending fade. It MUST be less than or equal to the total length of the media minus the **in**.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

103. \<xsd:complexType name="CT_MediaFade"\>

     \<xsd:attribute name="in" type="p14:ST_UniversalTimeOffset" use="optional" default="0"/\>

     \<xsd:attribute name="out" type="p14:ST_UniversalTimeOffset" use="optional" default="0"/\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.16 CT_MediaPlaybackEventRecord

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_ShowEventRecordList

A complex type that specifies a media playback event that occurs during **slide show**.

*Attributes:*

**time:** An **ST_UniversalTimeOffset** (section 2.3.4.6) attribute that specifies the time that the event occurs. The value is relative to the beginning of the slide timeline, which does not include side transition time.

**objId:** An **a:ST_DrawingElementId** attribute ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.5.1) that specifies the identifier of the media object on the slide that is the target of the event.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

107. \<xsd:complexType name="CT_MediaPlaybackEventRecord"\>

     \<xsd:attribute name="time" type="p14:ST_UniversalTimeOffset" use="required"/\>

     \<xsd:attribute name="objId" type="a:ST_DrawingElementId" use="required"/\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.17 CT_MediaSeekEventRecord

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_ShowEventRecordList

A complex type that specifies a media seek event that occurs during **slide show**.

*Attributes:*

**time:** An **ST_UniversalTimeOffset** attribute (section 2.3.4.6) that specifies the time that the event occurs. The value is relative to the beginning of the slide timeline, which does not include slide transition time.

**objId:** An **a:ST_DrawingElementId** attribute ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.5.1) that specifies the identifier of the media object on the slide that is the target of this event.

**seek:** An **ST_UniversalTimeOffset** attribute (section 2.3.4.6) that specifies the time of the position to seek to in the media stream.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

111. \<xsd:complexType name="CT_MediaSeekEventRecord"\>

     \<xsd:attribute name="time" type="p14:ST_UniversalTimeOffset" use="required"/\>

     \<xsd:attribute name="objId" type="a:ST_DrawingElementId" use="required"/\>

     \<xsd:attribute name="seek" type="p14:ST_UniversalTimeOffset" use="required"/\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.18 CT_MediaTrim

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_Media

A complex type that specifies the amount of time to be trimmed from the start and end of the media during playback.

*Attributes:*

**st:** An **ST_UniversalTimeOffset** attribute (section 2.3.4.6) that specifies a duration of time to be removed from the start of the media during playback. It MUST be less than the total length of the media minus the **end**.

**end:** An **ST_UniversalTimeOffset** attribute (section 2.3.4.6) that specifies a duration of time to be removed from the end of the media during playback. It MUST be less than the total length of the media minus the **st**.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

116. \<xsd:complexType name="CT_MediaTrim"\>

     \<xsd:attribute name="st" type="p14:ST_UniversalTimeOffset" use="optional" default="0"/\>

     \<xsd:attribute name="end" type="p14:ST_UniversalTimeOffset" use="optional" default="0"/\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.19 CT_NullEventRecord

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_ShowEventRecordList

A complex type that specifies an unknown event in **slide show**. This is for future extension.

*Attributes:*

**time:** An **ST_UniversalTimeOffset** attribute (section 2.3.4.6) that specifies the time that the event occurs. The value is relative to the beginning of the slide timeline, which does not include slide transition time.

**objId:** An **a:ST_DrawingElementId** attribute ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.5.1) that specifies the identifier of the shape object on the slide that is the target of this event.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

120. \<xsd:complexType name="CT_NullEventRecord"\>

     \<xsd:attribute name="time" type="p14:ST_UniversalTimeOffset" use="required"/\>

     \<xsd:attribute name="objId" type="a:ST_DrawingElementId" use="required"/\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.20 CT_PrismTransition

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* prism

A complex type that specifies the parameters for a prism transition.

*Attributes:*

**dir:** A **p:ST_TransitionSideDirectionType** attribute ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4) that specifies the direction **presentation slides** move during the transition.

**isContent:** An **xsd:boolean** attribute ([\[XMLSCHEMA2/2\]](https://go.microsoft.com/fwlink/?LinkId=90609) section 3.2.2) that specifies whether the visual representation draws presentation slide content and presentation slide background separately.

**isInverted:** An **xsd:boolean** attribute (\[XMLSCHEMA2/2\] section 3.2.2) that specifies whether the layout of the presentation slides during the transition is concave instead of convex.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

124. \<xsd:complexType name="CT_PrismTransition"\>

     \<xsd:attribute name="dir" type="p:ST_TransitionSideDirectionType" use="optional" default="l"/\>

     \<xsd:attribute name="isContent" type="xsd:boolean" use="optional" default="false"/\>

     \<xsd:attribute name="isInverted" type="xsd:boolean" use="optional" default="false"/\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.21 CT_RandomId

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* creationId, modId

A complex type that specifies a cryptographically strong random number.

*Attributes:*

**val:** An **xsd:unsignedInt** attribute ([\[XMLSCHEMA2/2\]](https://go.microsoft.com/fwlink/?LinkId=90609) section 3.3.22) that specifies the random number.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

129. \<xsd:complexType name="CT_RandomId"\>

     \<xsd:attribute name="val" type="xsd:unsignedInt" use="required"/\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.22 CT_RevealTransition

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* reveal

A complex type that specifies the parameters for a reveal transition.

*Attributes:*

**thruBlk:** An **xsd:boolean** attribute ([\[XMLSCHEMA2/2\]](https://go.microsoft.com/fwlink/?LinkId=90609) section 3.2.2) that specifies whether the transition fades through black.

**dir:** An **ST_TransitionLeftRightDirectionType** attribute (section 2.3.4.3) that specifies the direction **presentation slides** move during the transition.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

132. \<xsd:complexType name="CT_RevealTransition"\>

     \<xsd:attribute name="thruBlk" type="xsd:boolean" use="optional" default="false"/\>

     \<xsd:attribute name="dir" type="ST_TransitionLeftRightDirectionType" use="optional" default="l"/\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.23 CT_RippleTransition

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* ripple

A complex type that specifies the parameters for a ripple transition.

*Attributes:*

**dir:** An **ST_TransitionCornerAndCenterDirectionType** attribute (section 2.3.4.2) that specifies the direction **presentation slides** move during the transition.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

136. \<xsd:complexType name="CT_RippleTransition"\>

     \<xsd:attribute name="dir" type="ST_TransitionCornerAndCenterDirectionType" use="optional" default="center"/\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.24 CT_Section

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_SectionList

A complex type that specifies a section and its properties.

*Child Elements:*

**sldIdLst:** A **CT_SectionSlideIdList** element (section 2.3.3.26) that specifies a list of **presentation slides** in a section.

**extLst:** A **p:CT_ExtensionList** element ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.3) that specifies the extension list. All future extensions to the section will be stored in the extension list.

*Attributes:*

**name:** An **xsd:string** attribute ([\[XMLSCHEMA2/2\]](https://go.microsoft.com/fwlink/?LinkId=90609) section 3.2.1) that specifies the name of the section.

**id:** A **s:ST_Guid** attribute (\[ISO/IEC29500-4:2016\] section A.8.9) that specifies a **GUID** used to uniquely identify a section in the **CT_SectionList** (section 2.3.3.25).

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

139. \<xsd:complexType name="CT_Section"\>

     \<xsd:sequence\>

     \<xsd:element name="sldIdLst" type="CT_SectionSlideIdList" minOccurs="1" maxOccurs="1"/\>

     \<xsd:element name="extLst" type="p:CT_ExtensionList" minOccurs="0" maxOccurs="1"/\>

     \</xsd:sequence\>

     \<xsd:attribute name="name" type="xsd:string"/\>

     \<xsd:attribute name="id" type="s:ST_Guid"/\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.25 CT_SectionList

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* sectionLst

A complex type that specifies section properties for the document.

*Child Elements:*

**section:** A **CT_Section** element (section 2.3.3.24) that specifies a section in the document.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

147. \<xsd:complexType name="CT_SectionList"\>

     \<xsd:sequence\>

     \<xsd:element name="section" type="CT_Section" minOccurs="1" maxOccurs="unbounded"/\>

     \</xsd:sequence\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.26 CT_SectionSlideIdList

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_Section

A complex type that specifies the list of **presentation slides** that belong to a section.

*Child Elements:*

**sldId:** A **CT_SectionSlideIdListEntry** element (section 2.3.3.27) that specifies an identifier of a presentation slide.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

152. \<xsd:complexType name="CT_SectionSlideIdList"\>

     \<xsd:sequence\>

     \<xsd:element name="sldId" type="CT_SectionSlideIdListEntry" minOccurs="0" maxOccurs="unbounded"/\>

     \</xsd:sequence\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.27 CT_SectionSlideIdListEntry

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_SectionSlideIdList

A complex type that specifies the **presentation slide** identifier.

*Attributes:*

**id:** A **p:ST_SlideId** attribute ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4) that specifies the slide identifier.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

157. \<xsd:complexType name="CT_SectionSlideIdListEntry"\>

     \<xsd:attribute name="id" type="p:ST_SlideId" use="required"/\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.28 CT_ShowEventRecordList

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* showEvtLst

A complex type that specifies a list of **slide show** events that occurs during slide show.

*Child Elements:*

**triggerEvt:** A **CT_TriggerEventRecord** element (section 2.3.3.31) that specifies a triggered animation event.

**playEvt:** A **CT_MediaPlaybackEventRecord** element (section 2.3.3.16) that specifies an event targeting a media object that starts playback.

**stopEvt:** A **CT_MediaPlaybackEventRecord** element (section 2.3.3.16) that specifies an event targeting a media object that stops playback.

**pauseEvt:** A **CT_MediaPlaybackEventRecord** element (section 2.3.3.16) that specifies an event targeting a media object that pauses playback.

**resumeEvt:** A **CT_MediaPlaybackEventRecord** element (section 2.3.3.16) that specifies an event targeting a media object that assumes playback.

**seekEvt:** A **CT_MediaSeekEventRecord** element (section 2.3.3.17) that specifies an event targeting a media object that seeks to a specified time in the media stream.

**nullEvt:** A **CT_NullEventRecord** element (section 2.3.3.19) that specifies an unknown event.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

160. \<xsd:complexType name="CT_ShowEventRecordList"\>

     \<xsd:sequence\>

     \<xsd:choice minOccurs="0" maxOccurs="unbounded"\>

     \<xsd:element name="triggerEvt" type="CT_TriggerEventRecord"/\>

     \<xsd:element name="playEvt" type="CT_MediaPlaybackEventRecord"/\>

     \<xsd:element name="stopEvt" type="CT_MediaPlaybackEventRecord"/\>

     \<xsd:element name="pauseEvt" type="CT_MediaPlaybackEventRecord"/\>

     \<xsd:element name="resumeEvt" type="CT_MediaPlaybackEventRecord"/\>

     \<xsd:element name="seekEvt" type="CT_MediaSeekEventRecord"/\>

     \<xsd:element name="nullEvt" type="CT_NullEventRecord"/\>

     \</xsd:choice\>

     \</xsd:sequence\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.29 CT_ShowMediaControls

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* showMediaCtrls

A complex type that specifies the visibility of media user interface controls during **slide show**.

*Attributes:*

**val:** An **xsd:boolean** attribute ([\[XMLSCHEMA2/2\]](https://go.microsoft.com/fwlink/?LinkId=90609) section 3.2.2) that specifies the visibility of media user interface controls.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

173. \<xsd:complexType name="CT_ShowMediaControls"\>

     \<xsd:attribute name="val" type="xsd:boolean" use="required"/\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.30 CT_ShredTransition

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* shred

A complex type that specifies the parameters for a shred transition.

*Attributes:*

**pattern:** An **ST_TransitionShredPattern** attribute (section 2.3.4.5) that specifies the shape of the visuals used during the transition.

**dir:** A **p:ST_TransitionInOutDirectionType** attribute ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4) that specifies the direction **presentation slides** move during the transition.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

176. \<xsd:complexType name="CT_ShredTransition"\>

     \<xsd:attribute name="pattern" type="ST_TransitionShredPattern" use="optional" default="strip"/\>

     \<xsd:attribute name="dir" type="p:ST_TransitionInOutDirectionType" use="optional" default="in"/\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.3.31 CT_TriggerEventRecord

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_ShowEventRecordList

A complex type that specifies a triggered event that occurs during **slide show**.

*Attributes:*

**type:** A **p:ST_TLTriggerEvent** attribute ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4) that specifies the type of the triggered event.

**time:** An **ST_UniversalTimeOffset** attribute (section 2.3.4.6) that specifies the time that the event occurs. The value is relative to the beginning of the slide timeline, which does not include slide transition time.

**objId:** An **a:ST_DrawingElementId** attribute (\[ISO/IEC29500-4:2016\] section A.5.1) that specifies the identifier of the shape object on the slide that is the target of this event.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this complex type.

180. \<xsd:complexType name="CT_TriggerEventRecord"\>

     \<xsd:attribute name="type" type="p:ST_TLTriggerEvent" use="required"/\>

     \<xsd:attribute name="time" type="p14:ST_UniversalTimeOffset" use="required"/\>

     \<xsd:attribute name="objId" type="a:ST_DrawingElementId" use="required"/\>

     \</xsd:complexType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

### 2.3.4 Simple Types

#### 2.3.4.1 ST_TransitionCenterDirectionType

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* ST_TransitionCornerAndCenterDirectionType

A simple type that specifies a direction restricted to the center.

| Value  | Meaning                                         |
|--------|-------------------------------------------------|
| center | Direction of movement is to or from the center. |

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this simple type.

185. \<xsd:simpleType name="ST_TransitionCenterDirectionType"\>

     \<xsd:restriction base="xsd:token"\>

     \<xsd:enumeration value="center"/\>

     \</xsd:restriction\>

     \</xsd:simpleType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.4.2 ST_TransitionCornerAndCenterDirectionType

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_RippleTransition

A simple type that specifies a direction restricted to the corners and center.

This simple type is a union of the following types:

- The **p:ST_TransitionCornerDirectionType** simple type ([\[ISO/IEC29500-4:2016\]](https://go.microsoft.com/fwlink/?linkid=861068) section A.4)

- The **ST_TransitionCenterDirectionType** simple type (section 2.3.4.1)

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this simple type.

190. \<xsd:simpleType name="ST_TransitionCornerAndCenterDirectionType"\>

     \<xsd:union memberTypes="p:ST_TransitionCornerDirectionType ST_TransitionCenterDirectionType"/\>

     \</xsd:simpleType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.4.3 ST_TransitionLeftRightDirectionType

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_LeftRightDirectionTransition, CT_RevealTransition

A simple type that specifies a direction restricted to the values of left and right.

| Value | Meaning |
|-------|---------|
| l     | Left    |
| r     | Right   |

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this simple type.

193. \<xsd:simpleType name="ST_TransitionLeftRightDirectionType"\>

     \<xsd:restriction base="xsd:token"\>

     \<xsd:enumeration value="l"/\>

     \<xsd:enumeration value="r"/\>

     \</xsd:restriction\>

     \</xsd:simpleType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.4.4 ST_TransitionPattern

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_GlitterTransition

A simple type that specifies a geometric pattern that tiles together to fill a larger area.

| Value   | Meaning              |
|---------|----------------------|
| diamond | Diamond tile pattern |
| hexagon | Hexagon tile pattern |

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this simple type.

199. \<xsd:simpleType name="ST_TransitionPattern"\>

     \<xsd:restriction base="xsd:token"\>

     \<xsd:enumeration value="diamond"/\>

     \<xsd:enumeration value="hexagon"/\>

     \</xsd:restriction\>

     \</xsd:simpleType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.4.5 ST_TransitionShredPattern

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* CT_ShredTransition

A simple type that specifies a geometric shape that tiles together to fill a larger area.

| Value     | Meaning          |
|-----------|------------------|
| strip     | Vertical strips  |
| rectangle | Small rectangles |

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this simple type.

205. \<xsd:simpleType name="ST_TransitionShredPattern"\>

     \<xsd:restriction base="xsd:token"\>

     \<xsd:enumeration value="strip"/\>

     \<xsd:enumeration value="rectangle"/\>

     \</xsd:restriction\>

     \</xsd:simpleType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).

#### 2.3.4.6 ST_UniversalTimeOffset

*Target namespace:* http://schemas.microsoft.com/office/powerpoint/2010/main

*Referenced by:* dur, CT_MediaTrim, CT_MediaFade, CT_MediaBookmark, CT_LaserTracePoint, CT_TriggerEventRecord, CT_NullEventRecord, CT_MediaPlaybackEventRecord, CT_MediaSeekEventRecord, CT_ZoomObjectProperties

A simple type that specifies an amount of time. It MUST be a valid TIMEOFFSET as specified in the following ABNF [\[RFC5234\]](https://go.microsoft.com/fwlink/?LinkId=123096) grammar:

TIMEOFFSET = TIME \[UNITS\]

TIME = 1\*DIGIT \["." 1\*DIGIT\]

UNITS = "h" / "min" / "s" / "ms" / "µs" / "ns"

UNITS specifies one of the following units of time:

| Abbreviation | Unit of Time |
|--------------|--------------|
| h            | Hour         |
| min          | Minute       |
| s            | Second       |
| ms           | Millisecond  |
| µs           | Microsecond  |
| ns           | Nanosecond   |

If no unit is specified, the default is millisecond.

The following W3C XML Schema ([\[XMLSCHEMA1/2\]](https://go.microsoft.com/fwlink/?LinkId=90607) section 2.1) fragment specifies the contents of this simple type.

211. \<xsd:simpleType name="ST_UniversalTimeOffset"\>

     \<xsd:restriction base="xsd:string"/\>

     \</xsd:simpleType\>

See section 5.1 for the full W3C XML Schema (\[XMLSCHEMA1/2\] section 2.1).
