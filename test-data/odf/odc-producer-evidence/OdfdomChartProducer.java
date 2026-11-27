// SPDX-License-Identifier: Apache-2.0

import java.io.File;
import org.odftoolkit.odfdom.doc.OdfChartDocument;
import org.odftoolkit.odfdom.dom.element.chart.ChartAxisElement;
import org.odftoolkit.odfdom.dom.element.chart.ChartChartElement;
import org.odftoolkit.odfdom.dom.element.chart.ChartPlotAreaElement;
import org.odftoolkit.odfdom.dom.element.chart.ChartSeriesElement;
import org.odftoolkit.odfdom.dom.element.chart.ChartTitleElement;
import org.odftoolkit.odfdom.dom.element.text.TextPElement;
import org.odftoolkit.odfdom.incubator.meta.OdfOfficeMeta;
import org.w3c.dom.NodeList;

public final class OdfdomChartProducer {
  private static final String CHART =
      "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
  private static final String TEXT =
      "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

  private OdfdomChartProducer() {}

  private static ChartChartElement chart(OdfChartDocument document) throws Exception {
    NodeList charts = document.getContentRoot().getElementsByTagNameNS(CHART, "chart");
    if (charts.getLength() != 1) {
      throw new IllegalStateException("expected exactly one chart:chart");
    }
    return (ChartChartElement) charts.item(0);
  }

  private static ChartPlotAreaElement plot(ChartChartElement chart) {
    NodeList plots = chart.getElementsByTagNameNS(CHART, "plot-area");
    if (plots.getLength() != 1) {
      throw new IllegalStateException("expected exactly one chart:plot-area");
    }
    return (ChartPlotAreaElement) plots.item(0);
  }

  private static void clearLegacyTemplateIdentity(OdfChartDocument document) {
    OdfOfficeMeta metadata = document.getOfficeMetadata();
    metadata.setAutomaticUpdate(false);
    metadata.setGenerator(null);
    metadata.setCreator(null);
    metadata.setInitialCreator(null);
    metadata.setCreationInstant(null);
    metadata.setInstant(null);
    metadata.setEditingCycles(null);
    metadata.setEditingDuration(null);
  }

  private static void create(File output) throws Exception {
    try (OdfChartDocument document = OdfChartDocument.newChartDocument()) {
      ChartChartElement chart = chart(document);
      chart.setChartClassAttribute("chart:bar");
      ChartPlotAreaElement plot = plot(chart);

      ChartTitleElement title = chart.newChartTitleElement();
      TextPElement paragraph = title.newTextPElement();
      paragraph.newTextNode("ODFDOM standalone chart");
      chart.insertBefore(title, plot);

      ChartAxisElement x = plot.newChartAxisElement("x");
      x.setChartNameAttribute("axis-x");
      ChartAxisElement y = plot.newChartAxisElement("y");
      y.setChartNameAttribute("axis-y");
      ChartSeriesElement series = plot.newChartSeriesElement();
      series.setChartClassAttribute("chart:bar");
      series.setChartAttachedAxisAttribute("axis-y");
      series.newChartDataPointElement();
      clearLegacyTemplateIdentity(document);
      document.save(output);
    }
  }

  private static void change(File input, File output) throws Exception {
    try (OdfChartDocument document = OdfChartDocument.loadDocument(input)) {
      ChartChartElement chart = chart(document);
      NodeList paragraphs = chart.getElementsByTagNameNS(TEXT, "p");
      if (paragraphs.getLength() != 1) {
        throw new IllegalStateException("expected exactly one title paragraph");
      }
      ((TextPElement) paragraphs.item(0)).setTextContent("ODFDOM independently resaved chart");
      NodeList axes = chart.getElementsByTagNameNS(CHART, "axis");
      if (axes.getLength() != 2) {
        throw new IllegalStateException("expected exactly two axes");
      }
      ((ChartAxisElement) axes.item(0)).setChartNameAttribute("axis-x-resaved");
      clearLegacyTemplateIdentity(document);
      document.save(output);
    }
  }

  private static void verify(File input, String expectedTitle, String expectedAxis)
      throws Exception {
    try (OdfChartDocument document = OdfChartDocument.loadDocument(input)) {
      ChartChartElement chart = chart(document);
      String title = chart.getElementsByTagNameNS(TEXT, "p").item(0).getTextContent();
      String axis =
          ((ChartAxisElement) chart.getElementsByTagNameNS(CHART, "axis").item(0))
              .getChartNameAttribute();
      if (!expectedTitle.equals(title) || !expectedAxis.equals(axis)) {
        throw new IllegalStateException("semantic reopen mismatch: " + title + " / " + axis);
      }
      System.out.println(document.getMediaTypeString() + " | " + title + " | " + axis);
    }
  }

  public static void main(String[] args) throws Exception {
    switch (args[0]) {
      case "create" -> create(new File(args[1]));
      case "change" -> change(new File(args[1]), new File(args[2]));
      case "verify" -> verify(new File(args[1]), args[2], args[3]);
      default -> throw new IllegalArgumentException("unknown mode");
    }
  }
}
