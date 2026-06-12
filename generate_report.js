const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
        AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType, ShadingType,
        PageBreak } = require('docx');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, opts={}) {
  return new TableCell({
    borders,
    width: { size: opts.width || 1560, type: WidthType.DXA },
    shading: opts.header ? { fill: "1B5E20", type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.CENTER,
      children: [new TextRun({ text, bold: !!opts.header, color: opts.header ? "FFFFFF" : "000000", size: 20 })]
    })]
  });
}

function row(cells, header=false) {
  return new TableRow({ children: cells.map(c => cell(c, {header})) });
}

function img(path, w, h) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [new ImageRun({ data: fs.readFileSync(path), type: "png", transformation: { width: w, height: h } })]
  });
}

function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}
function p(text, opts={}) {
  return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] });
}
function bullet(text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [new TextRun(text)] });
}

const COLW = 1560; // 6 cols * 1560 = 9360

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "1B5E20" },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "333333" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 } },
    ]
  },
  numbering: {
    config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }]
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1440, bottom: 1080, left: 1440 } } },
    children: [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
        children: [new TextRun({ text: "Trader Performance vs. Bitcoin Market Sentiment", bold: true, size: 36, color: "1B5E20" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 },
        children: [new TextRun({ text: "An analysis of Hyperliquid trader data against the Fear & Greed Index", italics: true, size: 22, color: "666666" })] }),

      h1("1. Executive Summary"),
      p("This report examines 211,224 trades across 32 Hyperliquid accounts (Feb 2023 \u2013 May 2025), merged with the Bitcoin Fear & Greed Index by date (99.997% match rate). The goal is to identify how trader profitability, risk behavior, and positioning shift across market sentiment regimes."),
      p("Key takeaways:"),
      bullet("Traders are most profitable in absolute terms during \u201cFear\u201d and \u201cExtreme Greed\u201d days \u2014 total realized PnL of $3.36M and $2.72M respectively \u2014 but \u201cExtreme Greed\u201d delivers the highest average PnL per trade ($130) and the highest win rate (89.2%)."),
      bullet("\u201cExtreme Fear\u201d is the weakest regime: lowest win rate (76.2%) and lowest profit factor (2.16), suggesting panic-driven trading is the least reliable."),
      bullet("Average position sizes are largest during Fear ($7,375) and smallest during Greed ($4,234) \u2014 traders appear to size up when the market is fearful, possibly viewing dips as opportunities or being forced into larger defensive trades."),
      bullet("Short positions are opened proportionally more often during Greed (21.4% of opens vs 16.9% during Fear), consistent with contrarian/short-the-rally behavior; long positions dominate during Fear (29.8% of opens)."),
      bullet("Overall, Greed days produced the highest total PnL ($4.87M) and best profit factor (4.66), while Fear days were close behind ($4.10M total PnL, profit factor 4.33). Neutral days were the weakest broad regime (profit factor 4.32, total PnL $1.29M)."),

      h1("2. Dataset Overview"),
      h2("2.1 Data Sources"),
      bullet("Hyperliquid Historical Trader Data: 211,224 trade/order records across 32 accounts and 246 traded assets, with execution price, size (tokens & USD), side, direction, closed PnL, fees, and timestamps."),
      bullet("Bitcoin Fear & Greed Index: daily classification (Extreme Fear, Fear, Neutral, Greed, Extreme Greed), covering Feb 2018 \u2013 May 2025."),
      h2("2.2 Methodology"),
      bullet("Trade timestamps (IST) were converted to dates and joined to the daily Fear & Greed classification."),
      bullet("\u201cClosed trades\u201d are rows where Closed PnL \u2260 0 (i.e., realized PnL events) \u2014 used for profitability, win-rate, and profit-factor metrics. All 211,224 rows are used for position-sizing and directional analysis."),
      bullet("Broad sentiment groups combine Extreme Fear + Fear \u2192 \u201cFear\u201d, and Extreme Greed + Greed \u2192 \u201cGreed\u201d, with Neutral standing alone."),

      h1("3. Performance by Sentiment Classification"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1660,1340,1500,1300,1300,1480,1280],
        rows: [
          row(["Classification","Trades","Total PnL ($)","Avg PnL ($)","Median PnL ($)","Win Rate","Profit Factor"], true),
          row(["Extreme Fear","10,406","739,110","71.03","6.39","76.2%","2.16"]),
          row(["Fear","29,808","3,357,155","112.63","6.35","87.3%","6.66"]),
          row(["Neutral","18,159","1,292,921","71.20","4.58","82.4%","4.32"]),
          row(["Greed","25,176","2,150,129","85.40","4.93","76.9%","3.03"]),
          row(["Extreme Greed","20,853","2,715,171","130.21","8.53","89.2%","11.02"]),
        ]
      }),
      new Paragraph({ spacing: { before: 200 }, children: [] }),
      img("/home/claude/chart1_avg_pnl.png", 520, 330),
      p("Average PnL per trade is U-shaped-ish across the sentiment spectrum: it dips during ordinary Fear/Neutral/Greed but spikes sharply at \u201cExtreme Greed\u201d ($130.21) \u2014 likely reflecting strong trending markets where momentum trades pay off well. \u201cExtreme Fear\u201d trades are profitable on average but far less so ($71.03)."),
      img("/home/claude/chart2_winrate.png", 520, 330),
      p("Win rates follow a similar pattern: Extreme Greed (89.2%) and Fear (87.3%) are the most favorable regimes for closing winning trades, while Extreme Fear (76.2%) and Greed (76.9%) are noticeably weaker."),
      img("/home/claude/chart4_totalpnl.png", 520, 330),
      p("In aggregate dollar terms, \u201cFear\u201d days generated the largest pool of realized profit ($3.36M), followed by \u201cExtreme Greed\u201d ($2.72M) and \u201cGreed\u201d ($2.15M). \u201cExtreme Fear\u201d contributed the least ($739K), consistent with both lower trade volume and weaker per-trade economics in that regime."),

      h1("4. Position Sizing and Risk Behavior"),
      img("/home/claude/chart3_size.png", 520, 330),
      p("Average trade size (USD) is highest during Fear ($7,375) and Extreme Fear ($5,350 \u2014 see classification-level data), and lowest during Greed ($4,234). This suggests traders commit larger notional positions when sentiment is fearful \u2014 either viewing fear-driven dips as buying opportunities, or because volatility-driven price swings inflate position values."),

      h1("5. Long/Short Positioning by Sentiment"),
      img("/home/claude/chart5_longshort.png", 520, 330),
      p("During Fear, 63.9% of new directional positions opened are Longs vs 36.1% Shorts \u2014 a clear \u201cbuy the fear\u201d tilt. During Greed, the split narrows to 43.4% Long / 56.6% Short, indicating traders increasingly position for reversals or fade rallies as greed intensifies. Neutral days sit closer to the Fear pattern (61.6% Long / 38.4% Short)."),

      h1("6. Broad Sentiment Comparison (Fear vs Neutral vs Greed)"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1872,1872,1872,1872,1872],
        rows: [
          row(["Sentiment","Trades","Total PnL ($)","Avg PnL ($)","Profit Factor"], true),
          row(["Fear","83,237","4,096,266","49.21","4.33"]),
          row(["Neutral","37,686","1,292,921","34.31","4.32"]),
          row(["Greed","90,295","4,865,301","53.88","4.66"]),
        ]
      }),
      new Paragraph({ spacing: { before: 200 }, children: [] }),
      p("At the broad level, Greed days are marginally the most efficient (profit factor 4.66, avg PnL/trade $53.88), with Fear close behind (4.33 / $49.21). Neutral days are the laggard on both total and per-trade profitability \u2014 markets with no strong sentiment signal appear to offer the least edge."),

      h1("7. Recommendations for Trading Strategy"),
      bullet("Favor momentum / trend-following setups during Extreme Greed: this regime shows the highest win rate (89.2%) and profit factor (11.02), suggesting strong directional follow-through is more reliable here than in any other regime."),
      bullet("Treat Extreme Fear with caution or use it for smaller, higher-conviction entries: it has the lowest win rate (76.2%) and profit factor (2.16) despite traders sizing up ($5,350 avg) \u2014 larger bets in the worst-performing regime is a risk-management red flag worth investigating per-account."),
      bullet("Consider contrarian short bias during Greed and long bias during Fear, consistent with observed trader behavior (56.6% shorts in Greed vs 63.9% longs in Fear), which appears to align with the better-than-average profit factors in both regimes."),
      bullet("Investigate position sizing discipline: the gap between average size in Fear ($7,375) vs Greed ($4,234) is large. If this is driven by a few accounts, individual risk limits may need review to avoid concentration risk during volatile, low-win-rate periods."),
      bullet("Use Neutral-sentiment days as a baseline / lower-conviction trading window \u2014 both total and per-trade PnL are weakest here, implying fewer high-quality setups when sentiment lacks a clear extreme."),

      h1("8. Appendix: Data Notes"),
      bullet("211,218 of 211,224 trades (99.997%) matched a Fear & Greed Index date; 6 unmatched records were excluded from sentiment-based analysis."),
      bullet("\u201cClosed PnL = 0\u201d rows (order placements without a realization event) were excluded from PnL, win-rate, and profit-factor calculations but included in position-sizing and directional analysis."),
      bullet("Profit Factor = Gross Profit \u00f7 Gross Loss across closed trades."),
      bullet("Full summary tables are provided in the accompanying Excel workbook (trader_sentiment_analysis.xlsx)."),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => fs.writeFileSync("/home/claude/report/Trader_Sentiment_Analysis_Report.docx", buf));
