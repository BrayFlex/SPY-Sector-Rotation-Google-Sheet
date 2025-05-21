// --- Configuration ---
const SPREADSHEET_NAME_CONFIG = "$SPY Sector Dashboard";
const DASHBOARD_SHEET_NAME = "Sectors";
const STATS_SHEET_NAME = "Stats_Data";
const TEMP_FETCH_SHEET_NAME = "TempDataFetch";

const TICKERS_INFO_RAW_WITH_EMOJIS = [
    { symbol: "SPY", description: "📊SPDR S&P 500 ETF Trust", emoji: "📊" },
    { symbol: "XLK", description: "🤖Technology Select Sector SPDR Fund", emoji: "🤖" },
    { symbol: "XLF", description: "💰Financial Select Sector SPDR Fund", emoji: "💰" },
    { symbol: "XLV", description: "⚕️Health Care Select Sector SPDR Fund", emoji: "⚕️" },
    { symbol: "XLY", description: "🛍️Consumer Discretionary Select Sector SPDR Fund", emoji: "🛍️" },
    { symbol: "XLC", description: "📡Communica- tion Services Select Sector SPDR Fund", emoji: "📡" },
    { symbol: "XLI", description: "🏭Industrial Select Sector SPDR Fund", emoji: "🏭" },
    { symbol: "XLP", description: "🛒Consumer Staples Select Sector SPDR Fund", emoji: "🛒" },
    { symbol: "XLE", description: "🛢️Energy Select Sector SPDR Fund", emoji: "🛢️" },
    { symbol: "XLU", description: "💡Utilities Select Sector SPDR Fund", emoji: "💡" },
    { symbol: "XLRE", description: "🏠Real Estate Select Sector SPDR Fund", emoji: "🏠" },
    { symbol: "XLB", description: "⛏️Materials Select Sector SPDR Fund", emoji: "⛏️" },
];

const PHRASE_TO_REMOVE = "Select Sector SPDR Fund";
const TICKERS_INFO = TICKERS_INFO_RAW_WITH_EMOJIS.map(ticker => {
    const baseDescription = ticker.description
        .replace(ticker.emoji, "")
        .replace(PHRASE_TO_REMOVE, "")
        .trim()
        .replace(/  +/g, ' ');
    return {
        symbol: ticker.symbol,
        description: `${ticker.emoji} ${baseDescription}`.trim()
    };
});

const COLORS = {
    VERY_LOW: "#df6262", LOW: "#eaa6a5", MEDIUM_LOW: "#f2cecd",
    MEDIUM_HIGH: "#bbe3dc", HIGH: "#8cc9bd", VERY_HIGH: "#5aa794",
    TEXT_DARK: "#000000", TEXT_LIGHT: "#ffffff",
    HEADER_BG: "#4a86e8", HEADER_TEXT: "#ffffff", BORDER_COLOR: "#b7b7b7"
};

const STATS_LOOKBACK_COUNT = 252;
const HISTORICAL_DATA_FETCH_YEARS = 3;
const TRIGGER_FUNCTION_NAME = "automatedStatisticalRefresh";

// --- Main Function to be Run Manually Once ---
function createCompleteDashboardAndSetupTrigger() {
    Logger.log("Dashboard Creation & Trigger Setup Starting...");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss.getName() !== SPREADSHEET_NAME_CONFIG) {
        ss.rename(SPREADSHEET_NAME_CONFIG);
    }

    let dashboardSheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
    if (dashboardSheet) {
        dashboardSheet.clear();
    } else {
        dashboardSheet = ss.insertSheet(DASHBOARD_SHEET_NAME, 0);
    }
    ss.setActiveSheet(dashboardSheet);

    let statsSheet = ss.getSheetByName(STATS_SHEET_NAME);
    if (statsSheet) {
      const isHidden = statsSheet.isSheetHidden();
      if (isHidden) statsSheet.showSheet();
      statsSheet.clear();
      if (isHidden && statsSheet.getMaxRows() > 1) statsSheet.hideSheet();
      else if (statsSheet.getMaxRows() === 1 && statsSheet.getMaxColumns() === 1 && statsSheet.getName() === STATS_SHEET_NAME) {
          statsSheet.hideSheet();
      }
    } else {
        statsSheet = ss.insertSheet(STATS_SHEET_NAME, 1);
        statsSheet.hideSheet();
    }

    setupDashboardSheetLayoutFormatted(dashboardSheet); // Call setup first

    if (statsSheet.isSheetHidden()) statsSheet.showSheet();
    statsSheet.clear();
    calculateAndStoreStatisticalParameters(ss, statsSheet);

    populateDashboardFormulasSimple(dashboardSheet); // Populate data (which might affect auto-column sizing if not careful)

    if (!statsSheet.isSheetHidden() && statsSheet.getLastRow() > 0) {
        statsSheet.hideSheet();
    }

    dashboardSheet.clearConditionalFormatRules();
    applyConditionalFormattingDirect(dashboardSheet, STATS_SHEET_NAME);

    const lastRowDashboard = dashboardSheet.getMaxRows();
    const legendStartRow = TICKERS_INFO.length + 3;
    if (lastRowDashboard >= legendStartRow) {
        dashboardSheet.getRange(legendStartRow, 1, lastRowDashboard - legendStartRow + 1, dashboardSheet.getMaxColumns()).clearContent().clearFormat();
    }
    createSimpleLegend(dashboardSheet);

    let tempSheet = ss.getSheetByName(TEMP_FETCH_SHEET_NAME);
    if (tempSheet) {
        try { ss.deleteSheet(tempSheet); } catch (e) { Logger.log("Could not delete temp sheet on build: " + e); tempSheet.clear(); }
    }

    createOrUpdateDailyTriggerNoMenu();

    SpreadsheetApp.flush();
    Logger.log("Dashboard built successfully and daily trigger is set/updated!");
    SpreadsheetApp.getActiveSpreadsheet().toast("Dashboard setup complete. Daily refresh trigger active.", "Status", 5);
}

// --- Sheet Setup (Formatted with new row height, font sizes, column widths) ---
function setupDashboardSheetLayoutFormatted(sheet) {
    const headers = [
        "Ticker",        // 1 (A)
        "Description",   // 2 (B)
        "Price",         // 3 (C)
        "%Chg\n1D",       // 4 (D)
        "%Chg\n1W",       // 5 (E)
        "%Chg\n2W",       // 6 (F)
        "%Chg\n1M",       // 7 (G)
        "%Chg\n2M",       // 8 (H)
        "Volume",        // 9 (I)
        "Avg Vol",       // 10 (J)
        "% vs\n20 SMA",      // 11 (K) - Multi-line
        "% vs\n50 SMA",      // 12 (L) - Multi-line
        "% vs\n200 SMA"     // 13 (M) - Multi-line
    ];
    sheet.appendRow(headers); // Appending values with \n directly works for multi-line
    sheet.setFrozenRows(1);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold")
               .setBackground(COLORS.HEADER_BG)
               .setFontColor(COLORS.TEXT_LIGHT)
               .setHorizontalAlignment("center")
               .setVerticalAlignment("middle")
               .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Ensure wrapping is enabled

    sheet.setRowHeight(1, 45); // Slightly increased header row height for two lines

    const columnWidths = [
        52,  // A: Ticker
        112, // B: Description
        72,  // C: Price
        72, 72, 72, 72, 72, // D-H: %Chg
        96, 96, // I-J: Volume
        72, 72, 72  // K-M: %vsSMA
    ];
    columnWidths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
    Logger.log("Initial column widths set. Ticker (Col 1) set to: " + columnWidths[0]);


    const dataRows = TICKERS_INFO.length;
    sheet.getRange(2, 1, dataRows, headers.length).setVerticalAlignment("middle");

    sheet.getRange(2, 1, dataRows, 1).setFontSize(12).setHorizontalAlignment("center");
    sheet.getRange(2, 2, dataRows, 1).setFontSize(10).setWrap(true);

    TICKERS_INFO.forEach((_, index) => {
        sheet.setRowHeight(index + 2, 48);
    });

    const percentCols = [4,5,6,7,8,11,12,13];
    percentCols.forEach(colIdx => {
        sheet.getRange(2, colIdx, dataRows, 1).setNumberFormat("0.00%").setFontSize(10);
    });
    sheet.getRange(2,3,dataRows,1).setFontSize(10);
    sheet.getRange(2,9,dataRows,2).setFontSize(10);
}

// populateDashboardFormulasSimple remains the same
function populateDashboardFormulasSimple(sheet) {
    const startRow = 2;
    TICKERS_INFO.forEach((tickerInfo, index) => {
        const row = startRow + index;
        const tickerSymbol = tickerInfo.symbol;
        sheet.getRange(row, 1).setValue(tickerSymbol); 
        sheet.getRange(row, 2).setValue(tickerInfo.description);

        const priceCell = sheet.getRange(row, 3).setFormula(`=IFERROR(GOOGLEFINANCE("${tickerSymbol}","price"), NA())`).setNumberFormat("$0.00").getA1Notation();
        
        sheet.getRange(row, 4).setFormula(`=IFERROR(GOOGLEFINANCE("${tickerSymbol}","changepct")/100, NA())`);
        sheet.getRange(row, 5).setFormula(`=IFERROR((${priceCell}/INDEX(GOOGLEFINANCE("${tickerSymbol}","close",WORKDAY(TODAY(),-5)),2,2))-1, NA())`);
        sheet.getRange(row, 6).setFormula(`=IFERROR((${priceCell}/INDEX(GOOGLEFINANCE("${tickerSymbol}","close",WORKDAY(TODAY(),-10)),2,2))-1, NA())`);
        sheet.getRange(row, 7).setFormula(`=IFERROR((${priceCell}/INDEX(GOOGLEFINANCE("${tickerSymbol}","close",EDATE(TODAY(),-1)),2,2))-1, NA())`);
        sheet.getRange(row, 8).setFormula(`=IFERROR((${priceCell}/INDEX(GOOGLEFINANCE("${tickerSymbol}","close",EDATE(TODAY(),-2)),2,2))-1, NA())`);

        sheet.getRange(row, 9).setFormula(`=IFERROR(GOOGLEFINANCE("${tickerSymbol}","volume"), NA())`).setNumberFormat("#,##0");
        sheet.getRange(row, 10).setFormula(`=IFERROR(AVERAGE(QUERY(GOOGLEFINANCE("${tickerSymbol}","volume",WORKDAY(TODAY(),-45),TODAY()),"SELECT Col2 offset 1",0)), NA())`).setNumberFormat("#,##0");
        
        const smaBaseFormula = (period) => `AVERAGE(QUERY(GOOGLEFINANCE("${tickerSymbol}","close",WORKDAY(TODAY(),-${period+10}),TODAY()),"SELECT Col2 order by Col1 desc limit ${period} offset 1",0))`;
        sheet.getRange(row, 11).setFormula(`=IFERROR((${priceCell}/${smaBaseFormula(20)})-1, NA())`);
        sheet.getRange(row, 12).setFormula(`=IFERROR((${priceCell}/${smaBaseFormula(50)})-1, NA())`);
        sheet.getRange(row, 13).setFormula(`=IFERROR((${priceCell}/${smaBaseFormula(200)})-1, NA())`);
    });
}

// calculateAndStoreStatisticalParameters remains the same
function calculateAndStoreStatisticalParameters(ss, statsSheet) {
    Logger.log("Starting calculation of statistical parameters.");
    const today = new Date();
    const startDate = new Date(today);
    startDate.setFullYear(today.getFullYear() - HISTORICAL_DATA_FETCH_YEARS);

    const statCategories = ["1D", "1W", "2W", "1M", "2M"];
    const smaDevCategories = ["20DSMA", "50DSMA", "200DSMA"]; // These need to match what CF logic expects if it builds keys
    const statsHeaders = ["Ticker"];
    statCategories.forEach(cat => { statsHeaders.push(`CHG_${cat}_MU`, `CHG_${cat}_SIGMA`); });
    smaDevCategories.forEach(cat => { statsHeaders.push(`DEV_${cat}_MU`, `DEV_${cat}_SIGMA`); });
    statsSheet.appendRow(statsHeaders);
    SpreadsheetApp.flush();

    let tempSheet = ss.getSheetByName(TEMP_FETCH_SHEET_NAME);
    if (!tempSheet) {
        tempSheet = ss.insertSheet(TEMP_FETCH_SHEET_NAME);
        tempSheet.hideSheet();
    } else {
        tempSheet.clearContents();
        if (!tempSheet.isSheetHidden()) tempSheet.hideSheet();
    }

    TICKERS_INFO_RAW_WITH_EMOJIS.forEach(tickerInfo => {
        const ticker = tickerInfo.symbol;
        Logger.log(`Fetching and processing historical data for ${ticker}...`);
        const rowData = [ticker];
        let closes = [];
        try {
            const formula = `=GOOGLEFINANCE("${ticker}", "close", DATE(${startDate.getFullYear()},${startDate.getMonth()+1},${startDate.getDate()}), DATE(${today.getFullYear()},${today.getMonth()+1},${today.getDate()}))`;
            tempSheet.getRange("A1").setFormula(formula);
            SpreadsheetApp.flush();
            Utilities.sleep(2500);

            const fetchedRange = tempSheet.getDataRange();
            const rawData = fetchedRange.getValues();
            tempSheet.clearContents();

            if (rawData.length > 1) {
                for (let k = 1; k < rawData.length; k++) {
                    if (rawData[k][1] !== "" && !isNaN(rawData[k][1])) {
                        closes.push(parseFloat(rawData[k][1]));
                    }
                }
            }
            if (closes.length < (200 + 42 + STATS_LOOKBACK_COUNT)/3 ) {
                 Logger.log(`Warning: Insufficient historical data for ${ticker} (got ${closes.length} points). Stats may be less reliable.`);
                 statCategories.forEach(() => rowData.push(0, 0.0001));
                 smaDevCategories.forEach(() => rowData.push(0, 0.0001));
                 statsSheet.appendRow(rowData);
                 Utilities.sleep(500);
                 return;
            }
        } catch (e) {
            Logger.log(`Error fetching data for ${ticker}: ${e}. Skipping stats for this ticker.`);
            statCategories.forEach(() => rowData.push(0, 0.0001));
            smaDevCategories.forEach(() => rowData.push(0, 0.0001));
            statsSheet.appendRow(rowData);
            Utilities.sleep(500);
            return;
        }

        const periods = [{p:1, n:"1D"}, {p:5, n:"1W"}, {p:10, n:"2W"}, {p:21, n:"1M"}, {p:42, n:"2M"}];
        periods.forEach(periodInfo => {
            const changes = [];
            for (let i = periodInfo.p; i < closes.length; i++) {
                if (closes[i-periodInfo.p] !== 0) changes.push((closes[i] / closes[i-periodInfo.p]) - 1);
            }
            const relevantChanges = changes.slice(-STATS_LOOKBACK_COUNT);
            const { mean, stdev } = calculateMeanAndStdev(relevantChanges);
            rowData.push(mean, stdev === 0 ? 0.0001 : stdev);
        });

        // For DEV_MU and DEV_SIGMA, ensure the statBase used here matches what CF expects
        // CF uses DEV_20DSMA_SIGMA, DEV_50DSMA_SIGMA, DEV_200DSMA_SIGMA
        const smaCalcPeriods = [20, 50, 200];
        smaCalcPeriods.forEach((smaPeriod) => { // idx is not needed if we use smaPeriod to build the key
            const deviations = [];
            for (let i = smaPeriod -1; i < closes.length; i++) {
                let sumSma = 0;
                for (let j=0; j < smaPeriod; j++) sumSma += closes[i-j];
                const sma = sumSma / smaPeriod;
                if (sma !== 0) deviations.push((closes[i] / sma) - 1);
            }
            const relevantDeviations = deviations.slice(-STATS_LOOKBACK_COUNT);
            const { mean, stdev } = calculateMeanAndStdev(relevantDeviations);
            rowData.push(mean, stdev === 0 ? 0.0001 : stdev);
        });

        statsSheet.appendRow(rowData);
        SpreadsheetApp.flush();
        Utilities.sleep(1000);
        Logger.log(`Finished calculations for ${ticker}.`);
    });
    Logger.log("Statistical parameters calculation complete.");
}


// calculateMeanAndStdev remains the same
function calculateMeanAndStdev(dataArray) {
    if (!dataArray || dataArray.length < 2) return { mean: 0, stdev: 0.0001 };
    const n = dataArray.length;
    const mean = dataArray.reduce((a, b) => a + b, 0) / n;
    const variance = dataArray.reduce((sq, val) => sq + Math.pow(val - mean, 2), 0) / (n - 1);
    const stdev = Math.sqrt(variance);
    return { mean, stdev };
}

// applyConditionalFormattingDirect remains the same
function applyConditionalFormattingDirect(dashboardSheet, statsSheetName) {
    Logger.log("Applying conditional formatting directly referencing Stats_Data.");
    dashboardSheet.clearConditionalFormatRules();
    const numDataRows = TICKERS_INFO.length;
    const rules = [];

    const getStatCfFormula = (tickerCellA1Relative, statHeaderName) => {
        const indirectStatsSheetRange = `INDIRECT("'${statsSheetName}'!$A:$AZ")`;
        const indirectStatsTickerCol = `INDIRECT("'${statsSheetName}'!$A:$A")`;
        const indirectStatsHeaderRow = `INDIRECT("'${statsSheetName}'!$1:$1")`;
        return `IFERROR(INDEX(${indirectStatsSheetRange}, MATCH(${tickerCellA1Relative}, ${indirectStatsTickerCol}, 0), MATCH("${statHeaderName}", ${indirectStatsHeaderRow}, 0)), 0)`;
    };

    const colors = [COLORS.VERY_LOW, COLORS.LOW, COLORS.MEDIUM_LOW, COLORS.MEDIUM_HIGH, COLORS.HIGH, COLORS.VERY_HIGH];

    const chgCols = [
        { dataCol: 4, statBase: "1D" }, { dataCol: 5, statBase: "1W" },
        { dataCol: 6, statBase: "2W" }, { dataCol: 7, statBase: "1M" },
        { dataCol: 8, statBase: "2M" }
    ];

    chgCols.forEach(colInfo => {
        const range = dashboardSheet.getRange(2, colInfo.dataCol, numDataRows, 1);
        const cellA1 = dashboardSheet.getRange(2, colInfo.dataCol).getA1Notation();
        const tickerA1Relative = `$A2`;

        const muFormula = getStatCfFormula(tickerA1Relative, `CHG_${colInfo.statBase}_MU`);
        const sigmaFormula = getStatCfFormula(tickerA1Relative, `CHG_${colInfo.statBase}_SIGMA`);

        const cfConditions = [
            `AND(ISNUMBER(${cellA1}), ISNUMBER(${muFormula}), ISNUMBER(${sigmaFormula}), ${sigmaFormula}<>0, ${cellA1} < ${muFormula} - 2 * ${sigmaFormula})`,
            `AND(ISNUMBER(${cellA1}), ISNUMBER(${muFormula}), ISNUMBER(${sigmaFormula}), ${sigmaFormula}<>0, ${cellA1} >= ${muFormula} - 2 * ${sigmaFormula}, ${cellA1} < ${muFormula} - 1 * ${sigmaFormula})`,
            `AND(ISNUMBER(${cellA1}), ISNUMBER(${muFormula}), ISNUMBER(${sigmaFormula}), ${sigmaFormula}<>0, ${cellA1} >= ${muFormula} - 1 * ${sigmaFormula}, ${cellA1} < ${muFormula})`,
            `AND(ISNUMBER(${cellA1}), ISNUMBER(${muFormula}), ISNUMBER(${sigmaFormula}), ${sigmaFormula}<>0, ${cellA1} >= ${muFormula}, ${cellA1} < ${muFormula} + 1 * ${sigmaFormula})`,
            `AND(ISNUMBER(${cellA1}), ISNUMBER(${muFormula}), ISNUMBER(${sigmaFormula}), ${sigmaFormula}<>0, ${cellA1} >= ${muFormula} + 1 * ${sigmaFormula}, ${cellA1} < ${muFormula} + 2 * ${sigmaFormula})`,
            `AND(ISNUMBER(${cellA1}), ISNUMBER(${muFormula}), ISNUMBER(${sigmaFormula}), ${sigmaFormula}<>0, ${cellA1} >= ${muFormula} + 2 * ${sigmaFormula})`
        ];

        cfConditions.forEach((cond, i) => {
            rules.push(SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied(`=${cond}`)
                .setBackground(colors[i])
                .setRanges([range])
                .build());
        });
    });

    const smaCols = [ // These statBase names MUST match the headers in Stats_Data sheet
        { dataCol: 11, statBase: "20DSMA" },
        { dataCol: 12, statBase: "50DSMA" },
        { dataCol: 13, statBase: "200DSMA" }
    ];
    smaCols.forEach(colInfo => {
        const range = dashboardSheet.getRange(2, colInfo.dataCol, numDataRows, 1);
        const cellA1 = dashboardSheet.getRange(2, colInfo.dataCol).getA1Notation();
        const tickerA1Relative = `$A2`;
        const muReference = 0;
        const sigmaFormula = getStatCfFormula(tickerA1Relative, `DEV_${colInfo.statBase}_SIGMA`);

        const cfConditions = [
            `AND(ISNUMBER(${cellA1}), ISNUMBER(${sigmaFormula}), ${sigmaFormula}<>0, ${cellA1} < ${muReference} - 2 * ${sigmaFormula})`,
            `AND(ISNUMBER(${cellA1}), ISNUMBER(${sigmaFormula}), ${sigmaFormula}<>0, ${cellA1} >= ${muReference} - 2 * ${sigmaFormula}, ${cellA1} < ${muReference} - 1 * ${sigmaFormula})`,
            `AND(ISNUMBER(${cellA1}), ISNUMBER(${sigmaFormula}), ${sigmaFormula}<>0, ${cellA1} >= ${muReference} - 1 * ${sigmaFormula}, ${cellA1} < ${muReference})`,
            `AND(ISNUMBER(${cellA1}), ISNUMBER(${sigmaFormula}), ${sigmaFormula}<>0, ${cellA1} >= ${muReference}, ${cellA1} < ${muReference} + 1 * ${sigmaFormula})`,
            `AND(ISNUMBER(${cellA1}), ISNUMBER(${sigmaFormula}), ${sigmaFormula}<>0, ${cellA1} >= ${muReference} + 1 * ${sigmaFormula}, ${cellA1} < ${muReference} + 2 * ${sigmaFormula})`,
            `AND(ISNUMBER(${cellA1}), ISNUMBER(${sigmaFormula}), ${sigmaFormula}<>0, ${cellA1} >= ${muReference} + 2 * ${sigmaFormula})`
        ];
        cfConditions.forEach((cond, i) => {
            rules.push(SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied(`=${cond}`)
                .setBackground(colors[i])
                .setRanges([range])
                .build());
        });
    });

    if (rules.length > 0) {
      dashboardSheet.setConditionalFormatRules(rules);
    }
    Logger.log("Conditional formatting applied.");
}


// createSimpleLegend remains the same
function createSimpleLegend(sheet) {
    const legendStartRow = TICKERS_INFO.length + 3;
    sheet.getRange(legendStartRow, 1).setValue("Color Legend for % Change Columns:").setFontWeight("bold").setFontSize(10);
    const legendPctChange = [
        { color: COLORS.VERY_LOW, text: "Value < Historical μ - 2σ (Very Low)" },
        { color: COLORS.LOW, text: "Historical μ - 2σ ≤ Value < μ - σ (Low)" },
        { color: COLORS.MEDIUM_LOW, text: "Historical μ - σ ≤ Value < μ (Medium Low)" },
        { color: COLORS.MEDIUM_HIGH, text: "Historical μ ≤ Value < μ + σ (Medium High)" },
        { color: COLORS.HIGH, text: "Historical μ + σ ≤ Value < μ + 2σ (High)" },
        { color: COLORS.VERY_HIGH, text: "Value ≥ Historical μ + 2σ (Very High)" }
    ];

    let currentLegendRow = legendStartRow + 1;
    legendPctChange.forEach(item => {
        sheet.getRange(currentLegendRow, 1).setBackground(item.color);
        sheet.getRange(currentLegendRow, 2).setValue(item.text).setFontSize(9);
        sheet.getRange(currentLegendRow, 2, 1, 5).merge();
        currentLegendRow++;
    });
    sheet.getRange(currentLegendRow, 2).setValue("(μ = hist. mean, σ = hist. std dev for period's % changes)").setFontStyle("italic").setFontSize(8);
    sheet.getRange(currentLegendRow, 2, 1, 5).merge();
    currentLegendRow += 2;

    sheet.getRange(currentLegendRow, 1).setValue("Color Legend for % vs DSMA Columns:").setFontWeight("bold").setFontSize(10);
    const legendDsma = [
        { color: COLORS.VERY_LOW, text: "Deviation < 0 - 2σ_dev (Very Low / Far Below DMA)" },
        { color: COLORS.LOW, text: "0 - 2σ_dev ≤ Dev < 0 - σ_dev (Low / Below DMA)" },
        { color: COLORS.MEDIUM_LOW, text: "0 - σ_dev ≤ Dev < 0 (Medium Low / Slightly Below DMA)" },
        { color: COLORS.MEDIUM_HIGH, text: "0 ≤ Dev < 0 + σ_dev (Medium High / Slightly Above DMA)" },
        { color: COLORS.HIGH, text: "0 + σ_dev ≤ Dev < 0 + 2σ_dev (High / Above DMA)" },
        { color: COLORS.VERY_HIGH, text: "Dev ≥ 0 + 2σ_dev (Very High / Far Above DMA)" }
    ];
    currentLegendRow++;
    legendDsma.forEach(item => {
        sheet.getRange(currentLegendRow, 1).setBackground(item.color);
        sheet.getRange(currentLegendRow, 2).setValue(item.text).setFontSize(9);
        sheet.getRange(currentLegendRow, 2, 1, 5).merge();
        currentLegendRow++;
    });
    sheet.getRange(currentLegendRow, 2).setValue("(σ_dev = hist. std dev of % deviation from DSMA. Mean dev ref ~0%)").setFontStyle("italic").setFontSize(8);
    sheet.getRange(currentLegendRow, 2, 1, 5).merge();

    // sheet.setColumnWidth(1, 30); // For legend color swatch
    sheet.getRange(legendStartRow, 2, (currentLegendRow - legendStartRow +1), 1).setWrap(true);
}

// automatedStatisticalRefresh remains the same
function automatedStatisticalRefresh() {
    Logger.log("Automated statistical refresh started by trigger.");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const statsSheet = ss.getSheetByName(STATS_SHEET_NAME);
    const dashboardSheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);

    if (!statsSheet || !dashboardSheet) {
        Logger.log("Error: Dashboard or Stats sheet not found. Automated refresh cannot proceed.");
        return;
    }

    const isStatsHidden = statsSheet.isSheetHidden();
    if (isStatsHidden) statsSheet.showSheet();

    statsSheet.clear();
    calculateAndStoreStatisticalParameters(ss, statsSheet);

    if (isStatsHidden && statsSheet.getLastRow() > 0) statsSheet.hideSheet();

    dashboardSheet.clearConditionalFormatRules();
    applyConditionalFormattingDirect(dashboardSheet, STATS_SHEET_NAME);

    const lastRowDashboard = dashboardSheet.getMaxRows();
    const legendStartRowForClear = TICKERS_INFO.length + 3;
    if (lastRowDashboard >= legendStartRowForClear) {
        dashboardSheet.getRange(legendStartRowForClear, 1, lastRowDashboard - legendStartRowForClear + 1, dashboardSheet.getMaxColumns()).clearContent().clearFormat();
    }
    createSimpleLegend(dashboardSheet);

    let tempSheet = ss.getSheetByName(TEMP_FETCH_SHEET_NAME);
    if (tempSheet) {
      try { ss.deleteSheet(tempSheet); } catch (e) { Logger.log("Could not delete temp sheet on refresh: " + e); tempSheet.clearContents(); }
    }
    SpreadsheetApp.flush();
    Logger.log("Automated statistical refresh completed successfully.");
}

// createOrUpdateDailyTriggerNoMenu and deleteAllProjectTriggers remain the same
function createOrUpdateDailyTriggerNoMenu() {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
        if (trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
            ScriptApp.deleteTrigger(trigger);
            Logger.log(`Deleted existing trigger for ${TRIGGER_FUNCTION_NAME}.`);
        }
    }
    ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
        .timeBased()
        .atHour(4)
        .everyDays(1)
        .create();
    Logger.log(`Created new daily trigger for ${TRIGGER_FUNCTION_NAME} to run around 4 AM (script project timezone). ` +
               `Ensure script project timezone is set correctly (e.g., America/New_York in appsscript.json).`);
}

function deleteAllProjectTriggers() {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
        ScriptApp.deleteTrigger(trigger);
    }
    Logger.log(triggers.length > 0 ? "All project triggers deleted." : "No project triggers found to delete.");
}