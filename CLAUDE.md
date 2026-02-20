# TRADE.xlsm — Architecture Reference

## What This Project Is

TRADE.xlsm is a complete automated stock screening, technical analysis, and trading signal system built in Excel VBA. It supports multiple markets (Canada TSX/TSXV, USA, International, ETFs, AI stocks) and covers the full pipeline from data retrieval through signal generation, backtesting, reporting, and cloud backup.

**Main file:** `TRADE.xlsm` (51 MB, in this folder)
**VBA source exports:** 30 active `.bas` files in this folder (use these for code editing)
**External dependency:** `drive_upload.py` (Google Drive backup script, in this folder)

---

## System Overview — What It Does

1. **Fetches** historical OHLCV data for hundreds of stocks via Excel's `STOCKHISTORY` function
2. **Validates** data quality — detects anomalies using % change thresholds and 3-sigma statistical outliers
3. **Calculates** 14+ technical indicators: RSI, MACD, EMA, Bollinger Bands, ATR, DMI, SMI, Bressert DSS, Elder-Ray, Keltner Channels, Volume Profile, OBV
4. **Scores** stocks via a composite 0–100 system (8 weighted financial metrics, sector-adjusted percentile) and DCF fair value ratings (1–10)
5. **Filters** stocks by price range, minimum score, and **Origin** (USA / Canada / ETF / INTL / ALL) in batches of 50 tickers
6. **Generates** BUY/SELL/HOLD signals with stop loss, position sizing, and R/R ratios using ATR-based risk management
7. **Reports** results to worksheet dashboards and sends HTML email via Outlook
8. **Backtests** strategies by iterating over historical date ranges (weekly/daily modes)
9. **Schedules** automated runs at 3:00 AM on valid US trading days (holiday-aware)
10. **Backs up** data to Google Drive via an external Python script

---

## Data Flow

```
[Excel STOCKHISTORY] → DataHistory sheet → BackupAll sheet
                                                  │
                                    [ALL.bas: ProcessAll()]
                                                  │
                               TJX sheet (master ticker list + fundamentals)
                                                  │
                                         [Filtering.bas: FilterAndReport()]
                           Filter by price range + Origin (USA/Canada/ETF/INTL/ALL)
                           Process in batches of 50 tickers
                                                  │
                         ┌────────────────────────┼─────────────────────┐
                         ▼                        ▼                     ▼
                  [Indicators.bas]         [ATRCalculation.bas]   [Scores.bas / FairValue.bas]
                  RSI, MACD, EMA,          Volatility zones,      Composite 0-100 score,
                  Bollinger, DMI,          stop loss sizing       DCF intrinsic value,
                  Keltner, Volume Profile                         PE/PB ratings 1-10
                         │                        │                     │
                         └────────────────────────┼─────────────────────┘
                                                  │
                           [CompleteTrading.bas]
                           BUY/SELL/HOLD signal + strength (WEAK/MEDIUM/STRONG)
                           Stop loss = ±2×ATR, Target = ±4×ATR
                                                  │
                                        [REPORTING.bas: theReporter()]
                                        Dashboard update + HTML email (Outlook)
                                                  │
                              [Performance2.bas / TradeAnalysis.bas]
                              Backtesting over date range + equity curve
                                                  │
                                       [UploadToDrive_VBA.bas]
                                       Shell → Python → Google Drive
```

---

## Module Reference

### Core Orchestration

| Module | Purpose | Key Functions |
|--------|---------|---------------|
| `ALL.bas` | Main entry point. Fetches OHLCV via `STOCKHISTORY`, deduplicates into BackupAll, manages backup. | `ProcessAll`, `ProcessAllStocks`, `RetrieveStockHistory`, `BackupALLAndSort`, `OneStock`, `SelectedHistorical`, `StopButton_Click` |
| `Phoenix.bas` | Master workflow coordinator. Chains data loading → indicators → signals for weekly strategy. | `TradeMaster`, `MasterDataFromBackup`, `MasterWeeklyTradingStrategyAndSignals` |
| `Scheduling.bas` | Automated daily scheduling at 3:00 AM. Full US federal holiday calendar. | `ScheduledRun`, `RunScheduledTask`, `GetNextRunTime`, `IsHoliday`, `IsWorkday`, `GetNextWorkday`, `LogMessage` |

### Data Processing & Filtering

| Module | Purpose | Key Functions |
|--------|---------|---------------|
| `Filtering.bas` | Core filter engine. Batches 50 tickers, filters by price + score + Origin (USA/Canada/ETF/INTL/ALL), calls indicator scoring. **Primary active module.** | `FilterAndReport`, `ProcessTickersUltraFast`, `DataFromBackup`, `GroupbyPrice`, `DeleteNARows`, `GetUserInputs`, `SortArray` |
| `ValidateData.bas` | Multi-layered anomaly detection: percentage-change thresholds (100% price / 300% volume), 3-sigma statistical outliers with 10-day rolling window. Logs to AnomaliesList sheet. | `DetectAnomalies`, `DetectStatisticalAnomalies`, `ClearAnomalousCells`, `DeleteAnomalousRows`, `CalcPnL` |
| `UTILITIES.bas` | Miscellaneous helpers: data loading, filter toggles, speed testing, dropdown validation setup. | `UpdateDashFromReport`, `ModDataFromBackup`, `AddDataValidation`, `toggleTJXFilter`, `toggleDashFilter`, `speedTest` |
| `UtilityFiles.bas` | Extended file and sheet utility operations. | File management helpers |
| `EnhancedDS.bas` | Enhanced data structures and helper types. | Type definitions, helper functions |

### Technical Indicators

| Module | Purpose | Key Functions |
|--------|---------|---------------|
| `Indicators.bas` | Calculates 14+ indicators on cleaned per-ticker data. Runtime ~20 min for full run. | `CalculateIndicators`, `CalculateEMA`, `CalculateRSI`, `CalculateMACD`, `CalculateBollingerBands`, `Calculate_SMI`, `Calculate_Bressert_DSS`, `CalculateATR`, `CalculateElderRay`, `CalculateVolumeProfile`, `CalculateKeltnerChannels`, `toCalculateDMI` |
| `ATRCalculation.bas` | ATR volatility zones (LOW/NORMAL/HIGH/EXTREME), risk management, position sizing. | `CalculateATRWithSignals`, `CalculateTickerATR`, `GenerateATRZoneAndSignal`, `CalculateRiskManagement`, `GenerateATRSignals`, `OutputATRSignals` |
| `CompleteATRSystem.bas` | Standalone comprehensive ATR pipeline (self-contained). | Full ATR pipeline |
| `CompleteIndicatorCalculations.bas` | Comprehensive indicator calculation suite. | Full indicator set |

### Signal Generation

| Module | Purpose | Key Functions |
|--------|---------|---------------|
| `CompleteTrading.bas` | **Primary active signal engine.** Multi-factor qualification, indicator weighting (RSI 1.6×, MACD 1.4×, Volume 1.3×, ATR 1.1×), regime-adjusted thresholds. Outputs 17-column signal table. | `FilterAndReport_Enhanced`, `GenerateCompleteTradingSignals_Final`, `GenerateTradingSignalsWithRiskManagement`, `ProcessTickersUltraFast_WithCollection`, `GetTickerSignalData`, `CalculateBuyScoreSimple`, `CalculateSellScoreSimple`, `GenerateTradingSummaryArray` |
| `CompleteSignalsGeneration.bas` | Complete signal generation suite (setup + process + output + formatting). | Full signal pipeline |
| `WeeklySignals.bas` | Ultra-fast weekly signal generation using array pre-fetching. Outputs confidence scores 1–5. | `GenerateWeeklyTradingSignals_UltraFast`, `CalculateSignalScore_Fast`, `CalculateSignalConfidence_Array`, `CreateAndOutputSignalsSheet` |
| `testWeeklySignals.bas` | Test harness for weekly signal functions. | Testing versions |

### Scoring & Valuation

| Module | Purpose | Key Functions |
|--------|---------|---------------|
| `Scores.bas` | Composite 0–100 stock score. Weights: Price 10%, EPS 10%, PE 15%, Growth 20%, Fair Value 15%, ROE 15%, Debt 10%, FCF 5%. Sector-adjusted percentile ranking. | `CalculateStockScore`, `CalculateSectorPercentile`, `CalculateAllScores`, `GetSectorStats`, `GetTopPerformers`, `UpdateTJXFromTickerFile` |
| `oneScores.bas` | Consolidated single-file scoring variant. | Scoring pipeline |
| `FairValue.bas` | DCF intrinsic value calculation + industry-specific PE/PB/D/E ratings (1–10). Fetches live data via Alpha Vantage and Yahoo Finance APIs. | `CalculateTJXScore`, `CalculateFairValueRating`, `CalculateRating`, `RetrieveAlphaVantageData`, `RetrieveYahooFinanceData`, `ExtractValueFromJson`, `RunFullAnalysis`, `CreateRatingsDashboard` |

### Reporting & Output

| Module | Purpose | Key Functions |
|--------|---------|---------------|
| `REPORTING.bas` | Transforms filtered data into formatted reports. Sends HTML email via Outlook. Creates TradingView and Yahoo Finance hyperlinks. | `theReporter`, `OptimizedFilterAndCount`, `tidyReportsOptimized`, `ReportToDashOptimized`, `CreateTVHyperLinks`, `CreateHyperLinks`, `sendReport`, `LogReports`, `HistoryToDash` |
| `TradeAnalysis.bas` | Trade performance analytics. Calculates KPIs, equity curve, max drawdown, expectancy. Creates charts. | `OneStepPerformance`, `AnalyzeTrades`, `CalculateAllMetrics`, `UpdateEquityCurve`, `CalcMaxDrawdown`, `CalcExpectancy`, `CreateAllCharts`, `UpdateMarketRegimeAnalysis` |
| `Performance2.bas` | Backtesting orchestrator. Iterates `FilterAndReport` over weekly/daily date ranges. 12-month weekly or 90-day daily. | `TISystemPerformance`, `PerformanceHISTORY`, `SetupTradeLog`, `SetupPerformance`, `GetNextMonday`, `GetNextWorkday`, `ExtractAndIncrementGroupNumber` |
| `RecordResults.bas` | Records test group metrics, increments group counter (Test_Group_1, _2, …). | `RecordTestResults` |

### Currency & Exchange Rates

| Module | Purpose | Key Functions |
|--------|---------|---------------|
| `XchgeRates.bas` | Exchange rates via Bank of Canada API with JSON parsing. Creates ConversionRates sheet. | `UpdateExchangeRates`, `ConvertStockPrices` |
| `RateConversion.bas` | Converts TSX/TSXV/international prices to USD. Handles multiple exchange codes. | `UpdateExchangeRates`, `ConvertStockPrices`, `ConvertStockPricesToUSD`, `GetExchangeRate`, `ExtractExchange` |

### Infrastructure & Utilities

| Module | Purpose | Key Functions |
|--------|---------|---------------|
| `JsonConverter.bas` | VBA-JSON v2.3.1 by Tim Hall. Full JSON parse/serialize. Used by API modules. | `ParseJson`, `ConvertToJson` |
| `UploadToDrive_VBA.bas` | Fire-and-forget shell to Python script for Google Drive backup. | `UploadToDrive` |
| `CompleteCoreProcessing.bas` | Full core processing suite. | Core pipeline |
| `CompleteEssentialHelpers.bas` | Essential shared helper functions. Provides `GetMarketRegime`, `HasVolumeConfirmation`, `IsFalsePositive` used by multiple modules. | Shared helpers |
| `CWE.bas` | "Complete Wealth Enjoyment" — portfolio-level scoring/tracking. | Portfolio functions |

---

## Key Worksheets in TRADE.xlsm

| Sheet | Role |
|-------|------|
| **DashBoard** | Central control panel — date range, price min/max, score thresholds, Origin filter, market regime setting, top-50 results display |
| **TJX** | Master ticker list with sector, price, EPS, PE, growth rate, fair value, ROE, debt/equity, FCF margin, composite score |
| **BackupAll** | Consolidated historical OHLCV data — sorted by date/ticker, deduplicated, primary history store |
| **DataHistory** | Raw STOCKHISTORY formula output before backup processing |
| **Data** | Working dataset for current analysis run (50 most recent rows per ticker) |
| **Reports** | Filtered stock candidates for current report |
| **ReportHistory** | Archive of report outputs |
| **ReportLog** | Appended running log of all report outputs |
| **TRADE LOG** | Manual trade journal — setup, conviction, market regime, outcome, P&L |
| **TradingSignals** | Generated BUY/SELL signals with entry, stop loss, target, position size, R/R ratio |
| **PERFORMANCE** | Backtesting results per iteration (group number, win rate, P&L, best setup) |
| **Trades** | Trade log data transferred for analysis |
| **Metrics** | Calculated performance KPIs (win rate, expectancy, max drawdown, Sharpe) |
| **Charts** | Equity curve, setup performance, market regime, conviction analysis charts |
| **ATR Signals** | Volatility zone classification per ticker (LOW/NORMAL/HIGH/EXTREME) |
| **AnomaliesList** | Flagged data anomalies with date, ticker, type, deviation details |
| **Indicators** | Calculated indicator values per ticker (RSI, MACD, ATR, etc.) |
| **WeeklySignals** | Weekly trading signal output |
| **ConversionRates** | Live exchange rate lookup table (CAD/USD and others) |
| **SchedulerLog** | Scheduled task activity log |
| **cweSignals** | CWE (Wealth Enjoyment) signal output |

---

## Risk Management Parameters

| Parameter | Value |
|-----------|-------|
| Stop Loss (buy) | Entry − 2×ATR |
| Stop Loss (sell/short) | Entry + 2×ATR |
| Price Target (buy) | Entry + 4×ATR |
| Price Target (sell) | Entry − 4×ATR |
| Minimum R/R Ratio | 2:1 |
| Position size — low vol (ATR% < 2%) | 8% of portfolio |
| Position size — normal vol (ATR% 2–3%) | 6% of portfolio |
| Position size — high vol (ATR% 3–5%) | 4% of portfolio |
| Position size — extreme vol (ATR% > 5%) | 2% of portfolio |
| Max portfolio risk per trade | 1–2% |

### Indicator Weights (CompleteTrading.bas)

| Indicator | Weight |
|-----------|--------|
| RSI | 1.6× |
| MACD | 1.4× |
| Volume | 1.3× |
| ATR | 1.1× |
| Price Action | 1.0× |
| OBV | 0.3× |
| ADX | 0.6× |
| Stochastic / Williams %R / CCI | 0× (disabled) |

### Signal Strength Tiers

| Tier | Condition |
|------|-----------|
| STRONG | Score ≥ 2.0× threshold |
| MEDIUM | Score ≥ 1.5× threshold |
| WEAK | Score < 1.5× threshold |

---

## ATR Volatility Zones

| Zone | ATR% Condition | Position Size |
|------|---------------|---------------|
| LOW VOL | ATR% < 1.5% | 8% |
| NORMAL | ATR% 1.5–3% | 6% |
| HIGH | ATR% 3–5% | 4% |
| EXTREME | ATR% > 5% | 2% |

ATR Ratio signals: > 1.5 = VOLATILITY_SPIKE, < 0.7 = VOLATILITY_CONTRACTION

---

## Composite Score Weights (Scores.bas)

| Metric | Default Weight | Direction |
|--------|---------------|-----------|
| Current Price | 10% | Higher percentile = better |
| EPS | 10% | Higher = better |
| PE Ratio | 15% | Lower = better |
| Growth Rate | 20% | Higher = better |
| Fair Value Ratio | 15% | Higher = better |
| ROE | 15% | Higher = better |
| Debt/Equity | 10% | Lower = better |
| FCF Margin | 5% | Higher = better |

Score output: 0–100 (column V in TJX sheet). Color scale: > 80 = forest green, 60–80 = light green, 20–40 = yellow, < 20 = crimson.

---

## File Naming Conventions

| Prefix/Suffix | Meaning |
|---------------|---------|
| `Complete` prefix | Standalone complete implementation (self-contained) |
| `test` prefix | Test harness or experimental version |
| `2` suffix | Revised version (e.g. `Performance2.bas`) |
| `Bare` (in file names) | Minimal stripped-down workbook version |
| `HollowedOut` (in file names) | Workbook with stub modules only |

> **Note:** Modules with `ho*`, `Orig*`, `From*`, `other*`, `AllOne`, and `*fromNewBare`/`*fromLess` prefixes/suffixes were removed from TRADE.xlsm during the Feb 2026 cleanup — they were stub skeletons, legacy versions, or variant imports superseded by the active modules above.

---

## Global Variables (ALL.bas)

| Variable | Type | Purpose |
|----------|------|---------|
| `gStopMacro` | Boolean | Emergency stop flag — halts all macros |
| `pubNotice` | Boolean | Silent mode toggle (suppress popups) |
| `minScore` | Long | Minimum composite score threshold |
| `perfTest` | Boolean | Performance testing mode flag |
| `endDate` | Date | Analysis end date |
| `sigTest` | Boolean | Signal testing mode flag |
| `ExchangeRates` | Scripting.Dictionary | In-memory exchange rate cache |

---

## External Dependencies

| Dependency | Purpose |
|------------|---------|
| Excel `STOCKHISTORY` function | Historical OHLCV data retrieval |
| Microsoft Scripting Runtime | `Scripting.Dictionary` used throughout |
| Microsoft Outlook | HTML email sending (`sendReport`) |
| Alpha Vantage API | PE, P/B, D/E financial data (`FairValue.bas`) |
| Yahoo Finance API | Alternative financial data source (`FairValue.bas`) |
| Bank of Canada API | Live USD/CAD exchange rates (`XchgeRates.bas`, `RateConversion.bas`) |
| Python (`drive_upload.py`) | Google Drive backup upload |
| VBA-JSON v2.3.1 (Tim Hall) | JSON parsing for API responses (`JsonConverter.bas`) |

---

## Backtesting Modes (Performance2.bas)

| Mode | Date Range | Increment | Iterations |
|------|-----------|-----------|------------|
| Weekly | 12 months | 1 week | ~52 |
| Daily | 90 days | 1 workday | ~63 |

Each iteration: calls `FilterAndReport` → `PerformanceHISTORY`. Groups named `Test_Group_1`, `Test_Group_2`, etc.

---

## Python Companion Scripts

The project has a parallel Python ecosystem in the same folder for data fetching, scoring, and analysis outside of Excel.

| Script | Purpose |
|--------|---------|
| `drive_upload.py` | Google Drive upload (called by `UploadToDrive_VBA.bas`); latest version |
| `drive_upload_1/2/3/4.py` | Previous iterations of the Drive upload script |
| `TickerFund.py` | Main Python ticker data fetcher with full fundamental analysis |
| `TickerScore.py` | Composite stock scoring in Python |
| `TickerValue.py` | Valuation/DCF calculations in Python (mirrors FairValue.bas) |
| `TickLimit.py` | Ticker list management |
| `FinancialFetcher.py` | Financial data downloader (Alpha Vantage / Yahoo Finance) |
| `UpdatedFinancialFetcher.txt` | Updated version of FinancialFetcher |
| `download_fundamentals_alpha_vantage.py` | Alpha Vantage fundamental data downloader |
| `download_fundamentals_alpha_vantage_v3.py` | v3 of Alpha Vantage downloader |
| `fetchFin.py` | Financial data fetch utility |
| `fetchstockdata.py` | Stock price data fetcher |
| `APIData.py` | API data retrieval helper |
| `AnalyzeDashTickers.py` | Analysis of dashboard tickers |
| `FullAnalysis.py` / `FirstFullAnalysis.py` | Complete analysis pipeline in Python |
| `CompLoaded.py` / `CompPreLoad.py` | Composite score pre-loading |
| `PerformanceAnalyzer.py` | Python backtesting performance analyzer |
| `IndicatorCalculator.py` | Technical indicators in Python |
| `ScoringModel.py` | ML-based scoring model |
| `RegimeSolution.py` | Market regime detection |
| `wkbacktest.py` | Weekly backtest script |
| `AfricaScore.py` / `CanadaScore.py` | Region-specific scoring |
| `app.py` | Web application entry point (stock screener) |
| `morningschedule.bat` | Windows batch file to trigger morning run |

**Web Screener:** `stock-screener_*.html` / `TI Screener.html` — iteratively developed browser-based stock screener (34 versions), supplementing the Excel dashboard.

---

## Development Notes

- **Emergency stop**: `StopButton_Click()` in ALL.bas sets `gStopMacro = True`; all loops check this flag
- **Batch size**: 50 tickers per processing group (hardcoded in Filtering.bas)
- **Data frequency**: Weekly = same weekday as `endDate`; Daily = all weekdays
- **Deduplication**: BackupAll sorted by date+ticker; duplicates removed after each ProcessAll run
- **Screen updates disabled** during backtesting for performance (re-enabled on completion)
- **MACD note**: Known issue in Indicators.bas line ~217 calculates 26−26 instead of 12−26 EMA diff
- **API rate limiting**: FairValue.bas includes `Wait()` delays between Alpha Vantage calls
- **Origin filter values**: `USA`, `Canada`, `ETF`, `INTL`, `ALL` — set on DashBoard, read by Filtering.bas
- **Active signal module**: `CompleteTrading.bas` is the authoritative signal generator (confirmed correct column mapping). `GenerateSignals.bas` was removed — it had a column mapping bug (read composite score from col 14 = ATR instead of col 9)
- **Data sheet column layout**: Col 1=Date, 2=Open, 3=High, 4=Low, 5=Close, 6=Volume, 7=Ticker; indicators appended by Indicators.bas: col 9=CompositeScore, 10=RSI, 11=MACD, 12=MACDSignal, 13=PriceVsMA, 14=ATR, 15=ATR%
- **Feb 2026 cleanup**: 18 redundant modules removed from TRADE.xlsm (8 hollow stubs, 3 legacy originals, 4 ATR/filtering variants, GenerateSignals, AllOne, Performance)
