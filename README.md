# Danny Tran

**Project Controls Analyst** at Primoris T&D. I turn manual job cost accounting work into checked, repeatable automation in Python, PowerShell and Power BI, and I build my own tools on the side.

[github.com/space-cowboy2](https://github.com/space-cowboy2)

---

## Highlights

- **Built a fuel card fraud-detection system** that reviews 73,015 transactions, groups them into 65,815 fueling stops and flags 243 for review. Then audited my own first version, which had flagged 3,605 rows, and threw out the false positives.
- **Rescued and automated a regional revenue dashboard.** Found that a truncated source file (322K of 957K rows) was silently breaking it, rebuilt the 950K-row dataset, and replaced the manual daily refresh with a job that runs every morning.
- **Automated journal entry prep** for cost corrections and accruals, from shared-mailbox request to JE import file, with approval required before anything posts.
- **Validated payroll reporting to zero differences:** 261,514 rows across 24 weeks matched 17 source workbooks.
- **Wrote an AI-agent runbook for month-end job close** and used it for the September and October close.

---

## Work projects

Built at Primoris T&D, Southwest Distribution, in 2026. The code and data belong to the company, so these are write-ups, not repositories. Every figure below is a count; no company dollar amounts are published.

### Fuel Card Anomaly Detection
`Python` `pandas` `openpyxl` `Outlook COM` `EIA price data` · Nov 2026

The fleet's fuel card exports have no odometer and no product code, and card numbers are masked. I built a system that checks every fueling stop against what each truck can physically hold.

- Groups swipes on one vehicle within 30 minutes in the same city into a single stop, then runs 15 weighted checks: over the tank, refills that don't add up given burn rate, wrong fuel type, identical rows, quick refills, volume outliers and context flags.
- Works out gasoline vs diesel from pump price clusters per state and week, since the export has no product code.
- Builds an audit-ready Excel report: executive summary, department breakdown by director, weekly trends, money-at-risk estimates, case profiles, and a "How Each Number" tab where 34 headline figures each recompute from the detail rows with a live formula.
- One-click refresh: collects new exports, rejects non-exports, skips duplicate files, and counts overlapping weeks once.
- **Self-audit:** the first version flagged 3,605 rows. An independent review against the raw data showed masked cards were merging, pump-limit splits were counted as refills, and tank sizes were learned from bad data. The reviewed version flags 243 stops, and a separately written program recounted it to the same result.

### Billed/Unbilled Revenue Pipeline
`PowerShell` `Outlook COM` `Python` `Task Scheduler` `Power BI` · Oct–Nov 2026

- Diagnosed a Power BI refresh failure on the regional billed/unbilled aging dashboard: the source workbook had been truncated to a third of its rows.
- Rebuilt the 950K-row dataset from the six original exports and verified it before replacing the corrupt file.
- Automated the daily refresh. A scheduled job pulls the six exports from a shared mailbox, checks each file (zip integrity, expected columns, SHA-256), swaps them in through staging with retry and rollback, and emails an alert only if the final run fails.
- Built in self-recovery: Outlook connection check and restart, wake from sleep, a catch-up run after reboots, and a warning when a weekday brings no new reports.

### JCA Cost Correction & Accrual Processor
`PowerShell` `Excel/Outlook COM` `Power Query` `VBA` · Nov 2026

- Turns job cost adjustment requests from a shared mailbox into journal entry import files: queues each request, routes it by sheet content (cost move vs accrual), refreshes the processor's queries, runs its macros and files the output.
- Catches resubmitted files by content fingerprint (SHA-256), keeps a hold/release queue for requests waiting on corrections, and refuses to run if the drop folder holds the wrong file.
- Prints every blocking issue and anything worth a note before processing, and nothing posts without my approval.

### Power BI Reporting
`Power BI` `DAX` `Power Query` `PBIP/TMDL` `Python` · Sep–Nov 2026

- **Payroll Metrics:** moved the report to the Power BI project format so changes can be scripted and versioned. Found 10 defects in the model, including a row-count gap and a missing division. Added a job-hours-by-pay-period page. Validated 261,514 rows over 24 weeks against 17 source workbooks with zero differences.
- **Southwest Distribution dashboards:** moved open commitments to a new PO/subcontract export (10,550 lines, 38 duplicates removed) and added invoice-level detail to the job cost pages.

### Month-End Close
`Viewpoint Vista` `AI agents` `Excel` · Sep–Nov 2026

- **Job close:** wrote a runbook (v1.5) that has an AI agent run the Viewpoint reports and prepare PO and subcontract closure requests, a review workbook and a contract status PDF per job, under strict evidence rules ("missing is not zero"). People keep control of every close action and email. Used for the September soft close and the October close, where it identified five jobs needing follow-up.
- **Monthly project reports (MPRR):** rebuilt all 20 October reports from job cost detail and published them with backups and hash-checked copies.
- **JE cost tracker:** kept the South Region master log current through period 9, across about 45 batches.
- **Rate uploads:** built rate upload templates from customer rate sheets.

### Payroll Reporting Database Pilot
`PostgreSQL` `Python` `pgAdmin` `Power BI` · Nov 2026

- Planned moving Power BI source data off Excel and Access into PostgreSQL with history.
- Built an import/export validation test, network and remote-access checks, and a stakeholder demo package with one-click launchers, an operating guide and a presentation.

---

## Personal projects

| Project | What it is | Stack |
|---|---|---|
| [**Duet Gateway**](https://github.com/space-cowboy2/duet-gateway) | Remote MCP gateway that lets Claude reach tools on a home PC or Mac: TV control, read-only bank data, desktop control, a real browser and sandboxed Claude Code sessions. OAuth with hashed tokens, a kill switch, approval before clicks or typing, and a monitor that files issues from its own logs. | Node.js, GitHub Actions |
| [**Second Wind**](https://github.com/space-cowboy2/second-wind) | macOS menu bar app that keeps two Claude accounts signed in with one shared session list and switches between them. 42 unit tests and written design decisions. | Swift |
| [**Finance Tracker**](https://github.com/space-cowboy2/finance-tracker-public) | Personal cash-flow tracker: bank transactions through Plaid into SQLite, a 30-day forecast that flags shortfalls, a purchase check, and a review queue. 88 tests; hardened against DNS rebinding. | Python, Flask, SQLite, Plaid |
| [**TV Remote**](https://github.com/space-cowboy2/tv-remote) | Control a Samsung TV from any device, at home or away, by buttons or by conversation. No cloud and no open ports: devices reach the home hub over Tailscale. | Python, JavaScript, Swift |
| [**shellmon**](https://github.com/space-cowboy2/shellmon) | Terminal pet with a full-screen care game and a Claude-aware status line. | Python, Textual |
| [**LidAwake**](https://github.com/space-cowboy2/lidawake) | macOS menu bar toggle that keeps a closed MacBook awake for SSH. | Swift |
| Garmin Live *(private)* | Garmin stress, HRV, sleep and heart-rate data in local SQLite, served to agents over MCP, plus an iPhone app for live Bluetooth heart rate. | Python, SQLite, iOS |
| Quant Microstructure *(private)* | Limit order book research: matching engine, order-flow simulation, market making. Python first, then C++. | Python, C++ |

---

## Brag list

| When | What | Result |
|---|---|---|
| Aug 2026 | Started Finance Tracker | Plaid sync into SQLite; grew into a 30-day cash-flow forecast |
| Sep 2026 | Payroll Metrics moved to Power BI project format | Report changes can be scripted and versioned |
| Sep 2026 | September soft close | Contract status reviewed and audit list produced for the close batch |
| Sep–Oct 2026 | JE cost tracker | South Region master log current through period 9 |
| Oct 2026 | Rate upload templates | Customer rate sheets converted to upload-ready templates |
| Oct 2026 | Revenue dashboard rescue | Found the truncated source; rebuilt the 950K-row dataset |
| Oct 2026 | Daily billed/unbilled automation | Manual refresh replaced by a scheduled job with alerts and self-recovery |
| Oct 2026 | Payroll Metrics validation | 261,514 rows, 24 weeks, zero differences against 17 sources; 10 defects found |
| Oct 2026 | Intercompany job setup review | Found a wrong receivable type and overtime setting on an intercompany job in Viewpoint |
| Oct 2026 | Job close runbook v1.5 and October close | Closure requests and contract PDFs prepared; five jobs flagged for follow-up |
| Nov 2026 | MPRR | All 20 October project reports rebuilt and published |
| Nov 2026 | PostgreSQL pilot | Validation test and stakeholder demo package |
| Nov 2026 | JCA processor | Cost corrections and accruals scripted from mailbox to JE import file |
| Nov 2026 | Duet Gateway desktop control | 44 tools shipped behind a safety guard; CI green on Windows and Linux |
| Nov 2026 | Second Wind | Shipped a native macOS app in Swift with 42 unit tests |
| Nov 2026 | Fuel card anomaly detection | 73,015 transactions reviewed; 243 high-priority stops flagged |
| Nov 2026 | Fuel report self-audit | 3,605 flagged rows narrowed to 243 stops that hold up; independent recount matched |
| Nov 2026 | Fuel executive report | Department and director roll-up for leadership |
| Nov 2026 | Southwest Distribution dashboards | New PO/subcontract source, 38 duplicate lines removed, invoice detail added |

---

## Skills

**Data and reporting:** Power BI (DAX, Power Query, PBIP/TMDL), Excel (Power Query, VBA, live-formula tie-outs), SQL, PostgreSQL, SQLite
**Automation:** Python (pandas, openpyxl), PowerShell, Outlook and Excel COM automation, Windows Task Scheduler
**Accounting systems:** Viewpoint Vista, job cost, accruals, journal entries, month-end close
**Engineering:** Git and GitHub Actions, unit testing, Node.js, Swift, MCP servers, AI coding agents (Claude Code, Codex)
