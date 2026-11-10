# Sandhya Enterprises — MES Tender Matchmaker

## What is configured

This project includes the MES search profile supplied for Sandhya Enterprises:

- Department: Military Engineer Services (MES / E-IN-C Branch).
- Exact actual-worksite PIN allowlist: `411025`, `400029`, `400078`, `400095`, `400094`, `400067`.
- Estimated tender value: INR 2,500,000 to INR 5,000,000 inclusive (₹25 lakh to ₹50 lakh).
- Search terms: provision works; civil/electrical provision; SITC; supply, installation, testing and commissioning; electrical equipment/panels; LT/HT panels; transformers; DG sets; pumps/motors; lighting; furniture; plumbing; water supply; solar; fire alarm; CCTV; HVAC; prefabricated and turnkey supply/installation, plus related aliases.
- Manpower preference: small specialist installation team, prefabricated/modular and supply-led scopes. Labour-intensive scope is flagged for review, not silently treated as a legal disqualification.

## Why there are two JSON files

- `tender_profile.json` is the business filter: department, keywords, worksite PINs, value range, manpower and bid gates.
- `sources.json` is the technical source connector list: public automated sources and CAPTCHA/manual sources.

The JSON the operator supplied mixes both concerns. It has been normalized into these two files so a source type cannot accidentally be mistaken for a parser type. Do not paste the full business profile into the `TENDER_KEYWORDS` environment variable.

## Portal access boundary and current source behavior

The automated discovery sources are the public Defence eProcurement homepage (`https://defproc.gov.in/nicgep/app`) and the public CPPP homepage (`https://eprocure.gov.in/eprocure/app`). Both use the NIC eProcurement landing-page format and display a limited Latest Tenders list and a separate Latest Corrigendums list. The `nicgep_home` parser reads the Latest Tenders table only. It does not claim to enumerate every active MES tender, and it does not infer actual worksite PIN, estimated value, or tender-specific organisation eligibility from a title.

The full Defence and CPPP Active Tenders/Advanced Search flows ask the operator to enter a CAPTCHA. The Advanced Search page exposes Organisation, Pincode and Value Criteria fields, but the agent does not bypass CAPTCHA. Use those search pages manually, complete CAPTCHA yourself, inspect the tender and all attachments/corrigenda, then import a completed CSV.

## Bid decision rules

- `BID`: only if the record is explicitly verified as an MES tender, the actual worksite PIN is one of the six allowed PINs, the value is within ₹25–50 lakh inclusive, the deadline is still open, and eligibility, EMD/fee, NIT/BOQ/tender pack, corrigendum, profitability, materials, manpower plan and completion risk are all evidenced as satisfactory.
- `CONDITIONAL_BID`: one or more critical fields are missing or not verified. This is a review queue, not a qualified match.
- `NO_BID`: a confirmed hard mismatch, such as a different known worksite PIN, out-of-range known value, non-MES issuing organisation, expired deadline, or no configured category keyword.

**PIN safeguard:** only `actual_worksite_pin` or `actual_worksite_address` is used. A PIN in the issuing-office/location field is never accepted as the worksite PIN. If the worksite PIN is not known, the record remains conditional.

## Windows setup

1. Extract the ZIP into `C:\SandhyaTenderAgent`.
2. Open Command Prompt in that folder.
3. Run `py -3.12 -m venv .venv`.
4. Run `.venv\Scripts\activate`.
5. Run `python -m pip install -r requirements.txt`.
6. Copy `.env.example` to `.env` and fill only your Telegram and Google Apps Script credentials.
7. Keep `tender_profile.json` and `sources.json` in the project root. They are already included.
8. Run `python tender_agent.py --mode self-test`.
9. Run `python -m unittest discover -s tests -v` (19 offline tests).
10. Run `python tender_agent.py --mode check-sheets` after the Apps Script bridge is deployed.
11. Run `python tender_agent.py --mode monitor` for the public Defence homepage discovery pass.

## Manual tender import

1. Use the official portal. For Active Tenders or Advanced Search, enter CAPTCHA yourself.
2. Search the organisation `E-IN-C BRANCH - MILITARY ENGINEER SERVICES`.
3. Apply the target PIN and value criteria on the portal when those controls are available; do not treat the issuing office PIN as worksite PIN.
4. Open the tender and latest corrigenda. Verify the worksite address/PIN in the actual work item/NIT/BOQ, estimated value, deadline, EMD, fee, completion period, enlistment and technical eligibility.
5. Copy `import_template.csv` to `imports\my_tenders.csv`.
6. Fill the tender-specific fields. `actual_worksite_pin` must be the actual worksite PIN. Populate `estimated_value_inr` as a plain number (e.g. `3200000`).
7. Record evidence for eligibility, EMD, fee, NIT, BOQ, full tender pack, corrigendum, profitability, materials, manpower and completion risk. Never enter `YES` without checking the supporting record.
8. Run `python tender_agent.py --mode import-csv --file imports\my_tenders.csv`.
9. Review Google Sheets and Telegram. Only `BID` is a fully passed system gate; `CONDITIONAL_BID` needs human review; `NO_BID` is suppressed from Telegram.

## Useful commands

```powershell
python tender_agent.py --mode self-test
python -m unittest discover -s tests -v
python tender_agent.py --mode check-sheets
python tender_agent.py --mode monitor
python tender_agent.py --mode import-csv --file imports\my_tenders.csv
python tender_agent.py --mode compliance --file tender.pdf
```

## Manual source URLs

- Defence eProcurement System: https://defproc.gov.in/nicgep/app
- Defence eProcurement Advanced Search: https://defproc.gov.in/nicgep/app?page=FrontEndAdvancedSearch&service=page
- Defence eProcurement Active Tenders: https://defproc.gov.in/nicgep/app?page=FrontEndLatestActiveTenders&service=page
- CPPP eProcure: https://eprocure.gov.in/eprocure/app
- CPPP Advanced Search: https://eprocure.gov.in/eprocure/app?page=FrontEndAdvancedSearch&service=page

## Limitations and operator responsibility

- Scheduled GitHub Actions checks are periodic, not second-by-second real-time; hosted scheduling can be delayed.
- Public homepage listings are discovery-only and limited; do not treat them as complete active tender coverage.
- Tender field extraction and keyword filters are heuristics. Validate each tender's original NIT, BOQ, GCC/SCC, corrigenda and annexures.
- OCR is not included. Image-only/scanned tender PDFs require a separate OCR step and visual page checking.
- The `BID` gate is a screening aid, not an authorization to submit a bid. The contractor must approve the final bid and confirm current MES enlistment, legal/tax position, price, technical compliance and deadline.
- Do not commit `.env`, bot tokens, shared keys, private bid documents or confidential company records to GitHub.
