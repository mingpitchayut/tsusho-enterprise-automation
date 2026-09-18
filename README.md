# Enterprise Logistics & Trade Automation Dossier
**Engineer:** Pitchayut Boonporn  
**Affiliation:** B.Eng. in Computer Engineering, Mahidol University International College (MUIC)  
**Target:** Technical Certification & Production Portfolio

📄 **[Download Complete Executive PDF](Pitchayut_Boonporn_Engineering_Dossier.pdf)**

---

### Impact & Scale at a Glance

| Metric | Details |
| :--- | :--- |
| **Production Pipelines** | 20 Dedicated Automated Systems |
| **Codebase Scope** | 100+ Custom Python Production Modules |
| **Data Throughput** | 600+ Enterprise Documents (PDF, CSV, Excel) |
| **Data Privacy** | 100% Offline Air-Gapped Local LLM Inference |
| **Primary Domains** | Ocean Freight, Customs & Tax, FTA Compliance, Industrial Supply Chain |

---

### Systems Breakdown

#### 1. Multi-Carrier Logistics & Ocean Freight
* **BL-ONE-SITC (Ocean B/L Draft & Invoice Reconciler):** Automates data extraction from Ocean Network Express (ONE) and SITC Draft B/Ls and Tsusho multi-page commercial invoices. Solves address continuation overflows (`SH>` / `CN>`), container/seal pairing, and cargo weight/CBM cross-validation.  
  *Stack: Python, PyMuPDF (fitz), openpyxl, regex coordinate clipping*
* **CARGO NO (HS Code Extractor 1.0–1.2):** Extracts 8-to-10 digit Customs HS codes and part numbers from complex Excel import instruction sheets (`.xls` / `.xlsx`).  
  *Stack: Python, pandas, regex, PyInstaller*
* **IMPORT CARGO (Cross-Referencing Tool):** Reconciles and merges Excel Import Instructions with Chemiplas commercial invoice PDFs, extracting cargo codes, tariff rates, and CSM approval numbers.  
  *Stack: Python, pdfplumber, pandas, win32com.client, openpyxl*
* **VN (TTAST Steel Export PDF Extractor 1.0–1.3):** Extracts steel coil export specifications, JIS standards, Tax IDs, and dimensional data from TT Automotive Steel (Thailand) shipment PDFs into formatted Excel sheets.  
  *Stack: Python, pdfminer.six, openpyxl, pathlib*
* **VN-IPVL (Invoice & Packing List Dimension Extractor):** Parses steel manufacturing documents to isolate mill/slit edges, industrial grades (SPFH590, TIS 1884-2564), and thickness/width metrics.  
  *Stack: Python, pdfminer.six, openpyxl, regex tokenizer*

#### 2. Customs, Financial & Tax Extraction Engines
* **pdfToTax (Customs PDF to Tax Declaration 1.0–1.10):** High-volume pipeline processing 355+ Thai Customs declaration PDFs (`QDRS...`), extracting declaration numbers, CIF values, duty rates, and VAT breakdowns into audit-ready tax sheets.  
  *Stack: Python, pdfminer.six, openpyxl, regex, PyInstaller*
* **CHEMIPLAS INVOICE (Toyotsu Chemiplas Parser):** Batch converts multi-page commercial invoices and packing sheets into dual-sheet Excel workbooks with automated cell-highlight alerts for missing entries.  
  *Stack: Python, PyMuPDF (fitz), openpyxl (PatternFill), regex, AST*
* **Chemiplas_N (Composite PDF Splitter):** Splits composite master PDFs into distinct Invoices and Packing Lists using visual boundary coordinates (`coord_finder.py`, `redline_xAxis.py`).  
  *Stack: Python, pdfminer.six, PIL, PyMuPDF, openpyxl*
* **CUSTOM INV (Customs Declaration & Tax Bill Parser):** Automates data extraction from Thai Customs Department declaration entries (ใบขนสินค้า / ใบเสร็จรับเงิน), Tax IDs, VAT, customs fees, and payment dates.  
  *Stack: Python, PyMuPDF (fitz), pandas, openpyxl, Thai Unicode tokenizers*
* **CSV INVOICE PACKING (Packing List Aggregator 1.0–1.3):** Batch parses and consolidates 100+ export packing list CSV files (`EXPORTPAC_TCOM...`) into standardized master sheets.  
  *Stack: Python, pandas, xlsxwriter, zipfile, PyInstaller*
* **INVOICE EXTRACTOR (InvPackingList Extractor 1.0–4.1):** Multi-version extraction engine featuring dynamic layout detection and automated error logging.  
  *Stack: Python, PyMuPDF (fitz), pandas, openpyxl, tqdm*
* **OT Extractor (Customs Overtime & Receipt Fee Extractor):** Scans Thai Customs electronic receipts to extract officer overtime charges, clearance fees, and receipt numbers into summary cost reports.  
  *Stack: Python, pdfminer.six, openpyxl, argparse*

#### 3. Free Trade Agreements (FTA) & Regulatory Compliance
* **EPA JP (JTEPA Agreement Form JP Extractor 1.0–2.1):** Parses 15+ page Japan-Thailand Economic Partnership Agreement certificates, capturing product tables, importer/exporter details, and non-numerized part numbers.  
  *Stack: Python, PyMuPDF (fitz), openpyxl, regex*
* **EPA-ATG KH (Cambodia E-Form D ATIGA Extractor):** Maps HS codes and invoice values from Cambodia-Thailand ASEAN Trade in Goods Agreement certificates (NHK Spring / Toyota Tsusho).  
  *Stack: Python, PyMuPDF (fitz), openpyxl*
* **CO-JTCH (Certificate of Origin Item Extractor):** Extracts structured item entries, automotive steering components, clamp/spring codes, quantities, and HS codes from EPA Certificate of Origin PDFs.  
  *Stack: Python, PyMuPDF (fitz), pandas, openpyxl*

#### 4. Applied AI, Vision & OCR Pipelines
* **Ollama_Model (Local LLM Document Intelligence):** Employs local, air-gapped LLMs (`ministral3`, `ollama`) for zero-shot schema extraction, invoice header classification, and unstructured text normalization with 100% offline enterprise privacy.  
  *Stack: Python, ollama SDK, PyPDF2, PyMuPDF, rich, ThreadPoolExecutor*
* **OCR-Thing (Optical Character Recognition Engine):** Extracts tabular billing data from non-searchable, scanned paper invoices (e.g., Qingdao Special Iron & Steel) using computer vision filtering.  
  *Stack: Python, EasyOCR, pdf2image, numpy, OpenCV*

#### 5. Industrial Supply Chain & Specialized Manufacturing
* **RETURNABLE RACK (Keyword PDF Filter & Splitter 1.0–1.3):** Bulk filters, isolates, and splits across 70+ Delivery Order PDFs for returnable container logistics fleets.  
  *Stack: Python, PyPDF2, PyMuPDF (fitz), glob, PyInstaller*
* **JTEKT (Automotive Parts Manifest Extractor):** Automates classification and extraction of steering parts, packaging units, and line items tagged with custom product markers (`U` / `K`).  
  *Stack: Python, PyMuPDF (fitz), pdf2image, pandas, regex, PyInstaller*
* **RUBBER (Chemical Certificate & Rubber Batch Extractor 1.0–1.1):** Parses Thai chemical certificate forms (หนังสือรับรองยางผสมสารเคมี E-QC), batch certificates (`ECER...`), compound formulas, test dates, and percentage compositions.  
  *Stack: Python, PyMuPDF (fitz), pandas, regex*

---

### Core Engineering Competencies
`Python 3.x` • `PyMuPDF (fitz)` • `pdfminer.six` • `pandas` • `openpyxl` • `EasyOCR` • `OpenCV` • `Ollama` • `Local LLMs` • `PyInstaller` • `Regex Spatial Extraction`
