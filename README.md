# Email Extractor

### Turning Unorganised Email Evidence into an Organised Investigation Dataset

**Email Extractor** is an Excel VBA automation tool designed to process `.eml` files in bulk and convert an unorganised collection of emails into a structured Excel investigation index.

It addresses a simple but time-consuming problem:

> **Instead of opening emails one by one, identifying relevant emails and copying them into separate folders, investigators can review the entire email population through a structured Excel dataset.**

The original `.eml` files remain in their existing location, while the Excel workbook provides a structured review layer with direct links back to the original evidence.

---

## 🔐 Local Processing

The tool is designed for local processing using Microsoft Excel and VBA.

Email files and extracted information remain within the user's local environment during processing. No cloud-based email-processing service is required by the tool.

Users remain responsible for ensuring that they have appropriate authority to process the underlying data.

---

## 🔄 Before vs After

### Traditional Review

```text
Open Email
    ↓
Read
    ↓
Identify Relevant Email
    ↓
Copy / Move to Investigation Folder
    ↓
Repeat
```

### With Email Extractor

```text
Bulk Process EML Files
        ↓
Structured Excel Index
        ↓
Find / Filter / Sort
        ↓
Mark Relevant Rows
        ↓
Open Original Email When Required
```

**The focus shifts from managing email files to reviewing the email population.**

---

## 🔍 Search & Review

Once emails are converted into Excel rows, investigators can use standard Excel functionality to review the population.

### Find

Search email information for specific names, terms, phrases or identifiers.

### Filter & Sort

Filter or sort the population by fields such as:

- Date
- Sender
- Recipient
- Subject
- Attachment
- Status
- Email content

Relevant rows can then be **highlighted, coloured or categorised directly in Excel**.

The tool does not determine whether an email is relevant. It provides a structured population that allows the investigator to apply their own judgement.

---

## 📧 Email Information

The tool can extract:

- No.
- Date
- From
- From Email
- To
- CC
- BCC
- Subject
- Body
- Attachment(s)
- Open Attachment
- Open Email
- Status
- Error

Each email is represented as a row in the **Email Index**.

---

## 🔒 Evidence & Data Integrity

The tool is built around a simple principle:

> **Don't reorganise the evidence. Organise the review.**

The original `.eml` files are not moved or modified.

The Excel workbook acts as an **investigation index**, with each email represented as a row and links back to the original email.

This allows investigators to:

- Review the complete email population in one place
- Search and filter the extracted information
- Colour or mark relevant rows
- Open the original email directly when required
- Review extracted attachments without creating duplicate copies of the original evidence

The original files remain the underlying source evidence.

---

## 📎 Attachments

Attachments are extracted and organised separately.

```text
Attachments/
│
├── Email_00001/
│   ├── Invoice.pdf
│   └── Payment.xlsx
│
├── Email_00002/
│   └── Contract.pdf
│
└── Email_00003/
    └── Statement.xlsx
```

The Excel index provides links to the extracted attachments.

---

## 🕵️ Potential Applications

The tool can support email organisation and review in:

- Forensic audit
- Insolvency & restructuring investigations
- Internal investigations
- Due diligence
- Audit support
- Email evidence review

It is an **investigation-support tool**, not a system that independently determines fraud, misconduct or legal conclusions.

---

## 📊 Demonstration

This repository contains a **demonstration Excel workbook using 50 sample EML records**.

The demonstration is provided to showcase the workflow and output of the tool.

It allows the reviewer to see how the extracted email population can be:

**processed → structured → searched → filtered → reviewed → linked back to the original email.**

The full working tool and VBA implementation are **not distributed in this repository**.

The demonstration data is fictional and contains no real client or confidential information.

---

## ⚙️ Technology

- Microsoft Excel
- VBA
- Excel UserForms
- FileSystemObject
- ADODB.Stream
- MIME parsing
- Regular expressions

The tool processes `.eml` files directly and does not require Microsoft Outlook.

---

## 📦 Output

The generated workbook includes:

**Email Index**

Structured email information with links to original emails and attachments.

**Dashboard**

Overview of processing activity, including total emails, processed/unprocessed files, emails with attachments and processing status.

---

## ⚠️ Processing Exceptions

If an individual EML file cannot be processed, the tool records the exception rather than stopping the entire population.

The output can record the processing status, error number, error description and original file location.

---

## 📌 Project Status

**Version:** `v1.0.0`  
**Status:** Portfolio / Demonstration Project

---

## 👤 Author

**Dhruvi Bhatiya**

Areas of interest:

- Restructuring & Insolvency
- Forensic & Investigative Workflows
- Financial Analysis
- Process Automation
- Excel VBA
- Financial Technology

---

## License

See the accompanying `LICENSE` file for licensing information.
