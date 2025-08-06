# 📄 Job Application Automator

This script automates and streamlines your job application workflow by scraping job descriptions, organizing application materials, customizing documents using templates, extracting keywords, and tracking submissions in an Excel sheet — all from your terminal.

## 🚀 Features

- ✅ **Job scraping** via [Playwright](https://playwright.dev/)
- 📄 **Resume & cover letter generation** from template folders
- 🧠 **Keyword extraction** from job description using spaCy NLP
- 📊 **Excel tracking** of applied roles
- 📁 **Organized folders** for each job application (timestamped)
- 📝 Optional PDF conversion and trimming of resumes/cover letters

## 🖼️ Directory Structure

📁 your_project/

├── jobs/

│   └── CompanyName/

│       └── YYYY-MM-DD_HH-MM-SS/

│           ├── Resume.docx

│           ├── Resume.pdf

│           ├── Cover_Letter.docx

│           ├── Cover_Letter.pdf

│           └── Job_Description.docx

├── template/

│   ├── template_1/

│   │   ├── Resume.docx

│   │   └── Cover_Letter.docx

│   ├── template_2/

│   └── ...

├── skills.json

├── applications.xlsx

└── main.py

## 🧠 Prerequisites

Install the following before running the script:

```bash
pip install -r requirements.txt
```

Also install Playwright browsers:

```bash
python -m playwright install
```

And download the spaCy language model:

```bash
python -m spacy download en_core_web_sm
```

---

## 🛠️ Setup

1. Place your resume and cover letter templates under `template/template_1`, `template/template_2`, etc.
2. Create a `skills.json` file structured like this:

```json
{
  "Languages": ["Python", "JavaScript", "C++"],
  "Cloud": ["AWS", "Azure", "GCP"],
  "Operating Systems": ["Linux", "Windows", "macOS"],
  "Tools": ["Burp Suite", "Nmap", "Wireshark"]
}
```

3. Run the script using Python:

```bash
python main.py
```

## ✅ What It Does

1. Asks you for a company name and job posting URL
2. Scrapes the job title and description using Playwright
3. Lets you select a resume template
4. Copies and saves the templates in a new timestamped folder
5. Extracts keywords using spaCy and matches them against `skills.json`
6. Saves the job description and keyword matches to `Job_Description.docx`
7. Adds the application to an `applications.xlsx` log
8. Opens the resume and cover letter for redaction

## 📬 Contact

For feedback, reach out or contribute via GitHub Issues or Pull Requests.
