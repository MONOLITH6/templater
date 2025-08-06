# 📄 Job Application Automator

This script automates and streamlines your job application workflow by scraping job descriptions, organizing application materials, customizing documents using templates, extracting keywords, and tracking submissions in an Excel sheet — all from your terminal.

## 🚀 Features

- ✅ **Job scraping** via [Playwright](https://playwright.dev/)
- 📄 **Resume & cover letter generation** from template folders
- 🧠 **Keyword extraction** from job description using spaCy NLP
- 📊 **Excel tracking** of applied roles
- 📁 **Organized folders** for each job application (timestamped)

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

### Create Your Own Templates

Place your resume and cover letter templates in subfolders under the `template/` directory.

*Each folder’s name will appear as a category in the CLI selection prompt.*
For example:

```
template/
├── Red_Team/
├── SOC_Analyst/
└── Sec_Engineer/
```

Will display:

```
Select a template category:
1. Red_Team
2. SOC_Analyst
3. Sec_Engineer
```

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
python templater.py
```

## 🧪 What the Script Does

1. Prompts you for:
   * Company name
   * Job link
   * Template category
   * Other notes (optional)
2. Scrapes job title + full description via Playwright
3. Creates a timestamped folder for that application
4. Copies the selected resume and cover letter templates into the folder
5. Extracts keywords using spaCy and matches them to your `skills.json`
6. Saves job description and matched skills to `Job_Description.docx`
7. Updates the `applications.xlsx` file with metadata
8. Opens the resume and cover letter for redaction

## 📬 Contact

For feedback, reach out or contribute via GitHub Issues or Pull Requests.

## 💡 Future Ideas

* GPT-4 integration for rewriting resumes and auto-generating custom cover letters
* Bulk processing of job links
