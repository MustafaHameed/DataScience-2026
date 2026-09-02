# AI Agents Lecture Generator

This folder now includes `generate_ai_agents_handout.py`, which creates:

- `AI_Agents_Lecture_Handout.docx`
- `AI_Agents_Lecture_Handout.pdf`
- `AI_Agents_Lecture_Presentation.pptx`

## Requirements

Install dependencies:

```powershell
pip install -r requirements.txt
```

## Run

```powershell
python generate_ai_agents_handout.py
```

## Notes

- PDF generation uses `docx2pdf`, which typically requires Microsoft Word on Windows.
- The script writes all outputs into this folder.
