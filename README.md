# commcare-2-kobo

Convert CommCare XForm XML files to KoboToolbox XLSForm `.xlsx` files, with optional upload and deployment through the KoboToolbox API.

The script supports plain XForm XML files and CommCare XML exported through Word XML. Word-wrapped XForms are detected and unwrapped automatically.

## Safety defaults

Publishing behavior is controlled in the `CONFIG` block at the top of
`commcare_2_kobo.py`.

- `UPLOAD_TO_KOBO = True` uploads forms when you run the script.
- `KOBO_DEPLOY = True` deploys uploaded forms.
- `KOBO_API_TOKEN` can be pasted into the script, but for GitHub it is safer to keep it blank and put the token in `.env`.
- Use `--dry-run` to validate only, even when `UPLOAD_TO_KOBO = True`.

If a Kobo token was ever committed or shared, revoke it in KoboToolbox and create a new one.

## Requirements

- Python 3.10+
- A KoboToolbox account and API token, only if uploading

```bash
pip install -r requirements.txt
```

## Quick start

Place `.xml` files in `XML_INPUT_FOLDER/`, edit the `CONFIG` block in
`commcare_2_kobo.py`, then run:

```bash
python commcare_2_kobo.py
```

This parses each XML file, builds the XLSForm workbook, validates generated survey and choice names, saves `.xlsx` files to `XLS_OUTPUT/`, and uploads/deploys if enabled in the script.

To use a different input folder:

```bash
python commcare_2_kobo.py --input-folder "path/to/xml-files"
```

To validate without publishing:

```bash
python commcare_2_kobo.py --dry-run
```

To validate without saving workbooks:

```bash
python commcare_2_kobo.py --dry-run --no-save
```

## KoboToolbox publishing

The script can publish with only:

```bash
python commcare_2_kobo.py
```

when these config values are set:

```python
UPLOAD_TO_KOBO = True
KOBO_DEPLOY = True
KOBO_SERVER_URL = "https://eu.kobotoolbox.org"
KOBO_API_TOKEN = ""
```

For GitHub, keep `KOBO_API_TOKEN` blank in the script and set it in `.env`:

PowerShell:

```powershell
$env:KOBO_API_TOKEN = "your-token"
python commcare_2_kobo.py
```

Bash:

```bash
export KOBO_API_TOKEN="your-token"
python commcare_2_kobo.py
```

Use the global Kobo server instead of the EU server:

```bash
python commcare_2_kobo.py --kobo-server-url "https://kf.kobotoolbox.org"
```

## CommCare fetch mode

CommCare's supported Application Structure API requires a plan with API
access and the right app permissions. It can expose app/module/form schema,
but it does not provide the same raw XForm XML as the CommCare XML export.
For full conversion fidelity, export XML from CommCare into
`XML_INPUT_FOLDER/` and run folder mode.

If you still want to test API access, set:

```bash
$env:COMMCARE_DOMAIN = "your-domain"
$env:COMMCARE_USER = "you@example.org"
$env:COMMCARE_TOKEN = "your-commcare-api-token"
python commcare_2_kobo.py --commcare-fetch --dry-run
```

Limit the number of fetched forms:

```bash
python commcare_2_kobo.py --commcare-fetch --commcare-limit 5
```

## Environment variables

| Variable | Default | Description |
|---|---:|---|
| `XML_INPUT_FOLDER` | `./XML_INPUT_FOLDER` | Folder containing source `.xml` files |
| `XLSFORM_OUTPUT_FOLDER` | `./XLS_OUTPUT` | Folder for generated `.xlsx` files |
| `SAVE_XLSFORMS_LOCALLY` | `true` | Save generated workbooks unless `--no-save` is used |
| `KOBO_API_TOKEN` | empty | KoboToolbox API token, required for `--upload` |
| `KOBO_SERVER_URL` | `https://eu.kobotoolbox.org` | KoboToolbox server URL |
| `KOBO_DEPLOY` | `false` | Deploy after upload when true, or use `--deploy` |
| `COMMCARE_FETCH` | `false` | Fetch source XML from CommCare |
| `COMMCARE_DOMAIN` | empty | CommCare project domain |
| `COMMCARE_USER` | empty | CommCare API username/email |
| `COMMCARE_TOKEN` | empty | CommCare API token |
| `COMMCARE_LIMIT` | `0` | Maximum CommCare forms to fetch; `0` means all |

## What the script does

1. Reads CommCare XForm XML from a folder or the CommCare API.
2. Unwraps XForms embedded in Word XML when needed.
3. Parses questions, groups, repeats, choices, languages, skip logic, and calculations.
4. Sanitizes generated XLSForm survey, settings, and choice names.
5. Avoids duplicate survey names by adding numeric suffixes.
6. Builds a Kobo-compatible XLSForm workbook.
7. Validates the generated workbook structure.
8. Optionally uploads and deploys the form in KoboToolbox.

## Folder structure

```text
commcare-2-kobo/
├── commcare_2_kobo.py
├── requirements.txt
├── README.md
├── XML_INPUT_FOLDER/
│   ├── form1.xml
│   └── form2.xml
└── XLS_OUTPUT/
    ├── form1.xlsx
    └── form2.xlsx
```

`XLS_OUTPUT/` and Python cache files are generated artifacts and are ignored by Git.
