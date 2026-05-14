# commcare-2-kobo

Convert CommCare apps/forms to KoboToolbox XLSForm `.xlsx` files, with optional upload and deployment through the KoboToolbox API.

The script can fetch CommCare Application Structure API schema directly, including form questions and referenced lookup tables. It also supports plain XForm XML files and CommCare XML exported through Word XML. Word-wrapped XForms are detected and unwrapped automatically.

## Safety defaults

Runtime behavior is controlled by environment variables. For local use, copy
`.env.example` to `.env` and put secrets there.

- `UPLOAD_TO_KOBO=true` uploads forms when you run the script.
- `KOBO_DEPLOY=true` deploys uploaded forms.
- `KOBO_API_TOKEN` belongs in `.env`, not in `commcare_2_kobo.py`.
- Use `--dry-run` to validate only, even when `UPLOAD_TO_KOBO=true`.

Each boolean `.env` setting has a matching CLI flag that overrides it for a
single run: `--save/--no-save`, `--upload/--no-upload`, `--deploy/--no-deploy`,
and `--commcare-fetch/--no-commcare-fetch`.

If a Kobo token was ever committed or shared, revoke it in KoboToolbox and create a new one.

## Requirements

- Python 3.10+
- A CommCare account with API access and app permissions, when using `COMMCARE_FETCH=true`
- A KoboToolbox account and API token, only if uploading

```bash
pip install -r requirements.txt
```

## Quick start

Copy the example environment file and fill in your credentials:

```bash
cp .env.example .env
```

On Windows PowerShell:

```powershell
Copy-Item .env.example .env
```

Then run:

```bash
python commcare_2_kobo.py
```

This fetches CommCare form schema when `COMMCARE_FETCH=true`, or parses XML files from `XML_INPUT_FOLDER/` when `COMMCARE_FETCH=false`. It builds XLSForm workbooks, validates generated survey and choice names, saves `.xlsx` files to `XLS_OUTPUT/`, and uploads/deploys if enabled in `.env`.

Recommended first run:

```bash
python commcare_2_kobo.py --dry-run --no-save --no-upload
```

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

when these `.env` values are set:

```dotenv
UPLOAD_TO_KOBO=true
KOBO_DEPLOY=true
KOBO_SERVER_URL=https://eu.kobotoolbox.org
KOBO_API_TOKEN=your-token
```

You can also set values in the shell instead of `.env`:

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
access and the right app permissions. It exposes app/module/form schema,
including questions, labels, translations, choice options, groups, repeats,
required flags, relevance, constraints, and calculations. The script can
convert that API schema directly into XLSForm.

The API does not provide the same raw XForm XML as the CommCare XML export.
If you need full XML fidelity beyond the exposed schema fields, export XML
from CommCare into `XML_INPUT_FOLDER/` and run folder mode.

When a CommCare question uses a dynamic lookup table, the script fetches the
referenced Fixture/Lookup Table rows from CommCare and writes them into the
XLSForm `choices` sheet. Simple dependent lookup filters are translated into
Kobo `choice_filter` expressions.

Example dynamic lookup conversion:

```text
CommCare data source: tarjetas.ocho_ultimos_digitos
XLSForm type       : select_one num_8_tarejta
choices columns    : list_name, name, label::es, Cod_departamento, Identificacion, Numero_Tarjeta_completo, ocho_ultimos_digitos
choice_filter      : ocho_ultimos_digitos=${num_8_tarejta}
```

The script prints a lookup summary during API mode, for example:

```text
Fetched lookup table rows: tarjetas=2991, Mun_Col=278
```

If you still want to test API access, set:

```bash
$env:COMMCARE_DOMAIN = "your-domain"
$env:COMMCARE_USER = "you@example.org"
$env:COMMCARE_TOKEN = "your-commcare-api-token"
python commcare_2_kobo.py --commcare-fetch --dry-run
```

If `COMMCARE_FETCH=true` is set in `.env` and you want to force folder mode:

```bash
python commcare_2_kobo.py --no-commcare-fetch
```

Limit the number of fetched CommCare apps:

```bash
python commcare_2_kobo.py --commcare-fetch --commcare-limit 5
```

`COMMCARE_LIMIT=5` means fetch the first 5 CommCare apps and then convert all forms inside those apps. It is not a form limit. `COMMCARE_LIMIT=0` means all apps.

## Conversion notes

The API conversion preserves:

- app, module, and form names in generated form titles
- question labels and translations
- static select options
- dynamic lookup-table rows as XLSForm choices
- lookup-table fields as extra choices columns
- simple lookup filters as Kobo `choice_filter`
- groups and repeats
- required flags
- relevance, constraints, and calculations when they can be expressed in Kobo-compatible XLSForm syntax

Some CommCare-only features cannot be converted directly:

- `instance('casedb')` expressions are omitted with a warning because Kobo has no CommCare case database.
- `instance('commcaresession')` expressions are omitted with a warning because Kobo has no CommCare session object.
- external-instance expressions that are not resolved through fetched lookup tables are omitted or downgraded with a warning.

Warnings do not necessarily mean the XLSForm is invalid. They identify behavior that may need manual review for full parity with the original CommCare app.

## Limits and caveats

- `COMMCARE_LIMIT` limits CommCare apps, not forms. If one fetched app contains 20 forms, all 20 forms are converted.
- The Application Structure API returns schema, not full raw XForm XML. Folder mode with exported XML is still the highest-fidelity path for XML-specific details.
- Lookup-table choices are fetched from CommCare Fixture/Lookup Table rows and embedded into the XLSForm `choices` sheet. Large lookup tables can create large `.xlsx` files and slower Kobo imports.
- Lookup-table fields are copied into the `choices` sheet as extra columns. Reserved XLSForm headers such as `name`, `list_name`, `label`, or `image` are renamed with a `lookup_` prefix, for example CommCare field `name` becomes `lookup_name`.
- Simple lookup predicates like `departamento_id = /data/depto_registro` are converted to Kobo `choice_filter` syntax. Complex predicates that depend on CommCare case/session data are not preserved automatically.
- CommCare `casedb` and `commcaresession` expressions are omitted with warnings because Kobo does not have those runtime objects.
- A dry run validates generated workbooks locally, but only an upload/import/deploy attempt can catch every Kobo/PyXForm server-side rule.

## Environment variables

| Variable | Default | Description |
|---|---:|---|
| `XML_INPUT_FOLDER` | `./XML_INPUT_FOLDER` | Folder containing source `.xml` files |
| `XLSFORM_OUTPUT_FOLDER` | `./XLS_OUTPUT` | Folder for generated `.xlsx` files |
| `SAVE_XLSFORMS_LOCALLY` | `true` | Save generated workbooks; override per run with `--save`/`--no-save` |
| `UPLOAD_TO_KOBO` | `false` | Upload generated XLSForms to KoboToolbox |
| `KOBO_API_TOKEN` | empty | KoboToolbox API token, required for `--upload` |
| `KOBO_SERVER_URL` | `https://eu.kobotoolbox.org` | KoboToolbox server URL |
| `KOBO_DEPLOY` | `false` | Deploy after upload when true |
| `COMMCARE_FETCH` | `false` | Fetch source schema from CommCare |
| `COMMCARE_DOMAIN` | empty | CommCare project domain |
| `COMMCARE_USER` | empty | CommCare API username/email |
| `COMMCARE_TOKEN` | empty | CommCare API token |
| `COMMCARE_LIMIT` | `0` | Maximum CommCare apps to fetch; `0` means all |
| `COMMCARE_BASE_URL` | `https://www.commcarehq.org` | CommCare server URL |

## What the script does

1. Reads CommCare app/form schema from the API, or XForm XML from a folder.
2. Fetches referenced CommCare Fixture/Lookup Table rows in API mode.
3. Unwraps XForms embedded in Word XML when needed.
4. Parses questions, groups, repeats, choices, languages, skip logic, and calculations.
5. Converts dynamic lookup-table selects into XLSForm choices and simple `choice_filter` expressions.
6. Sanitizes generated XLSForm survey, settings, and choice names.
7. Avoids duplicate survey names by adding numeric suffixes.
8. Builds a Kobo-compatible XLSForm workbook.
9. Validates the generated workbook structure and reports conversion warnings.
10. Optionally uploads and deploys the form in KoboToolbox.

## Folder structure

```text
commcare-2-kobo/
|-- commcare_2_kobo.py
|-- requirements.txt
|-- README.md
|-- .env.example
|-- XML_INPUT_FOLDER/
|   |-- form1.xml
|   `-- form2.xml
`-- XLS_OUTPUT/
    |-- form1.xlsx
    `-- form2.xlsx
```

`XLS_OUTPUT/` and Python cache files are generated artifacts and are ignored by Git.
