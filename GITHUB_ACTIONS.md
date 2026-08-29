# Automating PDF Conversion with GitHub Actions

This guide explains how to run the **Flyingbee PDF Converter SDK (Windows CLI)** as an
automated, server-side conversion job using GitHub Actions. It builds on the
[`FPPDFConverter.exe`](FPPDFFramework_CLI/FPPDFConverter.exe) command-line tool shipped in
this repository.

## What it does

The workflow in [`.github/workflows/pdf-convert.yml`](.github/workflows/pdf-convert.yml)
converts a PDF in the repository into an editable Office format (`.docx`, `.xlsx`, `.pptx`,
or `.txt`) and uploads the result as a downloadable build artifact — no local install needed.

## Prerequisites

| Requirement | Notes |
| :--- | :--- |
| Windows runner | The converter is a Windows binary, so the job uses `runs-on: windows-latest`. |
| Valid license | The SDK bundled here is an **expired trial** (limited to **10 pages**). For production CI, purchase a license from Flyingbee and inject the license file via a repository **secret** before running. |
| SDK binaries | `FPPDFConverter.exe`, `FPPDFFramework.dll`, and `Resources.bundle/` are large. Do **not** commit them to git for real deployments — download them from a private GitHub Release or use `actions/cache` (a commented template is included in the workflow). |

## Usage

1. Push the workflow to your repository (it lives under `.github/workflows/`).
2. Open **Actions → PDF Convert (Flyingbee SDK)** and click **Run workflow**.
3. Fill in the inputs:

   | Input | Description | Default |
   | :--- | :--- | :--- |
   | `input_pdf` | Path to the PDF to convert (relative to repo root) | `FPPDFFramework_CLI/Test.pdf` |
   | `format` | Output format: `docx`, `xlsx`, `pptx`, `txt` | `docx` |
   | `pages` | Page range, e.g. `all` or `1-3` | `all` |

4. When the job finishes, download the `converted-<format>` artifact from the run summary.

## Example (manual trigger)

```yaml
# Triggered via Actions tab -> Run workflow
input_pdf: FPPDFFramework_CLI/Test.pdf
format:    xlsx
pages:     all
```

This produces `output/Test.xlsx`, uploaded as `converted-xlsx`.

## How the job works

```powershell
$tool  = "FPPDFFramework_CLI/FPPDFConverter.exe"
$in    = "${{ github.event.inputs.input_pdf }}"
$fmt   = "${{ github.event.inputs.format }}"
$pages = "${{ github.event.inputs.pages }}"
$out   = "output/$([System.IO.Path]::GetFileNameWithoutExtension($in)).$fmt"
New-Item -ItemType Directory -Force -Path output | Out-Null
& $tool -a PDF2Files -i "$in" -o "$out" -f $fmt -p $pages
```

The invocation is identical to running the tool locally, so the behavior you tested on your
machine matches what runs in CI.

## Notes & limits

- **Trial limit:** with the bundled expired license, only PDFs up to 10 pages convert.
- **Runner quota:** conversion is CPU-bound; mind the GitHub Actions minutes on free/limited plans.
- **Artifacts:** outputs are kept as workflow artifacts; configure retention as needed.
- **Security:** never hard-code a license or API key in the workflow file — use
  `secrets` and reference them with `${{ secrets.NAME }}`.
