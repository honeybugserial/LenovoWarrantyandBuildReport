# Lenovo Report Rich API

`lenovo_report_rich_api.py` — Generate Lenovo warranty + spec reports directly from Lenovo’s API.  

---

## Features
- Pulls **full warranty & spec info** from Lenovo API.
- Extracts and canonicalizes machine specifications.
- Derives a **Product Key** from subSeries info.
- Interactive report menu (save/quit) via [InquirerPy](https://github.com/kazhala/InquirerPy).

---

## Installation

Clone the repo and install dependencies:

```bash
git clone https://github.com/youruser/lenovo-report-rich-api.git
cd lenovo-report-rich-api
pip install -r requirements.txt
```

Or install directly:

```bash
pip install requests rich InquirerPy pyfiglet
```

---

## Usage

Run with a serial number:

```bash
python lenovo_report_rich_api.py -s GM079MHZ
```

Prompt for serial interactively:

```bash
python lenovo_report_rich_api.py
```

Disable Rich color/styling:

```bash
python lenovo_report_rich_api.py -s PF2V08GA --no-color
```

Automatically save to `./Reports` without menu:

```bash
python lenovo_report_rich_api.py -s PF2V08GA --autosave
```

---

## Example Output

![Lenovo Report Screenshot](https://raw.githubusercontent.com/honeybugserial/LenovoWarrantyandBuildReport/main/image.png)


---

## Command-Line Options

| Option          | Description |
|-----------------|-------------|
| `-s, --serial`  | Serial number to query (if omitted, prompts). |
| `--timeout`     | HTTP timeout in seconds (default: 15). |
| `--no-color`    | Disable Rich colors/styles. |
| `--autosave`    | Automatically save to `./Reports` and exit. |

Call with no options for menu/prompts

---

## Notes
- Only works with **Lenovo machines** that return data from Lenovo’s public API.

---

## License
MIT License — see your mother for details.
