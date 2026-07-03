# DSS Availability Visualiser

A single Jupyter notebook that turns the monthly DSS availability workbook into slide-ready images.
It shows each service's availability over a rolling window of months with the year to date and the
target alongside, and a one-glance summary of the current month.

Three images are produced from the same data:

- An executive tile view, one panel per service, recommended as the default. It shows a rolling
  window of months (12 by default) stitched across the financial-year boundary, so each tile carries
  a real trend. Tiles are ordered with the lowest current month first. Each tile has its own vertical
  scale capped at 100 percent, a dotted target line, a red triangle on any month below target, a red
  circle with a white centre on the lowest month (the triangle is kept where the lowest month is also
  a breach), a red shaded band where the year to date runs below target, and a vertical line marking
  the start of the new financial year in April.
- A grouped columns view for the current financial year, with each service's target drawn and any
  below-target bar flagged, useful for a straight month-to-month comparison.
- An executive summary table of the current month: counters for on target, watch and at risk, then a
  row per service with status, availability, year to date, margin over target and the target, sorted
  closest to target first. Intended to replace the old spreadsheet table on the slide.

## Project structure

- `notebook.ipynb` - the notebook, built as setup, ingestion, cleaning, the tile view, the columns
  view, the executive summary, and a closing summary
- `notebook.html` - a shareable HTML export that follows the operating system light or dark theme
- `outputs/` - generated figures, named after the source file stem so different datasets never
  overwrite each other
- `data/` - input files (gitignored)

## How to use with a new data file

1. Drop the updated workbook into the `data/` folder.
2. If the file name changed, open the notebook and update only the config cell at the top of the
   setup section. Nothing else needs editing.
3. Run all cells. The selected fiscal year, the latest month, the services present, and the
   vertical axis range all adapt to the data.

The config cell is the single point of change. It also lets you chart a specific fiscal year
instead of the latest one, and set how many months the rolling tile view shows.

## Dependencies

- Python 3
- pandas
- openpyxl (Excel reading)
- matplotlib

For the HTML export, additionally:

- nbconvert
- ipykernel

Install them with:

```
pip install pandas openpyxl matplotlib nbconvert ipykernel
```

## How to run

Open `notebook.ipynb` in VS Code or Jupyter, select the Python environment as the kernel, and run
all cells. Each visualisation cell saves its figure to `outputs/`.

## Run without Jupyter

`availability.py` is a standalone command-line version of the notebook. It produces the same three
images and is meant for a machine that has only Python installed, with no notebook environment.

1. Put the monthly workbook in the same folder as the script.
2. Run it:

   ```
   python availability.py
   ```

On the first run it installs anything it needs (pandas, matplotlib, openpyxl) automatically. It picks
up the workbook beside it even when the file name changes month to month. If there is more than one
workbook to choose from, pass the one you want:

```
python availability.py "path/to/workbook.xlsx"
```

The figures are written to `outputs/` next to the script, named after the source file stem, exactly
as the notebook does.

## Regenerating the HTML export

The export executes the notebook, writes `notebook.html`, and injects a dark-mode aware style.
That style does not survive a fresh conversion, so it is always re-applied. A one-step helper that
does both lives in the gitignored `data/` folder:

```
python data/build_html.py
```

Or run the steps manually:

```
python -m nbconvert --to html --execute notebook.ipynb --output notebook.html
```

and then re-apply the dark-mode style block to the head of `notebook.html`.
