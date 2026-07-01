# DSS Availability Visualiser

A single Jupyter notebook that turns the monthly DSS availability workbook into one slide-ready
chart. For the current fiscal year it shows each service's monthly availability from April to the
latest month, with the year to date alongside. Per-service targets differ and are deliberately
left off the chart.

Two chart styles are produced from the same data so the presentation style can be chosen:

- An executive tile view, one panel per service, recommended as the default. It shows a rolling
  window of months (12 by default) stitched across the financial-year boundary, so each tile carries
  a real trend. Tiles are ordered with the lowest current month first, each shows its monthly line,
  its year to date, and its target line, any month below target is marked in red, a vertical line
  marks the start of the new financial year in April, and the vertical axis never goes above 100
  percent while each tile scales itself so a dip stays visible.
- A grouped columns view for the current financial year, with each service's target drawn and any
  below-target bar flagged, useful for a straight month-to-month comparison.

## Project structure

- `notebook.ipynb` - the notebook, built as setup, ingestion, cleaning, the two visualisations,
  and a summary
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
