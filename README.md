# Payment Terms Estimator

Desktop Python app that reads pricing rules from `Payment Terms.xlsx` and calculates payment estimates for new and used buses.

## Features

- Loads rates by:
	- Points
	- Loan Type
	- Term-Months
	- Bus condition (new or used)
- Required exact match for:
	- Loan Type
	- Term-Months
	- Bus condition
- Points behavior:
	- Tries exact selected points first
	- Falls back to Points = 2 if no exact points row is found
- Monthly payment formula is aligned with Excel PMT-style logic used in the workbook
- Balloon amount is calculated from residual percent and financed amount where applicable
- Copy-friendly output panel with `Copy Results` button (clipboard support)

## Project Files

- `app.py`: Tkinter desktop UI
- `estimator.py`: Workbook loader and estimator logic
- `Payment Terms.xlsx`: Rate/rule source data
- `requirements.txt`: Python dependency list

## Requirements

- Python 3.10+

## Setup

1. Open a terminal in this folder.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Run

```bash
python app.py
```

## How To Use

1. Enter:
	 - Amount
	 - Deposit
	 - Term Months
	 - Bus Condition
	 - Loan Type
	 - Points
2. Click `Calculate`.
3. Review results:
	 - Rate
	 - Residual
	 - Balloon Amount
	 - Monthly Payment
	 - Total Payable
4. Copy summary text from the output box, or click `Copy Results`.

## Notes

- If workbook rates change, replace or update `Payment Terms.xlsx` and rerun the app.
- Keep sheet names containing `New` and `Used` so condition mapping continues to work.
