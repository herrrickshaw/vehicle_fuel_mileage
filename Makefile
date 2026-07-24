.PHONY: all model report data clean deps

deps:
	pip install -r requirements.txt
	npm install

# Extract the SIAM FE declaration PDF and run the fuel-efficiency analysis.
data:
	python scripts/extract_fe_pdf.py
	python scripts/analyze_fe_declarations.py

model:
	python scripts/build_model.py
	python scripts/recalc.py outputs/EV_CNG_Diesel_Ethanol_Vehicle_Cost_Model.xlsx

report:
	node scripts/build_report.js

all: data model report

clean:
	rm -f pg-*.jpg *.pdf
