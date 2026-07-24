# How to publish this to GitHub

The repo is already a git repository with an initial commit on the `main` branch.
You just need to create the GitHub repo and push. Pick either path.

## Option A — one command (GitHub CLI)

```bash
tar xzf vehicle_fuel_mileage.tar.gz
cd vehicle_fuel_mileage
gh repo create vehicle_fuel_mileage --public --source=. --remote=origin --push
```

## Option B — manual (create repo on github.com first)

1. Go to https://github.com/new and create an **empty** public repo named `vehicle_fuel_mileage`
   (no README/licence/.gitignore — the repo already has them).
2. Then:

```bash
tar xzf vehicle_fuel_mileage.tar.gz
cd vehicle_fuel_mileage
git remote add origin https://github.com/herrrickshaw/vehicle_fuel_mileage.git
git push -u origin main
```

## Verify it regenerates

```bash
pip install -r requirements.txt && npm install
make all      # rebuilds the Excel model + Word report
```

`python scripts/recalc.py outputs/EV_CNG_Diesel_Ethanol_Vehicle_Cost_Model.xlsx`
should print `"status": "success", "total_errors": 0` (needs LibreOffice for local recalc).
