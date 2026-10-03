# Data

## Option A: Demo (works immediately)
`data/sample/sample_churn_data.csv` is **synthetic** (2,000 rows, same columns as the Kaggle file).
It exists only so the project runs out of the box. Do not draw business conclusions from it.

## Option B: Real data (recommended)
Dataset: **Churn Modelling** (Bank Customer Churn), Kaggle.
Search "Churn Modelling" or "Bank Customer Churn Prediction" on kaggle.com and download `Churn_Modelling.csv`.

**Manual:** download the CSV and save it as `data/Churn_Modelling.csv`.

**Kaggle CLI:**
```bash
pip install kaggle
# put your kaggle.json in ~/.kaggle/ (never inside this repo)
kaggle datasets download -d <owner>/<dataset-slug> -p data --unzip
```
Replace `<owner>/<dataset-slug>` with the slug shown in the dataset's URL.

## Expected columns (14)
RowNumber, CustomerId, Surname, CreditScore, Geography, Gender, Age, Tenure, Balance,
NumOfProducts, HasCrCard, IsActiveMember, EstimatedSalary, Exited (1 = left the bank).
