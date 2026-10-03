"""Generate a SYNTHETIC sample dataset with the same columns as the Kaggle
'Churn_Modelling.csv' file. Used only so the project runs out of the box.
Replace it with the real Kaggle file for real insights (see data/README.md)."""
import argparse
import numpy as np
import pandas as pd

def make(n=2000, seed=42):
    rng = np.random.default_rng(seed)
    geo = rng.choice(["France", "Spain", "Germany"], n, p=[0.50, 0.25, 0.25])
    gender = rng.choice(["Female", "Male"], n, p=[0.45, 0.55])
    age = np.clip(rng.normal(39, 10, n).round(), 18, 85).astype(int)
    tenure = rng.integers(0, 11, n)
    credit = np.clip(rng.normal(650, 95, n).round(), 350, 850).astype(int)
    products = rng.choice([1, 2, 3, 4], n, p=[0.50, 0.46, 0.03, 0.01])
    active = rng.choice([0, 1], n, p=[0.48, 0.52])
    has_card = rng.choice([0, 1], n, p=[0.30, 0.70])
    balance = np.where(rng.random(n) < 0.36, 0.0,
                       np.clip(rng.normal(120000, 30000, n), 5000, 250000))
    balance = np.where((geo == "Germany") & (balance > 0), balance + 25000, balance).round(2)
    salary = rng.uniform(1000, 200000, n).round(2)

    z = (-3.2 + 0.065 * (np.minimum(age, 60) - 18) + 0.9 * (geo == "Germany")
         + 0.45 * (gender == "Female") - 1.0 * active
         + 1.1 * (products >= 3) + 2.5 * (products == 4) - 0.6 * (products == 2))
    exited = (rng.random(n) < 1 / (1 + np.exp(-z))).astype(int)

    surnames = rng.choice(["Smith", "Kumar", "Chen", "Garcia", "Mueller", "Dubois",
                           "Rossi", "Singh", "Lopez", "Brown"], n)
    return pd.DataFrame({
        "RowNumber": np.arange(1, n + 1),
        "CustomerId": rng.integers(15_600_000, 15_820_000, n),
        "Surname": surnames, "CreditScore": credit, "Geography": geo,
        "Gender": gender, "Age": age, "Tenure": tenure, "Balance": balance,
        "NumOfProducts": products, "HasCrCard": has_card, "IsActiveMember": active,
        "EstimatedSalary": salary, "Exited": exited})

if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--rows", type=int, default=2000)
    p.add_argument("--out", default="data/sample/sample_churn_data.csv")
    a = p.parse_args()
    df = make(a.rows)
    df.to_csv(a.out, index=False)
    print(f"Wrote {len(df):,} synthetic rows to {a.out} (churn rate {df.Exited.mean():.1%})")
