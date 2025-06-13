
import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
from sklearn.cluster import KMeans
from sklearn.metrics import mean_absolute_error, r2_score

# === Load Excel File ===
excel_path = r"/Users/nishoksmac/Desktop/Nishok/Swinburne/Semester 2 (Feb 2025)/COS60011 - Technology Design Project/Individual Research Report/Data/Data.xlsx"
xls = pd.ExcelFile(excel_path)

# === Load Relevant Sheets ===
sheets_to_load = [
    'Region_Smoker_Status',
    'Region_Alcohol',
    'Region_Physical_Activity',
    'Region_Waist_Circumference',
    'Region_Veg_Consumption',
    'Region_BMI',
    'Region_Chronic_Condition'
]

def clean_region_sheet(df):
    df = df.rename(columns={df.columns[0]: "Metric"})
    df = df.set_index("Metric").T.reset_index()
    df = df.rename(columns={"index": "State"})
    return df

region_data = {sheet: clean_region_sheet(xls.parse(sheet)) for sheet in sheets_to_load}

# === Merge All Sheets ===
df = region_data['Region_Smoker_Status']
for name in sheets_to_load[1:]:
    df = pd.merge(df, region_data[name], on="State", how="inner")

# === Feature Engineering ===
df["Total_Pop"] = df["Total persons, all ages"]
df["Smoking_%"] = df["Current daily smoker"] / df["Total_Pop"] * 100
df["Alcohol_%"] = df["Exceeded guideline(o)"] / df["Total_Pop"] * 100
df["Activity_%"] = df["Met guidelines"] / df["Total_Pop"] * 100
df["VegIntake_%"] = df["Daily consumption of fruit and vegetables — Did not meet at least 1 recommendation"] / df["Total_Pop"] * 100
df["Obese_%"] = df["Total Obese (30.00 or more)"] / df["Total_Pop"] * 100
df["WaistRisk_%"] = df["Substantially increased risk(u)"] / df["Total_Pop"] * 100
df["CVD_Risk_%"] = df["Has 2 or more selected chronic conditions"] / df["Total_Pop"] * 100

features = df[["Smoking_%", "Alcohol_%", "Activity_%", "VegIntake_%", "Obese_%", "WaistRisk_%"]]
target = df["CVD_Risk_%"]

# === Linear Regression ===
lin_reg = LinearRegression()
lin_reg.fit(features, target)
df["Linear_CVD_Risk_Prediction"] = lin_reg.predict(features)

mae = mean_absolute_error(target, df["Linear_CVD_Risk_Prediction"])
r2 = r2_score(target, df["Linear_CVD_Risk_Prediction"])
print(f"Linear Regression MAE: {mae:.2f}")
print(f"Linear Regression R^2 Score: {r2:.2f}")

# === KMeans Clustering ===
kmeans = KMeans(n_clusters=3, random_state=42, n_init=10)
df["CVD_Risk_Cluster"] = kmeans.fit_predict(features)

# === Export to CSV ===
# Linear Regression Output
linear_output = df[["State", "CVD_Risk_%", "Linear_CVD_Risk_Prediction"]]
linear_output.columns = ["State", "Actual CVD Risk (%)", "Linear Regression CVD Risk (%)"]
linear_output.to_csv("Linear_Regression_Output.csv", index=False)

# KMeans Output
kmeans_output = df[[
    "State", "CVD_Risk_Cluster", "Smoking_%", "Alcohol_%",
    "Activity_%", "VegIntake_%", "Obese_%", "WaistRisk_%"
]]
kmeans_output.columns = [
    "State", "Risk Cluster", "Smoking (%)", "Alcohol (%)",
    "Physical Activity (%)", "Veg Intake (%)", "Obese (%)", "Waist Risk (%)"
]
kmeans_output.to_csv("KMeans_Clustering_Output.csv", index=False)


