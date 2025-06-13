import pandas as pd

# Assigning file paths to variables
summary_data_region = r"/Users/nishoksmac/Desktop/Nishok/Swinburne/Semester 2 (Feb 2025)/COS60011 - Technology Design Project/Individual Research Report/Data/NHSDC02_Summary_Region wise.xlsx"
summary_data_age = r"/Users/nishoksmac/Desktop/Nishok/Swinburne/Semester 2 (Feb 2025)/COS60011 - Technology Design Project/Individual Research Report/Data/NHSDC01_Summary_Age wise.xlsx"
output_excel_path = r"/Users/nishoksmac/Desktop/Nishok/Swinburne/Semester 2 (Feb 2025)/COS60011 - Technology Design Project/Individual Research Report/Data/Data.xlsx"

# Load data
rsum_df = pd.read_excel(summary_data_region, sheet_name="Table 2.1_Estimates", skiprows=4, skipfooter=33)
asum_df = pd.read_excel(summary_data_age, sheet_name="Table 1.1_Estimates", skiprows=4, skipfooter=33)

def insert_empty_rows_and_split(df):
    df.reset_index(drop=True, inplace=True)

    # Step 1: Insert empty row before rows where column A is empty
    updated_rows = []
    for i in range(len(df)):
        if pd.isna(df.iloc[i, 0]):
            updated_rows.append(pd.Series([None] * len(df.columns), index=df.columns))
        updated_rows.append(df.iloc[i])

    df_step1 = pd.DataFrame(updated_rows).reset_index(drop=True)

    # Step 2: Insert empty row before rows where column B is empty
    final_rows = []
    for i in range(len(df_step1)):
        if pd.isna(df_step1.iloc[i, 1]):
            final_rows.append(pd.Series([None] * len(df_step1.columns), index=df_step1.columns))
        final_rows.append(df_step1.iloc[i])

    df_final = pd.DataFrame(final_rows).reset_index(drop=True)

    # Splitting the final dataframe into separate blocks (based on completely empty rows)
    block_dfs = []
    current_block = []

    for _, row in df_final.iterrows():
        if row.isnull().all():
            if current_block:
                block_dfs.append(pd.DataFrame(current_block, columns=df.columns))
                current_block = []
        else:
            current_block.append(row)

    if current_block:
        block_dfs.append(pd.DataFrame(current_block, columns=df.columns))

    return block_dfs

# Process both dataframes
rsum_blocks = insert_empty_rows_and_split(rsum_df)
asum_blocks = insert_empty_rows_and_split(asum_df)

# Save to an Excel file with blocks from both DataFrames in separate sheets
with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
    for i, block in enumerate(rsum_blocks, 1):
        block.to_excel(writer, sheet_name=f"Region_Block_{i}", index=False)
    for i, block in enumerate(asum_blocks, 1):
        block.to_excel(writer, sheet_name=f"Age_Block_{i}", index=False)

print(f"Done! Blocks saved to {output_excel_path}")
