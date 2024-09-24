import pandas as pd
import numpy as np
from scipy import stats

# Read the Excel file
data = pd.read_excel("F:/NISER internship/microbe and metbolite data analysis code/microbiome_phylum/phylum_week__NT.xlsx")

# Calculate and add the sample size column
data['Sample Size'] = data.iloc[:, 1:-1].count(axis=1)

# Print the Sample Size values
print("Sample Size:")
print(data['Sample Size'])

# Eliminate rows with Sample Size less than 2
data = data[data['Sample Size'] >=2]
# Calculate and add the sample size column
data['Sample Size'] = data.iloc[:, 1:-1].count(axis=1)

# Update the sample size for rows with exactly 3 data points
data.loc[data['Sample Size'] == 3, 'Sample Size'] = 3

# Calculate the mean abundance column
data['Average'] = data.mean(axis=1)

# Calculate the standard deviation column
data['Standard Deviation'] = data.iloc[:, 1:-1].std(axis=1)

# Set the desired confidence level
confidence_level = 0.95

# Calculate the margin of error
z_score = np.abs(stats.norm.ppf((1 - confidence_level) / 2))
margin_of_error = z_score * data['Standard Deviation'] / np.sqrt(data['Sample Size'])

# Calculate the lower and upper bounds of the confidence interval
data['Lower Bound'] = data['Average'] - margin_of_error
data['Upper Bound'] = data['Average'] + margin_of_error

# Calculate the sample variance column
data['CV'] = data.iloc[:, 1:-1].var(axis=1, ddof=1)

# Save the updated DataFrame to a new Excel file
data.to_excel("F:/NISER internship/microbe and metbolite data analysis code/microbiome_phylum/phylum_week_4_NT_avg.xlsx", index=False)
