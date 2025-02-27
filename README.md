# ROIAL-NMR
ROIAL-NMR can systematically identify potential metabolites from defined proton NMR spectral regions-of-interest (ROIs), which are identified from complex biological samples (i.e., human serum, saliva, sweat, urine, CSF, and tissues) using the Human Metabolome Database (HMDB) as a reference platform.


#  Python version and packages
1. [Python](https://www.anaconda.com/download/)>=3.9
2. XlsxWriter 3.2.2
3. pandas 2.2.3
4. PyQt5  5.15.11
5. openpyxl 3.1.5


#  How to use ROIAL-NMR?
Run python main.py

For usage, please refer the following instructions and the GUI video


##  Add analysis

### Step 1 Parameters setting: 
Follow the existing analysis template to input ROIs, trends, and significance levels:
1.	Use “+” for increase and “-” for decrease.
2.	Use “*” for Significance Level 1 and “!” for Significance Level 2.
3.	Select the sample type.
4.	Input the region to be analyzed.
5.	Highlight disease-relevant metabolites (optional).

###  Step 2: Abbreviations and Matching Results
After setting parameters and completing the calculation, open the “All Metabolites” window. Here, you can view identified metabolites along with their abbreviations, match ratios, matched regions, and concentration ranges. Users can modify or add abbreviations as needed.

###  Step 3: Final Results
Once calculations are complete in the “All Metabolites” window, navigate to the “Result Show” window to view the final results:The first table displays metabolites categorized by three significance levels. The second table shows metabolites within each ROI. All results are also saved in the “dataResult” file for further reference.


##  Combine Analysis
To compare two groups, select them in the left window and click “Combine Analysis.” The resulting table highlights metabolites identified in both groups:
1.	Overline presents the same trends for both comparison groups.
2.	Underline presents the opposite trends for both comparison groups.
3.	Users can export the table to a file for further analysis.
*In the output file of combine analysis, the double underline represents the same trend for both comparisons.

# License

This package is distributed under the MIT License.
