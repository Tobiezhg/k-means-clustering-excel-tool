# K-Means Clustering in Excel VBA

## What This Project Does

This project implements the K-Means clustering algorithm in Excel using VBA. It allows users to apply clustering on tabular data stored in Excel and visualize the results directly within the spreadsheet. The tool is designed for novice learners who want to understand clustering concepts without requiring advanced programming skills or external software.

## Why This Project Is Useful

- Provides an Excel-based clustering solution without needing external libraries.
- Helps beginners understand K-Means clustering through step-by-step execution.
- Supports standardization of continuous data for better clustering results.
- Enables users to track iterations and centroid movements for enhanced learning.
- Offers visualization of clustering results within Excel.
- Includes an option for Balanced K-Means clustering to ensure even cluster sizes.

## Getting Started

### Prerequisites

- Microsoft Excel (2016 or later recommended)
- Macros enabled in Excel

### Installation

1. Download the Excel file containing the VBA implementation.
2. Open the file and enable macros when prompted.

### Usage (For Users)

1. Load your data into the `tblRaw` table in the Excel workbook.
2. Ensure your dataset is suitable for clustering by cleaning it with `2. Clean Data` before running the clustering process.
3. Select the number of clusters and whether to use Balanced K-Means clustering and determine the most suitable number of clusters for the dataset using Silhouette Scores. In the preprocessing step, it is recommended that data be standardized to enhance efficiency and performance.
4. Click the "3. Apply Button" button to start the process.
5. View the clustering results in the summary window (`frmClusteringResultSummary`) and in the RESULT\_TABLE sheet.
6. Review the centroid coordination, assigned clusters, and visualizations in each iteration with the Iteration Log feature.
Please refer to this [instructional video](https://youtu.be/7hnCragi7Mc) to have an overview of the tool.
## For Programmers

### VBA Implementation Details

- The main clustering algorithm is implemented in a modular structure for easy maintenance.
- Standardization of continuous variables is performed using Excel’s built-in `STANDARDIZE` formula.
- The Euclidean distance function (`DistanceEuclidean`) is used to compute distances between data points and centroids.
- Iteration logs are maintained to track centroid movements and assignments while showing Excel formulas for educational purposes.
- Balanced K-Means clustering ensures an even distribution of points across clusters.

### Extending the Project

- Enhance the data preprocessing stage to ensure a well-structured and clean dataset while preserving its original characteristics.
- Modify the macro to accommodate additional clustering constraints.
- Integrate visualization enhancements such as interactive charts.
- Optimize performance for handling larger datasets efficiently.

## Where to Get Help

- Refer to the inline comments within the VBA code for detailed explanations.
- Check Excel's built-in VBA documentation for debugging and troubleshooting.
- Contact support or seek help from VBA/Excel forums if issues arise.

## Who Maintains and Contributes to the Project

This project was created through the collaborative efforts of a team of Bachelor of Economics and Business Administration students from VGU:

- The team—Ta Nguyen Minh Hang, Vuong Binh Nguyen, Nguyen Phan Hoang Nhi, Nguyen Ai Nhi, and Phan Tan Huy—collectively contributed to research, analysis, documentation, and presentation, ensuring the project was well-structured, accessible, and effectively communicated.
- Ta Nguyen Minh Hang: Develops and maintains the project, handles documentation, and contributes to researching different approaches, exploring various techniques and algorithms to enhance the clustering process.
- Vuong Binh Nguyen: Assists with testing and evaluating the tool, providing feedback to improve usability and effectiveness.

Contributions are welcome! If you find a bug or have suggestions for improvements, since this is an Excel-based project, modifications should be made directly within the downloaded file. If you encounter any issues or have suggestions, please open an issue or start a discussion on the GitHub repository.

