Purpose and Functionality:
This Python script utilizes the Pandas and Openpyxl libraries to effectively divide large Excel files into smaller, more manageable chunks. It addresses the issue of memory limitations and performance bottlenecks that arise when dealing with massive Excel spreadsheets.




Key Features:
Chunk Processing: The code employs a chunk-based approach, iterating through the original Excel file in blocks of a specified size (e.g., 10,000 rows). This significantly reduces memory consumption and improves processing efficiency.
Empty Row Removal: The script incorporates the dropna(how='all') function to eliminate empty rows from each data chunk. This ensures that only rows containing actual data are copied to the new split files.
Custom File Naming: The code allows for personalized file naming by initializing the numero_arquivo variable with a desired starting number (e.g., 15). The new file names are generated using the format numerosSalvos{numero_arquivo}.xlsx, ensuring consistent and organized file naming.




Applications and Usage:
This script is particularly useful for handling large Excel files that exceed memory limitations or cause performance issues during processing. It is well-suited for scenarios where:
Data Analysis: Working with massive datasets for statistical analysis or machine learning tasks can be challenging. This script can divide the data into smaller chunks for easier analysis and processing.
Data Archiving: When archiving large Excel files for storage or backup purposes, splitting them into smaller files can be more manageable and efficient.
Data Sharing: Sharing large Excel files can be problematic due to email size restrictions or recipient limitations. Splitting the files allows for easier distribution and accessibility.



Usage Instructions:
Prerequisites: Ensure you have the Pandas and Openpyxl libraries installed in your Python environment.
File Preparation: Save the Excel file you want to split with the appropriate name (e.g., "lista1.xlsx").
Code Execution: Run the Python script, ensuring the file path in the arquivo_original variable matches your file location.
Output: The script will create new Excel files named numerosSalvos15.xlsx, numerosSalvos16.xlsx, and so on, containing the split data segments.



Additional Considerations:
Data Integrity: The script maintains the integrity of the original data by copying the exact values and formatting from the source file to the split files.
Error Handling: Implement error handling mechanisms to gracefully handle potential issues such as file access errors or memory limitations.
Performance Optimization: For exceptionally large datasets, consider further optimizing the chunk size and utilizing parallel processing techniques if available.



Conclusion:
This Python code provides a robust and efficient solution for splitting large Excel files into smaller, more manageable chunks, addressing memory constraints and improving processing performance. It offers a valuable tool for data analysis, archiving, and sharing tasks involving massive Excel spreadsheets.
