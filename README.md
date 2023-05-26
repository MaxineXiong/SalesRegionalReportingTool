# Sales Regional Reporting Tool with VBA
[![GitHub][github_badge]][github_link]
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

The Sales Regional Reporting Tool is an Excel template with a VBA program designed to streamline and automate the monthly reporting process for a Sales Controller working at a construction company that owns four subsidiary companies across the USA and Europe. The tool addresses the challenge of **consolidating data from multiple sources**, **generating interactive Pivot-based regional reports for managers**, **creating a hardcoded regional overview report for top management**, and **exporting data as a CSV file for uploading to the central ERP system**.

![Automation Demo](./Input%20Sales%20Reports/ReportingTool.gif)

## Project Description

The Sales Regional Reporting Tool simplifies the reporting process by automating various tasks. It allows users to consolidate data from multiple sources, map codes to descriptions, and generate comprehensive reports tailored for regional managers and top management, with only a few clicks. The tool can save significant time of more than 4 hours and reduces the risk of manual errors associated with repetitive tasks.

## Repository Structure

The Sales Regional Reporting Tool repository is structured as follows:

- **Monthly_Sales_Reporting_Template_Tool.xlsm**: This is the main Excel template that comes with a series of Macros. It serves as the core tool for automating the reporting process. Users can customize and adapt the template according to their specific requirements.

- **SalesRegionalReporting.bas**: Contains the VBA code that powers all the Macros in the Excel workbook, enabling automation and report generation.

- **Input Sales Reports**: This folder contains individual sales reports across the USA and Europe. These reports serve as test data for the program.

- **Sample Output Reports**: This folder contains three types of sample output reports that can be generated from the reporting program. These reports showcase the formatting and layout of the generated reports.

- **README.md**: Provides an overview of this repository.

- **LICENSE**: The license file for the project.

## Getting Started

1. Clone or download the Sales Regional Reporting Tool from this repository.
2. Open the **Monthly_Sales_Reporting_Template_Tool.xlsm** file to access the Sales Regional Reporting Tool.
3. Enable macros if prompted and follow the provided step-by-step instructions to run macros one by one within the workbook.
4. Explore the interactive Pivot reports, hardcoded regional overview reports, and CSV export functionality.
5. Customize the tool according to your company's specific requirements and data sources. You can modify the workbook or edit the VBA code to tailor the tool to your needs. To edit the VBA code, open the VBA editor by pressing `Alt + F11` key. This will provide access to the underlying VBA code that powers the tool, allowing you to make changes and adjustments as necessary. To access the code related to a particular form control button, right-click the button and choose "Assign Macro" -> "Edit". This will open the VBA editor directly to the code associated with the button, allowing you to modify its functionality.
6. Utilize the tool to streamline your monthly reporting process, save time, and enhance efficiency.

## Requirements

- Microsoft Excel (version 2010 or later) with macros enabled.
- Basic knowledge of VBA and Microsoft Excel is recommended to make modifications to the tool.

## Contributing

Contributions to the Sales Regional Reporting Tool are welcome! If you have any suggestions, improvements, or bug fixes, please feel free to submit a pull request.

## License

The Sales Regional Reporting Tool is released under the [MIT License](https://choosealicense.com/licenses/mit/). Feel free to use, modify, and distribute the code in this repository.

Simplify your reporting process and drive efficiency with the Sales Regional Reporting Tool.

[github_badge]: https://badgen.net/badge/icon/GitHub?icon=github&color=black&label
[github_link]: https://github.com/MaxineXiong
