# Word Document Automation

## Overview
This project automates the generation of Word documents from Excel data using Python. It processes an Excel file, fills a Word template with relevant data, and updates tables accordingly. Additionally, the project runs a secondary script (`table_update.py`) to refine table data after document generation.

## Features
- Extracts data from an Excel file
- Fills placeholders in a Word template
- Updates only the matched tables with data
- Runs `table_update.py` automatically after generating documents

## Installation
To set up the project, install the required dependencies using:
```sh
pip install -r requirements.txt
```

## Usage
1. Ensure your input files (`template.docx`, `data.xlsx`) are in the project directory.
2. Run the main script:
```sh
python index.py
```
3. Generated documents will be saved in the `generated_docs` folder.

## Contributing
We welcome contributions to improve this project! Here’s how you can help:
- Fix bugs and optimize performance
- Improve documentation
- Add new features or enhance existing ones
- Report issues and suggest improvements

### How to Contribute
1. Fork the repository
2. Create a new branch (`git checkout -b feature-name`)
3. Commit your changes (`git commit -m "Description of changes"`)
4. Push to your branch (`git push origin feature-name`)
5. Open a pull request

## License
This project is open-source and available under the MIT License.

## Contact
For any questions or suggestions, feel free to reach out or open an issue in the repository.

