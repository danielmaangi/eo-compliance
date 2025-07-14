# EO Compliance - Keyword Search Tool

A Python-based document analysis tool for searching and extracting keyword-containing sentences from various document formats including PDF, Excel (.xlsx), and Word (.docx) files.

## Overview

This tool helps organizations perform compliance checks by searching for specific keywords related to gender, diversity, equity, inclusion, and key populations across multiple document types and generating detailed reports with exact sentence matches and their locations.

## Features

- **Multi-format Support**: Processes PDF, Excel (.xlsx), and Word (.docx) files
- **Sentence-level Extraction**: Extracts complete sentences containing target keywords
- **Detailed Location Tracking**: Records source type, page/sheet/paragraph numbers, and cell locations
- **Automated Report Generation**: Creates CSV and Excel reports with findings
- **Batch Processing**: Processes entire folders of documents at once

## Installation

### Prerequisites
- Python 3.8 or higher
- uv package manager

### Setup
1. Clone this repository:
```bash
git clone https://github.com/danielmaangi/eo-compliance.git
cd eo-compliance
```

2. Install dependencies using uv:
```bash
uv add ipykernel pandas pypdf2 openpyxl python-docx
```

## Usage

### 1. Prepare Your Documents
- Create a `docs` folder in the project directory (or it will be created automatically)
- Place all documents you want to analyze in the `docs` folder
- Supported formats: `.pdf`, `.xlsx`, `.docx`

### 2. Configure Keywords
The tool is pre-configured with a comprehensive list of compliance-related keywords:
```python
keywords = ['Gender', 'Transgender','transmen', 'transwomen' , 'LGBTQ', 'LGBT', ' DEI ', 'Diversity', 'Equity', 'Inclusion', 'gender',' GBV', 'trans-gender', 'trans-women', 'trans-men', 'disparity', 'pregnant people',
            'Gender', 'DEI', 'Diversity', 'Inclusion', 'disparity', 'equity', 'identity', 'inclusivity', 'binary', 'non-binary', 'prejudice',
            'pronouns', 'race', 'stereotype', 'tgw', 'transgender', 'tg', 'transgender women', 'trans', 'transgender', 'protecting women', 'key pops', 'key populations', 
            'MAT', 'hormone', 'gbv', 'dreams', 'abortion', 'fsw', 'female sex worker', 'food']
```
You can modify this list in `KeyWordSearch.ipynb` to add or remove keywords as needed for your specific compliance requirements.

### 3. Run the Analysis
Execute the notebook cells in order:

1. **Import Libraries**: Loads all required modules
2. **Extract Sentences**: Processes documents and extracts sentences
3. **Find Matches**: Searches for keyword-containing sentences
4. **Generate Report**: Creates comprehensive analysis report

### 4. Review Results
The tool generates two output files:
- `docs_keyword_summary.csv`: Detailed CSV report with all matches
- `docs Keywords Check.xlsx`: Excel summary with Partner, Keyword, and Content columns

## Output Format

### CSV Report Columns:
- **File Path**: Full path to the source document
- **Source Type**: Type of content (worksheet, page, paragraph)
- **Source Name**: Specific location (Sheet name, Page number, Paragraph number)
- **Location**: Detailed position (Row/Column for Excel files)
- **Keyword**: The matched keyword
- **Exact Sentence**: Complete sentence containing the keyword

### Excel Summary Columns:
- **No.**: Sequential index
- **Partner**: Extracted filename/partner name
- **Keyword**: The matched keyword
- **Content**: Complete sentence containing the keyword

## Document Processing Details

### Excel Files (.xlsx)
- Processes all worksheets in the file
- Extracts text from each cell
- Records row and column locations
- Handles multiple sheets automatically

### PDF Files (.pdf)
- Extracts text from all pages
- Processes page by page
- Records page numbers for reference

### Word Documents (.docx)
- Processes all paragraphs
- Records paragraph numbers
- Maintains document structure

## Error Handling

The tool includes robust error handling:
- Continues processing if individual files fail
- Reports errors for problematic files
- Ensures proper resource cleanup (file handles, memory)

## Validation

Built-in validation checks ensure:
- No missing partner names in output
- Proper filename extraction
- Data integrity across all processed files

## Customization

### Adding New Keywords
Modify the `keywords` list in the configuration section of `KeyWordSearch.ipynb` to add or remove keywords as needed for your specific compliance requirements.

### Changing Input Folder
Update the `folder_path` variable:
```python
folder_path = 'your_custom_folder'  # Default is 'docs'
```

### Output Customization
Modify the report generation section to change output formats or add new columns.

## Requirements

- pandas: Data manipulation and analysis
- PyPDF2: PDF text extraction
- openpyxl: Excel file processing
- python-docx: Word document processing
- ipykernel: Jupyter notebook support

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is open source and available under the MIT License.

## Support

For issues, questions, or contributions, please open an issue on the GitHub repository.
