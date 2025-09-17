# Document Similarity Comparator

This repository contains a Python script and a GitHub Actions workflow to compare a PDF document with a Word document and determine the percentage of similarity between them.

## Problem Description

When writing academic documents using Quarto/LaTeX, the output is a PDF with specific formatting. However, reviewers often require corrections to be made in a Word document. Converting a PDF to Word while maintaining the exact structure is challenging. This workflow provides a way to evaluate if the content of the Word document is equivalent to the PDF document.

## How it Works

The `compare_pdf_docx.py` script extracts the text from both the PDF and Word documents, normalizes it, and then performs a series of comparisons to calculate similarity metrics. It generates an HTML report with these metrics and a side-by-side diff of the two documents.

### Features

- **Text Extraction**: Extracts text from PDF and DOCX files.
- **Text Normalization**: Cleans and normalizes the text to improve comparison accuracy.
- **Similarity Metrics**:
    - Character-level similarity ratio.
    - Token-based fuzzy ratio.
    - Sentence coverage to see how many sentences from the PDF are found in the DOCX.
- **HTML Report**: Generates a detailed HTML report with the similarity metrics and a visual diff.

## Usage

1. **Place your documents**: Put the PDF and DOCX files you want to compare in the root of the repository.
2. **Run the script locally**:
   ```bash
   pip install -r requirements.txt
   python compare_pdf_docx.py my_document.pdf my_document.docx
   ```
3. **Use the GitHub Actions workflow**:
   - Push your PDF and DOCX files to the repository.
   - The workflow will automatically run, comparing the two documents.
   - A comment will be added to the commit with a summary of the similarity report.
   - The full HTML report will be available as a workflow artifact.

## GitHub Actions Workflow

The `.github/workflows/compare.yml` workflow is triggered on every push to the `main` branch. It performs the following steps:

1. **Sets up the environment**: Checks out the code and sets up Python.
2. **Installs dependencies**: Installs the required Python packages from `requirements.txt`.
3. **Runs the comparison**: Executes the `compare_pdf_docx.py` script on the specified PDF and DOCX files.
4. **Uploads the report**: Uploads the generated `compare_report.html` as a workflow artifact.
5. **Comments on the commit**: Posts a summary of the comparison results as a comment on the commit that triggered the workflow.
