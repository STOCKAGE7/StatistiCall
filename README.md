# StatistiCall - Automated Reporting Tool for RB Call Center

StatistiCall is a Python-based application designed to automate the process of generating daily appointment reports for RB Call Center. This tool extracts data from Excel files, processes key statistics, and generates a structured email report, saving time and reducing errors in manual calculations.

## Features

- **Drag and drop an Excel file**: Simply drop the file containing appointment data.
- **Automatic data analysis**: The application extracts and processes key information (appointments scheduled, cancellations, reschedules, transfers, etc.).
- **Instant email generation**: A structured report is automatically generated and ready to be sent.
- **Modern and intuitive interface**: Built with CustomTkinter for a better user experience, featuring a dark mode and sleek design.
- **Quick copy button**: Allows users to copy the generated text with one click for easy email sending.

## Technologies Used

- **Python**: For data processing and report generation.
- **Pandas**: For analyzing Excel files.
- **CustomTkinter**: For building the modern and responsive UI.
- **TkinterDnD2**: For implementing the drag-and-drop file feature.

## Requirements

Before running the application, ensure you have the following installed:

- Python 3.x
- pandas
- customtkinter
- tkinterdnd2
- Pillow (for image handling)

Install the required packages using pip:

```bash
pip install pandas customtkinter tkinterdnd2 Pillow
```

## How to Run

1. Clone the repository or download the files.
2. Install the required packages as mentioned above.
3. Run the `StatistiCall.py` file.
4. Drag and drop an Excel file with appointment data onto the application window.
5. The application will automatically generate a structured report based on the Excel data.

## Excel File Format

Ensure your Excel file follows this structure:

- **Sheet 1 (Default sheet)**: Contains the `Docteur` (doctor) column with the appointment data.
- **Sheet 2 (Déplacements)**: Contains the data related to reschedules or moves.
- **Sheet 3 (Annulations)**: Contains the cancellation data.
- **Sheet 4 (Transfert)**: Contains data on transfers.
- **Sheet 5 (Autre détail)**: Any other details you want to include in the report.

## Objective

This project was developed to enhance efficiency and accuracy at RB Call Center by automating a repetitive task, reducing human errors, and optimizing workflow.

## License

This project is licensed under the MIT License 

