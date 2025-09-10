# Add these utility functions to your existing OSTtessToPDF.py file
import glob
import re
import os
import json
import csv
from pathlib import Path


def extract_meta_info(meta_file_path: str) -> dict:
    """
    Extract number of references and biggest label number from meta JSON file

    Args:
        meta_file_path: Path to the meta JSON file

    Returns:
        Dictionary with meta information
    """
    meta_info = {
        "number_of_references": 0,
        "biggest_label_number": 0,
        "has_meta_file": False
    }

    try:
        if os.path.exists(meta_file_path):
            with open(meta_file_path, "r", encoding="utf-8") as f:
                meta_data = json.load(f)

            meta_info["has_meta_file"] = True

            # Extract number of references from the specific JSON structure
            # Look for content.references.number_of_references
            if "content" in meta_data and "references" in meta_data["content"]:
                references_data = meta_data["content"]["references"]
                if "number_of_references" in references_data:
                    meta_info["number_of_references"] = references_data["number_of_references"]

            # Extract biggest label number - find the highest reference number
            # Look through reference_blocks and reference_content for numbered labels
            if "content" in meta_data and "references" in meta_data["content"]:
                references_data = meta_data["content"]["references"]
                biggest_label = 0

                if "reference_blocks" in references_data:
                    for block in references_data["reference_blocks"]:
                        if "reference_content" in block:
                            for ref_content in block["reference_content"]:
                                if "label" in ref_content and ref_content["label"]:
                                    try:
                                        # Try to convert label to integer
                                        label_num = int(ref_content["label"])
                                        biggest_label = max(biggest_label, label_num)
                                    except (ValueError, TypeError):
                                        # If label is not a number, skip it
                                        continue

                meta_info["biggest_label_number"] = biggest_label

            print(f"Meta info extracted from {os.path.basename(meta_file_path)}:")
            print(f"  Number of references: {meta_info['number_of_references']}")
            print(f"  Biggest label number: {meta_info['biggest_label_number']}")

    except Exception as e:
        print(f"Error reading meta file {meta_file_path}: {e}")

    return meta_info


def extract_issue_number_from_filename(filename: str) -> str:
    """
    Extract issue number from filename

    Args:
        filename: The filename to extract issue number from

    Returns:
        Issue number as string
    """
    # Remove file extension and common suffixes
    base_name = os.path.splitext(filename)[0]
    base_name = base_name.replace("_footnotes", "").replace("_references", "")

    # Extract numbers from filename (get the last/largest number)
    numbers = re.findall(r'\d+', base_name)
    if numbers:
        # Return the last number found (usually the issue number)
        return numbers[-1]

    return "Unknown"


def extract_journal_name_from_path(file_path: str) -> str:
    """
    Extract journal name from file path

    Args:
        file_path: Full path to the file

    Returns:
        Journal name
    """
    path_lower = file_path.lower()

    journal_mapping = {
        "tarbiz": "Tarbiz",
        "meghillot": "Meghillot",
        "shenmishivri": "Shenmishivri",
        "sidra": "Sidra",
        "leshonenu": "Leshonenu",
        "zion": "Zion"
    }

    for pattern, name in journal_mapping.items():
        if pattern in path_lower:
            return name

    return "Unknown"


def create_csv_report(report_data: list, output_folder: str, journal_name: str):
    """
    Create CSV report with processing results

    Args:
        report_data: List of dictionaries with processing results
        output_folder: Folder to save the report
        journal_name: Name of the journal
    """
    if not report_data:
        print("No data to create report")
        return

    # Create filename with timestamp
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_filename = f"{journal_name}_processing_report_{timestamp}.csv"
    csv_path = os.path.join(output_folder, csv_filename)

    # Define CSV headers
    headers = [
        "Issue_Number",
        "Filename",
        "Meta_References_Count",
        "Meta_biggest_label_number",
        "Collected_Footnotes_Count",
        "Has_Meta_File",
        "Processing_Status",
        "Accuracy_Percentage"  # Added accuracy calculation
    ]

    try:
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=headers)
            writer.writeheader()

            for row in report_data:
                # Calculate accuracy percentage if both values are available
                accuracy = ""
                if (row.get("Meta_References_Count", 0) > 0 and
                        row.get("Collected_Footnotes_Count", 0) > 0):
                    meta_count = int(row["Meta_References_Count"])
                    collected_count = int(row["Collected_Footnotes_Count"])
                    accuracy = f"{(collected_count / meta_count * 100):.1f}%"

                row["Accuracy_Percentage"] = accuracy
                writer.writerow(row)

        print(f"CSV report saved to: {csv_path}")
        return csv_path

    except Exception as e:
        print(f"Error creating CSV report: {e}")
        return None


def print_processing_summary(report_data: list, journal_name: str):
    """
    Print a detailed summary of processing results

    Args:
        report_data: List of processing results
        journal_name: Name of the journal
    """
    if not report_data:
        print("No processing data available")
        return

    total_files = len(report_data)
    processed_files = len([r for r in report_data if r.get("Processing_Status") == "Completed"])
    skipped_files = len([r for r in report_data if r.get("Processing_Status") == "Skipped"])
    error_files = total_files - processed_files - skipped_files

    total_meta_refs = sum(int(r.get("Meta_References_Count", 0)) for r in report_data)
    total_collected = sum(int(r.get("Collected_Footnotes_Count", 0)) for r in report_data)

    print("\n" + "=" * 60)
    print(f"PROCESSING SUMMARY - {journal_name.upper()}")
    print("=" * 60)
    print(f"Total Files: {total_files}")
    print(f"  ├─ Processed: {processed_files}")
    print(f"  ├─ Skipped: {skipped_files}")
    print(f"  └─ Errors: {error_files}")
    print()
    print(f"References Summary:")
    print(f"  ├─ Total Meta References: {total_meta_refs}")
    print(f"  ├─ Total Collected Footnotes: {total_collected}")
    if total_meta_refs > 0:
        accuracy = (total_collected / total_meta_refs) * 100
        print(f"  └─ Overall Accuracy: {accuracy:.1f}%")
    print()

    if processed_files > 0:
        avg_meta = total_meta_refs / processed_files
        avg_collected = total_collected / processed_files
        print(f"Averages per File:")
        print(f"  ├─ Meta References: {avg_meta:.1f}")
        print(f"  └─ Collected Footnotes: {avg_collected:.1f}")

    print("=" * 60)


# Modified main function to use these utilities
def main_with_meta_analysis():
    """
    Enhanced main function with meta analysis and CSV reporting
    """
    # Your existing main function code...
    input_folder_path = r'C:/NikWorckSpase/Corpus/tarbiz/ocr-tess-scanned/'
    output_folder_path = r'C:/NikWorckSpase/Corpus/tarbiz/outputXML/'
    meta_folder_path = r'C:/NikWorckSpase/Corpus/tarbiz/meta/'

    xlsx_files = glob.glob(os.path.join(input_folder_path, "*.xlsx"))
    xlsx_files.sort()

    # Extract journal name
    journal_name = extract_journal_name_from_path(input_folder_path)

    # List to store report data
    report_data = []

    for xlsx_file in xlsx_files:
        filename = os.path.basename(xlsx_file)
        base_name = os.path.splitext(filename)[0]
        issue_number = extract_issue_number_from_filename(filename)

        # Extract meta information
        json_file = os.path.join(meta_folder_path, base_name + ".json")
        meta_info = extract_meta_info(json_file)

        # Initialize row data for report
        row_data = {
            "Issue_Number": issue_number,
            "Filename": filename,
            "Meta_References_Count": meta_info["number_of_references"],
            "Meta_biggest_label_number": meta_info["biggest_label_number"],
            "Collected_Footnotes_Count": 0,
            "Has_Meta_File": meta_info["has_meta_file"],
            "Processing_Status": "Processed"
        }

        # Check if file should be skipped
        if os.path.exists(json_file):
            with open(json_file, "r", encoding="utf-8") as jf:
                meta_data = json.load(jf)
            if meta_data.get("skipped", False) is True:
                print(f"File {filename} skipped according to meta file")
                row_data["Processing_Status"] = "Skipped"
                report_data.append(row_data)
                continue

        try:
            # Your existing processing code...
            # processor = footnoteProcessor(config)
            # all_footnotes, main_texts = processor.process_workbook(xlsx_file)
            # ... save files ...

            # Update collected footnotes count
            # row_data["Collected_Footnotes_Count"] = len(all_footnotes)

            print(f"File: {filename}")
            print(f"  Number of references (meta): {meta_info['number_of_references']}")
            print(f"  Last label number (meta): {meta_info['biggest_label_number']}")
            # print(f"  Collected footnotes: {len(all_footnotes)}")

        except Exception as e:
            print(f"Error processing {filename}: {e}")
            row_data["Processing_Status"] = f"Error: {str(e)}"

        report_data.append(row_data)

    # Create CSV report and print summary
    csv_path = create_csv_report(report_data, output_folder_path, journal_name)
    print_processing_summary(report_data, journal_name)

    if csv_path:
        print(f"\nDetailed report saved to: {csv_path}")


if __name__ == "__main__":
    main_with_meta_analysis()