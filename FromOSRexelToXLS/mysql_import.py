import mysql.connector
import os
import xml.etree.ElementTree as ET
import csv
import re
from datetime import datetime

# Database connection parameters
db_config = {
    "host": "localhost",
    "user": "nikita",
    "password": "nikita!##123",
    "database": "academic_journals",
    "charset": "utf8mb4",
    "collation": "utf8mb4_unicode_ci"
}


def connect_to_db():
    """Establish connection to the database"""
    try:
        conn = mysql.connector.connect(**db_config)
        print("Successfully connected to database")
        return conn
    except mysql.connector.Error as err:
        print(f"Connection error: {err}")
        return None


def get_journal_id(conn, journal_name):
    """Get journal ID by its name"""
    cursor = conn.cursor()
    # Case-insensitive search
    query = "SELECT journal_id FROM journals WHERE LOWER(journal_name) = LOWER(%s)"
    cursor.execute(query, (journal_name,))
    result = cursor.fetchone()
    cursor.close()

    if result:
        print(f"Found journal '{journal_name}' with ID: {result[0]}")
        return result[0]
    else:
        print(f"Journal '{journal_name}' not found in database")
        # List available journals for debugging
        cursor = conn.cursor()
        cursor.execute("SELECT journal_name FROM journals")
        available = cursor.fetchall()
        print(f"Available journals: {[j[0] for j in available]}")
        cursor.close()
        return None


def create_issue(conn, journal_id, issue_number):
    """Create a new journal issue"""
    cursor = conn.cursor()

    # Check if issue already exists
    check_query = "SELECT issue_id FROM issues WHERE journal_id = %s AND issue_number = %s"
    cursor.execute(check_query, (journal_id, issue_number))
    existing_issue = cursor.fetchone()

    if existing_issue:
        print(f"Issue {issue_number} already exists with ID: {existing_issue[0]}")
        cursor.close()
        return existing_issue[0]

    # Create new issue
    insert_query = "INSERT INTO issues (journal_id, issue_number) VALUES (%s, %s)"
    cursor.execute(insert_query, (journal_id, issue_number))
    conn.commit()

    issue_id = cursor.lastrowid
    print(f"Created new issue {issue_number} with ID: {issue_id}")
    cursor.close()
    return issue_id


def extract_journal_info(filename):
    """Extract journal name and issue number from filename"""
    file_base = os.path.splitext(filename)[0].replace("_references", "").replace("_footnotes", "")

    # Dictionary mapping filename patterns to journal names
    journal_mapping = {
        "leshonenu": "Leshonenu",
        "meghillot": "Meghillot",
        "shenmishivri": "Shenmishivri",
        "sidra": "Sidra",
        "tarbiz": "Tarbiz",
        "zion": "Zion"
    }

    journal_name = None
    for pattern, name in journal_mapping.items():
        if pattern in file_base.lower():
            journal_name = name
            break

    if not journal_name:
        print(f"Could not determine journal for file {filename}")
        return None, None

    # Extract issue number from filename
    issue_match = re.search(r'(\d+)', file_base)
    if issue_match:
        issue_number = issue_match.group(1)
    else:
        # If no number found, use the whole base name
        issue_number = file_base

    return journal_name, issue_number


def import_xml_file(conn, xml_file_path):
    """Import data from XML file"""
    try:
        file_name = os.path.basename(xml_file_path)
        print(f"\nProcessing XML file: {file_name}")

        # Extract journal and issue information
        journal_name, issue_number = extract_journal_info(file_name)
        if not journal_name:
            return

        # Get journal ID
        journal_id = get_journal_id(conn, journal_name)
        if not journal_id:
            return

        # Create issue
        issue_id = create_issue(conn, journal_id, issue_number)

        # Parse XML file
        tree = ET.parse(xml_file_path)
        root = tree.getroot()

        main_text_count = 0
        reference_count = 0

        # Process pages
        for page in root.findall('.//Page'):
            page_name = page.get('name')
            print(f"Processing page: {page_name}")

            # Import main text
            main_text_elem = page.find('.//MainText')
            if main_text_elem is not None and main_text_elem.text:
                cursor = conn.cursor()
                # Check if exists
                check_query = "SELECT text_id FROM main_texts WHERE issue_id = %s AND page_name = %s"
                cursor.execute(check_query, (issue_id, page_name))
                if not cursor.fetchone():
                    insert_query = "INSERT INTO main_texts (issue_id, page_name, content) VALUES (%s, %s, %s)"
                    cursor.execute(insert_query, (issue_id, page_name, main_text_elem.text))
                    conn.commit()
                    main_text_count += 1
                    print(f"  Added main text for page {page_name}")
                else:
                    print(f"  Main text for page {page_name} already exists")
                cursor.close()

            # Import references (looking for both Reference and footnote elements)
            for ref in page.findall('.//Reference'):
                ref_number = ref.get('number')
                ref_content = ref.text if ref.text else ""

                cursor = conn.cursor()
                # Check if exists
                check_query = "SELECT reference_id FROM document_references WHERE issue_id = %s AND reference_number = %s"
                cursor.execute(check_query, (issue_id, ref_number))
                if not cursor.fetchone():
                    insert_query = "INSERT INTO document_references (issue_id, page_name, reference_number, content) VALUES (%s, %s, %s, %s)"
                    cursor.execute(insert_query, (issue_id, page_name, ref_number, ref_content))
                    conn.commit()
                    reference_count += 1
                    print(f"  Added reference #{ref_number}")
                cursor.close()

            # Import footnotes if they exist
            for footnote in page.findall('.//footnote'):
                footnote_number = footnote.get('number')
                footnote_content = footnote.text if footnote.text else ""

                cursor = conn.cursor()
                check_query = "SELECT footnote_id FROM footnotes_table WHERE issue_id = %s AND footnote_number = %s"
                cursor.execute(check_query, (issue_id, footnote_number))
                if not cursor.fetchone():
                    insert_query = "INSERT INTO footnotes_table (issue_id, page_name, footnote_number, content) VALUES (%s, %s, %s, %s)"
                    cursor.execute(insert_query, (issue_id, page_name, footnote_number, footnote_content))
                    conn.commit()
                    print(f"  Added footnote #{footnote_number}")
                cursor.close()

        print(f"Successfully imported {file_name}: {main_text_count} main texts, {reference_count} references")

    except Exception as e:
        print(f"Error importing file {xml_file_path}: {e}")
        import traceback
        traceback.print_exc()


def import_csv_file(conn, csv_file_path):
    """Import data from CSV file with improved error handling and type checking"""
    try:
        file_name = os.path.basename(csv_file_path)
        print(f"\nProcessing CSV file: {file_name}")

        # Extract journal and issue information
        journal_name, issue_number = extract_journal_info(file_name)
        if not journal_name:
            return

        # Get journal ID
        journal_id = get_journal_id(conn, journal_name)
        if not journal_id:
            return

        # Create issue
        issue_id = create_issue(conn, journal_id, issue_number)

        main_text_count = 0
        reference_count = 0

        # Check if file exists and is readable
        if not os.path.exists(csv_file_path):
            print(f"File does not exist: {csv_file_path}")
            return

        # Read CSV file with proper encoding and error handling
        with open(csv_file_path, 'r', encoding='utf-8-sig', newline='') as f:
            # First, read a few lines to check the structure
            first_line = f.readline().strip()
            print(f"First line of CSV: {first_line}")

            # Reset file position
            f.seek(0)

            # Try to determine delimiter
            sample = f.read(1024)
            f.seek(0)

            # Check if it's tab-separated or comma-separated
            if '\t' in sample:
                delimiter = '\t'
            elif ',' in sample:
                delimiter = ','
            else:
                delimiter = ','

            print(f"Using delimiter: '{delimiter}'")

            try:
                csv_reader = csv.DictReader(f, delimiter=delimiter)

                # Print headers for debugging
                print(f"CSV headers: {csv_reader.fieldnames}")

                row_count = 0
                for row in csv_reader:
                    row_count += 1

                    # TYPE CHECK: Ensure row is a dictionary
                    if not isinstance(row, dict):
                        print(
                            f"Warning: Row {row_count} is not a dictionary, skipping. Type: {type(row)}, Value: {row}")
                        continue

                    # Debug: print first few rows
                    if row_count <= 3:
                        print(f"Row {row_count}: {row}")

                    # Handle different possible column names
                    content_type = None
                    page_name = None
                    content = None
                    ref_number = None

                    # Safely iterate through dictionary items
                    try:
                        for key, value in row.items():
                            if key is None or value is None:
                                continue

                            key_lower = str(key).lower().strip()
                            value_str = str(value).strip()

                            if key_lower in ['type']:
                                content_type = value_str
                            elif key_lower in ['page']:
                                page_name = value_str
                            elif key_lower in ['content']:
                                content = value_str
                            elif key_lower in ['number']:
                                ref_number = value_str

                    except AttributeError as e:
                        print(f"Error processing row {row_count}: {e}")
                        print(f"Row content: {row}")
                        continue

                    # Skip empty rows
                    if not content_type or not page_name or not content:
                        if row_count <= 10:  # Only show warning for first 10 rows
                            print(
                                f"Skipping row {row_count}: missing data - Type: {content_type}, Page: {page_name}, Content exists: {bool(content)}")
                        continue

                    # Clean the data
                    content_type = str(content_type).strip()
                    page_name = str(page_name).strip()
                    content = str(content).strip()

                    cursor = conn.cursor()

                    if content_type.lower() in ['maintext', 'main_text']:
                        # Check if exists
                        check_query = "SELECT text_id FROM main_texts WHERE issue_id = %s AND page_name = %s"
                        cursor.execute(check_query, (issue_id, page_name))
                        if not cursor.fetchone():
                            insert_query = "INSERT INTO main_texts (issue_id, page_name, content) VALUES (%s, %s, %s)"
                            cursor.execute(insert_query, (issue_id, page_name, content))
                            conn.commit()
                            main_text_count += 1
                            print(f"  Added main text for page {page_name}")

                    elif content_type.lower() in ['reference']:
                        if ref_number:
                            try:
                                ref_number = int(ref_number)
                                # Check if exists
                                check_query = "SELECT reference_id FROM document_references WHERE issue_id = %s AND reference_number = %s"
                                cursor.execute(check_query, (issue_id, ref_number))
                                if not cursor.fetchone():
                                    insert_query = "INSERT INTO document_references (issue_id, page_name, reference_number, content) VALUES (%s, %s, %s, %s)"
                                    cursor.execute(insert_query, (issue_id, page_name, ref_number, content))
                                    conn.commit()
                                    reference_count += 1
                                    print(f"  Added reference #{ref_number}")
                            except ValueError:
                                print(f"  Invalid reference number: {ref_number}")

                    cursor.close()

                print(
                    f"Successfully imported {file_name}: {main_text_count} main texts, {reference_count} references from {row_count} rows")

            except csv.Error as e:
                print(f"CSV parsing error: {e}")
                # Try alternative parsing method
                print("Attempting alternative parsing method...")
                f.seek(0)

                # Read line by line and parse manually
                lines = f.readlines()
                if len(lines) < 2:
                    print("File has insufficient data")
                    return

                headers = lines[0].strip().split(delimiter)
                print(f"Manual parsing headers: {headers}")

                for i, line in enumerate(lines[1:], 2):
                    values = line.strip().split(delimiter)
                    if len(values) != len(headers):
                        print(f"Row {i}: column count mismatch, skipping")
                        continue

                    row_dict = dict(zip(headers, values))
                    # Process row_dict similar to above...
                    # (You can add the same processing logic here if needed)

    except Exception as e:
        print(f"Error importing file {csv_file_path}: {e}")
        import traceback
        traceback.print_exc()


def debug_csv_structure(csv_file_path):
    """Debug function to check CSV file structure"""
    print(f"\n=== DEBUGGING CSV STRUCTURE: {os.path.basename(csv_file_path)} ===")

    try:
        with open(csv_file_path, 'r', encoding='utf-8-sig') as f:
            # Read first 5 lines
            lines = []
            for i in range(5):
                line = f.readline()
                if not line:
                    break
                lines.append(line.strip())

            print("First 5 lines:")
            for i, line in enumerate(lines, 1):
                print(f"Line {i}: {line}")

            # Reset and try to parse as CSV
            f.seek(0)

            # Try comma delimiter
            print("\nTrying comma delimiter:")
            csv_reader = csv.DictReader(f, delimiter=',')
            headers = csv_reader.fieldnames
            print(f"Headers: {headers}")

            # Read first row
            try:
                first_row = next(csv_reader)
                print(f"First data row: {first_row}")
                print(f"Row type: {type(first_row)}")
            except StopIteration:
                print("No data rows found")

            # Reset and try tab delimiter
            f.seek(0)
            print("\nTrying tab delimiter:")
            csv_reader = csv.DictReader(f, delimiter='\t')
            headers = csv_reader.fieldnames
            print(f"Headers: {headers}")

            try:
                first_row = next(csv_reader)
                print(f"First data row: {first_row}")
                print(f"Row type: {type(first_row)}")
            except StopIteration:
                print("No data rows found")

    except Exception as e:
        print(f"Error debugging CSV: {e}")


# ... (rest of the verification functions remain the same)

def main():
    """Main function for data import"""
    conn = connect_to_db()
    if not conn:
        return

    # Specify the path to the folder containing XML and CSV files
    folder_path = r'C:/NikWorckSpase/Corpus/tarbiz/outputXMLscanned/'

    if not os.path.exists(folder_path):
        print(f"Folder does not exist: {folder_path}")
        return

    print(f"Processing files in: {folder_path}")

    # Debug first CSV file if any exist
    csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]
    if csv_files:
        debug_csv_structure(os.path.join(folder_path, csv_files[0]))

    # Process all XML files in the folder
    xml_files = [f for f in os.listdir(folder_path) if f.endswith('.xml')]

    for xml_file in xml_files:
        xml_file_path = os.path.join(folder_path, xml_file)
        import_xml_file(conn, xml_file_path)


    for csv_file in csv_files:
        csv_file_path = os.path.join(folder_path, csv_file)
        import_csv_file(conn, csv_file_path)

    conn.close()
    print("\nData import completed")


if __name__ == "__main__":
    main()