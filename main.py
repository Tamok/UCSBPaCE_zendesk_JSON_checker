import os
import json
import logging
from datetime import datetime
from collections import defaultdict
import xlsxwriter


def setup_logger(log_file):
    """
    Set up logging configuration.
    """
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    # File handler
    fh = logging.FileHandler(log_file, mode='w')
    fh.setLevel(logging.INFO)
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    # Console handler
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(formatter)
    logger.addHandler(ch)


def process_file(file_path):
    """
    Process a single JSON file and return the data entries.
    """
    logging.info(f"Processing file: {os.path.basename(file_path)}")
    data_entries = []

    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                data = json.loads(line.strip())
                if isinstance(data, list):
                    data_entries.extend(data)
                else:
                    data_entries.append(data)
            except json.JSONDecodeError as e:
                logging.error(f"Error reading line in {file_path}: {e}")
    return data_entries


def analyze_data(data_entries):
    """
    Analyze the list of data entries and return statistics.
    """
    created_at_year_count = defaultdict(int)
    tags_count = defaultdict(int)
    program_area_count = defaultdict(int)
    segment_count = defaultdict(int)
    channel_count = defaultdict(int)
    type_of_inquiry_count = defaultdict(int)
    spam_ticket_count = 0
    unique_tags = set()
    unique_program_areas = set()
    unique_segments = set()
    unique_channels = set()
    unique_types_of_inquiry = set()

    # Mapping of custom field IDs to their names
    FIELD_ID_TO_NAME = {
        38954747: 'channel',
        38829288: 'type_of_inquiry',
        38830788: 'program_area',
        1500001654502: 'segment',
    }

    for entry in data_entries:
        # Count the number of tickets per "created_at" year
        if 'created_at' in entry:
            try:
                year = datetime.strptime(entry['created_at'], "%Y-%m-%dT%H:%M:%S.%fZ").year
                created_at_year_count[year] += 1
            except ValueError:
                try:
                    year = datetime.strptime(entry['created_at'], "%Y-%m-%dT%H:%M:%S.%f%z").year
                    created_at_year_count[year] += 1
                except ValueError:
                    logging.warning(f"Invalid date format for entry: {entry.get('created_at')}")

        # Collect tags and count occurrences
        tags = entry.get('tags', [])
        if tags:
            for tag in tags:
                tags_count[tag] += 1
                unique_tags.add(tag)

        # Collect program_area, segment, channel, and type_of_inquiry from custom_fields
        custom_fields = entry.get('custom_fields', [])
        program_area = None
        segment = None
        channel = None
        type_of_inquiry = None

        for field in custom_fields:
            field_id = field.get('id')
            field_value = field.get('value')
            field_name = FIELD_ID_TO_NAME.get(field_id)
            if field_name == 'program_area':
                program_area = field_value
                if program_area:
                    program_area_count[program_area] += 1
                    unique_program_areas.add(program_area)
            elif field_name == 'segment':
                segment = field_value
                if segment:
                    segment_count[segment] += 1
                    unique_segments.add(segment)
            elif field_name == 'channel':
                channel = field_value
                if channel:
                    channel_count[channel] += 1
                    unique_channels.add(channel)
            elif field_name == 'type_of_inquiry':
                type_of_inquiry = field_value
                if type_of_inquiry:
                    type_of_inquiry_count[type_of_inquiry] += 1
                    unique_types_of_inquiry.add(type_of_inquiry)

        if not program_area:
            spam_ticket_count += 1

    # Compile the analysis results
    analysis_stats = {
        'total_tickets': len(data_entries),
        'created_at_year_count': dict(created_at_year_count),
        'tags_count': dict(tags_count),
        'program_area_count': dict(program_area_count),
        'segment_count': dict(segment_count),
        'channel_count': dict(channel_count),
        'type_of_inquiry_count': dict(type_of_inquiry_count),
        'spam_ticket_count': spam_ticket_count,
        'unique_tags': unique_tags,
        'unique_program_areas': unique_program_areas,
        'unique_segments': unique_segments,
        'unique_channels': unique_channels,
        'unique_types_of_inquiry': unique_types_of_inquiry,
    }

    return analysis_stats


def write_analysis_to_excel(analysis_stats, analysis_file):
    """
    Write the analysis statistics to an Excel file with multiple sheets.
    """
    logging.info(f"Writing analysis to Excel file: {analysis_file}")
    workbook = xlsxwriter.Workbook(analysis_file)

    # Define formats
    header_format = workbook.add_format({
        'bold': True,
        'font_color': 'white',
        'bg_color': '#003660',
        'border': 1
    })
    cell_format = workbook.add_format({'border': 1})
    bold_format = workbook.add_format({'bold': True})
    percent_format = workbook.add_format({'num_format': '0.00%', 'border': 1})

    # Sheet 1: Overview
    overview_sheet = workbook.add_worksheet("Overview")
    row = 0

    # Write headers
    overview_sheet.write(row, 0, 'Metric', header_format)
    overview_sheet.write(row, 1, 'Value', header_format)
    row += 1

    # Total tickets and spam tickets
    overview_sheet.write(row, 0, 'Total number of tickets', cell_format)
    overview_sheet.write(row, 1, analysis_stats['total_tickets'], cell_format)
    row += 1

    overview_sheet.write(row, 0, 'Number of spam tickets', cell_format)
    overview_sheet.write(row, 1, analysis_stats['spam_ticket_count'], cell_format)
    row += 1

    # Yearly stats
    for year, count in sorted(analysis_stats['created_at_year_count'].items()):
        overview_sheet.write(row, 0, f'Number of tickets for year {year}', cell_format)
        overview_sheet.write(row, 1, count, cell_format)
        row += 1

    # Adjust column width
    overview_sheet.set_column('A:A', 50)
    overview_sheet.set_column('B:B', 20)

    # Sheet 2: Tag Report
    tag_sheet = workbook.add_worksheet("Tag Report")
    write_report_sheet(workbook, tag_sheet, "Tag", analysis_stats['tags_count'], header_format, cell_format)

    # Sheet 3: Program Area Report
    program_area_sheet = workbook.add_worksheet("Program Area Report")
    write_report_sheet(workbook, program_area_sheet, "Program Area", analysis_stats['program_area_count'], header_format, cell_format)

    # Sheet 4: Segment Report
    segment_sheet = workbook.add_worksheet("Segment Report")
    write_report_sheet(workbook, segment_sheet, "Segment", analysis_stats['segment_count'], header_format, cell_format)

    # Sheet 5: Channel Report
    channel_sheet = workbook.add_worksheet("Channel Report")
    write_report_sheet(workbook, channel_sheet, "Channel", analysis_stats['channel_count'], header_format, cell_format)

    # Sheet 6: Type Report
    type_sheet = workbook.add_worksheet("Type Report")
    write_report_sheet(workbook, type_sheet, "Type of Inquiry", analysis_stats['type_of_inquiry_count'], header_format, cell_format)

    # Close the workbook
    workbook.close()


def write_report_sheet(workbook, worksheet, column_title, data_dict, header_format, cell_format):
    """
    Helper function to write data to a worksheet.
    """
    worksheet.write(0, 0, column_title, header_format)
    worksheet.write(0, 1, 'Count', header_format)

    total = sum(data_dict.values())
    row = 1
    for key, count in sorted(data_dict.items(), key=lambda x: x[1], reverse=True):
        worksheet.write(row, 0, key, cell_format)
        worksheet.write(row, 1, count, cell_format)
        row += 1

    # Write total
    worksheet.write(row, 0, 'Total', header_format)
    worksheet.write(row, 1, total, header_format)

    # Adjust column width
    worksheet.set_column('A:A', 30)
    worksheet.set_column('B:B', 15)

    # Add a chart
    chart = workbook.add_chart({'type': 'pie'})
    chart.add_series({
        'name': f'{column_title} Distribution',
        'categories': [worksheet.name, 1, 0, row - 1, 0],
        'values':     [worksheet.name, 1, 1, row - 1, 1],
        'data_labels': {'percentage': True},
    })
    chart.set_title({'name': f'{column_title} Distribution'})
    chart.set_style(10)
    worksheet.insert_chart('D2', chart)


def collate_and_analyze_json_files(input_folder, output_file, combined_analysis_file):
    """
    Collate JSON files from the input folder, analyze both individual and combined data,
    and export the analysis to Excel files.
    """
    all_data = []
    split_files = [f for f in os.listdir(input_folder) if f.endswith('.json')]

    if not split_files:
        logging.error("No JSON files found in the input folder.")
        return

    # Process each split file individually
    for filename in split_files:
        file_path = os.path.join(input_folder, filename)
        data_entries = process_file(file_path)
        all_data.extend(data_entries)

        # Analyze and export analysis for each split file
        analysis_stats = analyze_data(data_entries)
        split_analysis_file = f"{os.path.splitext(filename)[0]}_analysis.xlsx"
        write_analysis_to_excel(analysis_stats, os.path.join(input_folder, split_analysis_file))

    # Validate the combined data count
    total_split_entries = len(all_data)
    existing_data_count = 0

    # Check if the output file already exists
    if os.path.exists(output_file):
        logging.info(f"{output_file} already exists. Verifying the integrity of the collated data...")
        with open(output_file, 'r', encoding='utf-8') as f:
            existing_data = json.load(f)
        existing_data_count = len(existing_data)

        if existing_data_count == total_split_entries:
            logging.info("The collated data is complete. Skipping re-collation.")
        else:
            logging.warning("The collated data does not match the split files. Re-collating data...")
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(all_data, f, indent=2)
    else:
        logging.info("Writing combined JSON data to output file...")
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(all_data, f, indent=2)

    # Analyze and export analysis for the combined data
    combined_analysis_stats = analyze_data(all_data)
    write_analysis_to_excel(combined_analysis_stats, combined_analysis_file)

    # Perform validation checks
    total_tickets = combined_analysis_stats['total_tickets']
    total_program_area_tickets = sum(combined_analysis_stats['program_area_count'].values())
    spam_ticket_count = combined_analysis_stats['spam_ticket_count']

    if total_tickets != (total_program_area_tickets + spam_ticket_count):
        logging.error(
            "Validation failed (1/2): Total tickets (%d) do not match the sum of program area tickets (%d) and spam tickets (%d).",
            total_tickets, total_program_area_tickets, spam_ticket_count)
        raise ValueError("Validation failed for combined data.")
    else:
        logging.info("Validation passed (1/2): Total tickets match the sum of program area tickets and spam tickets.")

    if total_tickets == spam_ticket_count:
        logging.error(
            "Validation failed (2/2): All tickets are considered spam (%d out of %d). Something went wrong.",
            spam_ticket_count, total_tickets)
        raise ValueError("All tickets are considered spam.")
    else:
        logging.info("Validation passed (2/2): Not all tickets are spam.")

    logging.info("Collation and analysis completed successfully.")


if __name__ == "__main__":
    input_folder = "./json_files"  # Folder containing JSON files
    output_file = "combined.json"  # Output file for combined JSON
    combined_analysis_file = "combined_analysis.xlsx"  # Output file for combined analysis
    log_file = "collate_json_files.log"  # Log file for verbose output

    setup_logger(log_file)
    logging.info("Starting collation and analysis...")
    try:
        collate_and_analyze_json_files(input_folder, output_file, combined_analysis_file)
    except Exception as e:
        logging.error(f"An error occurred during processing: {e}")
    logging.info("Process finished.")
