from dataclasses import dataclass
from datetime import datetime
from io import IOBase
from logging import getLogger
from typing import Iterator
from zipfile import ZipFile

import xml.etree.ElementTree as ET

# -----------------
# Constants
#------------------

CONTENT_XML_FILE_NAME = 'content.xml'

OFFICE_NS = 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
TABLE_NS = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'
TEXT_NS = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'

NAMESPACE_DICT = {
    'office': OFFICE_NS,
    'table': TABLE_NS,
    'text': TEXT_NS
}

VALUE_TYPE_ATTRIBUTE = f'{{{OFFICE_NS}}}value-type'
VALUE_ATTRIBUTE = f'{{{OFFICE_NS}}}value'
STRING_VALUE_ATTRIBUTE = f'{{{OFFICE_NS}}}string-value'

TABLE_NAME_ATTRIBUTE = f'{{{TABLE_NS}}}name'
TABLE_NUMBER_COLUMNS_REPEATED_ATTRIBUTE = f'{{{TABLE_NS}}}number-columns-repeated'
TABLE_NUMBER_ROWS_REPEATED_ATTRIBUTE = f'{{{TABLE_NS}}}number-rows-repeated'

TABLE_TABLE_TAG = f'{{{TABLE_NS}}}table'
TABLE_COLUMN_TAG = f'{{{TABLE_NS}}}table-column'
TABLE_ROW_TAG = f'{{{TABLE_NS}}}table-row'
TABLE_CELL_TAG = f'{{{TABLE_NS}}}table-cell'
TEXT_P_TAG = f'{{{TEXT_NS}}}p'

PARSED_ELEMENTS = {
    TABLE_TABLE_TAG,
    TABLE_CELL_TAG,
    TABLE_ROW_TAG,
}

TAG_START_EVENT = 'start'
TAG_END_EVENT = 'end'

logger = getLogger(__name__)

#------------------
# Options
#------------------

@dataclass(slots=True)
class ODSParserOptions:
    table: str | int = 0
    convert_values: bool = False
    take_n_rows: int | None = None
    skip_n_rows: int | None = None
    skip_empty_rows_at_start: bool = True
    verify_zip: bool = True

#------------------
# Parser
#------------------

class ODSParser():
    def __init__(self, default_options: ODSParserOptions | None = None):
        self.options = default_options or ODSParserOptions()

    def _parse_table_internal(self, ods_contents: IOBase, table: str | int) -> Iterator[tuple]:
        if ods_contents is None:
            raise ValueError("'ods_contents' was null")
        
        if not isinstance(table, (int, str)):
            raise ValueError("'table' must be an int or str")

        # Cached option values
        convert_values = self.options.convert_values
        take_n_rows = self.options.take_n_rows
        skip_n_rows = self.options.skip_n_rows
        skip_empty_rows_at_start = self.options.skip_empty_rows_at_start

        # Tracking variables for finding the right table
        seen_target_table = False
        number_of_tables_checked = 0

        # Row counting for the take/skip N rows functionality
        row_count = 0
        rows_taken = 0

        # Value accumulator for the current row
        current_row = []
        current_row_has_value = False

        for event, element in ET.iterparse(ods_contents, events=[TAG_START_EVENT, TAG_END_EVENT]):
            # Cached element properties
            iter_element = iter(element)
            child = next(iter_element, None)
            second = next(iter_element, None)

            tag_name = element.tag
            attrib_get_func = element.attrib.get

            # Check if the current table element matches the name or index provided
            if (not seen_target_table) and tag_name == TABLE_TABLE_TAG and event == TAG_START_EVENT:
                table_name = attrib_get_func(TABLE_NAME_ATTRIBUTE)
                
                if (isinstance(table, str) and table == table_name) or (isinstance(table, int) and table == number_of_tables_checked):
                    seen_target_table = True
                    continue

                number_of_tables_checked += 1
                element.clear()

            # Ignore element start events once the table is located
            if event == TAG_START_EVENT:
                continue

            # Handle </table:table-row> tag
            if tag_name == TABLE_ROW_TAG:
                row_repeat_amount = int(attrib_get_func(TABLE_NUMBER_ROWS_REPEATED_ATTRIBUTE, 1))
                
                for _ in range(row_repeat_amount):
                    # Increment row count by 1
                    row_count += 1

                    # Skip the requested amount of rows
                    if skip_n_rows and row_count <= skip_n_rows:
                        continue

                    # Skip the row if it's empty and the "skip_empty_rows_at_start" option is True
                    if (not current_row_has_value) and skip_empty_rows_at_start:
                        continue

                    # Clear the "skip_empty_rows_at_start" option when the first row with data is found
                    skip_empty_rows_at_start = False

                    yield tuple(current_row)

                    rows_taken += 1

                    # Stop iteration if the targeted number of rows have already been returned
                    if take_n_rows and rows_taken >= take_n_rows:
                        return

                current_row_has_value = False
                current_row = []

                element.clear()

            # Handle </table:table-cell> tag
            if tag_name == TABLE_CELL_TAG:
                cell_value = None
                
                string_value_attribute = attrib_get_func(STRING_VALUE_ATTRIBUTE)

                # Try finding the cell value through its' elements
                if not string_value_attribute is None:
                    cell_value = string_value_attribute
                else:
                    value_attribute = attrib_get_func(VALUE_ATTRIBUTE)

                    if not value_attribute is None:
                        cell_value = value_attribute

                if (not cell_value is None) and (not child is None):
                    # Exactly one child
                    if second is None:
                        if child.text and len(child) == 0:
                            # The text property of the child is defined and there are no elements in 
                            cell_value = child.text
                        else:
                            # The child element has no direct text node, but it contains children (which might have text)
                            cell_value = "".join(child.itertext())
                    else:
                        # Multiple children
                        cell_value = "".join(child.itertext())

                # Convert the cell value to the type specified in the cell 'value-type' attribute
                if (not cell_value is None) and convert_values:
                    value_type_attribute = attrib_get_func(VALUE_TYPE_ATTRIBUTE)

                    if value_type_attribute in ("float", "currency", "percentage"):
                        cell_value = float(cell_value)
                    elif value_type_attribute == "date":
                        cell_value = datetime.fromisoformat(cell_value)
                    elif cell_value is not None:
                        cell_value = str(cell_value)

                if not cell_value is None:
                    current_row_has_value = True

                # Append cell values to the current row
                column_repeat_amount = int(attrib_get_func(TABLE_NUMBER_COLUMNS_REPEATED_ATTRIBUTE, 1))

                if column_repeat_amount == 1:
                    current_row.append(cell_value)
                else:
                    current_row.extend([cell_value] * column_repeat_amount)
            
            if tag_name in PARSED_ELEMENTS:
                element.clear()

    def parse(self, path: str) -> Iterator[tuple]:
        if not path.endswith('.ods'):
            logger.warning('File does not have the .ods extension')

        with ZipFile(path, mode='r') as zip_stream:
            if self.options.verify_zip:
                bad_file = zip_stream.testzip()

                if bad_file == CONTENT_XML_FILE_NAME:
                    logger.warning('ODS file may be corrupted')

            with zip_stream.open(CONTENT_XML_FILE_NAME) as content_stream:
                yield from self._parse_table_internal(content_stream, self.options.table)