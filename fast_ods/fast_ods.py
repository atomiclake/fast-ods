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
    verify_zip: bool = True

#------------------
# Parser
#------------------

class ODSParser():
    def __init__(self, options: ODSParserOptions | None = None):
        self.options = options or ODSParserOptions()
    
    def _parse_core(self, ods_contents: IOBase, target_table: str | int) -> Iterator[tuple]:
        seen_tables = 0
        seeking_target_table = True
        current_row = []

        physical_row_count = 0
        expanded_row_count = 0

        # Cache options locally (hot path optimization)
        convert_values = self.options.convert_values
        skip_n_rows = self.options.skip_n_rows
        take_n_rows = self.options.take_n_rows

        for event, element in ET.iterparse(ods_contents, events=[TAG_START_EVENT, TAG_END_EVENT]):
            # Locate target table
            if seeking_target_table and event == TAG_START_EVENT:
                if element.tag != TABLE_TABLE_TAG:
                    continue

                attrib = element.attrib
                table_name = attrib.get(TABLE_NAME_ATTRIBUTE)

                if isinstance(target_table, int):
                    seeking_target_table = target_table != seen_tables
                elif isinstance(target_table, str) and table_name:
                    seeking_target_table = target_table != table_name

                seen_tables += 1

                if seeking_target_table:
                    element.clear()

                continue

            if event == TAG_START_EVENT:
                continue

            tag = element.tag

            if tag not in PARSED_ELEMENTS:
                element.clear()
                continue

            # ----------------------
            # Handle ROW
            # ----------------------
            if tag == TABLE_ROW_TAG:
                attrib = element.attrib
                row_repeat = int(attrib.get(TABLE_NUMBER_ROWS_REPEATED_ATTRIBUTE, 1))

                for _ in range(row_repeat):
                    physical_row_count += 1

                    if skip_n_rows and physical_row_count <= skip_n_rows:
                        continue

                    yield tuple(current_row)

                expanded_row_count += row_repeat
                current_row = []

                element.clear()

                if take_n_rows and expanded_row_count >= take_n_rows:
                    return

            # ----------------------
            # Handle CELL (HOT PATH)
            # ----------------------
            elif tag == TABLE_CELL_TAG:
                attrib = element.attrib

                # --- Extract value (FAST PATH FIRST) ---
                str_value = None

                xml_string = attrib.get(STRING_VALUE_ATTRIBUTE)

                if xml_string is not None:
                    str_value = xml_string
                else:
                    xml_value = attrib.get(VALUE_ATTRIBUTE)
                    if xml_value is not None:
                        str_value = xml_value
                    else:
                        # fallback to text:p
                        for child in element:
                            if child.tag == TEXT_P_TAG:
                                text = child.text
                                if text is not None:
                                    str_value = text
                                else:
                                    str_value = "".join(child.itertext())
                                break

                # --- Optional conversion (fast path) ---
                if convert_values:
                    value_type = attrib.get(VALUE_TYPE_ATTRIBUTE)

                    if value_type == "float" or value_type in ("currency", "percentage"):
                        try:
                            str_value = float(str_value)
                        except:
                            pass
                    elif value_type == "date":
                        try:
                            str_value = datetime.fromisoformat(str_value)
                        except:
                            pass
                    elif value_type == "string":
                        if str_value is not None:
                            str_value = str(str_value)

                # --- Handle column repetition ---
                repeat = int(attrib.get(TABLE_NUMBER_COLUMNS_REPEATED_ATTRIBUTE, 1))
                append = current_row.append

                for _ in range(repeat):
                    append(str_value)

                element.clear()

    def parse(self, path: str) -> Iterator[list]:
        if not path.endswith('.ods'):
            logger.warning('File does not have the .ods extension')

        with ZipFile(path, mode='r') as zip_stream:
            if self.options.verify_zip:
                bad_file = zip_stream.testzip()

                if bad_file == CONTENT_XML_FILE_NAME:
                    logger.warning('ODS file may be corrupted')

            with zip_stream.open(CONTENT_XML_FILE_NAME) as content_stream:
                yield from self._parse_core(content_stream, self.options.table)