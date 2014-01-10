# debug flag
DEBUG = False

# for handling command lines args
REQUIRED_NUM_ARGS = 2
OPTIONAL_NUM_ARGS = 1
EXCEL_ARG_INDEX = 1
XML_ARG_INDEX = 2
OUTPUT_XML_ARG_INDEX = 3

# XML tags for DDI schema
VAR_TAG = "var"
DESCRIPTION_TAG = "txt"
QUESTION_TAG = "qstn"
LITERAL_QUESTION_TAG = "qstnLit"

# special regex for using lxml
SEARCH_FULL_TREE = ".//"
IGNORE_NAMESPACE = "{*}"

# prefix for output XML file
OUTPUT_XML_PREFIX = "output_"

OLD_EXCEL_EXTENSION = "xls"

SURVEY_SHEET_NAME = "survey"

# keywords for Excel spreedsheet for XLSForm
DESCRIPTION_COLUMN_NAME = "name"
LITERAL_QUESTION_COLUMN_NAME = "label::English"

NULL_INDEX = -1

# debug names
EXAMPLE_FILES = "examples/"
DEBUG_XML_NAME = "example.xml"
DEBUG_XML = EXAMPLE_FILES + DEBUG_XML_NAME
DEBUG_OUTPUT_XML_NAME = OUTPUT_XML_PREFIX + DEBUG_XML_NAME
DEBUG_EXCEL_FILE_NAME = "example.xlsx"
DEBUG_EXCEL = EXAMPLE_FILES + DEBUG_EXCEL_FILE_NAME
