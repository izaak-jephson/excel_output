library(here)
library(readxl)
library(tidyverse)
library(openxlsx)
library(lubridate)

# Set publication date manually
publication_date <- dmy("18 June 2024")

source(here("R","read_meta.R"))
source(here("R","arrange_tables.R"))
source(here("R","output_tables.R"))
source(here("R","format_suppression.R"))


# Create outputs folder if needed
if (!(dir.exists("outputs"))) {
  dir.create("outputs")
  print("output folder created")
} else {
  print("output folder exists in project")
}

# Read metadata.xlsx
# This file contains information about the desired layout of the tables as well
# as the sheet and table titles, and publication notes. 

metadata_filepath <- here("inputs", "table_metadata.xlsx")

metadata <- read_metadata(metadata_filepath)


# Import data tables to be written to excel 
# This will be specific to each publication.

load(here("inputs","adp_2024_04.R"))

# data_tables must be a named list of publication tables, where names are 
# consistent with those in metadata.xlsx file. The tables in the list should 
# have numeric or % columns formatted as numeric. Values to be suppressed should
# be represented as -1, which is turned to "[c]" in the output tables. 
# format_suppression.R contains helper functions to convert tables to correct 
# format. 

data_tables <- tables_to_numeric(data) # may not be needed if tables in correct format.



# Set up tables according to instructions in metadata.xlsx
table_layout <- arrange_tables(metadata$sheet_layout, 
                               metadata$tables, 
                               metadata$table_list, 
                               metadata$notes, 
                               metadata$options, 
                               data_tables)

# Write output tables to excel


make_output_tables(table_layout = table_layout, 
                   notes_list = metadata$notes, 
                   publication_date = publication_date, 
                   contents_title = metadata$contents$contents_title, 
                   workbook_filename = "tables.xlsx")

# If multiple excel outputs are required (e.g. for raw, rounded and suppressed 
# tables), then repeat final two commands for each set of tables,
# changing data_tables and workbook_filename as appropriate.
