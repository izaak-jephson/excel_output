### Create tibble of sheet titles based on publication date

make_contents_table <- function(publication_date, table_layout) {
  contents <- table_layout %>% 
    select(sheet_name, sheet_title) %>% 
    rename(Sheet = sheet_name, Description = sheet_title) %>% 
    add_row(Sheet = "Notes", Description = "List of notes") %>% 
    slice(order(Sheet != "Notes"))
    
  contents
}

### Add contents sheet to give given workbook based on tibble of sheet titles

add_contents_sheet <- function(wb, contents, contents_title) {
  contents_table <-
    contents %>%
    slice(-1) %>% 
    mutate(`Table Number` = makeHyperlinkString(
      sheet = Sheet,
      row = 1,
      col = 1,
      text = Sheet
    ))
  
  class(contents_table$`Table Number`) <- c(class(contents_table$`Table Number`), "formula")
  
  modifyBaseFont(wb,
                 fontSize = 12,
                 fontColour = "black",
                 fontName = "Roboto"
  )
  
  addWorksheet(wb, "Contents")
  
  setColWidths(
    wb,
    "Contents",
    cols = 1:2,
    width = c(20, "auto"),
    ignoreMergedCells = TRUE
  )
  
  # Make table headings bold
  addStyle(
    wb,
    "Contents",
    rows = 1,
    cols = 1,
    style = createStyle(
      fontSize = 15,
      wrapText = FALSE,
      textDecoration = "bold"
    )
  )
  
  showGridLines(wb,
                "Contents",
                showGridLines = FALSE
  )
  
  # Format table numbers as hyperlinks
  addStyle(
    wb,
    "Contents",
    rows = 4:(nrow(contents) + 2),
    cols = 1,
    style = createStyle(fontColour = "blue", textDecoration = "underline")
  )
  
  # Add title
  writeData(wb,
            "Contents",
            x = c(paste(contents_title, "to", format(publication_date %m-% months(2), "%B %Y")), "Table of Contents")
  )
  
  # Add contents
  writeDataTable(
    wb,
    "Contents",
    startRow = 3,
    x = contents_table %>% select(`Table Number`, Description),
    tableStyle = "TableStyleLight1",
    withFilter = openxlsx_getOp("withFilter", FALSE)
  )
  
  return(wb)
}


### Add notes sheet to given workbook

add_notes_sheet <- function(wb, contents, notes_list) {
  addWorksheet(wb, sheetName = "Notes")
  
  notes_list <- notes_list %>% 
    rename(`Note number` = note_number,
           `Note text` = note_text)
  
  # Format notes table
  setColWidths(
    wb,
    "Notes",
    width = 12,
    cols = 1,
    ignoreMergedCells = TRUE
  )
  
  setColWidths(
    wb,
    "Notes",
    width = 150,
    cols = 2,
    ignoreMergedCells = TRUE
  )
  
  addStyle(
    wb,
    "Notes",
    rows = 1,
    cols = 1,
    style = createStyle(
      fontSize = 15,
      wrapText = FALSE,
      textDecoration = "bold"
    )
  )
  
  addStyle(
    wb,
    "Notes",
    rows = c(1:nrow(notes_list) + 4),
    cols = 2,
    style = createStyle(wrapText = TRUE)
  )
  
  # Add heading text
  writeData(wb, "Notes", x = "List of notes")
  writeData(wb, "Notes", startRow = 2, x = "This worksheet displays 1 table")
  writeData(wb, "Notes",
            startRow = 3,
            x = "The notes within this table are referred to in other worksheets of this workbook."
  )
  
  # Add notes table
  writeDataTable(
    wb,
    "Notes",
    startRow = 4,
    x = notes_list,
    tableStyle = "TableStyleLight1",
    withFilter = openxlsx_getOp("withFilter", FALSE)
  )
  
  return(wb)
}


format_columns <- function(wb, sheet_name, table, column, start_row, end_row) {
  # Format numbers with commas
  if (is_integer(table[[column]])) {
    addStyle(wb, sheet_name,
             rows = start_row:end_row,
             cols = column,
             style = createStyle(numFmt = "#,##0;-;0", halign = "right"),
             gridExpand = TRUE
    )
  
  
    
    # Format % columns
  } else if (str_detect(colnames(table[column]), "[Pp]ercent|[Pp]ercentage|[Pp]roportion")) {
    addStyle(wb, sheet_name,
             rows = start_row:end_row,
             cols = column,
             style = createStyle(numFmt = "0%;-;0%", halign = "right"),
             gridExpand = TRUE
    )
  
  }
}

format_rows <- function(wb, sheet_name, table, table_row, sheet_row, start_col, end_col) {
  # Make total rows bold
  if (table[[table_row, 1]] == "Total" |
      str_starts(table[[table_row, 1]],"Financial")) {
    addStyle(wb, sheet_name,
             rows = sheet_row,
             cols = start_col:end_col,
             style = createStyle(textDecoration = "bold"),
             gridExpand = TRUE,
             stack = TRUE
    )
  }
}



# Add data tables to existing worksheet
add_data_tables <- function(wb, sheet_name, sheet_title, tables, header_rows) {
  tables %>%
    select(table, start_row, end_row, table_title, table_name) %>%
    pmap(function(table, start_row, end_row, table_title, table_name) {
      add_data_table(wb, sheet_name, sheet_title, table, start_row, end_row, header_rows, table_name)
      if(nrow(tables) > 1){ 
        writeData(wb, sheet_name,
                  x = table_title,
                  startRow = start_row + header_rows - 1,
                  startCol = 1
        )
        addStyle(wb,
                 sheet_name,
                 rows = start_row + header_rows - 1,
                 cols = 1,
                 style = createStyle(
                   wrapText = FALSE,
                   textDecoration = "bold"
                 )
        )
      }
    })
}

# Create table sheet with header wording, tables and notes
add_data_sheet <- function(wb, sheet_name, sheet_title, sheet_tables, notes_list, note_mapping, n_tables, notes_start) {
  # Define start and end points for tables
  header_rows <- 5
  
  # end_row <- header_rows + nrow(table[[1]]) + 1
  
  # Add data table sheet
  addWorksheet(wb, sheetName = sheet_name)
  
  # Format header section
  showGridLines(wb, sheet_name, showGridLines = FALSE)
  
  addStyle(wb,
           sheet_name,
           rows = 1,
           cols = 1,
           style = createStyle(
             fontSize = 15,
             wrapText = FALSE,
             textDecoration = "bold"
           )
  )
  
  # Add title and header text
  writeData(wb, sheet_name,
            x = paste(
              sheet_title,
              paste0("[note ", note_mapping, "]",
                     collapse = " "
              )
            )
  )
  
  writeData(wb, sheet_name,
            x = paste(
              "This worksheet contains",
              n_tables,
              if_else(n_tables == 1, "table.", "tables.")
            ),
            startRow = 2,
            startCol = 1
  )
  
  writeData(wb, sheet_name,
            x = paste0("Banded rows are used in ",if_else(n_tables == 1, "this table", "these tables"),". To remove them, highlight the table, go to the Design tab and uncheck the banded rows box."),
            startRow = 3,
            startCol = 1
  )
  
  # writeData(wb, sheet_name,
  #           x = paste0("Data bars are used in ",if_else(n_tables == 1, "this table", "these tables"),". To remove these, select the table, go to the Home tab, click on Conditional Formatting and select Clear Rules from This Table."),
  #           startRow = 4,
  #           startCol = 1
  # )
  
  writeData(wb, sheet_name,
            x = paste0(
              "Notes are located below the",
              if_else(n_tables == 1, " table ", " tables "),
              "beginning in cell A",
              notes_start + header_rows,
              " and in the notes sheet of this document."
            ),
            startRow = 4,
            startCol = 1
  )
  
  

    
   writeData(wb, sheet_name,
              x = "Some rows between tables are left blank in this sheet to improve readability.",
              startRow = 5,
              startCol = 1
  )
    
    
    
    # Add additional notes to header
    
    
   
  
  add_data_tables(wb, sheet_name, sheet_title, sheet_tables, header_rows)
  
  # Add notes
  writeData(wb, sheet_name,
            startRow = notes_start + header_rows,
            x = tibble(note = paste0("[note ", note_mapping, "]"),
                       note_text = notes_list$note_text[note_mapping]),
            colNames = FALSE
            )

  
  
  return(wb)
}

add_data_table <- function(wb, sheet_name, sheet_title, table, start_row, end_row, header_rows, table_name) {
  # Format data table
  addStyle(wb,
           sheet_name,
           rows = header_rows + start_row:header_rows + end_row - 1,
           cols = 1,
           style = createStyle(wrapText = FALSE)
  )
  
  
  # Add data table
  writeDataTable(wb, sheet_name,
                 startRow = header_rows + start_row,
                 tableName = str_replace_all(table_name, " ","_"),
                 x = table,
                 colNames = TRUE,
                 rowNames = FALSE,
                 keepNA = TRUE,
                 na.string = "n/a",
                 tableStyle = "TableStyleLight1", headerStyle = createStyle(wrapText = TRUE),
                 stack = TRUE,
                 withFilter = openxlsx_getOp("withFilter", FALSE)
  )
  
  # Format table headers
  addStyle(wb, sheet_name,
           rows = header_rows + start_row,
           cols = 1:length(table),
           style = createStyle(halign = "center", wrapText = TRUE)
  )
  
  # Format columns
  setColWidths(
    wb,
    sheet_name,
    cols = 1:length(table),
    width = 20,
    ignoreMergedCells = TRUE
  )
  
  walk(
    1:length(table),
    ~ format_columns(wb, sheet_name, table, .x, header_rows + 1 + start_row, header_rows + end_row - 1)
  )
  
  walk(
    1:nrow(table),
    ~ format_rows(wb, sheet_name, table, .x, start_row + .x + header_rows, 1, length(table))
  )
  
  negative_to_c <- function (wb, sheet_name, table, column, row, start_row, end_row){
    if (table[[column]][[row]] == -1 & (!is.na(table[[column]][[row]]))){
      writeData(wb,
                sheet_name,
                x = "[c]",
                startCol = column,
                startRow = row + start_row + header_rows
      )
    }
  }
  
  walk(1:length(table),
       function(y) walk(1:nrow(table),
        function(x) negative_to_c(wb, sheet_name, table, y, x, start_row, end_row)))

  
  return(wb)
}



add_data_sheets <- function(wb, table_layout, notes_list) {
  walk(
    1:nrow(table_layout),
    ~ add_data_sheet(
      wb,
      table_layout[[1]][[.x]],
      table_layout[[2]][[.x]],
      table_layout[[5]][[.x]],
      notes_list,
      table_layout[[4]][[.x]],
      table_layout[[6]][[.x]],
      table_layout[[7]][[.x]]
    )
  )
  
  return(wb)
}

### Specific changes to individual sheet formatting. 
### Could be made more intelligently automated in future or for larger publications.

tweak_formatting <- function(wb) {
  
  return(wb)
}

### Create excel tables

make_output_tables <- function(table_layout, notes_list, publication_date, contents_title, workbook_filename) {
  
  wb <- createWorkbook()
  
  modifyBaseFont(wb, fontSize = 12, fontColour = "black", fontName = "Roboto")
  
  contents <- make_contents_table(publication_date, table_layout)
  
  wb <- add_contents_sheet(wb, contents, contents_title)
  
  wb <- add_notes_sheet(wb, contents, notes_list)
  
  wb <- add_data_sheets(wb, table_layout, notes_list)
  
  wb <- tweak_formatting(wb)
  
  saveWorkbook(wb, here("outputs",workbook_filename), overwrite = TRUE)
  
  here("outputs",workbook_filename)
}
