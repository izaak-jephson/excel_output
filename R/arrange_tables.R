arrange_tables <- function(sheet_layout, 
                           tables, 
                           table_list, 
                           notes, 
                           options, 
                           data_tables){
  
  sheet_layout %>%
    select (-sheet_number) %>% 
    mutate(tables = str_split(tables,",")) %>% 
    mutate(notes = str_split(notes,",") %>% map(.,as.numeric)) %>% 
    mutate(
      sheet_tables = 
        map(tables, 
            ~ tibble(
              table_name = .x,
              table = data_tables[.x],
              table_title = table_list %>% 
                filter(table_name %in% .x) %>% 
                pull(table_title)
            )
        )
      
    ) %>%
    mutate(
      sheet_tables =
        map(
          sheet_tables,
          ~ .x %>%
            mutate(
              n_rows = case_when(nrow(.) > 1 ~ map(table, ~.x %>% nrow) %>% as.numeric() + 1 + options$padding_rows_multi,
                                 nrow(.) == 1 ~ map(table, ~.x %>% nrow) %>% as.numeric() + 1 + options$padding_rows_single),
              n_cols = map(table, length) %>% as.numeric(),
              end_row = cumsum(n_rows),
              notes_start = end_row,
              start_row = case_when(nrow(.) > 1 ~ end_row - n_rows + options$padding_rows_multi,
                                    nrow(.) == 1 ~ end_row - n_rows + options$padding_rows_single),
              start_col = 1,
              end_col = n_cols,
            )
        )
    ) %>%
    mutate(
      n_tables = map(sheet_tables, nrow) %>% as.numeric(),
      notes_start = map(sheet_tables, ~ .x %>%
                          select(n_rows) %>%
                          map(~ sum(.x)) %>%
                          as.numeric())
    )
  
  
}
