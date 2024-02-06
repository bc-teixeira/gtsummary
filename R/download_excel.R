#' Convert a list of gtsummary tables or data frames into a single excel file with multiple tables
#'
#' Function converts a list of gtsummary tables or data frames into a single
#'  excel file with multiple tables. The function is a wrapper around the
#'  as_hux_xlsx function from the gtsummary package and the write.xlsx function
#'  from the openxlsx package.
#'


export_objects_to_excel <- function(x) {
  # Initialize two lists: 'data_frames' to store data frames and 'gt_tables' to
  # store gt tables.
  .assert_class(x, "list")
  assert_package("huxtable", "as_hux_table()")
  assert_package("openxlsx")
  data_frames <- list()
  gt_tables <- list()

  # Iterate over each object in the 'x' list.
  for (name in names(x)) {
    object <- x[[name]]
    # If the object is a data frame, add it to the 'data_frames' list.
    if (is.data.frame(object)) {
      data_frames[[name]] <- object
      # If the object is a gt table (using 'tbl_summary'), add it to the
      #'gt_tables' list.
    } else if (inherits(object, "tbl_summary")) {
      gt_tables[[name]] <- object
    }
  }

  # Prepare a variable to hold the path to the Excel file that will be created.
  excel_path <- NULL
  # If there are any data frames to write, create a temporary Excel file and
  #write them to it.
  if (length(data_frames) > 0) {
    excel_path <- tempfile(fileext = ".xlsx")
    write.xlsx(data_frames, file = excel_path)
  }

  # Process each gt table and merge them into the Excel file.
  for (name in names(gt_tables)) {
    # Create a temporary Excel file for the gt table.
    gt_table_path <- tempfile(fileext = ".xlsx")
    # Convert the gt table to an Excel format and save it to the temporary file.
    as_hux_xlsx(gt_tables[[name]], file = gt_table_path)
    # Load the workbook from the temporary file.
    wb <- loadWorkbook(gt_table_path)
    # Rename the first worksheet with the gt table's name.
    renameWorksheet(wb, 1, name)
    # Save the changes to the temporary file.
    saveWorkbook(wb, gt_table_path, overwrite = TRUE)

    # If the main Excel file path is null, set it to the current gt table path.
    if (is.null(excel_path)) {
      excel_path <- gt_table_path
    } else {
      # Otherwise, load the main Excel workbook.
      wb <- loadWorkbook(excel_path)
      # Get the sheet names from the temporary gt table file.
      sheets_to_add <- getSheetNames(gt_table_path)
      # Iterate over each sheet in the gt table workbook.
      for (sheet in sheets_to_add) {
        # Read the gt table workbook.
        gtwb <- readWorkbook(gt_table_path, sheet)
        # Add a new worksheet to the main workbook with the gt table's name.
        addWorksheet(wb = wb, sheetName = name)
        # Write the data from the gt table workbook to the new worksheet.
        writeData(wb, name, gtwb)
      }
      # Save the updated main workbook with the added gt table data.
      saveWorkbook(wb, excel_path, overwrite = TRUE)
    }
  }

  # Return the path of the final Excel file containing all data frames and
  #gt tables.
  return(excel_path)
}
