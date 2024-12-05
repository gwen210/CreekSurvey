
# Clear all variables and close any plots before running the app
rm(list = ls())  # Remove all objects from the environment
# Safely check if a graphical device is open before calling dev.off()
if (dev.cur() > 1) {
  dev.off()  # Close active graphical device
}
graphics.off()    # Close all open graphics windows
gc()              # Run garbage collection to free up memory


# Load and install required packages. Checks to see if library is already downloaded. if not it downloads it then installs the libraries
loadandinstall <- function(mypkg) {
  if (!is.element(mypkg, installed.packages()[, 1])) {
    install.packages(mypkg)
  }
  library(mypkg, character.only = TRUE)
}

# Load necessary libraries using loadandinstall
loadandinstall("shiny")
loadandinstall("readxl")
loadandinstall("dplyr")
loadandinstall("tidyr")
loadandinstall("ggplot2")
loadandinstall("plotly")
loadandinstall("ggforce")




# Function to safely extract the instrument elevation values using row and column index search
extract_inst_ele <- function(header, search_string) {
  # Use apply to find the row and column that contains the search string
  row_index <- apply(header, 1, function(row) any(grepl(search_string, row, ignore.case = TRUE)))
  col_index <- apply(header, 2, function(col) any(grepl(search_string, col, ignore.case = TRUE)))
  # Ensure that both row and column were found
  if (sum(row_index) > 0 & sum(col_index) > 0) {
    row_number <- which(row_index)  # Row where the phrase appears
    col_number <- which(col_index)  # Column where the phrase appears
    # Extract the value, ensuring that you get the correct column (col_number - 1)
    # This assumes the value is in the column before the match
    inst_ele <- as.numeric(header[[row_number, col_number - 1]])
    return(inst_ele)
  } else {
    # If not found, return NA
    return(NA)
  }
}


######################################################################3
#Set up the GUI screen. This is the code that specifies the graphical interface.
ui <- fluidPage(
  titlePanel("Interactive Data Upload and Processing"),
  sidebarLayout(
    #Stuff on left side panel
    sidebarPanel(
      # File input for uploading the Excel file
      fileInput("file_input", "Upload Excel File", accept = ".xlsx"),
      # Dropdown for selecting sheet name
      uiOutput("sheet_select_ui"),
      # Text input for selecting shot numbers to skip (comma-separated, e.g., "1, 5, 7")
      textInput("shots_to_skip", "Enter the shot numbers to skip (comma-separated):", value = ""),
      # Text input for selecting shot numbers to skip (comma-separated, e.g., "1, 5, 7"),
      numericInput("extra_distance", 
                   label = "Enter additional distance to add to Corrected Distance (in feet):", 
                   value = 0, 
                   step = 1),
      # Dropdown for selecting the x-axis variable
      selectInput("distance_type", 
                  label = "Select the distance type to plot:",
                  choices = c("Corrected Distance" = "Corrected_Distance", 
                              "Cumulative Distance" = "Cumulative_Distance"),
                  selected = "Corrected_Distance"),  # Default choice
      # Checkbox group to select columns for download
      checkboxGroupInput("columns_to_download", 
                         "Select columns to download", 
                         choices = NULL,  # This will be populated dynamically
                         selected = NULL),  # Default to no columns selected
      
      # Download Button for CSV export
      downloadButton("download_data", "Download Data as CSV")),
    #Stuff to put on the main panel
    mainPanel(
      # Instructions text for status or instructions
      textOutput("instructions"),
      br(),
      # Plot Output placed in the main panel (centered area)
      plotlyOutput("elevation_plot"),  # First plot for elevation
      br(),
      # Add new plot for X vs Y
      plotlyOutput("xy_plot"),  # Second plot for X vs Y
      br(),
      # Table Output for displaying the cleaned data
      tableOutput("dat_table"),
      
    )
  )
)


#######################################################################
# Server functions for shiny app
#Start the shiny app This part has the R code that does the analysis, makes plots, makes table, etc.

server <- function(input, output, session) {
  
  # Reactive expression to load and process data from uploaded file
  #upload the data based on the input file name and the sheet name from the dropdown menu
  dat_raw <- reactive({
    req(input$file_input, input$sheet_name)  # Ensure file and sheet are selected
    # Load the Excel file
    dat_path <- input$file_input$datapath
    
    # Read the header (first 7 rows) of the selected sheet
    #This contains info like the height of the instrument and the pins
    header <- read_xlsx(dat_path, n_max = 7, sheet = input$sheet_name, trim_ws = TRUE, col_names = FALSE)
    
    # Check if the header is empty
    if (nrow(header) == 0 || all(is.na(header))) {
      # If the sheet is empty or contains only NA values, return an empty data frame or a message
      return(NULL)  # Or you can return a message or empty data frame here as per your needs
    }
    # Return the raw header data before any manipulation (first 7 rows)
    return(header)
  })
  
  # Reactive output for sheet names (based on uploaded file)
  output$sheet_select_ui <- renderUI({
    req(input$file_input)  # Ensure file is uploaded
    
    # Read sheet names from the Excel file
    sheets <- excel_sheets(input$file_input$datapath)
    
    # Create a dropdown UI for selecting the sheet
    selectInput("sheet_name", "Select a sheet:", choices = sheets)
  })
  
  # Show the header (first 7 rows) of the raw data (before any processing)
  output$raw_header_preview <- renderTable({
    req(dat_raw())  # Ensure raw header data is available before displaying
    
    # If dat_raw is NULL, show a message instead of a table
    if (is.null(dat_raw())) {
      return(data.frame("Message" = "The selected sheet is empty or contains no valid data"))
    }
    dat_raw()
  }) 
  
  # Reactive expression to load and process data from uploaded file
  dat_clean <- reactive({
    req(input$file_input)  # Ensure the file is uploaded
    # Load the Excel file and check for empty sheet
    dat_path <- input$file_input$datapath
    #this makes it read the first 7 rows, which typically makes up the header for the file. Might need to be adjusted
    #there are no column names. 
    header <- read_xlsx(dat_path, n_max = 7, sheet = input$sheet_name, trim_ws = TRUE, col_names = FALSE)
    
    # Check if the header is empty and return NULL
    if (nrow(header) == 0 || all(is.na(header))) {
      return(NULL)  # Return NULL or an appropriate empty data frame
    }
    
    # Safely extract instrument elevation values, assuming there are 6 or less instrument movements
    #can adjust if needed to include more instruments
    #might want to adust so the 3rd and 4th characters can be anything.. and might be able to make it a for loop
    inst_ele_1 <- extract_inst_ele(header, "1st instrument ele")
    inst_ele_2 <- extract_inst_ele(header, "2nd instrument ele")
    inst_ele_3 <- extract_inst_ele(header, "3rd instrument ele")
    inst_ele_4 <- extract_inst_ele(header, "4th instrument ele")
    inst_ele_5 <- extract_inst_ele(header, "5th instrument ele")
    inst_ele_6 <- extract_inst_ele(header, "6th instrument ele")
    
    # Read the actual data
    dat1 <- read_xlsx(dat_path, skip = 7, sheet = input$sheet_name, trim_ws = TRUE)
    
    colnames(dat1) <- gsub("\\.\\.\\..*", "", colnames(dat1))  # Clean column names
    Description_column_number <- which(tolower(colnames(dat1)) == tolower("Description"))
    dat1 =(dat1[1:Description_column_number[1]])
    dat1 <- dat1[, !colnames(dat1) %in% ""]
    ################
    dat1 <- dat1 %>% 
      mutate(OpticalDistance = (U.S. - L.S.) * 100,
             X = OpticalDistance * sin(Bearing * pi / 180),
             Y = OpticalDistance * cos(Bearing * pi / 180)) %>% 
      mutate(Xlag = lag(X),
             Ylag = lag(Y))
    
    
    ################
    Intra_df <- dat1 %>%
      filter(grepl("Instr", Description, ignore.case = TRUE)) %>%
      select(Shot, Description, X,Y, Xlag, Ylag) %>%
      mutate(RowNumber = row_number(),
             Instrument = paste0("inst_ele_", RowNumber+1),
             Xlagged = (X- Xlag),
             Ylagged = (Y- Ylag)) # Append 'inst_ele_' to RowNumber
    
    dat1 <- dat1 %>%
      left_join(Intra_df %>% select(Shot, Instrument , Xlagged, Ylagged), by = "Shot")
    
    dat1 <- dat1 %>%
      mutate(
        Instrument = case_when(
          grepl("inst_ele_2", Description) ~ "inst_ele_2",  # Assign inst_ele_2 where found in Description
          grepl("inst_ele_3", Description) ~ "inst_ele_3",  # Assign inst_ele_3 where found in Description
          grepl("inst_ele_4", Description) ~ "inst_ele_4",  # Assign inst_ele_4 where found in Description
          grepl("inst_ele_5", Description) ~ "inst_ele_5",  # Assign inst_ele_5 where found in Description
          grepl("inst_ele_6", Description) ~ "inst_ele_6",  # Assign inst_ele_6 where found in Description
          TRUE ~ Instrument  # Keep the current instrument value if already assigned
        )) %>%
      # Use fill() to propagate instrument labels downwards
      fill(Instrument, .direction = "down")
    dat1$Instrument[is.na(dat1$Instrument)] <- "inst_ele_1"
    dat1$Inst_Ele <- sapply(dat1$Instrument, function(x) get(x))
    
    dat1 <- dat1 %>%
      tidyr::fill(Inst_Ele, .direction = "down")  # Fill missing values down
    
    
    dat1 <- dat1 %>%
      mutate(Xlagged = ifelse(is.na(Xlagged), NA, Xlagged),
             Ylagged = ifelse(is.na(Ylagged), NA, Ylagged)) %>%  # Ensure Xlaged is a column in dat1
      tidyr::fill(Xlagged, .direction = "down")%>%  # Ensure Xlaged is a column in dat1
      tidyr::fill(Ylagged, .direction = "down")
    
    dat1$Xlagged[is.na(dat1$Xlagged)] <- "0"
    dat1$Ylagged[is.na(dat1$Ylagged)] <- "0"    
    dat1<- dat1 %>%
      mutate(Xlagged = ifelse(is.na(Xlagged), 0, Xlagged),
             Ylagged = ifelse(is.na(Ylagged), 0, Ylagged))
    
    dat1 <- dat1 %>%
      mutate(Xlagged = ifelse(Xlagged != as.numeric(lag(Xlagged)),
                              as.numeric(Xlagged) + as.numeric(lag(Xlagged)), NA))%>%
      mutate(Ylagged = ifelse(Ylagged != as.numeric(lag(Ylagged)),
                              as.numeric(Ylagged) + as.numeric(lag(Ylagged)),NA))%>%  # Ensure Xlaged is a column in dat1
      tidyr::fill(Xlagged, .direction = "down")%>%  # Ensure Xlaged is a column in dat1
      tidyr::fill(Ylagged, .direction = "down")
    dat1<- dat1 %>%
      mutate(Xlagged = ifelse(is.na(Xlagged), 0, Xlagged),
             Ylagged = ifelse(is.na(Ylagged), 0, Ylagged))
    
    
    # Split the input into a list of shot numbers to skip
    shots_to_skip <- as.integer(unlist(strsplit(input$shots_to_skip, ",")))
    shots_to_skip <- na.omit(shots_to_skip)  # Remove NA values in case of bad input
    
    # Filter out the selected shot numbers to skip
    if (length(shots_to_skip) > 0) {
      dat1 <- dat1 %>%
        filter(!Shot %in% shots_to_skip)  # Keep only rows where Shot is not in the skipped list
    }
    
    # Add QC columns, flags, and calculations
    dat1$S_dif <- round((dat1$U.S. - dat1$L.S.) / 2, 3)
    dat1$US_check <- round(dat1$U.S. - dat1$Level, 3)
    dat1$LS_check <- round(dat1$Level - dat1$L.S., 3)
    
    dat1 <- dat1 %>%
      mutate(
        US_flag = ifelse(near(US_check, S_dif, tol = 0.15), "1", "0"),
        LS_flag = ifelse(near(LS_check, S_dif, tol = 0.15), "1", "0"),
        Elevation = Inst_Ele - Level,
        Bearinglag = lag(Bearing),
        OpticalDistancelag = lag(OpticalDistance)
      )
    
    # Remove rows where the 'Description' column contains the word "Pin"
    #dat1 <- dat1[!grepl("pin", dat1$Description, ignore.case = TRUE), ]
    
    
    # Filter rows where Description contains "Thalweg"
    dat_pin <- dat1 %>%
      filter(
        grepl("backshot", Description, ignore.case = TRUE) |  # Keep rows with "backshot"
          !(grepl("Thalweg|Start of survey", Description, ignore.case = TRUE))  # Exclude rows with "Thalweg" or "Start of survey"
      )
    
    # Filter rows where Description contains "Thalweg"
    dat <- dat1 %>%
      filter(grepl("Thalweg|Start of survey", Description, ignore.case = TRUE))  # Keep only rows where Description contains "Thalweg"
    
    # Filter rows where US_flag is 1
    dat <- dat %>%
      filter(US_flag == "1")%>%  # Keep only rows where US_flag is 1
      filter(!grepl("Backshot", Description, ignore.case = TRUE)) 
    
    
    # Create an empty data frame to store the result with 'Description' and 'Elevation' columns
    pins_df <- data.frame("Description" = character(0), "Elevation" = numeric(0))
    
    # Loop through the header matrix to search for "Left bank pin"
    for (i in 1:nrow(header)) {
      for (j in 1:ncol(header)) {
        # Check if the phrase "Left bank pin" is present in the current cell
        if (grepl("Left bank pin", header[i, j], ignore.case = TRUE)) {
          # Check if the previous column (j-1) has a valid elevation value
          if (!is.na(header[i, j - 1])) {
            elevation_value <- as.numeric(header[i, j - 1])  # Adjust if necessary based on your data structure
            Description_value <- as.character(header[i, j])
            # If a match is found, add the 'Description' and 'Elevation' to the pins_df
            pins_df <- rbind(pins_df, data.frame("Description" = Description_value, "Elevation" = elevation_value))
          }
        }
      }
    }
    
    # Loop through the header matrix to search for "Right bank pin"
    for (i in 1:nrow(header)) {
      for (j in 1:ncol(header)) {
        # Check if the phrase "Right bank pin" is present in the current cell
        if (grepl("Right bank pin", header[i, j], ignore.case = TRUE)) {
          # Check if the previous column (j-1) has a valid elevation value
          if (!is.na(header[i, j - 1])) {
            elevation_value <- as.numeric(header[i, j - 1])  # Adjust if necessary based on your data structure
            Description_value <- as.character(header[i, j])
            # If a match is found, add the 'Description' and 'Elevation' to the pins_df
            pins_df <- rbind(pins_df, data.frame("Description" = Description_value, "Elevation" = elevation_value))
          }
        }
      }
    }
    colnames(pins_df) <- c("Description", "Elevation")
    
    
    # Replace Elevation based on a condition
    # Define a threshold for similarity (e.g., values between 0 and 1, where 1 is an exact match)
    # Define a threshold for similarity (e.g., values between 0 and 1, where 1 is an exact match)
    similarity_threshold <- 0.6  # Adjust this value depending on how strict you want the match
    
    # Iterate through each row of the larger dataframe (dat_pin)
    for (i in 1:nrow(dat_pin)) {
      # For each Description in the larger dataframe, check for a similar match in pins_df
      for (j in 1:nrow(pins_df)) {
        # Calculate string similarity (Jaro-Winkler similarity)
        similarity <- stringdist::stringdist(dat_pin$Description[i], pins_df$Description[j], method = "jw")
        
        # Ensure similarity is a numeric value and not NA or a vector
        if (length(similarity) == 1 && !is.na(similarity)) {
          # Convert the similarity score to a range [0, 1] where 1 is perfect match
          similarity_score <- 1 - similarity  # Similarity score (higher means more similar)
          
          # If the similarity score is above the threshold, consider it a match
          if (similarity_score >= similarity_threshold) {
            # Replace the Elevation in the larger dataframe with the value from pins_df
            dat_pin$Elevation[i] <- pins_df$Elevation[j]
            break  # Exit the loop once the first sufficient match is found
          }
        } else {
          # Skip this comparison if similarity is NA or not a valid number
          next
        }
      }
    }
    
    
    # Calculate the distance between consecutive shots using the Law of Cosines formula
    dat <- dat %>%
      mutate(
        # A is the previous row's OpticalDistance
        A = lag(OpticalDistance),
        # B is the current row's OpticalDistance
        B = OpticalDistance,
        # C is the difference between current and previous row's Bearing
        C = Bearing - lag(Bearing),
        # Distance_ft is calculated using the Law of Cosines formula
        Distance_ft = sqrt((A^2) + (B^2) - (2 * A * B) * cos(C * pi / 180))
      ) %>%
      # Replace NA values (first row) with 0 for A and C, since there's no previous row for comparison
      replace_na(list(A = NA, C = NA, Distance_ft = 0))
    
    dat <- dat %>%
      mutate(
        C = ifelse(
          Instrument != lag(Instrument), 
          Bearing - Bearinglag ,  # Subtract Bearing from dat1 where Shot == Shot - 1
          C  # Otherwise, keep the current value of C
        ),
        A = ifelse(
          Instrument != lag(Instrument), 
          OpticalDistancelag   ,  # Subtract Bearing from dat1 where Shot == Shot - 1
          A  # Otherwise, keep the current value of C
        ) ,
        
        Distance_ft = ifelse(
          Instrument != lag(Instrument), 
          sqrt((A^2) + (B^2) - (2 * A * B) * cos(C * pi / 180)) ,  # Subtract Bearing from dat1 where Shot == Shot - 1
          Distance_ft ))  # Otherwise, keep the current value of C
    # Replace NA values (first row) with 0 for A and C, since there's no previous row for comparison
    
    
    # Calculate the cumulative distance
    dat <- dat %>%
      replace_na(list(A = NA, C = NA, Distance_ft = 0)) %>% 
      mutate(X_orig = X,
             Y_orig = Y,
             X = as.numeric(X) - as.numeric(Xlagged),
             Y = as.numeric(Y) - as.numeric(Ylagged),
             Cumulative_Distance = cumsum(Distance_ft))
    
    # Calculate the Corrected Distance based on the total length (TL_TW)
    extra_distance <- as.numeric(input$extra_distance)  # Get the additional distance from the user input
    a = dat$X[1] - dat$X[length(dat$X)]
    b = dat$Y[1] - dat$Y[length(dat$Y)]
    c = sqrt(a^2 + b^2)
    TL_TW = c / dat$Cumulative_Distance[length(dat$Cumulative_Distance)]
    
    
    dat <-dat %>%
      replace_na(list(X = 0 )) %>%  
      mutate(Corrected_Distance = Cumulative_Distance * TL_TW + extra_distance,
             TL_TW_rat = TL_TW)
    
    
    return(dat)
  })
  
  # Generate the plot output
  output$elevation_plot <- renderPlotly({
    req(dat_clean())  # Ensure the data is available before plotting
    # Dynamically select the distance variable based on user input
    x_variable <- input$distance_type  # Get the selected distance type
    
    # Create the Plotly plot using the line_shape = "spline" for smooth lines
    p <- plot_ly(data = dat_clean(), 
                 x = as.formula(paste("~", x_variable)),  # Dynamically choose x variable 
                 y = ~Elevation,
                 type = 'scatter', 
                 mode = 'lines+markers',  # Show both lines and markers
                 line = list(shape = 'spline', color = 'black', width = 1), # Smooth spline line
                 marker = list(color = 'red', size = 5),  # Red markers
                 text = ~paste("Shot Number: ", Shot),  # Add shot number to hover text
                 hoverinfo = 'x+y+text'  # Show x, y, and text (shot number) on hover
    ) %>%
      layout(
        title = "Thalweg Profile",
        xaxis = list(title = "Distance (ft)"),  # Set x-axis label
        yaxis = list(title = "Elevation (ft)"),  # Set y-axis label
        showlegend = FALSE  # Hide legend if not needed
      )
    
    p  # Return the plot
  })
  
  # Generate the second plot output (X vs Y plot)
  output$xy_plot <- renderPlotly({
    req(dat_clean())  # Ensure the data is available before plotting
    
    # Create the Plotly scatter plot for X vs Y
    p2 <- plot_ly(data = dat_clean(), 
                  x = ~X, 
                  y = ~Y, 
                  type = 'scatter', 
                  mode = 'lines+markers',  # Show both lines and markers
                  line = list(shape = 'spline', color = 'black', width = 1), # Smooth spline line
                  marker = list(color = 'blue', size = 5),  # Blue markers
                  text = ~paste("Shot Number: ", Shot),  # Add shot number to hover text
                  hoverinfo = 'x+y+text'  # Show x, y, and text (shot number) on hover
    ) %>%
      layout(
        title = "Thalweg Plan View",
        xaxis = list(title = "X"),  # Set x-axis label for X
        yaxis = list(title = "Y"),  # Set y-axis label for Y
        showlegend = FALSE  # Hide legend if not needed
      )
    
    p2  # Return the second plot
  })
  
  
  
  # Update the checkbox options based on the columns of the cleaned data
  observe({
    dat <- dat_clean()  # Get the cleaned data
    colnames_to_select <- colnames(dat)  # Get column names
    updateCheckboxGroupInput(session, "columns_to_download", 
                             choices = colnames_to_select, 
                             selected = colnames_to_select)  # Select all columns by default
  })
  
  # Display the cleaned data table
  output$dat_table <- renderTable({
    dat <- dat_clean()  # Get the cleaned data
    dat  # Return the cleaned data frame
  })
  
  # Download handler for exporting the selected columns to CSV
  output$download_data <- downloadHandler(
    filename = function() {
      paste("cleaned_data_", Sys.Date(), ".csv", sep = "")  # Create a filename with current date
    },
    content = function(file) {
      # Get the selected columns for download
      selected_columns <- input$columns_to_download
      dat <- dat_clean()  # Get the cleaned data
      if (length(selected_columns) > 0) {
        # Write only the selected columns to CSV
        write.csv(dat[, selected_columns, drop = FALSE], file, row.names = FALSE)
      } else {
        # If no columns selected, show a message in the file
        write.csv(data.frame("No columns selected"), file, row.names = FALSE)
      }
    }
  )
  
  # Instructions text
  output$instructions <- renderText({
    if (is.null(input$file_input)) {
      "Please upload an Excel file to get started."
    } else {
      "Excel file uploaded. Please enter shot numbers to skip and view the plot and table."
    }
  })
}

# Run the Shiny app
shinyApp(ui = ui, server = server)