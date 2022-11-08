# Author: Chase Conner
# Date: 10/8/2021
# Description: Cleans data missing the user hours and check-out time cells
#
# Sources: 
#
#   1) https://www.geeksforgeeks.org/how-to-filter-r-dataframe-by-values-in-a-column/#:~:text=The%20filter%20%28%29%20method%20in%20R%20can%20be,as%20NA%20value%20check%20against%20the%20column%20values.
#   2) https://www.statology.org/r-export-to-excel/#:~:text=The%20easiest%20way%20to%20export%20a%20data%20frame,data%20frame%20to%20an%20Excel%20file%20in%20R.
#   3) https://www.tutorialkart.com/r-tutorial/concatenate-two-or-more-strings-in-r/#:~:text=To%20concatenate%20strings%20in%20r%20programming%2C%20use%20paste,comma.%20Any%20number%20of%20strings%20can%20be%20given.
#   4) https://www.journaldev.com/52445/write-data-to-excel-using-r
#   5) https://stackoverflow.com/questions/34731382/add-sheet-to-excel-file
#   6) https://stackoverflow.com/questions/22590597/r-filtering-a-dataframe-by-a-column-contating-a-keyword
#   7) https://www.r-bloggers.com/2019/08/how-to-select-multiple-columns-using-grep-r/
#   8) https://stackoverflow.com/questions/28556658/why-does-an-empty-dataframe-fail-an-is-null-test

library(readxl) # Import the functions that can read Excel Worksheet files
library(writexl) # Import the functions that can write the data frame to Excel Worksheet files
library(xlsx) # Import functions needed to create and write a new sheet to an existing Excel file
library(tidyverse) # Import functions that manipulate data

## Remove unnecessary columns from data frame and keep only the requested needed ones
filter_columns <- function(df) {
  keep_needed_columns <- df %>% select(Event.Start.Time, Event.End.Time, Event.Name, User.username, 
                                       User.Slots.Used, User.Hours, Earliest.Check.In.Time, 
                                       Latest.Check.Out.Time, Full.Name.FirstName, FullName.LastName, 
                                       Email.Value, Date.of.Birth.Value)
  
}

## Function for extracting the volunteers with no user hours, no earliest check-in time,
## and no latest check-out time.
extract_volunteers_with_no_user_hours_and_no_check_in_time_and_no_check_out_time <- function(df) {
  no_user_hours_df <- df[is.na(df$User.Hours),] # Find and store the volunteers with no user hours
  
  # Find and store the volunteers with no check out times from the no_user_hours_df
  no_check_out_time_and_no_user_hours_df <- no_user_hours_df[is.na(no_user_hours_df$Latest.Check.Out.Time),]
  
  # Find and store volunteers with no check-in time from the no_check_out_time_and_no_user_hours_df
  no_earliest_check_in_time_df <- no_check_out_time_and_no_user_hours_df[is.na(no_check_out_time_and_no_user_hours_df$Earliest.Check.In.Time),]
}

## Function for extracting the volunteers with no user hours and no check-out time, but
## do have an earliest check in time
extract_volunteers_with_no_user_hours_and_with_a_check_in_time_and_no_check_out_time <- function(df) {
  no_user_hours_df <- df[is.na(df$User.Hours),] # Find and store the volunteers with no user hours
  
  # Find and store the volunteers with no check out times from the no_user_hours_df
  no_check_out_time_and_no_user_hours_with_check_in_time_df <- no_user_hours_df[is.na(no_user_hours_df$Latest.Check.Out.Time),]
 
  has_check_in_time_df  <- no_check_out_time_and_no_user_hours_with_check_in_time_df[!is.na(no_check_out_time_and_no_user_hours_with_check_in_time_df$Earliest.Check.In.Time),]
  
  # Filter out unnecessary columns
  filtered_columns <- filter_columns(has_check_in_time_df)
  
  # Write output to new Excel workbook
  write.xlsx(filtered_columns, "List of Volunteers with Check In Time But No User Hours and No Check Out Times.xlsx", 
              append = TRUE, showNA = FALSE)
}

## Categorize the volunteers with no user hours, no check-in time, and no check-out time
## by department
categorize_volunteers_with_no_user_hours_and_no_check_in_time_and_no_check_out_time_by_dept <- function(df, fileName) {
  depts <- c("Resale", "Employment Services", "Food Shelf", "Client Services") # Volunteer departments except for "Random"

  # Filter the volunteers by department
  for (val in depts) {
    aDept <- toString(val) # Get name of the department
    aDept_df <- df[grep(aDept, df$Event.Group.Path),] # See if Event Group Path contains the name of the department
    removed_offsite_tracking <- aDept_df[!grepl("Offsite", aDept_df$Event.Name),] # Extract only rows that are not offsite tracking
    keep_needed_columns <- filter_columns(removed_offsite_tracking) # Keep only requested columns in data frame

    # If data frame is not empty, write to Excel
    if (is.data.frame(keep_needed_columns) && nrow(keep_needed_columns) != 0) {
      write.xlsx(keep_needed_columns, file = fileName, sheetName = aDept, append = TRUE, showNA = FALSE) # Create a sheet named after the dept, and write the volunteers there
    } else {
      write.xlsx("No volunteers worked in this department this month", file = fileName, sheetName = aDept, append = TRUE, showNA = FALSE) # Notify that this sheet has no cells
    }
  }
  
  # Find other departments not listed in depts list
  otherDepts_df <- df[!grepl("Resale|Employment Services|Food Shelf|Client Services", df$Event.Group.Path),]
  removed_offsite_tracking <- otherDepts_df[!grepl("Offsite", otherDepts_df$Event.Name),] # Extract only rows that are not offsite tracking
  keep_needed_columns <- filter_columns(removed_offsite_tracking) # Keep only requested columns in data frame

  # If data frame is not empty, write to Excel
  if (is.data.frame(keep_needed_columns) && nrow(keep_needed_columns) != 0) {
    write.xlsx(keep_needed_columns, file = fileName, sheetName = "Other", append = TRUE, showNA = FALSE) # Create a sheet named after the dept, and write the volunteers there
  } else {
    write.xlsx("No volunteers worked in this department this month", file = fileName, sheetName = aDept, append = TRUE, showNA = FALSE)
  }
}


## Set the working directory and read the data from the file
setup <- function(wd, file) {
  setwd(wd) # Set the current working directory
  data <- read_xlsx(file) # Read and store the data into variable
  df <- data.frame(data) # Turn the data into a data frame
}

currentWD = "C:/Users/chase/OneDrive/Interfaith Outreach/Data Cleaning R Scripts" # The current working directory
file = "All Data - Removed Confidential Info.xlsx" # The file
df <- setup(currentWD, file) # The data frame

# Find volunteers with no user hours and no latest check-out time
volunteers_with_no_user_hours_and_no_check_in_times_and_no_check_out_time_df <- extract_volunteers_with_no_user_hours_and_no_check_in_time_and_no_check_out_time(df) 

# Categorize the volunteers by dept and put in separate sheets
categorize_volunteers_with_no_user_hours_and_no_check_in_time_and_no_check_out_time_by_dept(volunteers_with_no_user_hours_and_no_check_in_times_and_no_check_out_time_df, file)
 
#Extract the volunteers that do have an earliest check in time, but no user hours and no check out time
extract_volunteers_with_no_user_hours_and_with_a_check_in_time_and_no_check_out_time(df)