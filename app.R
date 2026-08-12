# The webpage containing the version of this code submitted with the
# hackathon can be found at https://github.com/KoryKeough-USDA/Hackathon-AI-Email-Summaries/commit/5d8d2224342c75dd78c64af71c5eca290940a5fc

library(shiny)
library(shinyjs)
library(shinyFiles)
library(htmltools)
library(tidyverse)
library(taskscheduleR)
library(lubridate)
library(remotes)
library(devtools)
# RDCOMClient can be installed by running pak::pak('omegahat/RDCOMClient')
library(RDCOMClient)
library(curl)
library(ellmer)
library(xml2)
library(rvest)

source("hackathon_r_file.R")
source("task_scheduler.R")

input_to <- textInput(inputId = "to_text", label = NULL, value = " to ")
input_to$children[[2]] <- tagAppendAttributes(input_to$children[[2]], readonly = "readonly")

input_colon <- textInput(inputId = "colon", label = NULL, value = ":")
input_colon$children[[2]] <- tagAppendAttributes(input_colon$children[[2]], readonly = "readonly")

# Define UI for application
ui <- fluidPage(
  useShinyjs(),
  tags$head(
    tags$style(HTML("
      hr {
        border: none;
        height: 3px;
        background-color: black;
      }
    "))
  ),
  titlePanel("AI Email Summaries Shiny App"),
  br(),
  textInput(inputId = "email", label = "Outlook Email"),
  br(),
  shinyDirButton('folder', 'Select a desktop folder to save AI summaries to', 'Please select a folder', icon = icon("folder-open"), multiple = FALSE),
  verbatimTextOutput("path_display"),
  br(),
  selectInput("folderType", "Which Outlook folder should the AI summarize emails from?", 
              choices = c("Inbox", "Archives")),
  hr(),
  br(),
  
  fluidRow(
    
    column(width = 6,
           h3("Run AI Email Summaries", style = "font-weight: bold;"),
           br(),
           h4("From which dates should the AI summarize emails?", style = "font-weight: bold;"),
              # Create a date input with label, default value, and limits
           fluidRow(
             column(width = 5, dateInput(
                                 inputId = "start_date",          
                                 label = "Start Date:",    
                                 value = Sys.Date(),    # Default date (today)
                                 min = NULL,          
                                 max = Sys.Date(),        
                                 format = "mm/dd/yyyy",
                                 startview = "month"
                   )
           ),
           column(width = 2, input_to
                 ),
           column(width = 5, dateInput(
                                inputId = "end_date",          
                                label = "End Date:",    
                                value = Sys.Date(),    # Default date (today)
                                min = NULL,          
                                max = Sys.Date(),         
                                format = "mm/dd/yyyy",
                                startview = "month"
           )
        )
    ),
             hr(), # Visual divider line
    actionButton("run_ai_summary", "Run AI Summary", icon = icon("play")),
           ),
    column(width = 6,
           wellPanel(
             h3("AI Email Summary Task Scheduler", style = "font-weight: bold;"),
             selectInput("freq", "Frequency", 
                         choices = c("DAILY", "WEEKLY", "MONTHLY")),
             br(),
             h4("Set Time to Run Task Scheduler Using 24-hour Time Format", style = "font-weight: bold;"),
             fluidRow(
               column(width = 5, numericInput(inputId = "hours", label = NULL, value = 0, min = 0, max = 23, step = 1)),
               column(width = 2, input_colon
                    ),
               column(width = 5, numericInput(inputId = "minutes", label = NULL, value = 0, min = 0, max = 59, step = 1)),
             ),
             br(),
             h4("When Should the Task Scheduler First Run?", style = "font-weight: bold;"),
             # Create a date input with label, default value, and limits
             fluidRow(
               column(width = 8, dateInput(
                      inputId = "start_taskscheduler_date",          
                      label = "Date:",    
                      value = NULL,
                      min = Sys.Date(),          
                      max = NULL,         
                      format = "mm/dd/yyyy",
                      startview = "month"
             )
             ),
             column(width = 4, checkboxInput(
                      inputId = "last_day_of_month", 
                      label = "Last Day of Each Month", 
                      value = FALSE
             )
             )
             ),
             hr(),
               actionButton("run_task_scheduler", "Run Task Scheduler", class = "btn-success"),
             br(),
             br(),
               actionButton("remove_task_scheduler", "Stop Performing Scheduled Tasks", class = "btn-danger"),
           )
        )
    )
)


server <- function(input, output, session) {
  roots <- getVolumes()()
  shinyDirChoose(input, 'folder', roots = roots, session = session)
  folder_dir <- reactive({
    req(input$folder)
    parseDirPath(roots, input$folder)
  })
  week_day <- reactive({
    if (input$freq == "WEEKLY") {
      str_trunc(toupper(weekdays(input$start_taskscheduler_date)), 3, ellipsis = "")
    }
  })
           
  output$path_display <- renderText({
    if (length(folder_dir()) == 0) {
      "No folder selected yet."
    } else {
      folder_dir()
    }
  })
  
  run_email_summary_once_selections <- observe({
    run_email_summary_once(
      sharedEmail = input$email,
      folder_dir = folder_dir(),
      folderType = input$folderType,
      start_date  = input$start_date,
      end_date   = input$end_date
    )
  }) %>%
    bindEvent(input$run_ai_summary)
  
  start_taskscheduler_day <- reactiveVal(value = c("*", "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN", 1:31))
  modifier <- reactiveVal(value = "")
  
  observe({
    if (input$freq == "DAILY") {
      shinyjs::reset("last_day_of_month")
      shinyjs::disable("last_day_of_month")
    } else if (input$freq == "WEEKLY") {
      shinyjs::reset("last_day_of_month")
      shinyjs::disable("last_day_of_month")
    } else if (input$freq == "MONTHLY") {
      shinyjs::enable("last_day_of_month")
    }
  }) %>%
    bindEvent(input$freq)
  
  observe({
    if (input$last_day_of_month) {
      shinyjs::disable("start_taskscheduler_date")
      modifier("LASTDAY")
    } else if (!isTruthy(input$last_day_of_month) && input$freq == "MONTHLY") {
      shinyjs::enable("start_taskscheduler_date")
      new_value <- day(as.Date(input$start_taskscheduler_date))
    } else if (input$freq == "WEEKLY") {
        new_value <- week_day()
    }
    else if (input$freq == "DAILY") {
      new_value <- c("*", "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN", 1:31)
    }
    start_taskscheduler_day(new_value)
    }) %>%
    bindEvent(input$last_day_of_month, input$start_taskscheduler_date, input$freq)
  
  observe({
    saveRDS(list(
      sharedEmail = input$email,
      folderType <- input$folderType,
      folder_dir <- folder_dir(),
      start_taskscheduler_date  = input$start_taskscheduler_date,
      freq = input$freq),
      file = "scheduler_config.rds")
    }) %>%
    bindEvent(input$run_task_scheduler)
  
  start_time <- reactiveVal()
  
  observe({
    start_time(paste0(sprintf("%02d", input$hours), ":", sprintf("%02d", input$minutes)))
  })
  
  observe({
    if (input$last_day_of_month) {
    task_scheduler_emails(
      freq = input$freq,
      start_time = start_time(),
      start_taskscheduler_date = input$start_taskscheduler_date,
      last_day = modifier()
    )
  } else if (!isTruthy(input$last_day_of_month)) {
      task_scheduler_emails(
        freq = input$freq,
        start_time = start_time(),
        start_taskscheduler_date = input$start_taskscheduler_date,
        start_taskscheduler_day = start_taskscheduler_day(),
        last_day = ""
      )
    }
  }) %>%
      bindEvent(input$run_task_scheduler)
  
  observe({
    taskscheduler_delete(taskname = "Hackathon_Task_Scheduler")
  }) %>%
  bindEvent(input$remove_task_scheduler)
}


shinyApp(ui, server)