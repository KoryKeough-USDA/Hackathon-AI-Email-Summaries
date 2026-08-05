library(shiny)
library(shinyjs)
library(shinyFiles)
library(htmltools)

# Define UI for application
ui <- fluidPage(
  useShinyjs(),
  titlePanel("AI Email Summaries Shiny App"),
  br(),
  textInput(inputId = "email", label = "Outlook Email"),
  br(),
  shinyDirButton('folder', 'Select a desktop folder to save AI summaries to', 'Please select a folder', multiple = FALSE),
  br(),
  selectInput("folderType", "Which Outlook folder should the AI summarize emails from?", 
              choices = c("Inbox", "Archives")),
  br(),
  
  fluidRow(
    
    column(width = 6,
           h3("Run AI Email Summaries", style = "font-weight: bold;"),
           br(),
           h4("From which dates should the AI summarize emails?", style = "font-weight: bold;"),
              # Create a date input with label, default value, and limits
           fluidRow(
             column(width = 2, dateInput(
                                 inputId = "start_date",          
                                 label = "Date:",    
                                 value = Sys.Date(),    # Default date (today)
                                 min = NULL,          
                                 max = Sys.Date(),         
                                 format = "mm/dd/yyyy",
                                 startview = "month"
                   )
           ),
           column(width = 2, tagAppendAttributes(
             textInput(inputId = "to_text", label = NULL, value = " to "),
             readonly = "readonly")
                 ),
           column(width = 2, dateInput(
                                inputId = "end_date",          
                                label = "Date:",    
                                value = Sys.Date(),    # Default date (today)
                                min = NULL,          
                                max = Sys.Date(),         
                                format = "mm/dd/yyyy",
                                startview = "month"
           )
        )
    ),
             hr(), # Visual divider line
    downloadButton("run_ai_summary", "Run AI Summary", icon = icon("play")),
           ),
    column(width = 6,
           wellPanel(
             h3("AI Email Summary Task Scheduler", style = "font-weight: bold;"),
             selectInput("freq", "Frequency", 
                         choices = c("DAILY", "WEEKLY", "MONTHLY")),
             br(),
             h4("Set Time to Run Task Scheduler Using 24-hour Time Format", style = "font-weight: bold;"),
             fluidRow(
             column(width = 2, numericInput(inputId = "hours", label = NULL, value = 0, min = 0, max = 23, step = 1)),
             column(width = 2, tagAppendAttributes(
                                                  textInput(inputId = "colon", label = NULL, value = ":"),
                                                  readonly = "readonly")
                    ),
             column(width = 2, numericInput(inputId = "minutes", label = NULL, value = 0, min = 0, max = 59, step = 1)),
             ),
             br(),
             h4("When Should the Task Scheduler First Run?", style = "font-weight: bold;"),
             # Create a date input with label, default value, and limits
             column(width = 4, dateInput(
                      inputId = "start_taskscheduler_date",          
                      label = "Date:",    
                      value = Sys.Date(),    # Default date (today)
                      min = Sys.Date(),          
                      max = NULL,         
                      format = "yyyy-mm-dd",
                      startview = "month"
             )
             ),
             column(width = 2, checkboxInput(
                      inputId = "last_day_of_month", 
                      label = "Last Day of Each Month", 
                      value = FALSE
             )
             ),
             hr(),
               actionButton("run_task_scheduler", "Run Task Scheduler", class = "btn-success"),
             br(),
               actionButton("remove_task_scheduler", "Remove Scheduled Tasks", class = "btn-danger"),
           )
        )
    )
)


message("Outputting and saving summary")
server <- function(input, output, session) {
  roots <- getVolumes()()
  shinyDirChoose(input, 'folder', roots = roots, session = session)
  folder_dir <- reactive({
    req(input$folder)
    parseDirPath(roots, input$folder)
  })
  run_email_summary_once_selections <- reactive({
    run_email_summary_once(
      sharedEmail = input$email,
      folderType = input$folderType,
      start_date  = input$start_date,
      end_date   = input$end_date
    )
  }) %>%
    bindEvent(input$run_ai_summary, ignoreInit = TRUE)
  output$run_ai_summary <- downloadHandler(
    filename = function() {
      paste0("Groq_email_summary_", format(Sys.Date(), "%m%d%Y"), ".txt")
    },
    content = function(file) {
      req(run_email_summary_once_selections())
      req(folder_dir())
      writeLines(total_summary, file)
      writeLines(total_summary, paste0(folder_dir(), "/Groq_email_summary", todays_date, ".txt"))
    }
  )
  schtasks_extra <- reactiveVal(value = "")
  observe({
    if (input$freq == "DAILY" | input$freq == "WEEKLY") {
      shinyjs::disable("last_day_of_month")
    } else if (input$freq == "MONTHLY") {
      shinyjs::enable("last_day_of_month")
    }
  }) %>%
    bindEvent(input$freq)
  observe({
    if (input$last_day_of_month) {
      shinyjs::disable("start_taskscheduler_date")
      schtasks_extra("/mo LASTDAY /m *")
    } else {
      shinyjs::enable("start_taskscheduler_date")
    }
  }) %>%
    bindEvent(input$last_day_of_month)
  observe({
      sharedEmail <- input$email
      folder_dir <- folder_dir()
      folderType <- input$folderType
      start_taskscheduler_date  = input$start_taskscheduler_date
      freq = input$freq
      source("hackathon_r_file_scheduler.R", local = TRUE)
    })
  
  observe({
    task_scheduler_emails(
      freq = input$freq,
      start_time = paste0(sprintf("%02d", input$hours), ":", sprintf("%02d", input$minutes)),
      start_taskscheduler_date  = input$start_taskscheduler_date,
      start_taskscheduler_day  = input$start_taskscheduler_day,
      schtasks_extra   = input$schtasks_extra
    )
  }) %>%
    bindEvent(input$run_task_scheduler)
  
  observe({
    taskscheduler_delete(taskname = "hackathon_r_file_scheduler")
  }) %>%
  bindEvent(input$remove_task_scheduler)
}


shinyApp(ui, server)