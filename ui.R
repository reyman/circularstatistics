library(shiny)

# Define UI for dataset viewer application
shinyUI(fluidPage(
  
  # Application title
  titlePanel("Saint Cesaire"),
  
  # Sidebar with controls to provide a caption, select a dataset,
  # and specify the number of observations to view. Note that
  # changes made to the caption in the textInput control are
  # updated in the output area immediately as you type
  sidebarLayout(
    sidebarPanel(
      sliderInput("orientation", "Orientation:", min = 0, max=360, step =1,dragRange=FALSE,value=360),
      selectInput("classes", "Classes:", choices = c(20,30)),
      selectInput("columns", "Colonnes:", choices = ""),
      selectInput("criteria", "Critères:", choices = "")
    ),
    
    # Show the caption, a summary of the dataset and an HTML 
    # table with the requested number of observations
    mainPanel(
      tabsetPanel(
        tabPanel("result", formattable::formattableOutput("result")),
        tabPanel("data", dataTableOutput('data'))
      )
    )
  )
))
