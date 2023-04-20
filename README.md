## PowerPoint Slide Builder - `prrresent`. 

> This R package provides a set of functions to automate the process of creating PowerPoint slides in a consistent and efficient manner. The package is designed to work with a predefined PowerPoint template, offering a range of slide layouts and designs.  

### Features   

The PowerPoint Slide Builder package includes the following functions to create various types of slides:  
- `create_agenda_slide()`: Creates an agenda slide with a list of topics.  
- `create_column_slide()`: Creates a slide with multiple columns.  
- `create_decision_point_slide()`: Creates a slide with a graphic and decision points.  
- `create_graphic_slide()`: Creates a slide with a graphic or image.  
- `create_objectives_slide()`: Creates a slide with objectives.  
- `create_point_slide()`: Creates a point-based slide.  
- `create_problem_statement_slide()`: Creates a slide with a problem statement.  
- `create_section_slide()`: Creates a section slide with a title and subtitle.  
- `create_table_slide()`: Creates a table slide with a title, table, caption, body text, and section.  
- `create_title_slide()`: Creates a title slide with a title and subtitle.  

The package also includes several utility functions and global variables to aid in slide creation:

- `globals.R`: Contains global variables and configurations.  
- `ph_disclaimer()`: Adds a disclaimer to the slides.  
- `ph_slide_number()`: Adds slide numbers to the slides.  
- `read_ppt_template()`: Reads a PowerPoint template file.  
- `slide_builder_view()`: A function to preview the slides.   

### Installation  
To install the `prrresent` package, you can use the following code:  
  
```r  
# Install from GitHub  
# install.packages("devtools")  
devtools::install_github("EV-Advisory/prrresent")    
```  

### Usage  

Before using the functions to create slides, you will need to load the package and read a PowerPoint template:  
  
```r  
require(prrresent)  
doc <- read_ppt_template(System.file("default-eva-template.potx",package = "prrresent"))   
```    

Then, use the slide creation functions to create the desired slides:   

```r  
doc <- create_title_slide(doc, "Title", "Subtitle"). 
doc <- create_section_slide(doc, "Section Title", "Section Subtitle"). 
# Add more slides as needed. 
```  
After creating all the slides, save the final PowerPoint file:
```r  
officer::print.rpptx(doc, "path/to/save/final_presentation.pptx")
```

### Contributing   

Contributions are welcome! Please feel free to submit issues or pull requests to help improve the package.  


### License  

This project is licensed by EV Advisory for distribution with the corresponding deck template. We are not responsible for uses that result in damages or losses, specifically when the product is not explicitly permitted for use by other companies or private label contractors.  

For more information, please reach out to the [maintainer](//mailto:info@evadvisory.ca)
