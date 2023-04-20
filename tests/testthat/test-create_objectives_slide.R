# test-create_objectives_slide.R

context("create_objectives_slide function")

# Test that a slide is created successfully
test_that("create_objectives_slide creates a slide successfully", {
  doc <- read_ppt_template()
  table_data <- data.frame(Objective = c("Objective 1", "Objective 2"),
                           KeyResult = c("Key Result 1", "Key Result 2"))
  slide <- create_objectives_slide(
    doc,
    slide_table = table_data
  )
  expect_silent(officer:::print.rpptx(slide, tempfile(fileext = ".pptx")))
})

# Test that the function throws an error when the layout doesn't exist
test_that("create_objectives_slide throws an error when the layout doesn't exist", {
  doc <- read_ppt_template()
  table_data <- data.frame(Objective = c("Objective 1", "Objective 2"),
                           KeyResult = c("Key Result 1", "Key Result 2"))
  expect_error(create_objectives_slide(doc, layout_name = "Non-existent layout", slide_table = table_data))
})

# Test that the function throws an error when the layout is not an objectives slide
test_that("create_objectives_slide throws an error when the layout is not an objectives slide", {
  doc <- read_ppt_template()
  table_data <- data.frame(Objective = c("Objective 1", "Objective 2"),
                           KeyResult = c("Key Result 1", "Key Result 2"))
  expect_error(create_objectives_slide(doc, layout_name = "Title - Black", slide_table = table_data))
})

# Add more test cases as needed
